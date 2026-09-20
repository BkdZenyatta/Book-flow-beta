import datetime
import hashlib
import hmac
import ipaddress
import logging
import os
import re
import socket
import sqlite3
import urllib.parse
from contextlib import contextmanager
from html import unescape
from io import BytesIO
from pathlib import Path

import requests
import streamlit as st
from docx import Document
from google import genai
from pypdf import PdfReader

try:
  from ddgs import DDGS  # pacote novo (pip install ddgs)
except ImportError:
  from duckduckgo_search import DDGS  # pacote antigo, como fallback

st.set_page_config(page_title="Book Flow", page_icon="📚", layout="wide")

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("bookflow")

# --- CONSTANTES ---
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

DB_PATH = Path(__file__).parent / "bookflow.db"
MAX_PDF_BYTES = 15 * 1024 * 1024  # 15 MB
MAX_HTML_BYTES = 1_500_000  # limite ao ler páginas HTML em busca de PDFs
MAX_TEXTO_PDF = 6000  # texto guardado para checar relevância
TEXTO_PROMPT_MAX = 2500  # trecho enviado ao Gemini / preview
MAX_GERACOES_SESSAO = 10  # limite de resumos por visitante (chave do servidor)

# Sites consultados na busca de PDFs (além do archive.org, que usa API própria)
SITES_PREFERIDOS = [
    "baixelivros.com.br",
    "clubedolivrodesatolep.wordpress.com",
]

MODELOS = ["gemini-2.5-flash", "gemini-2.0-flash"]  # mantenha atualizado
CODES_TENTAR_OUTRO = {404, 429, 500, 503}

# Textos das mensagens de erro (usados também para limpar lixo antigo do banco)
MARCADORES_ERRO_ANTIGOS = [
    "**Sem Conexão / Falha de DNS:**",
    "**Servidor Ocupado:**",
    "**Modelo Indisponível:**",
    "**Chave Inválida:**",
    "**Instabilidade na Conexão:**",
    "**Chave de API não inserida.**",
]


# --- SECRETS ---
def get_secret(nome, padrao=""):
  try:
    return st.secrets.get(nome, padrao) or padrao
  except Exception:
    return padrao


# --- SENHAS (hash com salt) ---
def hash_senha(pwd, salt=None):
  salt = salt or os.urandom(16)
  h = hashlib.pbkdf2_hmac("sha256", pwd.encode(), salt, 200_000)
  return f"{salt.hex()}:{h.hex()}"


def verificar_senha(pwd, armazenado):
  try:
    salt_hex, h_hex = armazenado.split(":")
    h = hashlib.pbkdf2_hmac(
        "sha256", pwd.encode(), bytes.fromhex(salt_hex), 200_000
    )
  except (ValueError, AttributeError):
    return False
  return hmac.compare_digest(h.hex(), h_hex)


# --- BANCO DE DADOS LOCAL (SQLite) ---
@contextmanager
def db():
  conn = sqlite3.connect(DB_PATH, timeout=10)
  try:
    yield conn
    conn.commit()
  except Exception:
    conn.rollback()
    raise
  finally:
    conn.close()


def _garantir_coluna(conn, tabela, coluna, tipo):
  cols = [r[1] for r in conn.execute(f"PRAGMA table_info({tabela})")]
  if coluna not in cols:
    conn.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {tipo}")


@st.cache_resource
def init_db():
  with db() as conn:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS resumos (
            titulo TEXT PRIMARY KEY,
            resumo TEXT,
            fonte TEXT,
            pdf_url TEXT,
            autor TEXT,
            paginas INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS erros_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT,
            livro TEXT,
            erro_original TEXT,
            mensagem_amigavel TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            username TEXT PRIMARY KEY,
            password TEXT,
            role TEXT
        )
    """)

    # Migração de bancos antigos
    _garantir_coluna(conn, "resumos", "autor", "TEXT")
    _garantir_coluna(conn, "resumos", "paginas", "INTEGER")
    _garantir_coluna(conn, "erros_log", "livro", "TEXT")

    # Remove mensagens de erro que foram salvas por engano como resumo
    for marcador in MARCADORES_ERRO_ANTIGOS:
      conn.execute("DELETE FROM resumos WHERE instr(resumo, ?) > 0", (marcador,))

    # Remove usuários com senha antiga em texto puro (sem "salt:hash")
    conn.execute("DELETE FROM usuarios WHERE password NOT LIKE '%:%'")

    # Admin: a senha vem do secrets.toml (ADMIN_PASSWORD)
    admin_pwd = get_secret("ADMIN_PASSWORD")
    if admin_pwd:
      conn.execute(
          "INSERT OR REPLACE INTO usuarios (username, password, role)"
          " VALUES (?, ?, ?)",
          ("admin", hash_senha(admin_pwd), "admin"),
      )
    else:
      log.warning(
          "ADMIN_PASSWORD não definido no secrets: painel admin desativado."
      )


def autenticar_usuario(user, pwd):
  with db() as conn:
    row = conn.execute(
        "SELECT password, role FROM usuarios WHERE username = ?",
        (user.strip(),),
    ).fetchone()
  if row and verificar_senha(pwd, row[0]):
    return row[1]
  return None


def _chave(titulo):
  return " ".join(titulo.lower().split())


def salvar_no_banco(titulo, resumo, fonte="IA", pdf_url="", autor="", paginas=0):
  with db() as conn:
    conn.execute(
        """INSERT OR REPLACE INTO resumos
           (titulo, resumo, fonte, pdf_url, autor, paginas)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (_chave(titulo), resumo, fonte, pdf_url, autor, paginas),
    )


def atualizar_pdf_no_banco(titulo, pdf_url, autor, paginas):
  """Atualiza só os dados do PDF, sem mexer no resumo. Retorna True se existia."""
  with db() as conn:
    cur = conn.execute(
        "UPDATE resumos SET pdf_url = ?, autor = ?, paginas = ? WHERE titulo = ?",
        (pdf_url, autor, paginas, _chave(titulo)),
    )
    return cur.rowcount > 0


def buscar_no_banco(titulo):
  with db() as conn:
    row = conn.execute(
        "SELECT resumo, fonte, pdf_url, autor, paginas FROM resumos"
        " WHERE titulo = ?",
        (_chave(titulo),),
    ).fetchone()
  if row:
    return {
        "resumo": row[0],
        "fonte": row[1],
        "pdf_url": row[2],
        "autor": row[3],
        "paginas": row[4],
    }
  return None


def listar_livros_salvos():
  with db() as conn:
    rows = conn.execute(
        "SELECT titulo FROM resumos ORDER BY titulo ASC"
    ).fetchall()
  return [r[0].title() for r in rows]


def traduzir_e_registrar_erro(erro, titulo="Desconhecido"):
  erro = erro or RuntimeError("Erro desconhecido")
  erro_str = str(erro)
  code = getattr(erro, "code", None)

  if (
      "getaddrinfo" in erro_str
      or "Name or service not known" in erro_str
      or type(erro).__name__ in ("ConnectError", "ConnectTimeout", "ReadTimeout")
  ):
    msg = (
        "🌐 **Sem Conexão / Falha de DNS:** Não foi possível conectar aos"
        " servidores. Verifique sua conexão com a internet."
    )
  elif code == 503 or "UNAVAILABLE" in erro_str:
    msg = (
        "⚠️ **Servidor Ocupado:** O modelo do Gemini está com alta demanda"
        " (Erro 503). Tente novamente."
    )
  elif code == 429 or "RESOURCE_EXHAUSTED" in erro_str:
    msg = (
        "⏳ **Cota Excedida:** O limite de uso da API foi atingido (Erro 429)."
        " Aguarde um pouco e tente novamente."
    )
  elif code == 404 or "NOT_FOUND" in erro_str:
    msg = (
        "❌ **Modelo Indisponível:** O modelo de IA solicitado não foi"
        " localizado (Erro 404)."
    )
  elif (
      code in (401, 403)
      or "API_KEY_INVALID" in erro_str
      or "API key not valid" in erro_str
  ):
    msg = (
        "🔑 **Chave Inválida:** A chave de API do Gemini é inválida, expirou"
        " ou não tem permissão."
    )
  elif isinstance(erro, ValueError):
    msg = f"🚫 **Resposta Vazia:** {erro_str}"
  else:
    msg = (
        "⚠️ **Instabilidade na Conexão:** Ocorreu uma falha ao comunicar com"
        f" a IA: {erro_str[:120]}..."
    )

  try:
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with db() as conn:
      conn.execute(
          "INSERT INTO erros_log (data_hora, livro, erro_original,"
          " mensagem_amigavel) VALUES (?, ?, ?, ?)",
          (data_hora, titulo.title(), erro_str, msg),
      )
  except Exception:
    log.exception("Falha ao registrar erro no banco")

  return msg


def obter_logs_erros():
  with db() as conn:
    return conn.execute(
        "SELECT id, data_hora, livro, mensagem_amigavel, erro_original FROM"
        " erros_log ORDER BY id DESC LIMIT 15"
    ).fetchall()


init_db()


# --- PDF (com proteção contra SSRF e arquivos gigantes) ---
def _url_segura(url):
  p = urllib.parse.urlparse(url)
  if p.scheme not in ("http", "https") or not p.hostname:
    return False
  try:
    for info in socket.getaddrinfo(p.hostname, None):
      ip = ipaddress.ip_address(info[4][0])
      if not ip.is_global:  # bloqueia localhost, rede interna, reservados
        return False
  except (socket.gaierror, ValueError):
    return False
  return True


def _baixar_bytes(url, limite, truncar=False):
  """Baixa uma URL com proteção contra SSRF e limite de tamanho."""
  res = None
  for _ in range(4):  # valida cada redirecionamento
    if not _url_segura(url):
      raise ValueError("URL não permitida.")
    res = requests.get(
        url, headers=HEADERS, timeout=12, stream=True, allow_redirects=False
    )
    if res.is_redirect:
      url = urllib.parse.urljoin(url, res.headers.get("Location", ""))
      res.close()
      continue
    break
  else:
    raise ValueError("Redirecionamentos demais.")

  buf = BytesIO()
  try:
    res.raise_for_status()
    for chunk in res.iter_content(65536):
      buf.write(chunk)
      if buf.tell() > limite:
        if truncar:
          break
        raise ValueError("Arquivo muito grande.")
  finally:
    res.close()
  return buf.getvalue()[:limite] if truncar else buf.getvalue()


def baixar_pdf_web(url):
  dados = _baixar_bytes(url, MAX_PDF_BYTES)
  if not dados.startswith(b"%PDF"):
    raise ValueError("O link fornecido não aponta para um arquivo PDF válido.")

  reader = PdfReader(BytesIO(dados))
  texto = "\n".join(p.extract_text() or "" for p in reader.pages[:5])
  meta = reader.metadata
  autor = meta.author if meta and meta.author else "Não especificado"
  return autor, len(reader.pages), texto[:MAX_TEXTO_PDF]


def _texto_relevante(titulo, texto):
  """Confere se o texto do PDF realmente parece ser do livro buscado."""
  t = titulo.lower()
  palavras = re.findall(r"\w{4,}", t) or re.findall(r"\w+", t)
  if not palavras:
    return False
  alvo = texto.lower()
  return sum(p in alvo for p in palavras) / len(palavras) >= 0.5


def _tentar_pdf(titulo, url):
  """Baixa o PDF e só o aceita se parecer o livro. Retorna dict ou None."""
  try:
    autor, paginas, texto = baixar_pdf_web(url)
  except Exception as e:
    log.info("PDF ignorado (%s): %s", url, e)
    return None
  if not _texto_relevante(titulo, texto):
    log.info("PDF descartado por não parecer o livro: %s", url)
    return None
  return {
      "url": url,
      "autor": autor,
      "paginas": paginas,
      "texto_preview": texto[:TEXTO_PROMPT_MAX],
  }


def _extrair_links_pdf(pagina_url, maximo=2):
  """Abre uma página HTML e devolve os links diretos para .pdf encontrados."""
  dados = _baixar_bytes(pagina_url, MAX_HTML_BYTES, truncar=True)
  html = dados.decode("utf-8", errors="ignore")
  links = re.findall(
      r"""href=["']([^"']+?\.pdf(?:\?[^"']*)?)["']""", html, flags=re.I
  )
  absolutos = [urllib.parse.urljoin(pagina_url, unescape(l)) for l in links]
  return list(dict.fromkeys(absolutos))[:maximo]


def buscar_pdf_archive_org(titulo):
  """Busca via API oficial do archive.org (mais confiável que scraping)."""
  termo = " ".join(re.findall(r"\w+", titulo))
  if not termo:
    return None

  r = requests.get(
      "https://archive.org/advancedsearch.php",
      params={
          "q": f"({termo}) AND mediatype:texts",
          "fl[]": ["identifier", "title"],
          "rows": 5,
          "output": "json",
      },
      headers=HEADERS,
      timeout=10,
  )
  r.raise_for_status()
  docs = r.json().get("response", {}).get("docs", [])

  for d in docs:
    ident = d.get("identifier")
    if not ident:
      continue
    try:
      meta = requests.get(
          f"https://archive.org/metadata/{ident}", headers=HEADERS, timeout=10
      ).json()
      info = meta.get("metadata", {})
      if info.get("access-restricted-item") == "true":
        continue  # livro só para empréstimo, sem download livre

      nomes = [f.get("name", "") for f in meta.get("files", [])]
      pdf = next((n for n in nomes if n.lower().endswith(".pdf")), None)
      if not pdf:
        continue

      base = f"https://archive.org/download/{ident}"
      texto = ""
      djvu = next((n for n in nomes if n.endswith("_djvu.txt")), None)
      if djvu:  # texto pronto: evita baixar o PDF inteiro
        texto = _baixar_bytes(
            f"{base}/{urllib.parse.quote(djvu)}", MAX_TEXTO_PDF, truncar=True
        ).decode("utf-8", errors="ignore")

      if not _texto_relevante(titulo, f"{d.get('title', '')} {texto}"):
        continue

      autor = info.get("creator", "Não especificado")
      if isinstance(autor, list):
        autor = ", ".join(autor)
      paginas = info.get("imagecount")
      return {
          "url": f"{base}/{urllib.parse.quote(pdf)}",
          "autor": autor,
          "paginas": int(paginas) if str(paginas).isdigit() else 0,
          "texto_preview": texto[:TEXTO_PROMPT_MAX],
      }
    except Exception as e:
      log.info("Item do archive.org ignorado (%s): %s", ident, e)
  return None


def buscar_pdf_nos_sites(titulo):
  """Pesquisa nos sites da lista SITES_PREFERIDOS e procura PDFs nas páginas."""
  for dominio in SITES_PREFERIDOS:
    try:
      with DDGS(timeout=8) as ddgs:
        paginas = list(ddgs.text(f"site:{dominio} {titulo}", max_results=3))
    except Exception:
      log.warning("Busca em %s falhou", dominio, exc_info=True)
      continue

    for item in paginas:
      url = item.get("href", "")
      if not url:
        continue
      try:
        if url.lower().split("?")[0].endswith(".pdf"):
          candidatos = [url]
        else:
          candidatos = _extrair_links_pdf(url)
      except Exception as e:
        log.info("Página ignorada (%s): %s", url, e)
        continue

      for pdf_url in candidatos:
        achado = _tentar_pdf(titulo, pdf_url)
        if achado:
          return achado
  return None


def buscar_pdf_web_geral(titulo):
  """Último recurso: busca geral por PDFs na web."""
  with DDGS(timeout=8) as ddgs:
    resultados = list(ddgs.text(f"{titulo} filetype:pdf", max_results=4))
  for item in resultados:
    url = item.get("href", "")
    if url:
      achado = _tentar_pdf(titulo, url)
      if achado:
        return achado
  return None


def buscar_pdf_relacionado(titulo):
  """Tenta as fontes em ordem: archive.org, sites preferidos, web geral."""
  for busca in (
      buscar_pdf_archive_org,
      buscar_pdf_nos_sites,
      buscar_pdf_web_geral,
  ):
    try:
      achado = busca(titulo)
    except Exception:
      log.warning("Fonte %s falhou", busca.__name__, exc_info=True)
      continue
    if achado:
      return achado
  return None


# --- IA (Gemini) ---
def gerar_resumo_gemini(titulo, texto_base, api_key):
  """Retorna (ok: bool, texto: str)."""
  if not api_key or not api_key.strip():
    return False, (
        "🔑 **Chave de API não inserida.** Digite uma chave no menu lateral."
    )

  prompt = f"""
    Você é um especialista em literatura. Faça um resumo conciso, envolvente e bem estruturado do livro '{titulo}'.

    Estrutura desejada:
    1. **Visão Geral e Contexto**
    2. **Principais Tópicos / Capítulos-Chave**
    3. **Conclusão e Ensinamento Central**

    Trecho extraído de um PDF (use apenas se for realmente do livro; caso contrário, ignore):
    {texto_base if texto_base else "Nenhum texto extraído."}
    """

  client = genai.Client(api_key=api_key.strip())
  ultimo_erro = None
  for mod in MODELOS:
    try:
      resp = client.models.generate_content(model=mod, contents=prompt)
      if resp.text:
        return True, resp.text
      ultimo_erro = ValueError("Resposta vazia ou bloqueada pelo modelo.")
    except Exception as e:
      ultimo_erro = e
      if (
          getattr(e, "code", None) not in CODES_TENTAR_OUTRO
          and "getaddrinfo" not in str(e)
      ):
        break

  return False, traduzir_e_registrar_erro(ultimo_erro, titulo=titulo)


# --- WORD ---
def _add_runs(par, texto):
  """Converte **negrito** inline em runs em negrito."""
  for i, parte in enumerate(re.split(r"\*\*(.+?)\*\*", texto)):
    if parte:
      par.add_run(parte).bold = i % 2 == 1


def criar_arquivo_docx(titulo, conteudo):
  doc = Document()
  doc.add_heading("Book Flow - Resumo", level=0)
  doc.add_heading(f"Obra: {titulo.title()}", level=2)

  for linha in conteudo.split("\n"):
    l = linha.strip()
    if not l:
      continue
    if m := re.match(r"^(#{1,3})\s+(.*)", l):
      doc.add_heading(
          m.group(2).replace("**", ""), level=min(len(m.group(1)) + 1, 3)
      )
    elif re.match(r"^(\d+\.\s+)?\*\*[^*]+\*\*:?$", l):
      doc.add_heading(re.sub(r"^\d+\.\s+|\*\*|:$", "", l), level=3)
    elif re.match(r"^[-*•]\s+", l):
      _add_runs(doc.add_paragraph(style="List Bullet"), re.sub(r"^[-*•]\s+", "", l))
    else:
      _add_runs(doc.add_paragraph(), l)

  buffer = BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# --- YOUTUBE ---
def _buscar_audiobooks(query):
  videos_encontrados = []
  try:
    with DDGS(timeout=8) as ddgs:
      resultados = list(
          ddgs.text(f"site:youtube.com {query} audiobook completo", max_results=6)
      )
    for r in resultados:
      href = r.get("href", "")
      if "youtube.com/watch" in href or "youtu.be/" in href:
        vid_id = ""
        if "v=" in href:
          vid_id = href.split("v=")[1].split("&")[0]
        elif "youtu.be/" in href:
          vid_id = href.split("youtu.be/")[1].split("?")[0]

        thumb = (
            f"https://img.youtube.com/vi/{vid_id}/hqdefault.jpg"
            if vid_id
            else None
        )
        videos_encontrados.append({
            "title": r.get("title", f"Audiobook - {query}"),
            "link": href,
            "views": "Resultado da busca",
            "image": thumb,
        })
  except Exception:
    log.warning("Busca de audiobooks (DDG) falhou", exc_info=True)

  if not videos_encontrados:
    try:
      search_url = (
          "https://www.youtube.com/results?search_query="
          f"{urllib.parse.quote(query + ' audiobook completo')}"
      )
      res = requests.get(search_url, headers=HEADERS, timeout=6)
      video_ids = re.findall(r"watch\?v=([a-zA-Z0-9_-]{11})", res.text)
      for v_id in list(dict.fromkeys(video_ids))[:4]:
        videos_encontrados.append({
            "title": f"Audiobook Completo: {query.title()}",
            "link": f"https://www.youtube.com/watch?v={v_id}",
            "views": "Sugestão do YouTube",
            "image": f"https://img.youtube.com/vi/{v_id}/hqdefault.jpg",
        })
    except Exception:
      log.warning("Busca de audiobooks (scraping) falhou", exc_info=True)

  return videos_encontrados


@st.cache_data(ttl=3600, show_spinner=False)
def _audiobooks_cacheado(query):
  videos = _buscar_audiobooks(query)
  if not videos:
    raise LookupError("sem resultados")  # exceção não entra no cache
  return videos


def buscar_audiobooks_youtube(query):
  try:
    return _audiobooks_cacheado(query)
  except LookupError:
    return []


# --- INTERFACE STREAMLIT ---
st.title("📚 Book Flow")

for chave_estado, valor in (
    ("usuario_logado", None),
    ("usuario_role", None),
    ("resumo_ok", False),
    ("geracoes", 0),
):
  st.session_state.setdefault(chave_estado, valor)

# --- SIDEBAR & AREA DE LOGIN ADMIN ---
st.sidebar.header("⚙️ Configurações")

gemini_key = get_secret("GEMINI_API_KEY")
USANDO_CHAVE_DO_SERVIDOR = bool(gemini_key)

if gemini_key:
  st.sidebar.success("🔑 Chave do Gemini configurada!")
else:
  gemini_key = st.sidebar.text_input("Chave de API do Gemini:", type="password")

st.sidebar.divider()


def pode_gerar_com_ia():
  if not USANDO_CHAVE_DO_SERVIDOR or st.session_state["usuario_role"] == "admin":
    return True
  return st.session_state["geracoes"] < MAX_GERACOES_SESSAO


with st.sidebar.expander("🔐 Área do Administrador"):
  if st.session_state["usuario_logado"] is None:
    user_input = st.text_input("Usuário:", key="login_user")
    pass_input = st.text_input("Senha:", type="password", key="login_pass")
    if st.button("Entrar"):
      role = autenticar_usuario(user_input, pass_input)
      if role:
        st.session_state["usuario_logado"] = user_input.strip()
        st.session_state["usuario_role"] = role
        st.rerun()
      else:
        st.error("Usuário ou senha incorretos.")
  else:
    st.write(f"Sessão ativa: **{st.session_state['usuario_logado']}**")
    if st.button("Sair (Logout)"):
      st.session_state["usuario_logado"] = None
      st.session_state["usuario_role"] = None
      st.rerun()

if st.session_state["usuario_role"] == "admin":
  with st.sidebar.expander("🛠️ Registro de Erros por Livro"):
    logs = obter_logs_erros()
    if logs:
      for l in logs:
        st.caption(f"**[{l[1]}]** 📖 *{l[2]}*")
        st.write(l[3])
        st.code(l[4][:120], language="text")
        st.divider()
    else:
      st.write("Nenhum erro registrado até o momento.")


# --- COMPONENTE DE BUSCA ---
with st.expander("🗂️ Pesquisar no Acervo Salvo ou Digitar Novo", expanded=True):
  livros_disponiveis = listar_livros_salvos()

  termo_busca_input = st.selectbox(
      "🔍 Pesquisar ou Selecionar Obra:",
      options=[""] + livros_disponiveis,
      format_func=lambda x: "Selecione ou digite um livro..." if x == "" else x,
  )

  query_livre = st.text_input(
      "Ou digite o nome de um novo livro para buscar online:",
      value="",
      max_chars=200,
  )

  query_final = (
      query_livre.strip()
      if query_livre.strip()
      else (termo_busca_input if termo_busca_input else "")
  )

  if st.button("🚀 Carregar / Processar Livro", type="primary"):
    if not query_final:
      st.warning("Por favor, selecione ou digite o nome de um livro.")
    else:
      st.session_state["search_active"] = True
      st.session_state["current_query"] = query_final

      with st.spinner("Processando informações do livro..."):
        cache = buscar_no_banco(query_final)
        pdf_info = {"url": "", "autor": "", "paginas": 0, "texto_preview": ""}

        if cache and cache["resumo"]:
          # Tudo vem do banco local, sem rede
          resumo_gerado, resumo_ok = cache["resumo"], True
          pdf_info.update(
              url=cache["pdf_url"] or "",
              autor=cache["autor"] or "",
              paginas=cache["paginas"] or 0,
          )
          st.info("⚡ Dados carregados diretamente do banco de dados local.")
        elif not pode_gerar_com_ia():
          resumo_gerado, resumo_ok = (
              "⏳ **Limite de resumos por sessão atingido.** Tente novamente"
              " mais tarde.",
              False,
          )
        else:
          achado = buscar_pdf_relacionado(query_final)
          if achado:
            pdf_info = achado

          resumo_ok, resumo_gerado = gerar_resumo_gemini(
              query_final, pdf_info["texto_preview"], gemini_key
          )
          st.session_state["geracoes"] += 1

          if resumo_ok:
            salvar_no_banco(
                query_final,
                resumo_gerado,
                fonte="IA",
                pdf_url=pdf_info["url"],
                autor=pdf_info["autor"],
                paginas=pdf_info["paginas"],
            )

        st.session_state["resumo"] = resumo_gerado
        st.session_state["resumo_ok"] = resumo_ok
        st.session_state["pdf_info"] = pdf_info
        st.session_state["audiobooks"] = buscar_audiobooks_youtube(query_final)


# --- EXIBIÇÃO DE RESULTADOS (3 TABS) ---
if st.session_state.get("search_active"):
  q_atual = st.session_state.get("current_query", "")
  st.markdown(f"### Resultados para: **{q_atual.title()}**")

  tab1, tab2, tab3 = st.tabs(
      ["📝 1. Resumo & Word", "📄 2. Leitor de PDF", "🎧 3. Audiobooks & Vídeos"]
  )

  # ABA 1: RESUMO & WORD
  with tab1:
    resumo_txt = st.session_state.get("resumo", "Nenhum resumo gerado.")
    st.markdown("#### 📖 Análise e Resumo Sintético")
    st.markdown(resumo_txt)

    st.markdown("---")
    st.markdown("#### 💾 Exportar Documento")
    if st.session_state.get("resumo_ok"):
      nome_arquivo = re.sub(r"[^\w\-]+", "_", q_atual).strip("_") or "livro"
      st.download_button(
          label="📄 Baixar Resumo em Word (.DOCX)",
          data=criar_arquivo_docx(q_atual, resumo_txt),
          file_name=f"Resumo_{nome_arquivo}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    else:
      st.info("Resumo indisponível para download no momento.")

  # ABA 2: LEITOR DE PDF
  with tab2:
    st.markdown("#### 📄 Documento PDF do Livro")
    pdf_data = st.session_state.get("pdf_info", {})

    if pdf_data.get("url"):
      st.success("✅ PDF associado a este livro!")
      col_info1, col_info2 = st.columns(2)
      with col_info1:
        st.write(f"**Autor:** {pdf_data.get('autor') or 'N/A'}")
        st.write(f"**Páginas:** {pdf_data.get('paginas') or 'N/A'}")
      with col_info2:
        st.markdown(f"🔗 **[Abrir / Baixar PDF na Fonte]({pdf_data['url']})**")

      if pdf_data.get("texto_preview"):
        with st.expander("🔍 Visualizar trecho extraído do PDF"):
          st.text(pdf_data["texto_preview"])
    else:
      st.info("Nenhum PDF associado a esta obra no momento.")

    st.markdown("---")
    st.markdown("#### 📥 Vincular Link Manual de PDF")
    manual_url = st.text_input(
        "Cole aqui o link do arquivo PDF:",
        placeholder="https://exemplo.com/livro.pdf",
    )

    if st.button("Salvar PDF no Banco"):
      if manual_url.strip():
        with st.spinner("Validando PDF e salvando no banco..."):
          try:
            url_manual = manual_url.strip()
            autor, paginas, texto = baixar_pdf_web(url_manual)
            salvo = atualizar_pdf_no_banco(q_atual, url_manual, autor, paginas)

            st.session_state["pdf_info"] = {
                "url": url_manual,
                "autor": autor,
                "paginas": paginas,
                "texto_preview": texto[:TEXTO_PROMPT_MAX],
            }
            if salvo:
              st.toast("PDF vinculado e salvo no banco local!", icon="✅")
            else:
              st.toast(
                  "PDF vinculado só nesta sessão: gere o resumo do livro"
                  " primeiro para salvá-lo no banco.",
                  icon="⚠️",
              )
            st.rerun()
          except Exception as e:
            st.error(f"O link informado não pôde ser lido: {e}")

  # ABA 3: AUDIOBOOKS & VÍDEOS
  with tab3:
    st.markdown("#### 🎧 Opções de Audiobooks no YouTube")
    vids = st.session_state.get("audiobooks", [])

    if vids:
      for video in vids:
        with st.container(border=True):
          col_v1, col_v2 = st.columns([1, 2])
          with col_v1:
            if video.get("image"):
              st.image(video["image"])
            else:
              st.markdown("🎬 **Vídeo do YouTube**")
          with col_v2:
            st.subheader(video["title"])
            st.caption(f"Status: {video.get('views', 'Disponível')}")

            link_vid = video.get("link")
            if link_vid:
              st.markdown(f"👉 **[Assistir / Ouvir no YouTube]({link_vid})**")
              with st.expander("▶️ Player do Vídeo"):
                st.video(link_vid)
    else:
      st.info("Nenhum audiobook encontrado diretamente.")

    st.markdown("---")
    yt_direct_link = (
        "https://www.youtube.com/results?search_query="
        f"{urllib.parse.quote(q_atual + ' audiobook completo')}"
    )
    st.markdown(
        f"🔗 **[Pesquisar '{q_atual} Audiobook' diretamente no YouTube]({yt_direct_link})**"
    )