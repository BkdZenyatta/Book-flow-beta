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
from concurrent.futures import ThreadPoolExecutor
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

# Idiomas: nome usado no prompt, região do buscador, códigos do archive.org
# e a expressão usada na busca de audiobooks.
IDIOMAS = {
    "Português": {
        "nome": "português do Brasil",
        "ddg": "br-pt",
        "archive": ["por", "Portuguese"],
        "audiobook": "audiobook completo",
    },
    "English": {
        "nome": "English",
        "ddg": "us-en",
        "archive": ["eng", "English"],
        "audiobook": "full audiobook",
    },
    "Español": {
        "nome": "español",
        "ddg": "es-es",
        "archive": ["spa", "Spanish"],
        "audiobook": "audiolibro completo",
    },
    "Français": {
        "nome": "français",
        "ddg": "fr-fr",
        "archive": ["fra", "fre", "French"],
        "audiobook": "livre audio complet",
    },
    "Deutsch": {
        "nome": "Deutsch",
        "ddg": "de-de",
        "archive": ["deu", "ger", "German"],
        "audiobook": "Hörbuch komplett",
    },
    "Italiano": {
        "nome": "italiano",
        "ddg": "it-it",
        "archive": ["ita", "Italian"],
        "audiobook": "audiolibro completo",
    },
}
IDIOMA_PADRAO = "Português"

# Modelos do Gemini: o app descobre os disponíveis via API; esta lista só é
# usada se a consulta falhar. Os nomes mudam rápido, por isso a busca dinâmica.
MODELOS_FALLBACK = [
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
]
MODELOS_EXCLUIR = (
    "image", "live", "tts", "audio", "transcribe", "embedding",
    "robotics", "computer", "native", "learnlm",
)
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
            paginas INTEGER,
            idioma TEXT
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
    _garantir_coluna(conn, "resumos", "idioma", "TEXT")
    _garantir_coluna(conn, "erros_log", "livro", "TEXT")

    # Importa o banco antigo, caso ele estivesse em outra pasta (a pasta onde
    # o Streamlit foi iniciado). Nada é apagado nem duplicado.
    antigo = Path.cwd() / "bookflow.db"
    try:
      if antigo.exists() and antigo.resolve() != DB_PATH.resolve():
        conn.execute("ATTACH DATABASE ? AS antigo", (str(antigo),))
        conn.execute(
            "INSERT OR IGNORE INTO resumos (titulo, resumo, fonte, pdf_url)"
            " SELECT titulo, resumo, fonte, pdf_url FROM antigo.resumos"
        )
        conn.commit()
        conn.execute("DETACH DATABASE antigo")
    except Exception:
      log.warning("Não foi possível importar o banco antigo", exc_info=True)

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


def salvar_no_banco(
    titulo, resumo, fonte="IA", pdf_url="", autor="", paginas=0,
    idioma=IDIOMA_PADRAO,
):
  with db() as conn:
    conn.execute(
        """INSERT OR REPLACE INTO resumos
           (titulo, resumo, fonte, pdf_url, autor, paginas, idioma)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (_chave(titulo), resumo, fonte, pdf_url, autor, paginas, idioma),
    )


def salvar_pdf_manual(titulo, pdf_url, autor, paginas, idioma=IDIOMA_PADRAO):
  """Guarda o PDF vinculado à mão, mesmo que o resumo ainda não exista."""
  with db() as conn:
    conn.execute(
        """INSERT INTO resumos
           (titulo, resumo, fonte, pdf_url, autor, paginas, idioma)
           VALUES (?, '', 'manual', ?, ?, ?, ?)
           ON CONFLICT(titulo) DO UPDATE SET
             pdf_url = excluded.pdf_url,
             autor = excluded.autor,
             paginas = excluded.paginas""",
        (_chave(titulo), pdf_url, autor, paginas, idioma),
    )


def buscar_no_banco(titulo):
  with db() as conn:
    row = conn.execute(
        "SELECT resumo, fonte, pdf_url, autor, paginas, idioma FROM resumos"
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
        "idioma": row[5] or IDIOMA_PADRAO,  # registros antigos eram em pt
    }
  return None


def listar_livros_salvos():
  with db() as conn:
    rows = conn.execute(
        "SELECT titulo FROM resumos WHERE COALESCE(resumo, '') <> ''"
        " ORDER BY titulo ASC"
    ).fetchall()
  return [r[0].title() for r in rows]


def contar_livros():
  with db() as conn:
    return conn.execute(
        "SELECT COUNT(*) FROM resumos WHERE COALESCE(resumo, '') <> ''"
    ).fetchone()[0]


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
        "❌ **Modelo Indisponível:** Nenhum modelo do Gemini respondeu"
        " (Erro 404). Atualize o SDK com `pip install -U google-genai` e"
        " reinicie o app."
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


# --- DOWNLOAD SEGURO (proteção contra SSRF e arquivos gigantes) ---
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
  """Confere se o início do texto (capa / título) bate com o livro buscado.

  Títulos curtos (até 3 palavras) exigem todas as palavras; os maiores,
  pelo menos 70%.
  """
  t = titulo.lower()
  palavras = re.findall(r"\w{4,}", t) or re.findall(r"\w+", t)
  if not palavras:
    return False
  alvo = texto[:1500].lower()
  minimo = 1.0 if len(palavras) <= 3 else 0.7
  return sum(p in alvo for p in palavras) / len(palavras) >= minimo


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


# --- BUSCA DE PDFs (archive.org > sites preferidos > web geral) ---
def buscar_pdf_archive_org(titulo, idioma):
  """Busca via API oficial do archive.org, filtrando pelo idioma escolhido."""
  termo = " ".join(re.findall(r"\w+", titulo))
  if not termo:
    return None

  langs = " OR ".join(f"language:{c}" for c in IDIOMAS[idioma]["archive"])
  r = requests.get(
      "https://archive.org/advancedsearch.php",
      params={
          "q": f"({termo}) AND mediatype:texts AND ({langs})",
          "fl[]": ["identifier", "title"],
          "sort[]": "downloads desc",  # mais populares primeiro
          "rows": 6,
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


def buscar_pdf_nos_sites(titulo, idioma):
  """Pesquisa nos sites da lista SITES_PREFERIDOS e procura PDFs nas páginas."""
  regiao = IDIOMAS[idioma]["ddg"]
  for dominio in SITES_PREFERIDOS:
    try:
      with DDGS(timeout=8) as ddgs:
        paginas = list(
            ddgs.text(f"site:{dominio} {titulo}", region=regiao, max_results=3)
        )
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


def buscar_pdf_web_geral(titulo, idioma):
  """Último recurso: busca geral por PDFs na web, na região do idioma."""
  with DDGS(timeout=8) as ddgs:
    resultados = list(
        ddgs.text(
            f"{titulo} filetype:pdf",
            region=IDIOMAS[idioma]["ddg"],
            max_results=4,
        )
    )
  for item in resultados:
    url = item.get("href", "")
    if url:
      achado = _tentar_pdf(titulo, url)
      if achado:
        return achado
  return None


def buscar_pdf_relacionado(titulo, idioma):
  """Tenta as fontes em ordem: archive.org, sites preferidos, web geral."""
  for busca in (
      buscar_pdf_archive_org,
      buscar_pdf_nos_sites,
      buscar_pdf_web_geral,
  ):
    try:
      achado = busca(titulo, idioma)
    except Exception:
      log.warning("Fonte %s falhou", busca.__name__, exc_info=True)
      continue
    if achado:
      return achado
  return None


# --- IA (Gemini) ---
@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _modelos_disponiveis(api_key):
  """Pergunta ao Google quais modelos Flash existem hoje (os nomes mudam)."""
  try:
    client = genai.Client(api_key=api_key)
    candidatos = []
    for m in client.models.list():
      nome = (m.name or "").replace("models/", "")
      acoes = getattr(m, "supported_actions", None) or ["generateContent"]
      if (
          nome.startswith("gemini-")
          and "flash" in nome
          and "generateContent" in acoes
          and not any(x in nome for x in MODELOS_EXCLUIR)
      ):
        v = re.search(r"gemini-(\d+(?:\.\d+)?)", nome)
        versao = float(v.group(1)) if v else 0.0
        # estáveis antes de preview, versão mais nova antes, "lite" por último
        candidatos.append(
            (("preview" not in nome), versao, ("lite" not in nome), nome)
        )
    candidatos.sort(reverse=True)
    nomes = [c[3] for c in candidatos][:4]
    if nomes:
      return nomes
  except Exception:
    log.warning("Não foi possível listar os modelos do Gemini", exc_info=True)
  return MODELOS_FALLBACK


def gerar_resumo_gemini(titulo, texto_base, api_key, idioma=IDIOMA_PADRAO):
  """Retorna (ok, texto, trecho_valido).

  trecho_valido: True/False se a IA confirmou/negou que o trecho do PDF é do
  livro; None se não houve resposta sobre isso.
  """
  if not api_key or not api_key.strip():
    return False, (
        "🔑 **Chave de API não inserida.** Digite uma chave no menu lateral."
    ), None

  nome_idioma = IDIOMAS[idioma]["nome"]
  prompt = f"""
    Você é um especialista em literatura. Escreva TODA a resposta em {nome_idioma}.
    Faça um resumo conciso, envolvente e bem estruturado do livro '{titulo}'.
    Se o título for ambíguo (por exemplo, apenas um número), considere a obra literária mais conhecida com esse título.

    Trecho extraído de um PDF (use apenas se for realmente do livro; caso contrário, ignore):
    {texto_base if texto_base else "Nenhum texto extraído."}

    Formato obrigatório da resposta:
    - Linha 1: exatamente "TRECHO_DO_LIVRO: SIM" se o trecho acima pertence a esse livro, ou "TRECHO_DO_LIVRO: NAO" se não pertence ou se não há trecho.
    - Depois, uma linha em branco e o resumo com esta estrutura (traduza os títulos das seções):
    1. **Visão Geral e Contexto**
    2. **Principais Tópicos / Capítulos-Chave**
    3. **Conclusão e Ensinamento Central**
    """

  client = genai.Client(api_key=api_key.strip())
  ultimo_erro = None
  for mod in _modelos_disponiveis(api_key.strip()):
    try:
      resp = client.models.generate_content(model=mod, contents=prompt)
      txt = (resp.text or "").strip()
      valido = None
      m = re.match(r"\s*TRECHO_DO_LIVRO:\s*(SIM|NAO|NÃO)\b[^\n]*\n?", txt, re.I)
      if m:
        valido = m.group(1).upper() == "SIM"
        txt = txt[m.end():].strip()
      if txt:
        return True, txt, valido
      ultimo_erro = ValueError("Resposta vazia ou bloqueada pelo modelo.")
    except Exception as e:
      ultimo_erro = e
      if (
          getattr(e, "code", None) not in CODES_TENTAR_OUTRO
          and "getaddrinfo" not in str(e)
      ):
        break

  return False, traduzir_e_registrar_erro(ultimo_erro, titulo=titulo), None


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


# --- YOUTUBE (só vídeos que existem, no idioma e com o título certo) ---
def _video_id(href):
  m = re.search(r"(?:v=|youtu\.be/|/shorts/|/embed/)([A-Za-z0-9_-]{11})", href)
  return m.group(1) if m else None


def _validar_videos(query, ids):
  """Confere no YouTube (oEmbed) se cada vídeo existe e bate com o livro."""

  def checar(vid):
    link = f"https://www.youtube.com/watch?v={vid}"
    try:
      r = requests.get(
          "https://www.youtube.com/oembed",
          params={"url": link, "format": "json"},
          headers=HEADERS,
          timeout=5,
      )
      if r.status_code != 200:  # removido, privado ou bloqueado
        return None
      d = r.json()
    except Exception:
      return None

    titulo_video = d.get("title", "")
    if not _texto_relevante(query, titulo_video):
      return None
    return {
        "title": titulo_video,
        "link": link,
        "views": f"Canal: {d.get('author_name', 'YouTube')}",
        "image": d.get("thumbnail_url")
        or f"https://img.youtube.com/vi/{vid}/hqdefault.jpg",
    }

  with ThreadPoolExecutor(max_workers=6) as ex:
    return [v for v in ex.map(checar, ids) if v]


def _buscar_audiobooks(query, idioma):
  cfg = IDIOMAS[idioma]
  ids = []

  try:
    with DDGS(timeout=8) as ddgs:
      resultados = list(
          ddgs.text(
              f"site:youtube.com {query} {cfg['audiobook']}",
              region=cfg["ddg"],
              max_results=10,
          )
      )
    for r in resultados:
      vid = _video_id(r.get("href", ""))
      if vid:
        ids.append(vid)
  except Exception:
    log.warning("Busca de audiobooks (DDG) falhou", exc_info=True)

  if len(ids) < 3:  # complementa com a busca do próprio YouTube
    try:
      termo_yt = f"{query} {cfg['audiobook']}"
      search_url = (
          "https://www.youtube.com/results?search_query="
          + urllib.parse.quote(termo_yt)
      )
      res = requests.get(search_url, headers=HEADERS, timeout=6)
      ids += re.findall(r"watch\?v=([a-zA-Z0-9_-]{11})", res.text)
    except Exception:
      log.warning("Busca de audiobooks (scraping) falhou", exc_info=True)

  ids = list(dict.fromkeys(ids))[:12]
  return _validar_videos(query, ids)[:5]


@st.cache_data(ttl=3600, show_spinner=False)
def _audiobooks_cacheado(query, idioma):
  videos = _buscar_audiobooks(query, idioma)
  if not videos:
    raise LookupError("sem resultados")  # exceção não entra no cache
  return videos


def buscar_audiobooks_youtube(query, idioma=IDIOMA_PADRAO):
  try:
    return _audiobooks_cacheado(query, idioma)
  except LookupError:
    return []


# --- INTERFACE STREAMLIT ---
st.title("📚 Book Flow")

for chave_estado, valor in (
    ("usuario_logado", None),
    ("usuario_role", None),
    ("resumo_ok", False),
    ("geracoes", 0),
    ("idioma_busca", IDIOMA_PADRAO),
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

st.sidebar.caption(f"📦 Acervo: {contar_livros()} livro(s) salvo(s)")
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
  with st.sidebar.expander("🗄️ Banco de Dados"):
    st.code(str(DB_PATH), language="text")
    st.write(f"Livros com resumo: **{contar_livros()}**")
    if DB_PATH.exists():
      st.caption(f"Tamanho: {DB_PATH.stat().st_size / 1024:.0f} KB")

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
  idioma = st.selectbox(
      "🌐 Idioma da pesquisa e do resumo:",
      options=list(IDIOMAS),
      index=list(IDIOMAS).index(IDIOMA_PADRAO),
  )

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
      help="Dica: título + autor (ex.: 1984 George Orwell) melhora a precisão.",
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
      st.session_state["idioma_busca"] = idioma

      precisa_recarregar = False
      with st.spinner("Processando informações do livro..."):
        cache = buscar_no_banco(query_final)
        pdf_info = {"url": "", "autor": "", "paginas": 0, "texto_preview": ""}

        if cache and cache["resumo"] and cache["idioma"] == idioma:
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
          pdf_salvo = bool(cache and cache["pdf_url"])
          if pdf_salvo:
            # PDF já vinculado antes (ex.: manualmente): reaproveita
            achado = {
                "url": cache["pdf_url"],
                "autor": cache["autor"] or "",
                "paginas": cache["paginas"] or 0,
                "texto_preview": "",
            }
            try:
              achado["texto_preview"] = baixar_pdf_web(cache["pdf_url"])[2][
                  :TEXTO_PROMPT_MAX
              ]
            except Exception:
              log.info("Não foi possível reler o PDF salvo", exc_info=True)
          else:
            achado = buscar_pdf_relacionado(query_final, idioma)
          if achado:
            pdf_info = achado

          resumo_ok, resumo_gerado, trecho_valido = gerar_resumo_gemini(
              query_final, pdf_info["texto_preview"], gemini_key, idioma
          )
          st.session_state["geracoes"] += 1

          if achado and not pdf_salvo and trecho_valido is False:
            # A IA leu o trecho e disse que não é o livro: descarta o PDF
            log.info("IA rejeitou o PDF encontrado: %s", achado["url"])
            pdf_info = {"url": "", "autor": "", "paginas": 0, "texto_preview": ""}

          if resumo_ok:
            salvar_no_banco(
                query_final,
                resumo_gerado,
                fonte="IA",
                pdf_url=pdf_info["url"],
                autor=pdf_info["autor"],
                paginas=pdf_info["paginas"],
                idioma=idioma,
            )
            precisa_recarregar = True

        st.session_state["resumo"] = resumo_gerado
        st.session_state["resumo_ok"] = resumo_ok
        st.session_state["pdf_info"] = pdf_info
        st.session_state["audiobooks"] = buscar_audiobooks_youtube(
            query_final, idioma
        )

      if precisa_recarregar:  # atualiza a lista do acervo com o livro novo
        st.rerun()


# --- EXIBIÇÃO DE RESULTADOS (3 TABS) ---
if st.session_state.get("search_active"):
  q_atual = st.session_state.get("current_query", "")
  idioma_res = st.session_state.get("idioma_busca", IDIOMA_PADRAO)
  st.markdown(f"### Resultados para: **{q_atual.title()}**")
  flash = st.session_state.pop("flash", None)
  if flash:
    st.toast(flash)

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
            salvar_pdf_manual(
                q_atual, url_manual, autor, paginas, idioma_res
            )

            st.session_state["pdf_info"] = {
                "url": url_manual,
                "autor": autor,
                "paginas": paginas,
                "texto_preview": texto[:TEXTO_PROMPT_MAX],
            }
            st.session_state["flash"] = "✅ PDF vinculado e salvo no banco local!"
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
            st.caption(video.get("views", "Disponível"))

            link_vid = video.get("link")
            if link_vid:
              st.markdown(f"👉 **[Assistir / Ouvir no YouTube]({link_vid})**")
              with st.expander("▶️ Player do Vídeo"):
                st.video(link_vid)
    else:
      st.info("Nenhum audiobook disponível encontrado para esta busca.")

    st.markdown("---")
    termo_yt = q_atual + " " + IDIOMAS[idioma_res]["audiobook"]
    yt_direct_link = (
        "https://www.youtube.com/results?search_query="
        + urllib.parse.quote(termo_yt)
    )
    st.markdown(
        f"🔗 **[Pesquisar '{q_atual} Audiobook' diretamente no YouTube]({yt_direct_link})**"
    )
