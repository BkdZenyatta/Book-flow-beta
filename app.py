import datetime
import re
import sqlite3
import urllib.parse
from io import BytesIO
from docx import Document
from duckduckgo_search import DDGS
from google import genai
from pypdf import PdfReader
import requests
import streamlit as st

st.set_page_config(page_title="Book Flow", page_icon="📚", layout="wide")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# --- BANCO DE DADOS LOCAL (SQLite) ---
def init_db():
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()

  # Tabela de Resumos
  c.execute("""
        CREATE TABLE IF NOT EXISTS resumos (
            titulo TEXT PRIMARY KEY,
            resumo TEXT,
            fonte TEXT,
            pdf_url TEXT
        )
    """)

  # Tabela de Erros para Admin (incluindo o campo livro)
  c.execute("""
        CREATE TABLE IF NOT EXISTS erros_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT,
            livro TEXT,
            erro_original TEXT,
            mensagem_amigavel TEXT
        )
    """)

  try:
    c.execute("ALTER TABLE erros_log ADD COLUMN livro TEXT")
  except Exception:
    pass

  # Tabela de Usuários para Controle de Acesso
  c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            username TEXT PRIMARY KEY,
            password TEXT,
            role TEXT
        )
    """)

  c.execute("SELECT username FROM usuarios WHERE username = 'admin'")
  if not c.fetchone():
    c.execute(
        "INSERT INTO usuarios (username, password, role) VALUES (?, ?, ?)",
        ("admin", "7070", "admin"),
    )

  conn.commit()
  conn.close()


def autenticar_usuario(user, pwd):
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()
  c.execute(
      "SELECT role FROM usuarios WHERE username = ? AND password = ?",
      (user.strip(), pwd.strip()),
  )
  res = c.fetchone()
  conn.close()
  return res[0] if res else None


def normalizar_dados_livro(dados_brutos, origem="online"):
  """Padroniza o dicionário de dados garantindo a chave 'resumo'."""
  if origem == "online":
    autores = dados_brutos.get("authors", dados_brutos.get("autor", []))
    if isinstance(autores, list):
      autores_str = ", ".join(autores) if autores else "Autor não informado"
    else:
      autores_str = str(autores) if autores else "Autor não informado"

    resumo = (
        dados_brutos.get("description")
        or dados_brutos.get("synopsis")
        or dados_brutos.get("resumo")
        or "Sem resumo disponível."
    )

    return {
        "titulo": dados_brutos.get("title", "Título desconhecido"),
        "autor": autores_str,
        "resumo": resumo,
        "pdf_url": dados_brutos.get("pdf_url", ""),
    }
  return dados_brutos


def salvar_no_banco(titulo, resumo, fonte="IA", pdf_url=""):
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()
  c.execute(
      """INSERT OR REPLACE INTO resumos (titulo, resumo, fonte, pdf_url) 
              VALUES (?, ?, ?, ?)""",
      (titulo.lower().strip(), resumo, fonte, pdf_url),
  )
  conn.commit()
  conn.close()


def buscar_no_banco(titulo):
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()
  c.execute(
      "SELECT resumo, fonte, pdf_url FROM resumos WHERE titulo = ?",
      (titulo.lower().strip(),),
  )
  resultado = c.fetchone()
  conn.close()
  if resultado:
    return {
        "resumo": resultado[0],
        "fonte": resultado[1],
        "pdf_url": resultado[2],
    }
  return None


def listar_livros_salvos():
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()
  c.execute("SELECT titulo FROM resumos ORDER BY titulo ASC")
  livros = [row[0].title() for row in c.fetchall()]
  conn.close()
  return livros


def traduzir_e_registrar_erro(erro, titulo="Desconhecido"):
  erro_str = str(erro)

  if "getaddrinfo failed" in erro_str or "11001" in erro_str:
    msg = (
        "🌐 **Sem Conexão / Falha de DNS:** Não foi possível conectar aos"
        " servidores. Verifique sua conexão com a internet."
    )
  elif (
      "503" in erro_str
      or "UNAVAILABLE" in erro_str
      or "high demand" in erro_str
  ):
    msg = (
        "⚠️ **Servidor Ocupado:** O modelo do Gemini está com alta demanda"
        " (Erro 503). Tente novamente."
    )
  elif "404" in erro_str or "NOT_FOUND" in erro_str:
    msg = (
        "❌ **Modelo Indisponível:** O modelo de IA solicitado não foi"
        " localizado (Erro 404)."
    )
  elif (
      "API_KEY_INVALID" in erro_str
      or "400" in erro_str
      or "INVALID_ARGUMENT" in erro_str
  ):
    msg = (
        "🔑 **Chave Inválida:** A chave de API do Gemini inserida é inválida"
        " ou expirou."
    )
  else:
    msg = (
        "⚠️ **Instabilidade na Conexão:** Ocorreu uma falha ao comunicar com"
        f" a IA: {erro_str[:120]}..."
    )

  try:
    conn = sqlite3.connect("bookflow.db")
    c = conn.cursor()
    data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute(
        "INSERT INTO erros_log (data_hora, livro, erro_original,"
        " mensagem_amigavel) VALUES (?, ?, ?, ?)",
        (data_hora, titulo.title(), erro_str, msg),
    )
    conn.commit()
    conn.close()
  except Exception:
    pass

  return msg


def obter_logs_erros():
  conn = sqlite3.connect("bookflow.db")
  c = conn.cursor()
  c.execute(
      "SELECT id, data_hora, livro, mensagem_amigavel, erro_original FROM"
      " erros_log ORDER BY id DESC LIMIT 15"
  )
  logs = c.fetchall()
  conn.close()
  return logs


init_db()


# --- FUNÇÕES DE PROCESSAMENTO ---
def baixar_pdf_web(url):
  res = requests.get(url, headers=HEADERS, timeout=12, allow_redirects=True)
  res.raise_for_status()
  content_type = res.headers.get("Content-Type", "").lower()

  if "application/pdf" in content_type or res.content.startswith(b"%PDF"):
    pdf_file = BytesIO(res.content)
    reader = PdfReader(pdf_file)
    texto = "".join([p.extract_text() or "" for p in reader.pages[:5]])
    meta = reader.metadata
    autor = meta.author if meta and meta.author else "Não especificado"
    return autor, len(reader.pages), texto[:2500]
  raise ValueError("O link fornecido não aponta para um arquivo PDF válido.")


def gerar_resumo_gemini(titulo, texto_base, api_key):
  key_limpa = api_key.strip() if api_key else ""
  if not key_limpa:
    return "⚠️ **Chave de API não inserida.** Digite uma chave válida no menu lateral."

  prompt = f"""
    Você é um especialista em literatura. Faça um resumo conciso, envolvente e bem estruturado do livro '{titulo}'.
    
    Estrutura desejada:
    1. **Visão Geral e Contexto**
    2. **Principais Tópicos / Capítulos Claves**
    3. **Conclusão e Ensinamento Central**
    
    Trecho extraído do PDF (se disponível):
    {texto_base if texto_base else "Nenhum texto extraído diretamente."}
    """

  modelos_para_tentar = [
      "gemini-2.5-flash",
      "gemini-2.0-flash",
      "gemini-1.5-flash",
  ]

  ultimo_erro = None
  for mod in modelos_para_tentar:
    try:
      client = genai.Client(api_key=key_limpa)
      response = client.models.generate_content(
          model=mod,
          contents=prompt,
      )
      return response.text
    except Exception as e:
      ultimo_erro = e
      if (
          "503" in str(e)
          or "UNAVAILABLE" in str(e)
          or "getaddrinfo" in str(e)
      ):
        continue
      break

  return traduzir_e_registrar_erro(ultimo_erro, titulo=titulo)


def criar_arquivo_docx(titulo, conteudo_resumo):
  doc = Document()
  doc.add_heading("Book Flow - Resumo", level=0)
  doc.add_heading(f"Obra: {titulo.title()}", level=2)
  doc.add_paragraph("")

  for paragrafo in conteudo_resumo.split("\n"):
    p = paragrafo.strip()
    if p:
      if p.startswith("**") and p.endswith("**"):
        doc.add_heading(p.replace("**", ""), level=3)
      else:
        doc.add_paragraph(p)

  buffer = BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


def buscar_audiobooks_youtube(query):
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
              "views": "Resultado verificado",
              "image": thumb,
          })
  except Exception:
    pass

  if not videos_encontrados:
    try:
      search_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(query + ' audiobook completo')}"
      res = requests.get(search_url, headers=HEADERS, timeout=6)
      video_ids = re.findall(r"watch\?v=([a-zA-Z0-9_-]{11})", res.text)
      vids_unicos = list(dict.fromkeys(video_ids))[:4]

      for v_id in vids_unicos:
        videos_encontrados.append({
            "title": f"Audiobook Completo: {query.title()}",
            "link": f"https://www.youtube.com/watch?v={v_id}",
            "views": "Opção em áudio disponível no YouTube",
            "image": f"https://img.youtube.com/vi/{v_id}/hqdefault.jpg",
        })
    except Exception:
      pass

  return videos_encontrados


# --- INTERFACE STREAMLIT ---
st.title("📚 Book Flow")

# --- SIDEBAR & AREA DE LOGIN ADMIN ---
st.sidebar.header("⚙️ Configurações")

gemini_key = ""
try:
  if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    gemini_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 Chave do Gemini configurada!")
except Exception:
  pass

if not gemini_key:
  gemini_key = st.sidebar.text_input("Chave de API do Gemini:", type="password")

st.sidebar.divider()

if "usuario_logado" not in st.session_state:
  st.session_state["usuario_logado"] = None

with st.sidebar.expander("🔐 Área do Administrador"):
  if st.session_state["usuario_logado"] is None:
    user_input = st.text_input("Usuário:", key="login_user")
    pass_input = st.text_input("Senha:", type="password", key="login_pass")
    if st.button("Entrar"):
      role = autenticar_usuario(user_input, pass_input)
      if role:
        st.session_state["usuario_logado"] = user_input
        st.session_state["usuario_role"] = role
        st.success(f"Bem-vindo, {user_input}!")
        st.rerun()
      else:
        st.error("Usuário ou senha incorretos.")
  else:
    st.write(f"Sessão ativa: **{st.session_state['usuario_logado']}**")
    if st.button("Sair (Logout)"):
      st.session_state["usuario_logado"] = None
      st.session_state["usuario_role"] = None
      st.rerun()

if st.session_state.get("usuario_role") == "admin":
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


# --- COMPONENTE DE BUSCA INTELIGENTE (ESTILO AUTO-COMPLETE) ---
with st.expander(
    "🗂️ Pesquisar no Acervo Salvo ou Digitar Novo", expanded=True
):
  livros_disponiveis = listar_livros_salvos()

  # Campo único estilo barra de busca inteligente do Google
  termo_busca_input = st.selectbox(
      "🔍 Pesquisar ou Selecionar Obra:",
      options=[""] + livros_disponiveis,
      format_func=lambda x: "Selecione ou digite um livro..."
      if x == ""
      else x,
  )

  # Fallback caso queira digitar um termo livre que não está na lista pronta do selectbox
  query_livre = st.text_input(
      "Ou digite o nome de um novo livro para buscar online:", value=""
  )

  # Define qual o termo ativo final considerando qual campo foi preenchido
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
        # 1. Tenta buscar do banco local primeiro
        cache = buscar_no_banco(query_final)

        texto_pdf = ""
        pdf_found_url = cache.get("pdf_url", "") if cache else ""
        pdf_autor = ""
        pdf_paginas = 0

        # Se já tem o resumo salvo no banco local, normaliza e utiliza
        if cache and cache.get("resumo"):
          resumo_gerado = cache["resumo"]
          st.info("⚡ Dados carregados diretamente do banco de dados local.")
        else:
          # 2. Caso contrário, faz a busca online por PDF/Web e gera com IA
          try:
            with DDGS(timeout=8) as ddgs:
              resultados_pdf = list(
                  ddgs.text(f"{query_final} filetype:pdf", max_results=4)
              )
              for item in resultados_pdf:
                try:
                  autor, paginas, texto = baixar_pdf_web(item["href"])
                  texto_pdf = texto
                  pdf_found_url = item["href"]
                  pdf_autor = autor
                  pdf_paginas = paginas
                  break
                except Exception:
                  continue
          except Exception:
            pass

          resumo_gerado = gerar_resumo_gemini(
              query_final, texto_pdf, gemini_key
          )

          # Normaliza e salva no banco local
          if (
              "Erro" not in resumo_gerado
              and "Chave de API" not in resumo_gerado
              and "⚠️" not in resumo_gerado
          ):
            salvar_no_banco(
                query_final, resumo_gerado, fonte="IA", pdf_url=pdf_found_url
            )

        # Processa informações extras de PDF se houver URL válida
        if pdf_found_url and not pdf_autor:
          try:
            autor, paginas, texto = baixar_pdf_web(pdf_found_url)
            pdf_autor, pdf_paginas, texto_pdf = autor, paginas, texto
          except Exception:
            pass

        audiobooks_lista = buscar_audiobooks_youtube(query_final)

        st.session_state["resumo"] = resumo_gerado
        st.session_state["pdf_info"] = {
            "url": pdf_found_url,
            "autor": pdf_autor,
            "paginas": pdf_paginas,
            "texto_preview": texto_pdf,
        }
        st.session_state["audiobooks"] = audiobooks_lista


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
    if (
        "Erro" not in resumo_txt
        and "⚠️" not in resumo_txt
        and "🌐" not in resumo_txt
    ):
      docx_file = criar_arquivo_docx(q_atual, resumo_txt)
      st.download_button(
          label="📄 Baixar Resumo em Word (.DOCX)",
          data=docx_file,
          file_name=f"Resumo_{q_atual.replace(' ', '_')}.docx",
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
        st.write(f"**Autor:** {pdf_data.get('autor', 'N/A')}")
        st.write(f"**Páginas lidas:** {pdf_data.get('paginas')}")
      with col_info2:
        st.markdown(
            f"🔗 **[Abrir / Baixar PDF na Fonte]({pdf_data['url']})**"
        )

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
      if manual_url:
        with st.spinner("Validando PDF e salvando no banco..."):
          try:
            autor, paginas, texto = baixar_pdf_web(manual_url)
            resumo_atual = st.session_state.get("resumo", "")
            salvar_no_banco(q_atual, resumo_atual, pdf_url=manual_url)

            st.session_state["pdf_info"] = {
                "url": manual_url,
                "autor": autor,
                "paginas": paginas,
                "texto_preview": texto,
            }
            st.success("✅ PDF vinculado e salvo no banco local!")
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
              st.image(video["image"], use_container_width=True)
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
    yt_direct_link = f"https://www.youtube.com/results?search_query={urllib.parse.quote(q_atual + ' audiobook completo')}"
    st.markdown(
        f"🔗 **[Pesquisar '{q_atual} Audiobook' diretamente no YouTube]({yt_direct_link})**"
    )