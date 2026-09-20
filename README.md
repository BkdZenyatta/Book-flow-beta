Título do Projeto: Book Flow — Assistente Inteligente de Leitura, Resumos e Audiobooks

Visão Geral:

O Book Flow é uma aplicação web interativa em Python desenvolvida com Streamlit que centraliza a experiência de busca e consumo de livros. O sistema utiliza a API do Google Gemini para gerar resumos concisos e estruturados, extrai textos de arquivos PDF públicos na web, organiza recomendações de audiobooks do YouTube e permite exportar resumos formatados diretamente em arquivos do Microsoft Word (.docx). Todos os resumos e fontes pesquisadas são salvos em um banco de dados local (SQLite) para consultas instantâneas offline.

🚀 Principais Funcionalidades
Busca Inteligente & Cache Local:

Busca com preenchimento/auto-complete de livros já salvos no acervo local.

Sistema de verificação inteligente: se o livro já foi pesquisado, ele recarrega o resumo do banco local sem gastar cota da API.

Organização em 3 Abas Principais:

📝 1. Resumo & Word: Exibe a síntese gerada pela IA (Visão Geral, Tópicos Principais e Lição Central) e oferece um botão para exportar o resumo formatado em .docx.

📄 2. Leitor de PDF: Busca automática de PDFs públicos do livro na web (ou inserção manual de URL pelo usuário), extraindo metadados como autor e total de páginas e associando o link permanente ao banco de dados.

🎧 3. Audiobooks & Vídeos: Busca e exibe cards com miniaturas, títulos e players incorporados dos melhores audiobooks disponíveis no YouTube.

Painel do Administrador & Sistema de Logs:

Sistema de autenticação na barra lateral para usuário admin.

Registro detalhado de erros por livro na tabela de logs para monitoramento de falhas de conexão ou cota de API.

Mensagens de erro amigáveis tratadas e traduzidas para o usuário (ex: instabilidades de rede, indisponibilidade temporária de modelo 503).

🛠️ Tecnologias Utilizadas
Linguagem: Python 3.10+

Interface Gráfica: Streamlit

Inteligência Artificial: Google GenAI SDK (google-genai / Gemini 2.5/2.0/1.5 Flash)

Banco de Dados: SQLite3 (bookflow.db)

Processamento de Documentos: pypdf (leitura de PDFs) e python-docx (geração de arquivos Word)

Busca e Web Scraping: duckduckgo_search e requests

💾 Estrutura do Banco de Dados (bookflow.db)
resumos: Armazena titulo, resumo, fonte e pdf_url.

erros_log: Registra data_hora, livro, erro_original e mensagem_amigavel.

usuarios: Controla o acesso à área administrativa (username, password, role).
