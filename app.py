import streamlit as st
import pandas as pd
import io
import os
import tempfile
from datetime import datetime
import time
import base64
from pathlib import Path
import sys
import re
import zipfile
import subprocess
import tempfile
temp_dir = tempfile.gettempdir()

# --- FUNÇÕES AUXILIARES (Mantidas exatamente como no original) ---

def extrair_conteudo_odt(arquivo_bytes):
    """Extrai o conteúdo de um arquivo ODT"""
    with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_file:
        temp_file.write(arquivo_bytes)
        temp_path = temp_file.name

    try:
        with zipfile.ZipFile(temp_path, 'r') as zip_ref:
            content_xml = zip_ref.read('content.xml').decode('utf-8')
        os.unlink(temp_path)
        return content_xml
    except Exception as e:
        st.error(f"Erro ao extrair conteúdo do arquivo ODT: {str(e)}")
        if os.path.exists(temp_path):
             os.unlink(temp_path)
        return None
    finally:
        # Garante que o arquivo temporário seja removido mesmo se ocorrer um erro inesperado antes do unlink
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass # Ignora erros se o arquivo já foi removido

def substituir_no_xml(content_xml, substituicoes):
    """Substitui texto no conteúdo XML do arquivo ODT"""
    texto_modificado = content_xml
    substituicoes_feitas = 0

    # Mapeamento dos nomes das colunas para os placeholders
    mapeamento_colunas = {
        "Cliente": "<Cliente>", "Cidade": "<Cidade>", "Estado": "<Estado>",
        "Número": "<Número>", "Nome": "<Nome>", "Telefone": "<Telefone>",
        "Email": "<Email>", "Modelo": "<Modelo>", "TIPO DE MÁQUINA": "<TIPO DE MÁQUINA>",
        "MODELO DE MÁQUINA": "<MODELO DE MÁQUINA>", "Valor Rompedor": "<Valor Rompedor>",
        "Valor Kit": "<Valor Kit>", "Condição de pagamento": "<Condição de pagamento>",
        "FRETE": "<FRETE>", "Data": "<Data>"
    }

    # Primeiro, substituir os placeholders no formato de database-display
    for coluna, placeholder in mapeamento_colunas.items():
        if placeholder in substituicoes:
            padrao = f'<text:database-display[^>]*text:column-name="{re.escape(coluna)}"[^>]*>([^<]*)</text:database-display>'
            # Usamos uma função lambda para preservar a estrutura original da tag, apenas mudando o conteúdo
            texto_modificado, num_subs = re.subn(
                padrao,
                lambda m: f'<text:database-display text:column-name="{coluna}" text:table-name="Planilha1" text:table-type="table" text:database-name="Formulário propostas Rompedor1">{substituicoes[placeholder]}</text:database-display>',
                texto_modificado
            )
            substituicoes_feitas += num_subs

    # Depois, substituir os placeholders como texto simples (se existirem)
    for placeholder, valor in substituicoes.items():
        padrao_simples = re.escape(placeholder)
        texto_modificado, num_subs_simples = re.subn(padrao_simples, str(valor), texto_modificado)
        substituicoes_feitas += num_subs_simples

    return texto_modificado, substituicoes_feitas


def criar_odt_modificado(arquivo_original_bytes, content_xml_modificado):
    """Cria um novo arquivo ODT com o conteúdo modificado"""
    temp_original_path = None
    temp_modificado_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_original:
            temp_original.write(arquivo_original_bytes)
            temp_original_path = temp_original.name

        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_modificado:
            temp_modificado_path = temp_modificado.name

        with zipfile.ZipFile(temp_original_path, 'r') as zip_original:
            with zipfile.ZipFile(temp_modificado_path, 'w', zipfile.ZIP_DEFLATED) as zip_modificado: # Usar compressão
                for item in zip_original.infolist():
                    if item.filename == 'content.xml':
                        zip_modificado.writestr('content.xml', content_xml_modificado.encode('utf-8')) # Garantir encoding utf-8
                    else:
                        zip_modificado.writestr(item, zip_original.read(item.filename))

        with open(temp_modificado_path, 'rb') as f:
            conteudo_modificado = f.read()

        return conteudo_modificado

    except Exception as e:
        st.error(f"Erro ao criar arquivo ODT modificado: {str(e)}")
        return None
    finally:
        # Limpeza robusta dos arquivos temporários
        if temp_original_path and os.path.exists(temp_original_path):
            os.unlink(temp_original_path)
        if temp_modificado_path and os.path.exists(temp_modificado_path):
            os.unlink(temp_modificado_path)

def converter_para_pdf(odt_bytes, nome_arquivo_base):
    """Converte ODT para PDF usando LibreOffice"""
    libreoffice_path = None
    paths_to_try = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        "/usr/bin/libreoffice",
        "/usr/bin/soffice"
    ]

    for path in paths_to_try:
        if os.path.exists(path):
            libreoffice_path = path
            break

    if not libreoffice_path:
        st.error("⚠️ **LibreOffice não encontrado.** Verifique a instalação ou o caminho no código.")
        return None

    temp_odt_path = None
    temp_pdf_dir = None
    pdf_path = None # Inicializa pdf_path

    try:
        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_odt:
            temp_odt.write(odt_bytes)
            temp_odt_path = temp_odt.name

        temp_pdf_dir = tempfile.mkdtemp()

        comando = [
            libreoffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', temp_pdf_dir,
            temp_odt_path
        ]

        # Usar Popen para melhor controle, especialmente no Windows
        process = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=(os.name == 'nt'))
        stdout, stderr = process.communicate(timeout=120) # Timeout aumentado

        if process.returncode != 0:
            error_message = stderr.decode('utf-8', errors='ignore')
            # Tentar extrair mensagem mais útil do erro do LibreOffice
            if "Error: source file could not be loaded" in error_message:
                 raise Exception("Erro do LibreOffice: O arquivo ODT de origem não pôde ser carregado (pode estar corrompido ou ter permissões incorretas).")
            elif "error while loading shared libraries" in error_message:
                 raise Exception(f"Erro do LibreOffice: Falta de bibliotecas compartilhadas. Detalhes: {error_message}")
            else:
                 raise Exception(f"Erro na conversão (código {process.returncode}): {error_message}")


        # O nome do arquivo PDF gerado pelo LibreOffice será o mesmo do ODT, mas com extensão .pdf
        pdf_filename = os.path.basename(temp_odt_path).replace('.odt', '.pdf')
        pdf_path = os.path.join(temp_pdf_dir, pdf_filename)


        if not os.path.exists(pdf_path):
             # Adicionar verificação do stdout para pistas
             output_message = stdout.decode('utf-8', errors='ignore')
             raise Exception(f"Arquivo PDF não foi gerado em '{temp_pdf_dir}'. Output: {output_message}")


        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()

        return pdf_bytes

    except subprocess.TimeoutExpired:
        st.error("⏳ A conversão para PDF demorou muito (timeout). Tente novamente ou verifique o arquivo ODT.")
        return None
    except Exception as e:
        st.error(f"Falha na conversão para PDF: {str(e)}")
        # Adicionar log extra para depuração
        st.error(f"Comando executado: {' '.join(comando)}")
        if 'stderr' in locals() and stderr: st.error(f"Saída de erro do processo: {stderr.decode('utf-8', errors='ignore')}")
        return None
    finally:
        # Limpeza final
        if temp_odt_path and os.path.exists(temp_odt_path):
            os.unlink(temp_odt_path)
        if pdf_path and os.path.exists(pdf_path):
             os.unlink(pdf_path)
        if temp_pdf_dir and os.path.exists(temp_pdf_dir):
             try:
                  os.rmdir(temp_pdf_dir)
             except OSError:
                  # Pode falhar se o LibreOffice ainda tiver algum lock, mas tentamos
                  st.warning(f"Não foi possível remover o diretório temporário {temp_pdf_dir}. Pode ser necessário remover manualmente.")
                  pass


def formatar_valor_monetario(valor):
    """Formata um valor como moeda brasileira (R$)"""
    try:
        # Tenta converter para float, tratando vírgula como separador decimal se necessário
        if isinstance(valor, str):
            valor = valor.replace('.', '').replace(',', '.')
        valor_float = float(valor)
        # Formatação padrão brasileira
        return f"R$ {valor_float:,.2f}".replace(',', 'v').replace('.', ',').replace('v', '.')
    except (ValueError, TypeError):
        return "R$ 0,00" # Retorna R$ 0,00 se a conversão falhar


def criar_substituicoes(dados):
    """Prepara dicionário de substituições a partir de uma linha (dict) do DataFrame"""
    substituicoes = {}
    data_hoje = datetime.today().strftime("%d/%m/%Y")

    # Mapeamento dos placeholders para as colunas (considerando nomes exatos)
    mapeamento_placeholders = {
        "<Cliente>": "Cliente", "<Cidade>": "Cidade", "<Estado>": "Estado",
        "<Número>": "Número", "<Nome>": "Nome", "<Telefone>": "Telefone",
        "<Email>": "Email", "<Modelo>": "Modelo", "<TIPO DE MÁQUINA>": "TIPO DE MÁQUINA",
        "<MODELO DE MÁQUINA>": "MODELO DE MÁQUINA", "<Valor Rompedor>": "Valor Rompedor",
        "<Valor Kit>": "Valor Kit", "<Condição de pagamento>": "Condição de pagamento",
        "<FRETE>": "FRETE", "<Data>": "Data"
    }

    for placeholder, coluna in mapeamento_placeholders.items():
        valor = dados.get(coluna, "") # Pega o valor da coluna correspondente

        # Tratamento especial para valores monetários
        if coluna in ["Valor Rompedor", "Valor Kit"]:
            valor_formatado = formatar_valor_monetario(valor)
            substituicoes[placeholder] = valor_formatado
        # Tratamento especial para Data
        elif coluna == "Data":
             if pd.isna(valor) or valor == "":
                  substituicoes[placeholder] = data_hoje
             elif isinstance(valor, datetime):
                  substituicoes[placeholder] = valor.strftime("%d/%m/%Y")
             else:
                  # Tenta converter string para data, se falhar usa o valor como está ou data de hoje
                  try:
                       data_obj = pd.to_datetime(valor, errors='coerce')
                       if pd.isna(data_obj):
                            substituicoes[placeholder] = str(valor) if valor else data_hoje
                       else:
                            substituicoes[placeholder] = data_obj.strftime("%d/%m/%Y")
                  except Exception:
                       substituicoes[placeholder] = str(valor) if valor else data_hoje
        # Para outros campos, apenas converte para string
        else:
            substituicoes[placeholder] = str(valor)

    return substituicoes

# --- Configuração da Página Streamlit ---
st.set_page_config(
    page_title="Gerador de Propostas Jardim Equipamentos",
    page_icon="https://i.postimg.cc/cHVj6Mk6/logo.png?text=AJCE+BRASIL",  # Pode usar URL diretamente
    layout="wide",
    initial_sidebar_state="collapsed" # Começa com sidebar recolhida
)

# --- Estilos CSS Customizados (Mantidos) ---
def load_css():
    st.markdown("""
    <style>
        /* Estilos gerais */
        .main .block-container {
             padding-top: 2rem; /* Adiciona espaço no topo */
             padding-bottom: 2rem;
             padding-left: 2rem;
             padding-right: 2rem;
        }
        .main {
             background-color: #f0f2f6; /* Um cinza um pouco mais claro */
        }

        /* Cabeçalho com logo */
        .header-container {
            display: flex;
            flex-direction: column; /* Empilha logo e título */
            align-items: center; /* Centraliza horizontalmente */
            justify-content: center;
            margin-bottom: 2rem; /* Espaço abaixo do header */
            background-color: #ffffff; /* Fundo branco para o header */
            padding: 1rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }

        .logo-img {
             height: 70px; /* Ajuste o tamanho do logo */
             margin-bottom: 0.5rem; /* Espaço entre logo e título */
        }

        .header-title {
            text-align: center;
            margin: 0;
            padding: 0;
            color: #2c3e50; /* Cor escura para o título */
            font-size: 1.8rem; /* Tamanho do título */
            font-weight: 600;
        }

        /* Estilos para as abas */
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            background-color: #e9ecef; /* Fundo da barra de abas */
            padding: 5px;
            border-radius: 6px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 45px; /* Altura da aba */
            white-space: pre-wrap; /* Permite quebra de linha se necessário */
            background-color: #f8f9fa; /* Fundo da aba inativa */
            border-radius: 4px;
            padding: 10px 20px;
            transition: background-color 0.3s ease, color 0.3s ease;
            border: none; /* Remove borda padrão */
            font-weight: 500;
        }
        .stTabs [aria-selected="true"] {
             background-color: #007bff; /* Azul mais vibrante para aba ativa */
             color: white;
             box-shadow: 0 2px 4px rgba(0, 123, 255, 0.3);
        }
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #dee2e6; /* Cor ao passar o mouse */
             color: #333;
        }
         .stTabs [aria-selected="true"]:hover {
             background-color: #0056b3; /* Cor mais escura ao passar mouse na ativa */
             color: white;
         }

        /* Estilos para botões */
        .stButton>button {
            border-radius: 5px;
            padding: 10px 20px; /* Botões maiores */
            font-size: 1rem;
            transition: all 0.2s ease-in-out;
            border: none; /* Remove borda padrão */
        }
        .stButton>button[kind="primary"] {
            background-color: #28a745; /* Verde para botões primários */
            color: white;
        }
         .stButton>button[kind="primary"]:hover {
            background-color: #218838; /* Verde mais escuro no hover */
             box-shadow: 0 2px 5px rgba(40, 167, 69, 0.4);
             transform: translateY(-1px);
        }
         .stButton>button[kind="secondary"] {
            background-color: #6c757d; /* Cinza para botões secundários */
            color: white;
        }
        .stButton>button[kind="secondary"]:hover {
            background-color: #5a6268;
             box-shadow: 0 2px 5px rgba(108, 117, 125, 0.4);
             transform: translateY(-1px);
        }
        /* Estilo específico para botão de download se necessário */
        .stDownloadButton>button {
            background-color: #17a2b8; /* Azul claro para download */
            color: white;
            width: 100%; /* Ocupar largura total do container */
        }
         .stDownloadButton>button:hover {
            background-color: #138496;
             box-shadow: 0 2px 5px rgba(23, 162, 184, 0.4);
             transform: translateY(-1px);
         }


        /* Melhorar aparência de st.container(border=True) */
        [data-testid="stVerticalBlock"] > [style*="border: 1px solid rgba(49, 51, 63, 0.2)"] {
            border-radius: 8px; /* Bordas arredondadas */
            padding: 1.5rem 1.5rem 1rem 1.5rem; /* Espaçamento interno */
            background-color: #ffffff; /* Fundo branco */
            box-shadow: 0 1px 3px rgba(0,0,0,0.05); /* Sombra suave */
            margin-bottom: 1.5rem; /* Espaço abaixo do container */
        }

        /* Rodapé */
        .footer {
            margin-top: 3rem;
            padding: 1.5rem 0;
            border-top: 1px solid #dee2e6; /* Linha divisória mais sutil */
            text-align: center;
            color: #6c757d; /* Cinza escuro para o texto */
            font-size: 0.85rem;
        }
        .footer p {
            margin-bottom: 0.3rem; /* Menos espaço entre parágrafos no footer */
        }
    </style>
    """, unsafe_allow_html=True)

load_css()

# --- Cabeçalho com Logo ---
def render_header():
    st.markdown("""
    <div class="header-container">
        <img class="logo-img" src="https://i.postimg.cc/cHVj6Mk6/logo.png?text=AJCE+BRASIL" alt="Logo Jardim Equipamentos">
        <h1 class="header-title">Gerador de Propostas</h1>
    </div>
    """, unsafe_allow_html=True)

render_header()

# --- Inicialização do Estado da Sessão ---
# Necessário para controlar a aba ativa e armazenar dados entre abas
if 'current_tab' not in st.session_state:
    st.session_state['current_tab'] = "Upload" # Usar nomes descritivos
if 'planilha_data' not in st.session_state:
    st.session_state['planilha_data'] = None # Armazenar o DataFrame aqui
if 'planilha_nome' not in st.session_state:
    st.session_state['planilha_nome'] = None
if 'modelos_info' not in st.session_state:
    st.session_state['modelos_info'] = {} # Dicionário {nome: bytes}
if 'dados_linha_selecionada' not in st.session_state:
    st.session_state['dados_linha_selecionada'] = None
if 'modelo_selecionado_nome' not in st.session_state:
    st.session_state['modelo_selecionado_nome'] = None

# --- Criação das Abas ---
tab_upload, tab_selecao, tab_geracao = st.tabs([
    "📤 1. Upload de Arquivos",
    "📊 2. Seleção de Dados",
    "🖨️ 3. Gerar Proposta"
])


# --- Aba 1: Upload de Arquivos ---
with tab_upload:
    st.header("Passo 1: Faça o Upload dos Arquivos Necessários")
    st.markdown("---") # Linha divisória

    # Seção para Upload da Planilha
    with st.container(border=True):
        st.subheader("Planilha de Propostas (.ods, .xlsx, .xls)")
        st.caption("Selecione a planilha que contém os dados para preencher as propostas.")
        arquivo_planilha = st.file_uploader(
            "Upload da Planilha",
            type=["ods", "xlsx", "xls"],
            key="planilha_upload_widget", # Chave única para o widget
            label_visibility="collapsed"
        )
        if arquivo_planilha:
             # Processa e armazena na sessão imediatamente após o upload
             try:
                  planilha_bytes = arquivo_planilha.getvalue()
                  engine = 'odf' if arquivo_planilha.name.endswith('.ods') else None
                  df = pd.read_excel(io.BytesIO(planilha_bytes), engine=engine)
                  st.session_state['planilha_data'] = df
                  st.session_state['planilha_nome'] = arquivo_planilha.name
                  st.success(f"✅ Planilha '{arquivo_planilha.name}' carregada com sucesso ({len(df)} linhas).")
             except Exception as e:
                  st.error(f"❌ Erro ao ler a planilha: {e}")
                  st.session_state['planilha_data'] = None # Limpa em caso de erro
                  st.session_state['planilha_nome'] = None


    st.divider() # Divisor visual

    # Seção para Upload dos Modelos
    with st.container(border=True):
        st.subheader("Modelos de Proposta (.odt)")
        st.caption("Selecione um ou mais arquivos de modelo no formato ODT.")
        arquivos_modelo = st.file_uploader(
            "Upload de Modelos ODT",
            type=["odt"],
            accept_multiple_files=True,
            key="modelos_upload_widget", # Chave única
            label_visibility="collapsed"
        )
        if arquivos_modelo:
             # Limpa modelos antigos e armazena os novos
             st.session_state['modelos_info'] = {modelo.name: modelo.getvalue() for modelo in arquivos_modelo}
             st.success(f"✅ {len(arquivos_modelo)} modelo(s) ODT carregado(s): {', '.join(st.session_state['modelos_info'].keys())}")


    st.divider()

    # Botão para avançar (só habilita se ambos os uploads foram feitos)
    if st.session_state['planilha_data'] is not None and st.session_state['modelos_info']:
        if st.button("Avançar para Seleção de Dados →", type="primary", key="goto_selecao"):
            st.session_state['current_tab'] = "Seleção"
            st.rerun() # Força a reexecução para mudar de aba
    else:
         st.info("ℹ️ Por favor, faça o upload da planilha E de pelo menos um modelo ODT para continuar.")


# --- Aba 2: Seleção de Dados ---
with tab_selecao:
    st.header("Passo 2: Selecione os Dados para a Proposta")
    st.markdown("---")

    # Verifica se os dados necessários da aba anterior existem
    if st.session_state['planilha_data'] is None or not st.session_state['modelos_info']:
        st.warning("⚠️ Volte ao Passo 1 e faça o upload da planilha e dos modelos ODT.")
        if st.button("← Voltar para Upload", key="back_to_upload_selecao"):
            st.session_state['current_tab'] = "Upload"
            st.rerun()
    else:
        df = st.session_state['planilha_data']

        # Visualização da Planilha Carregada
        with st.expander("👁️ Visualizar Planilha Carregada", expanded=False):
             st.dataframe(df, use_container_width=True, height=300) # Limita altura

        st.divider()

        # Seleção da Linha e do Modelo em Colunas
        col_linha, col_modelo = st.columns(2)

        with col_linha:
            with st.container(border=True):
                st.subheader("Selecione a Linha da Planilha")
                st.caption("Escolha a linha que contém os dados para esta proposta específica.")

                # Input para número da linha (base 1 para o usuário)
                linha_selecionada_usuario = st.number_input(
                     f"Número da linha (de 2 a {len(df) + 1}):", # Mostra o total de linhas + 1 (porque a primeira linha é header)
                     min_value=2,
                     max_value=len(df) + 1,
                     value=st.session_state.get('last_selected_line', 2), # Lembra última linha selecionada
                     step=1,
                     key="linha_input_selecao"
                 )

                # Converte para índice baseado em zero (0 = primeira linha de dados)
                linha_indice_zero = linha_selecionada_usuario - 2

                if 0 <= linha_indice_zero < len(df):
                     # Armazena os dados da linha selecionada (como dicionário) e a linha selecionada
                     st.session_state['dados_linha_selecionada'] = df.iloc[linha_indice_zero].fillna('').to_dict()
                     st.session_state['last_selected_line'] = linha_selecionada_usuario # Salva para próxima vez

                     # Mostra um preview dos dados selecionados
                     with st.expander("🔍 Pré-visualizar Dados da Linha Selecionada", expanded=True):
                          # Mostra alguns campos chave
                          preview_data = {k: v for k, v in st.session_state['dados_linha_selecionada'].items() if k in ['Cliente', 'Modelo', 'Valor Rompedor', 'Valor Kit', 'Data']}
                          st.dataframe(pd.Series(preview_data).astype(str), use_container_width=True)

                else:
                     st.error(f"❌ Linha {linha_selecionada_usuario} inválida. Selecione um valor entre 2 e {len(df) + 1}.")
                     st.session_state['dados_linha_selecionada'] = None # Limpa se a linha for inválida


        with col_modelo:
             with st.container(border=True):
                 st.subheader("Selecione o Modelo ODT")
                 st.caption("Escolha qual modelo ODT será usado para esta proposta.")
                 nomes_modelos = list(st.session_state['modelos_info'].keys())

                 if nomes_modelos:
                      modelo_selecionado = st.selectbox(
                           "Modelos Disponíveis:",
                           options=nomes_modelos,
                           index=nomes_modelos.index(st.session_state.get('modelo_selecionado_nome', nomes_modelos[0])) if st.session_state.get('modelo_selecionado_nome') in nomes_modelos else 0, # Lembra último selecionado
                           key="modelo_select_widget"
                      )
                      st.session_state['modelo_selecionado_nome'] = modelo_selecionado # Armazena o nome do modelo selecionado
                      st.info(f"📄 Modelo selecionado: **{modelo_selecionado}**")
                 else:
                      st.error("Nenhum modelo ODT encontrado. Volte ao Passo 1.")
                      st.session_state['modelo_selecionado_nome'] = None

        st.divider()

        # Botões de Navegação
        col_btn_back, col_btn_next = st.columns(2)
        with col_btn_back:
            if st.button("← Voltar para Upload", key="back_to_upload_selecao_2", use_container_width=True):
                st.session_state['current_tab'] = "Upload"
                st.rerun()

        with col_btn_next:
             # Habilita o botão de avançar apenas se linha e modelo válidos foram selecionados
             if st.session_state['dados_linha_selecionada'] is not None and st.session_state['modelo_selecionado_nome'] is not None:
                 if st.button("Avançar para Gerar Proposta →", type="primary", key="goto_geracao", use_container_width=True):
                     st.session_state['current_tab'] = "Geração"
                     st.rerun()
             else:
                  st.button("Avançar para Gerar Proposta →", type="primary", key="goto_geracao_disabled", use_container_width=True, disabled=True) # Botão desabilitado


# --- Aba 3: Gerar Proposta ---
with tab_geracao:
    st.header("Passo 3: Revise e Gere a Proposta em PDF")
    st.markdown("---")

    # Verifica se os dados necessários das abas anteriores existem
    if st.session_state.get('dados_linha_selecionada') is None or st.session_state.get('modelo_selecionado_nome') is None:
        st.warning("⚠️ Por favor, complete os Passos 1 e 2 primeiro (selecione uma linha válida e um modelo).")
        if st.button("← Voltar para Seleção", key="back_to_selecao_geracao"):
            st.session_state['current_tab'] = "Seleção"
            st.rerun()
    else:
        dados_linha = st.session_state['dados_linha_selecionada']
        nome_modelo_selecionado = st.session_state['modelo_selecionado_nome']
        modelo_bytes = st.session_state['modelos_info'].get(nome_modelo_selecionado)

        if not modelo_bytes:
             st.error(f"❌ Erro: Modelo ODT '{nome_modelo_selecionado}' não encontrado na memória. Volte ao Passo 1.")
        else:

            # Container para Revisão
            with st.container(border=True):
                st.subheader("Revisão das Informações")
                substituicoes = criar_substituicoes(dados_linha) # Gera as substituições

                # Mostra informações chave em colunas
                col_rev1, col_rev2 = st.columns(2)
                with col_rev1:
                     st.markdown("**Cliente e Contato:**")
                     st.text_input("Cliente:", value=substituicoes.get("<Cliente>", ""), disabled=True, key="rev_cliente")
                     st.text_input("Contato (Nome):", value=substituicoes.get("<Nome>", ""), disabled=True, key="rev_nome")
                     st.text_input("Local:", value=f"{substituicoes.get('<Cidade>', '')}/{substituicoes.get('<Estado>', '')}", disabled=True, key="rev_local")

                with col_rev2:
                     st.markdown("**Produto e Valores:**")
                     st.text_input("Modelo Proposta:", value=substituicoes.get("<Modelo>", ""), disabled=True, key="rev_modelo_prod")
                     st.text_input("Valor Rompedor:", value=substituicoes.get("<Valor Rompedor>", ""), disabled=True, key="rev_val_romp")
                     st.text_input("Valor Kit:", value=substituicoes.get("<Valor Kit>", ""), disabled=True, key="rev_val_kit")


                # Expander para ver todas as substituições
                with st.expander("Ver todas as substituições que serão feitas no documento"):
                     substituicoes_df = pd.DataFrame({
                         'Placeholder no Documento': list(substituicoes.keys()),
                         'Valor a ser Inserido': [str(v) for v in substituicoes.values()] # Garante que tudo é string
                     })
                     st.dataframe(substituicoes_df, hide_index=True, use_container_width=True)


            st.divider()

            # Botão para Gerar o PDF
            if st.button("🚀 Gerar Documento PDF Agora", type="primary", key="generate_pdf_final", use_container_width=True):
                 # ---> INÍCIO DA MODIFICAÇÃO <---
                 pdf_bytes_result = None  # Variável para guardar os bytes do PDF gerado
                 pdf_filename_result = None # Variável para guardar o nome do arquivo

                 # Usar st.status para mostrar o progresso
                 with st.status("⚙️ Iniciando geração da proposta...", expanded=True) as status:
                    try:
                        status.update(label="1/4 - Extraindo conteúdo do modelo ODT...")
                        st.write(f"📄 Usando modelo: {nome_modelo_selecionado}")
                        content_xml = extrair_conteudo_odt(modelo_bytes)
                        if not content_xml:
                             raise ValueError("Falha ao extrair 'content.xml' do modelo ODT.")
                        st.write("✅ Conteúdo extraído.")

                        status.update(label="2/4 - Aplicando substituições nos dados...")
                        content_xml_modificado, num_substituicoes = substituir_no_xml(content_xml, substituicoes)
                        if num_substituicoes == 0:
                             st.warning("⚠️ Nenhuma substituição foi feita. Verifique placeholders no modelo ODT.")
                        # Mesmo que 0 substituições, continua o processo, pode ser intencional
                        st.write(f"✅ {num_substituicoes} substituições realizadas.")


                        status.update(label="3/4 - Recriando arquivo ODT modificado...")
                        documento_odt_modificado = criar_odt_modificado(modelo_bytes, content_xml_modificado)
                        if not documento_odt_modificado:
                             raise ValueError("Falha ao recriar o arquivo ODT modificado.")
                        st.write("✅ Documento ODT modificado criado.")

                        # Define o nome do arquivo PDF (FUNCIONALIDADE ORIGINAL MANTIDA)
                        nome_base = dados_linha.get("NOME DO ARQUIVO")
                        if not nome_base:
                             nome_cliente = str(dados_linha.get('Cliente', 'Proposta')).replace(' ', '_').replace('/','-')
                             nome_base = f"Proposta_{nome_cliente}_{datetime.now().strftime('%Y%m%d')}"
                        nome_arquivo_pdf = f"{nome_base}.pdf"

                        status.update(label=f"4/4 - Convertendo para PDF ('{nome_arquivo_pdf}')... (pode levar alguns segundos)")
                        pdf_bytes = converter_para_pdf(documento_odt_modificado, nome_base)
                        if not pdf_bytes:
                             raise ValueError("Falha ao converter o documento ODT para PDF usando LibreOffice.")

                        # Armazena os resultados nas variáveis locais se tudo deu certo
                        pdf_bytes_result = pdf_bytes
                        pdf_filename_result = nome_arquivo_pdf

                        # Atualiza o status para sucesso (recolhido)
                        status.update(label="🎉 Proposta gerada com sucesso!", state="complete", expanded=False)
                        # NÃO coloca o botão de download aqui dentro

                    except (ValueError, Exception) as e:
                        # Se qualquer etapa falhar, atualiza o status para erro (expandido)
                        status.update(label=f"❌ Erro ao gerar proposta: {str(e)}", state="error", expanded=True)
                        # st.error(f"Detalhes: {str(e)}") # O erro já aparece no status expandido

                 # --- Botão de Download FORA do bloco 'with st.status' ---
                 # Verifica se as variáveis de resultado foram preenchidas (indicando sucesso)
                 if pdf_bytes_result and pdf_filename_result:
                      st.success(f"✅ Documento '{pdf_filename_result}' pronto!") # Mensagem de sucesso visível
                      st.download_button(
                           label=f"📥 Baixar {pdf_filename_result}",
                           data=pdf_bytes_result,
                           file_name=pdf_filename_result,
                           mime="application/pdf",
                           key="download_pdf_final_btn", # Pode manter a mesma chave
                           use_container_width=True,
                           type="primary" # Destaca o botão de download
                      )

            st.divider()

            # Botões de Navegação inferiores
            col_btn_back_geracao, col_btn_new_geracao = st.columns(2)
            with col_btn_back_geracao:
                 if st.button("← Voltar para Seleção", key="back_to_selecao_geracao_2", use_container_width=True):
                      st.session_state['current_tab'] = "Seleção"
                      st.rerun()
            with col_btn_new_geracao:
                 if st.button("✨ Iniciar Nova Proposta (Voltar ao Início)", key="new_proposal_geracao", use_container_width=True):
                      # Limpa estado da sessão relevante para uma nova proposta, mas mantém modelos carregados
                      st.session_state['current_tab'] = "Upload"
                      st.session_state['planilha_data'] = None
                      st.session_state['planilha_nome'] = None
                      # Mantém st.session_state['modelos_info']
                      st.session_state['dados_linha_selecionada'] = None
                      st.session_state['modelo_selecionado_nome'] = None
                      if 'last_selected_line' in st.session_state: del st.session_state['last_selected_line'] # Reseta a linha lembrada
                      st.rerun()


# --- Rodapé (Mantido) ---
st.markdown("---") # Linha divisória antes do rodapé
st.markdown("""
<div class="footer">
    <p>Jardim Equipamentos - Gerador de Propostas Comerciais</p>
    <p>© 2025 - Todos os direitos reservados - Desenvolvido por Rodrigo Ferreira</p>
</div>
""", unsafe_allow_html=True)

# --- Script para Navegação entre Tabs (CORRIGIDO) ---
# Este script JS ainda é uma forma comum de controlar as abas programaticamente
# Mapeia os nomes das abas para índices (0, 1, 2)
tab_map = {"Upload": 0, "Seleção": 1, "Geração": 2}
current_tab_index = tab_map.get(st.session_state['current_tab'], 0) # Pega o índice da aba atual

if st.session_state['current_tab'] != "Upload": # Só executa se não for a primeira aba
    js = f"""
    <script>
        function selectTab() {{
            const tabIndex = {current_tab_index};
            // Seleciona os botões das abas DENTRO do iframe pai onde o Streamlit renderiza
            const tabs = parent.document.querySelectorAll('button[data-baseweb="tab"]');

            // Verifica se o número de abas encontrado é maior que o índice desejado
            if (tabs && tabs.length > tabIndex) {{
                // Clica na aba correta
                tabs[tabIndex].click();
            }} else {{
                 // Log de aviso no console do navegador se a aba não for encontrada
                 // CORRIGIDO: Usando vírgulas para separar argumentos no console.warn
                 console.warn('Streamlit Tabs:', 'Tab index', tabIndex, 'not found or tabs not rendered yet. Available tabs:', tabs ? tabs.length : 0);
             }}
        }}
        // Executa a função 'selectTab' após um pequeno atraso (150ms)
        // para dar tempo ao Streamlit de renderizar os elementos das abas no DOM.
         if (window.parent) {{ // Garante que estamos em um contexto de iframe
            setTimeout(selectTab, 150);
         }}
    </script>
    """
    st.components.v1.html(js, height=0, width=0)
