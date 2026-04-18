import streamlit as st
import requests
from msal import ConfidentialClientApplication
import pandas as pd
import io
from datetime import datetime
import sqlite3
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os

# Configuração da página
st.set_page_config(
    page_title="Sistema de Avaliação - Rezende Energia",
    page_icon="📋",
    layout="wide"
)

# CSS customizado com as cores da empresa
st.markdown("""
    <style>
    .main {
        background-color: #ffffff;
    }
    .stButton>button {
        background-color: #F7931E;
        color: #000000;
        font-weight: bold;
        border: 2px solid #000000;
        border-radius: 5px;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #000000;
        color: #F7931E;
        border: 2px solid #F7931E;
    }
    h1, h2, h3 {
        color: #000000;
    }
    .highlight {
        background-color: #F7931E;
        color: #000000;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Credenciais Azure AD (usando st.secrets)
try:
    CLIENT_ID = st.secrets["azure"]["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["azure"]["CLIENT_SECRET"]
    TENANT_ID = st.secrets["azure"]["TENANT_ID"]
    LOGO_PATH = st.secrets["paths"]["LOGO_PATH"]
except KeyError as e:
    st.error(f"⚠️ Configuração faltando no secrets: {e}")
    st.info("Por favor, configure o arquivo .streamlit/secrets.toml")
    st.stop()
except FileNotFoundError:
    st.error("⚠️ Arquivo secrets.toml não encontrado!")
    st.info("Crie o arquivo .streamlit/secrets.toml na raiz do projeto")
    st.stop()


# Função para gerar PDF da avaliação
def gerar_pdf_avaliacao(dados_avaliacao, nome_arquivo=None):
    """
    Gera um PDF da avaliação com a logo da empresa
    dados_avaliacao: dicionário com os dados da avaliação
    """
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"Avaliacao_{dados_avaliacao['colaborador'].replace(' ', '_')}_{timestamp}.pdf"

    # Criar buffer para o PDF
    buffer = io.BytesIO()

    # Configurar documento
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    # Container para elementos do PDF
    elements = []

    # Estilos
    styles = getSampleStyleSheet()

    # Estilo customizado para título
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#000000'),
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )

    # Estilo para subtítulos
    subtitulo_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#F7931E'),
        spaceAfter=12,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )

    # Estilo para texto normal
    texto_style = ParagraphStyle(
        'CustomBody',
        parent=styles['BodyText'],
        fontSize=10,
        textColor=colors.HexColor('#000000'),
        alignment=TA_LEFT,
        fontName='Helvetica'
    )

    # Adicionar logo se existir
    if os.path.exists(LOGO_PATH):
        try:
            logo = Image(LOGO_PATH, width=3 * cm, height=1.5 * cm)
            logo.hAlign = 'CENTER'
            elements.append(logo)
            elements.append(Spacer(1, 0.5 * cm))
        except Exception as e:
            st.warning(f"Não foi possível adicionar a logo: {e}")

    # Título
    elements.append(Paragraph("FICHA DE AVALIAÇÃO DE EXPERIÊNCIA", titulo_style))
    elements.append(Spacer(1, 0.5 * cm))

    # Informações básicas
    data_atual = datetime.now().strftime('%d/%m/%Y')

    info_basica = [
        ['Data da Avaliação:', data_atual],
        ['Tipo de Avaliação:', dados_avaliacao['tipo_avaliacao']],
        ['', ''],
        ['Avaliador:', dados_avaliacao['avaliador']],
        ['Cargo do Avaliador:', dados_avaliacao['cargo_avaliador']],
        ['Região do Avaliador:', dados_avaliacao.get('regiao_avaliador', '')],
        ['', ''],
        ['Colaborador:', dados_avaliacao['colaborador']],
        ['Cargo do Colaborador:', dados_avaliacao['cargo']],
        ['Região do Colaborador:', dados_avaliacao.get('regiao_colaborador', '')],
    ]

    table_info = Table(info_basica, colWidths=[5 * cm, 12 * cm])
    table_info.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F7931E')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#000000')),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 8),
    ]))

    elements.append(table_info)
    elements.append(Spacer(1, 0.8 * cm))

    # Critérios de avaliação
    elements.append(Paragraph("CRITÉRIOS DE AVALIAÇÃO", subtitulo_style))
    elements.append(Spacer(1, 0.3 * cm))

    criterios = [
        ('ADAPTAÇÃO AO TRABALHO', dados_avaliacao['adaptacao']),
        ('INTERESSE', dados_avaliacao['interesse']),
        ('RELACIONAMENTO SOCIAL', dados_avaliacao['relacionamento']),
        ('CAPACIDADE DE APRENDIZAGEM', dados_avaliacao['capacidade']),
    ]

    for titulo, resposta in criterios:
        elements.append(Paragraph(f"<b>{titulo}</b>", texto_style))
        elements.append(Spacer(1, 0.2 * cm))

        # Criar tabela para a resposta
        resposta_table = Table([[resposta]], colWidths=[17 * cm])
        resposta_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F5F5F5')),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('PADDING', (0, 0), (-1, -1), 10),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elements.append(resposta_table)
        elements.append(Spacer(1, 0.4 * cm))

    # Classificação e Definição
    elements.append(Spacer(1, 0.3 * cm))

    classificacao_def = [
        ['Classificação Geral:', dados_avaliacao['classificacao']],
        ['Definição:', dados_avaliacao['definicao']],
    ]

    table_final = Table(classificacao_def, colWidths=[5 * cm, 12 * cm])
    table_final.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F7931E')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#000000')),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 8),
    ]))

    elements.append(table_final)
    elements.append(Spacer(1, 1.5 * cm))

    # Assinaturas
    assinaturas = [
        ['_' * 40, '_' * 40],
        ['Assinatura do Avaliador', 'Assinatura do Presidente'],
    ]

    table_assinatura = Table(assinaturas, colWidths=[8.5 * cm, 8.5 * cm])
    table_assinatura.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    elements.append(table_assinatura)

    # Construir PDF
    doc.build(elements)

    # Retornar buffer
    buffer.seek(0)
    return buffer, nome_arquivo


# Inicializar banco de dados
def init_db():
    conn = sqlite3.connect('avaliacoes.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS avaliacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            avaliador TEXT NOT NULL,
            colaborador TEXT NOT NULL,
            cargo TEXT,
            cargo_avaliador TEXT,
            regional TEXT,
            tipo_avaliacao TEXT,
            adaptacao TEXT,
            interesse TEXT,
            relacionamento TEXT,
            capacidade TEXT,
            classificacao TEXT,
            definicao TEXT,
            regiao_avaliador TEXT,
            regiao_colaborador TEXT,
            data_avaliacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Verificar e adicionar colunas se não existirem
    try:
        c.execute("SELECT cargo_avaliador FROM avaliacoes LIMIT 1")
    except sqlite3.OperationalError:
        c.execute("ALTER TABLE avaliacoes ADD COLUMN cargo_avaliador TEXT")
        conn.commit()
    
    try:
        c.execute("SELECT regiao_avaliador FROM avaliacoes LIMIT 1")
    except sqlite3.OperationalError:
        c.execute("ALTER TABLE avaliacoes ADD COLUMN regiao_avaliador TEXT")
        conn.commit()
    
    try:
        c.execute("SELECT regiao_colaborador FROM avaliacoes LIMIT 1")
    except sqlite3.OperationalError:
        c.execute("ALTER TABLE avaliacoes ADD COLUMN regiao_colaborador TEXT")
        conn.commit()

    conn.close()


# Salvar avaliação no banco
def salvar_avaliacao(dados):
    conn = sqlite3.connect('avaliacoes.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO avaliacoes (
            avaliador, colaborador, cargo, cargo_avaliador, regional, tipo_avaliacao,
            adaptacao, interesse, relacionamento, capacidade, 
            classificacao, definicao, regiao_avaliador, regiao_colaborador
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', dados)
    conn.commit()
    conn.close()


# Buscar avaliações do banco
def buscar_avaliacoes():
    conn = sqlite3.connect('avaliacoes.db')
    df = pd.read_sql_query("SELECT * FROM avaliacoes ORDER BY data_avaliacao DESC", conn)
    conn.close()
    return df


# Verificar se colaborador já foi avaliado
def ja_foi_avaliado(colaborador, tipo_avaliacao):
    conn = sqlite3.connect('avaliacoes.db')
    c = conn.cursor()
    c.execute('''
        SELECT COUNT(*) FROM avaliacoes 
        WHERE colaborador = ? AND tipo_avaliacao = ?
    ''', (colaborador, tipo_avaliacao))
    count = c.fetchone()[0]
    conn.close()
    return count > 0


# Baixar dados do SharePoint
@st.cache_data(ttl=3600)
def download_excel_sharepoint():
    try:
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )

        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            headers = {"Authorization": f"Bearer {result['access_token']}"}

            search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='Base Colaboradores - Rezende Energia.xlsx')"
            site_response = requests.get(site_url, headers=headers)

            if site_response.status_code == 200:
                site_data = site_response.json()
                site_id = site_data['id']

                search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='Base de Colaboradores - Rezende Energia')"
                search_response = requests.get(search_url, headers=headers)

                if search_response.status_code == 200:
                    search_data = search_response.json()
                    files_found = search_data.get('value', [])

                    for item in files_found:
                        if 'Base de Colaboradores - Rezende Energia' in item['name']:
                            download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                            download_response = requests.get(download_url, headers=headers)

                            if download_response.status_code == 200:
                                df = pd.read_excel(io.BytesIO(download_response.content), sheet_name="COLABORADORES ATIVOS")
                                return df
        return None
    except Exception as e:
        st.error(f"Erro ao baixar dados: {e}")
        return None


# Identificar avaliadores - ATUALIZADO COM NOVOS COORDENADORES
def identificar_avaliadores(df):
    # Verificar se df é válido
    if df is None or df.empty:
        return []
    
    cargos_avaliadores = [
        'SUPERVISOR', 
        'LIDER DE FROTA', 
        'GERENTE OPERACIONAL', 
        'COORDENADOR OPERACIONAL', 
        'ANALISTA FINANCEIRO', 
        'GERENTE DE QSMS',
        'GESTORA DE DEPARTEMENTO PESSOAL/ RECURSOS HUMANOS',
        'GESTORA DE DEPARTAMENTO PESSOAL/ RECURSOS HUMANOS', 'GERENTE GERAL'  # variação de escrita
    ]
    
    try:
        # Buscar avaliadores pelos cargos na planilha
        avaliadores = df[df.iloc[:, 8].str.upper().isin(cargos_avaliadores)]
        lista_avaliadores = avaliadores.iloc[:, 0].tolist()
    except Exception as e:
        st.warning(f"Erro ao buscar avaliadores na planilha: {e}")
        lista_avaliadores = []
    
    # Adicionar avaliadores fixos (garantindo que sempre apareçam)
    avaliadores_fixos = [
        'GABRIELLE ELLIBOX DE LIRA',
        'VINICIUS OLIVEIRA AMARAL DE SOUZA'
    ]
    
    # Combinar listas e remover duplicatas
    lista_completa = list(set(lista_avaliadores + avaliadores_fixos))
    
    return sorted(lista_completa)


# Identificar colaboradores para avaliação
def identificar_colaboradores_para_avaliacao(df):
    hoje = datetime.now()
    colaboradores_40_dias = []
    colaboradores_80_dias = []

    for idx, row in df.iterrows():
        try:
            nome = row.iloc[0]
            cargo = str(row.iloc[8]) if pd.notna(row.iloc[8]) else "Cargo não informado"
            regiao = str(row.iloc[12]) if pd.notna(row.iloc[12]) else "Região não informada"
            data_admissao = pd.to_datetime(row.iloc[9])
            dias_desde_admissao = (hoje - data_admissao).days

            if 37 <= dias_desde_admissao <= 43:
                colaboradores_40_dias.append({
                    'nome': nome,
                    'cargo': cargo,
                    'regiao': regiao,
                    'data_admissao': data_admissao.strftime('%d/%m/%Y'),
                    'dias_empresa': dias_desde_admissao
                })
            elif 77 <= dias_desde_admissao <= 83:
                colaboradores_80_dias.append({
                    'nome': nome,
                    'cargo': cargo,
                    'regiao': regiao,
                    'data_admissao': data_admissao.strftime('%d/%m/%Y'),
                    'dias_empresa': dias_desde_admissao
                })
        except:
            continue

    return colaboradores_40_dias, colaboradores_80_dias


# Inicializar banco de dados
init_db()

# Header
st.title("📋 Sistema de Avaliação de Experiência")
st.markdown("### Rezende Energia")
st.markdown("---")

# Sidebar - Menu
menu = st.sidebar.selectbox(
    "Menu",
    ["Dashboard", "Nova Avaliação", "Histórico de Avaliações"]
)

# Carregar dados
with st.spinner("Carregando dados do SharePoint..."):
    df = download_excel_sharepoint()

if df is None:
    st.error("❌ Erro ao carregar dados do SharePoint.")
    st.warning("**Possíveis causas:**")
    st.markdown("""
    - Credenciais Azure AD incorretas ou expiradas
    - Problemas de conexão com a internet
    - Arquivo 'Base de Colaboradores - Rezende Energia' não encontrado no SharePoint
    - Permissões insuficientes no SharePoint
    """)
    st.info("💡 **Soluções:**")
    st.markdown("""
    1. Verifique o arquivo `.streamlit/secrets.toml`
    2. Confirme as credenciais: CLIENT_ID, CLIENT_SECRET, TENANT_ID
    3. Verifique se o arquivo existe no SharePoint: Sites > Intranet
    4. Tente recarregar a página
    """)
    
    # Mostrar botão para recarregar
    if st.button("🔄 Tentar Recarregar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    st.stop()

if df.empty:
    st.error("❌ O arquivo do SharePoint está vazio.")
    st.stop()

# DASHBOARD
if menu == "Dashboard":
    st.header("📊 Dashboard de Avaliações")

    if df is None or df.empty:
        st.error("❌ Não foi possível carregar os dados. Verifique a conexão com o SharePoint.")
        st.stop()

    avaliadores = identificar_avaliadores(df)
    colab_40, colab_80 = identificar_colaboradores_para_avaliacao(df)

    # Métricas
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("👥 Avaliadores", len(avaliadores))

    with col2:
        st.metric("📋 Avaliações 40 dias", len(colab_40))

    with col3:
        st.metric("📋 Avaliações 80 dias", len(colab_80))

    with col4:
        total_avaliacoes = len(buscar_avaliacoes())
        st.metric("✅ Avaliações Realizadas", total_avaliacoes)

    st.markdown("---")

    # Colaboradores pendentes
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🕐 Avaliações de 40 dias pendentes")
        if colab_40:
            for col in colab_40:
                avaliado = ja_foi_avaliado(col['nome'], "40 dias")
                status = "✅" if avaliado else "⏳"
                st.write(
                    f"{status} **[{col['regiao']}] [{col['cargo']}] {col['nome']}** - Admitido em {col['data_admissao']} ({col['dias_empresa']} dias)")
        else:
            st.info("Nenhum colaborador no período de 40 dias")

    with col2:
        st.subheader("🕐 Avaliações de 80 dias pendentes")
        if colab_80:
            for col in colab_80:
                avaliado = ja_foi_avaliado(col['nome'], "80 dias")
                status = "✅" if avaliado else "⏳"
                st.write(
                    f"{status} **[{col['regiao']}] [{col['cargo']}] {col['nome']}** - Admitido em {col['data_admissao']} ({col['dias_empresa']} dias)")
        else:
            st.info("Nenhum colaborador no período de 80 dias")

# NOVA AVALIAÇÃO
elif menu == "Nova Avaliação":
    st.header("📝 Nova Avaliação de Experiência")

    if df is None or df.empty:
        st.error("❌ Não foi possível carregar os dados. Verifique a conexão com o SharePoint.")
        st.stop()

    avaliadores = identificar_avaliadores(df)
    todos_colaboradores = sorted(df.iloc[:, 0].dropna().tolist())

    st.subheader("Informações Básicas")

    col1, col2 = st.columns(2)

    with col1:
        avaliador = st.selectbox("Supervisor/Coordenador (Avaliador) *", avaliadores)
        # Buscar cargo e região do avaliador
        cargo_avaliador = ""
        regiao_avaliador = ""
        if avaliador:
            # Verificar se é um dos avaliadores fixos
            if avaliador == 'GABRIELLE ELLIBOX DE LIRA':
                cargo_avaliador = 'GESTORA DE DEPARTEMENTO PESSOAL/ RECURSOS HUMANOS'
                # Buscar região na planilha se existir
                linha_avaliador = df[df.iloc[:, 0] == avaliador]
                if not linha_avaliador.empty:
                    regiao_avaliador = str(linha_avaliador.iloc[0, 12]) if pd.notna(linha_avaliador.iloc[0, 12]) else ""
            elif avaliador == 'VINICIUS OLIVEIRA AMARAL DE SOUZA':
                cargo_avaliador = 'SUPERVISOR'
                # Buscar região na planilha se existir
                linha_avaliador = df[df.iloc[:, 0] == avaliador]
                if not linha_avaliador.empty:
                    regiao_avaliador = str(linha_avaliador.iloc[0, 12]) if pd.notna(linha_avaliador.iloc[0, 12]) else ""
            else:
                # Buscar na planilha
                linha_avaliador = df[df.iloc[:, 0] == avaliador]
                if not linha_avaliador.empty:
                    cargo_avaliador = str(linha_avaliador.iloc[0, 8]) if pd.notna(linha_avaliador.iloc[0, 8]) else ""
                    regiao_avaliador = str(linha_avaliador.iloc[0, 12]) if pd.notna(linha_avaliador.iloc[0, 12]) else ""
        
        st.text_input("Cargo do Avaliador", value=cargo_avaliador, disabled=True, key=f"cargo_avaliador_{avaliador}")
        st.text_input("Região do Avaliador", value=regiao_avaliador, disabled=True, key=f"regiao_avaliador_{avaliador}")

    with col2:
        colaborador = st.selectbox("Nome do colaborador *", todos_colaboradores)
        # Buscar cargo e região do colaborador selecionado automaticamente
        cargo_colaborador = ""
        regiao_colaborador = ""
        if colaborador:
            linha_colaborador = df[df.iloc[:, 0] == colaborador]
            if not linha_colaborador.empty:
                cargo_colaborador = str(linha_colaborador.iloc[0, 8]) if pd.notna(linha_colaborador.iloc[0, 8]) else ""
                regiao_colaborador = str(linha_colaborador.iloc[0, 12]) if pd.notna(linha_colaborador.iloc[0, 12]) else ""
        
        st.text_input("Cargo do Colaborador *", value=cargo_colaborador, disabled=True, key=f"cargo_colaborador_{colaborador}")
        st.text_input("Região do Colaborador", value=regiao_colaborador, disabled=True, key=f"regiao_colaborador_{colaborador}")

    tipo_avaliacao = st.radio("Avaliação de:", ["40 dias", "80 dias"])

    with st.form("formulario_avaliacao"):
        cargo = cargo_colaborador  # Usar o cargo já identificado

        st.markdown("---")
        st.subheader("Critérios de Avaliação")

        # Adaptação ao Trabalho
        st.markdown("**ADAPTAÇÃO AO TRABALHO**")
        adaptacao = st.radio(
            "Selecione uma opção:",
            [
                "Está plenamente identificado com as atividades do seu cargo, e integrou-se perfeitamente às normas da empresa.",
                "Tem feito o possível para integrar-se não só ao próprio trabalho, como também às características da empresa.",
                "Precisa modificar radicalmente suas características pessoais para conseguir integrar-se ao trabalho e aos requisitos administrativos da empresa.",
                "Mantém um comportamento oposto ao solicitado para o seu cargo e demonstra ter sérias dificuldades de aceitação das características da empresa."
            ],
            key="adaptacao"
        )

        # Interesse
        st.markdown("**INTERESSE**")
        interesse = st.radio(
            "Selecione uma opção:",
            [
                "Apresenta um entusiasmo adequado, tendo em vista o seu pouco tempo de casa.",
                "Parece muito interessado(a) por seu novo emprego.",
                "Passa a impressão de ser um colaborador(a) que no futuro necessitará de constante estímulo para poder interessar-se por seu trabalho.",
                "É indiferente, apresentando uma falta total de entusiasmo e vontade de trabalhar."
            ],
            key="interesse"
        )

        # Relacionamento Social
        st.markdown("**RELACIONAMENTO SOCIAL**")
        relacionamento = st.radio(
            "Selecione uma opção:",
            [
                "Apresentou grande habilidade em conseguir amigos, mesmo com pouco tempo de casa, todos já gostam muito dele(a).",
                "Entrosou-se bem com os demais, foi aceito(a) sem resistência.",
                "Está fazendo muita força para conseguir maior integração social com os colegas.",
                "Sente-se perdido(a) entre os colegas, parece não ter sido aceito(a) pelo grupo de trabalho."
            ],
            key="relacionamento"
        )

        # Capacidade de Aprendizagem
        st.markdown("**CAPACIDADE DE APRENDIZAGEM**")
        capacidade = st.radio(
            "Selecione uma opção:",
            [
                "Parece habilitado(a) para o cargo em que está, tem facilidade para aprender, permitindo-lhe executar sem falhas.",
                "Parece adequado(a) para o cargo ao qual foi encaminhado(a), aprende suas tarefas sem problemas.",
                "Consegue aprender o que lhe foi ensinado à custa de grande esforço pessoal, necessário repetir-se a mesma coisa várias vezes.",
                "Parece não ter a mínima capacidade para o trabalho."
            ],
            key="capacidade"
        )

        # Classificação Geral
        st.markdown("**De maneira geral como o colaborador (a) pode ser classificado?**")
        classificacao = st.radio(
            "Selecione uma opção:",
            [
                "Trata-se de excelente aquisição para a empresa",
                "Constitui Elemento com boas possibilidades futuras",
                "Tem possibilidades Rotineiras",
                "Fraco"
            ],
            key="classificacao"
        )

        # Definição
        st.markdown("**Qual a definição a ser tomada?**")
        definicao = st.radio(
            "Selecione uma opção:",
            [
                "Prorrogar o contrato de trabalho",
                "Encaminhá-lo para treinamento",
                "Demitir"
            ],
            key="definicao"
        )

        st.markdown("---")
        submitted = st.form_submit_button("💾 Salvar Avaliação e Gerar PDF", use_container_width=True)

    # Processar fora do formulário
    if submitted:
        if not cargo:
            st.error("⚠️ Por favor, selecione um colaborador válido!")
        else:
            # Salvar no banco
            dados = (
                avaliador, colaborador, cargo, cargo_avaliador, "", tipo_avaliacao,
                adaptacao, interesse, relacionamento, capacidade,
                classificacao, definicao, regiao_avaliador, regiao_colaborador
            )
            salvar_avaliacao(dados)

            # Gerar PDF
            dados_pdf = {
                'avaliador': avaliador,
                'cargo_avaliador': cargo_avaliador,
                'regiao_avaliador': regiao_avaliador,
                'colaborador': colaborador,
                'cargo': cargo,
                'regiao_colaborador': regiao_colaborador,
                'tipo_avaliacao': tipo_avaliacao,
                'adaptacao': adaptacao,
                'interesse': interesse,
                'relacionamento': relacionamento,
                'capacidade': capacidade,
                'classificacao': classificacao,
                'definicao': definicao
            }

            try:
                pdf_buffer, pdf_nome = gerar_pdf_avaliacao(dados_pdf)

                st.success(f"✅ Avaliação de {colaborador} salva com sucesso!")
                st.balloons()

                # Botão de download do PDF
                st.download_button(
                    label="📄 Download PDF da Avaliação",
                    data=pdf_buffer,
                    file_name=pdf_nome,
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Erro ao gerar PDF: {e}")
                st.info("A avaliação foi salva, mas o PDF não pôde ser gerado.")

# HISTÓRICO DE AVALIAÇÕES
elif menu == "Histórico de Avaliações":
    st.header("📚 Histórico de Avaliações")

    avaliacoes_df = buscar_avaliacoes()

    if len(avaliacoes_df) > 0:
        st.markdown(f"**Total de avaliações registradas:** {len(avaliacoes_df)}")

        # Filtros
        col1, col2, col3 = st.columns(3)

        with col1:
            filtro_avaliador = st.multiselect(
                "Filtrar por Avaliador",
                options=avaliacoes_df['avaliador'].unique()
            )

        with col2:
            filtro_tipo = st.multiselect(
                "Filtrar por Tipo",
                options=avaliacoes_df['tipo_avaliacao'].unique()
            )

        with col3:
            filtro_definicao = st.multiselect(
                "Filtrar por Definição",
                options=avaliacoes_df['definicao'].unique()
            )

        # Aplicar filtros
        df_filtrado = avaliacoes_df.copy()

        if filtro_avaliador:
            df_filtrado = df_filtrado[df_filtrado['avaliador'].isin(filtro_avaliador)]

        if filtro_tipo:
            df_filtrado = df_filtrado[df_filtrado['tipo_avaliacao'].isin(filtro_tipo)]

        if filtro_definicao:
            df_filtrado = df_filtrado[df_filtrado['definicao'].isin(filtro_definicao)]

        st.markdown("---")

        # Mostrar detalhes das avaliações
        for idx, row in df_filtrado.iterrows():
            with st.expander(f"📋 {row['colaborador']} - {row['tipo_avaliacao']} (Avaliado por: {row['avaliador']})"):
                col1, col2 = st.columns(2)

                with col1:
                    st.write(f"**Cargo:** {row['cargo']}")
                    st.write(f"**Região Colaborador:** {row.get('regiao_colaborador', 'N/A')}")
                    st.write(f"**Data:** {row['data_avaliacao']}")
                    st.write(f"**Classificação:** {row['classificacao']}")

                with col2:
                    st.write(f"**Avaliador:** {row['avaliador']}")
                    st.write(f"**Região Avaliador:** {row.get('regiao_avaliador', 'N/A')}")
                    st.write(f"**Definição:** {row['definicao']}")
                    st.write(f"**Adaptação:** {row['adaptacao'][:50]}...")

                # Botão para gerar PDF da avaliação histórica
                if st.button(f"📄 Gerar PDF", key=f"pdf_{idx}"):
                    dados_pdf = {
                        'avaliador': row['avaliador'],
                        'cargo_avaliador': row.get('cargo_avaliador', ''),
                        'regiao_avaliador': row.get('regiao_avaliador', ''),
                        'colaborador': row['colaborador'],
                        'cargo': row['cargo'],
                        'regiao_colaborador': row.get('regiao_colaborador', ''),
                        'tipo_avaliacao': row['tipo_avaliacao'],
                        'adaptacao': row['adaptacao'],
                        'interesse': row['interesse'],
                        'relacionamento': row['relacionamento'],
                        'capacidade': row['capacidade'],
                        'classificacao': row['classificacao'],
                        'definicao': row['definicao']
                    }

                    try:
                        pdf_buffer, pdf_nome = gerar_pdf_avaliacao(dados_pdf)

                        st.download_button(
                            label="⬇️ Download PDF",
                            data=pdf_buffer,
                            file_name=pdf_nome,
                            mime="application/pdf",
                            key=f"download_pdf_{idx}"
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar PDF: {e}")

        st.markdown("---")

        # Baixar histórico em Excel
        if st.button("📥 Baixar Histórico (Excel)", use_container_width=True):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Avaliações')

            st.download_button(
                label="⬇️ Download",
                data=output.getvalue(),
                file_name=f"historico_avaliacoes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Nenhuma avaliação registrada ainda.")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>Sistema de Avaliação de Experiência - Rezende Energia © 2025</div>",
    unsafe_allow_html=True
)







