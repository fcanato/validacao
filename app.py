import streamlit as st
import pandas as pd
import datetime
import os
from validator import load_excel_data, validate_equipment_serial
from export import generate_excel_report
import db

# Inicializar o banco de dados (Supabase ou SQLite local)
db.init_db()

@st.cache_resource
def get_shared_data_store():
    """Planilha carregada compartilhada entre todas as sessões/usuários."""
    return {"data": None, "file_name": None}

shared_store = get_shared_data_store()

# Configuração da página Streamlit
st.set_page_config(
    page_title="Sistema de Validação de Equipamentos - SGM x Vicky",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada (Modern & Premium Visuals)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        margin-bottom: 1.5rem;
    }
    
    .main-header h1 {
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
        color: #f8fafc;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    
    .main-header p {
        margin: 4px 0 0 0;
        color: #94a3b8;
        font-size: 0.95rem;
    }
    
    .badge-cloud {
        background-color: #10b981;
        color: #ffffff;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    
    .badge-local {
        background-color: #64748b;
        color: #ffffff;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    
    /* Cartões de Métricas / KPIs */
    .kpi-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 1rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
    }
    
    /* Status Badges */
    .badge-approved {
        background-color: #dcfce7;
        color: #15803d;
        border: 1px solid #86efac;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
        font-weight: 700;
        font-size: 1.2rem;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    
    .badge-blocked {
        background-color: #fee2e2;
        color: #b91c1c;
        border: 1px solid #fca5a5;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
        font-weight: 700;
        font-size: 1.2rem;
        display: inline-flex;
        align-items: center;
        gap: 8px;
    }
    
    .badge-warning {
        background-color: #fef3c7;
        color: #b45309;
        border: 1px solid #fde047;
        padding: 0.5rem 1rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.9rem;
        margin-top: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Inicialização do estado da sessão (Session State)
if 'history' not in st.session_state:
    st.session_state.history = []
if 'last_result' not in st.session_state:
    st.session_state.last_result = None

# Obter métricas consolidadas do banco de dados (Supabase ou SQLite)
db_metrics = db.get_db_metrics()
is_cloud = db_metrics.get('is_cloud', False)

db_status_badge = '<span class="badge-cloud">🌐 Supabase Nuvem Conectado</span>' if is_cloud else '<span class="badge-local">💾 SQLite Local (Demonstração)</span>'

# Título do aplicativo
st.markdown(f"""
<div class="main-header">
    <h1>🛡️ Validação de Equipamentos (SGM x Vicky)</h1>
    <p>Validação em tempo real com gravação centralizada no banco de dados para almoxarifado em campo &nbsp; {db_status_badge}</p>
</div>
""", unsafe_allow_html=True)

# --- PAINEL LATERAL (Configurações e Entrada de Dados) ---
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    
    # Operador
    operador = st.text_input("Operador do Almoxarifado:", value="Almoxarife", key="operador_input")
    data_validacao = datetime.date.today().strftime("%d/%m/%Y")
    st.info(f"📅 **Data da Validação:** {data_validacao}")
    
    st.markdown("---")
    st.subheader("🗄️ Status do Banco de Dados")
    if is_cloud:
        st.success("🌐 **Conectado ao Supabase Cloud**")
    else:
        st.info("💾 **Modo SQLite Local** (Para conectar à nuvem, configure as chaves SUPABASE_URL e SUPABASE_KEY).")
        
    st.markdown(f"💾 **Total no Banco:** {db_metrics['total']} registros")
    st.markdown(f"✅ **Aprovados no Banco:** {db_metrics['aprovados']}")
    st.markdown(f"🚫 **Bloqueados no Banco:** {db_metrics['bloqueados']}")
    
    st.markdown("---")
    st.subheader("📁 Seleção de Planilha Excel")
    
    default_file_exists = os.path.exists("dados.xlsx")
    source_choice = st.radio(
        "Origem do Arquivo:",
        ["Usar 'dados.xlsx' local", "Fazer Upload de Planilha"],
        index=0 if default_file_exists else 1
    )
    
    excel_source = None
    origem_nome = "dados.xlsx"
    if source_choice == "Usar 'dados.xlsx' local":
        if default_file_exists:
            excel_source = "dados.xlsx"
            origem_nome = "dados.xlsx"
            st.success("Planilha 'dados.xlsx' localizada!")
        else:
            st.error("Arquivo 'dados.xlsx' não encontrado na pasta raiz.")
    else:
        uploaded_file = st.file_uploader("Escolha a planilha (.xlsx)", type=["xlsx"])
        if uploaded_file is not None:
            excel_source = uploaded_file
            origem_nome = uploaded_file.name

    if excel_source is not None:
        if st.session_state.get('file_path_loaded') != str(excel_source) or shared_store["data"] is None:
            with st.spinner("Carregando bases SGM e Vicky..."):
                res = load_excel_data(excel_source)
                if res.get('loaded'):
                    shared_store["data"] = res
                    shared_store["file_name"] = origem_nome
                    st.session_state['file_path_loaded'] = str(excel_source)
                    st.toast("Planilha carregada com sucesso!", icon="✅")
                else:
                    st.error(f"Erro ao carregar planilha: {res.get('error')}")

    if shared_store["data"] and shared_store["data"].get('loaded'):
        df_sgm = shared_store["data"]['df_sgm']
        df_vicky = shared_store["data"]['df_vicky']
        file_loaded_name = shared_store["file_name"] or "planilha"
        st.success(f"✅ Planilha ativa: **{file_loaded_name}**")
        st.markdown(f"📊 **Base SGM:** {len(df_sgm)} registros")
        st.markdown(f"📊 **Base Vicky:** {len(df_vicky)} registros")
    
    st.markdown("---")
    if st.button("🗑️ Limpar Histórico da Sessão"):
        st.session_state.history = []
        st.session_state.last_result = None
        st.rerun()

# Verificar se a planilha foi carregada
if not shared_store["data"] or not shared_store["data"].get('loaded'):
    st.warning("⚠️ Nenhuma planilha válida foi carregada. Utilize o painel lateral para selecionar a planilha Excel contendo as abas **SGM** e **Vicky**.")
    st.stop()

# --- INDICADORES / KPIS EM TEMPO REAL DA SESSÃO ---
history = st.session_state.history

total_lidos = len(history)
total_aprovados = sum(1 for h in history if h.get('status') == 'LIBERADO')
total_bloqueados = sum(1 for h in history if h.get('status') == 'BLOQUEADO')

div_sgm = sum(1 for h in history if h.get('categoria') == 'Divergência SGM')
div_vicky = sum(1 for h in history if h.get('categoria') == 'Divergência Vicky')
div_produto = sum(1 for h in history if h.get('categoria') == 'Divergência de Produto')
div_tecnico = sum(1 for h in history if h.get('categoria') == 'Divergência de Técnico')

col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5, col_kpi6, col_kpi7 = st.columns(7)

with col_kpi1:
    st.metric("Sessão: Lidos", total_lidos)
with col_kpi2:
    st.metric("Sessão: Aprovados", total_aprovados)
with col_kpi3:
    st.metric("Sessão: Bloqueados", total_bloqueados)
with col_kpi4:
    st.metric("Div. SGM", div_sgm)
with col_kpi5:
    st.metric("Div. Vicky", div_vicky)
with col_kpi6:
    st.metric("Div. Produto", div_produto)
with col_kpi7:
    st.metric("Div. Técnico", div_tecnico)

st.markdown("<br>", unsafe_allow_html=True)

# --- PAINEL DE LEITURA E RESULTADO ---
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("🔍 Painel de Leitura de Equipamento")
    st.caption("Leitor de código de barras, QR code ou digitação. Pressione ENTER para validar e gravar na Nuvem/Banco.")
    
    with st.form(key="barcode_form", clear_on_submit=True):
        serial_input = st.text_input(
            "Número de Série (Serial):",
            placeholder="Aguardando leitura do leitor ou digitação...",
            key="serial_field"
        )
        submit_clicked = st.form_submit_button(" Validar e Gravar no Banco (ENTER)", use_container_width=True)
        
        if submit_clicked and serial_input:
            now = datetime.datetime.now()
            result = validate_equipment_serial(serial_input, shared_store["data"])
            result['hora'] = now.strftime("%H:%M:%S")
            result['data'] = now.strftime("%d/%m/%Y")
            result['operador'] = operador
            
            # Persistir no Banco de Dados (Supabase Nuvem / SQLite Local)
            row_id = db.save_validation(result, operador=operador, origem_planilha=origem_nome)
            result['db_id'] = row_id
            
            st.session_state.history.insert(0, result)
            st.session_state.last_result = result
            st.rerun()

with col_right:
    st.subheader("📋 Resultado da Validação")
    
    last = st.session_state.last_result
    if last is None:
        st.info("Aguardando a leitura do equipamento para validar e salvar no banco de dados.")
    else:
        is_liberado = last.get('status') == 'LIBERADO'
        
        if is_liberado:
            st.markdown(f"""
            <div class="badge-approved">
                ✔ LIBERADO PARA RECOLHIMENTO
            </div>
            """, unsafe_allow_html=True)
            st.success(f"**Motivo:** {last.get('motivo')}")
        else:
            st.markdown(f"""
            <div class="badge-blocked">
                ✖ RECOLHIMENTO BLOQUEADO
            </div>
            """, unsafe_allow_html=True)
            st.error(f"**Motivo:** {last.get('motivo')}")
            
        if last.get('alerta_classificacao'):
            st.markdown(f"""
            <div class="badge-warning">
                ⚠️ {last.get('alerta_classificacao')}
            </div>
            """, unsafe_allow_html=True)
            
        dest_str = "Supabase Nuvem" if is_cloud else "SQLite Local"
        st.caption(f"💾 **Gravado em [{dest_str}] com ID:** #{last.get('db_id', 'N/A')}")
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Grid de detalhes do equipamento lido
        res_col1, res_col2 = st.columns(2)
        with res_col1:
            st.markdown(f"**Número de Série:** `{last.get('serial', 'N/A')}`")
            st.markdown(f"**Técnico:** `{last.get('tecnico', last.get('tecnico_sgm', 'N/A'))}`")
            if last.get('tecnico_vicky') and last.get('tecnico_vicky') != last.get('tecnico_sgm'):
                st.markdown(f"**Técnico (Vicky):** `{last.get('tecnico_vicky')}`")
            st.markdown(f"**Código Produto:** `{last.get('codigo_produto', 'N/A')}`")
            st.markdown(f"**SKU Vicky:** `{last.get('sku', 'N/A')}`")
            
        with res_col2:
            st.markdown(f"**Descrição:** `{last.get('descricao', 'N/A')}`")
            st.markdown(f"**Classificação (SGM):** `{last.get('classificacao', 'N/A')}`")
            st.markdown(f"**Estado (Vicky):** `{last.get('estado', 'N/A')}`")
            st.markdown(f"**Tipo do Material:** `{last.get('tipo_material', 'N/A')}`")

st.markdown("<hr>", unsafe_allow_html=True)

# --- ABAS DE HISTÓRICO: SESSÃO ATUAL vs BANCO DE DADOS (SUPABASE / SQLITE) ---
tab_sessao, tab_banco = st.tabs(["📜 Histórico da Sessão Atual", "🗄️ Histórico Geral do Banco de Dados"])

with tab_sessao:
    hist_header_col1, hist_header_col2 = st.columns([3, 1])
    with hist_header_col1:
        st.subheader("Equipamentos Lidos nesta Sessão")
    with hist_header_col2:
        if len(history) > 0:
            excel_bytes = generate_excel_report(history)
            st.download_button(
                label="📥 Exportar Sessão em Excel",
                data=excel_bytes,
                file_name=f"relatorio_sessao_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    if len(history) == 0:
        st.caption("Nenhum equipamento foi lido nesta sessão até o momento.")
    else:
        df_hist = pd.DataFrame([
            {
                'ID DB': str(h.get('db_id', '') or ''),
                'Hora': str(h.get('hora', '') or ''),
                'Serial': str(h.get('serial', '') or ''),
                'Técnico': str(h.get('tecnico', h.get('tecnico_sgm', '')) or ''),
                'Produto': str(h.get('codigo_produto', h.get('sku', '')) or ''),
                'Resultado': str(h.get('status', '') or ''),
                'Categoria': str(h.get('categoria', '') or ''),
                'Observação / Motivo': str(h.get('motivo', '') or '')
            }
            for h in history
        ])
        
        def highlight_status(val):
            color = '#dcfce7' if val == 'LIBERADO' else '#fee2e2'
            text_color = '#15803d' if val == 'LIBERADO' else '#b91c1c'
            return f'background-color: {color}; color: {text_color}; font-weight: bold;'

        st.dataframe(
            df_hist.style.map(highlight_status, subset=['Resultado']),
            use_container_width=True,
            height=350
        )

with tab_banco:
    banco_tipo_str = "Supabase Nuvem (PostgreSQL)" if is_cloud else "SQLite Local"
    st.subheader(f"Registros Persistidos no Banco: {banco_tipo_str}")
    st.caption("Histórico centralizado de recolhimentos. Permite rastreabilidade em tempo real de múltiplos almoxarifes em campo.")
    
    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1])
    
    df_db_all = db.load_validations(limit=5000)
    operadores_unicos = ["Todos"] + sorted(list(df_db_all['operador'].dropna().unique())) if not df_db_all.empty and 'operador' in df_db_all.columns else ["Todos"]
    
    with filter_col1:
        op_filtro = st.selectbox("Filtrar por Operador:", operadores_unicos)
    with filter_col2:
        st_filtro = st.selectbox("Filtrar por Status:", ["Todos", "LIBERADO", "BLOQUEADO"])
    with filter_col3:
        search_serial = st.text_input("Buscar por Número de Série (Serial):", "")
        
    df_filtered = db.load_validations(limit=2000, operador_filter=op_filtro, status_filter=st_filtro)
    
    if search_serial and not df_filtered.empty and 'serial' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['serial'].astype(str).str.contains(search_serial.strip().upper())]
        
    db_export_col1, db_export_col2 = st.columns([3, 1])
    with db_export_col1:
        st.markdown(f"Exibindo **{len(df_filtered)}** registros do banco de dados [{banco_tipo_str}].")
    with db_export_col2:
        excel_db_bytes = db.export_db_to_excel()
        st.download_button(
            label="📥 Exportar Banco Completo em Excel",
            data=excel_db_bytes,
            file_name=f"relatorio_banco_validacoes_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    if df_filtered.empty:
        st.info("Nenhum registro encontrado no banco de dados para os filtros selecionados.")
    else:
        cols_to_show = [c for c in ['id', 'timestamp', 'data_validacao', 'hora_validacao', 'operador', 'serial', 'tecnico', 'codigo_produto', 'status', 'categoria', 'motivo', 'tipo_material'] if c in df_filtered.columns]
        st.dataframe(
            df_filtered[cols_to_show],
            use_container_width=True,
            height=400
        )
