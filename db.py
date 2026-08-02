import os
import sqlite3
import pandas as pd
import datetime
import io

# Timezone Brasil (UTC-3)
TZ_BRASIL = datetime.timezone(datetime.timedelta(hours=-3))

# Tentar importar cliente Supabase
try:
    from supabase import create_client, Client
    HAS_SUPABASE_LIB = True
except ImportError:
    HAS_SUPABASE_LIB = False

DB_FILE = "validacao_equipamentos.db"


def _get_secrets():
    """
    Obtém credenciais do Supabase a partir de st.secrets (Cloud) ou variáveis de ambiente (local).
    Retorna (url, key) ou (None, None).
    """
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_KEY")

    if not url or not key:
        try:
            import streamlit as st
            if hasattr(st, 'secrets'):
                url = st.secrets.get("SUPABASE_URL", url)
                key = st.secrets.get("SUPABASE_KEY", key)
        except Exception:
            pass

    return url, key


def get_supabase_client():
    """
    Retorna o cliente Supabase cacheado via st.cache_resource.
    Se o Streamlit não estiver disponível (ex: testes), cria diretamente.
    """
    try:
        import streamlit as st

        @st.cache_resource
        def _create_client():
            url, key = _get_secrets()
            if url and key and HAS_SUPABASE_LIB:
                try:
                    return create_client(url, key)
                except Exception as e:
                    st.error(f"❌ Erro ao conectar ao Supabase: {e}")
                    return None
            return None

        return _create_client()
    except Exception:
        # Fallback sem Streamlit (testes, scripts)
        url, key = _get_secrets()
        if url and key and HAS_SUPABASE_LIB:
            try:
                return create_client(url, key)
            except Exception:
                return None
        return None


def is_using_supabase():
    """
    Retorna True se estiver utilizando o banco de dados Supabase na nuvem.
    """
    return get_supabase_client() is not None


def _now_brasil():
    """
    Retorna datetime no fuso horário do Brasil (UTC-3).
    """
    return datetime.datetime.now(TZ_BRASIL)


# --- BANCO DE DADOS SQLITE LOCAL (fallback para desenvolvimento) ---

def get_sqlite_conn(db_path=DB_FILE):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path=DB_FILE):
    """
    Inicializa o banco de dados SQLite local. Ignorado se Supabase estiver ativo.
    """
    if is_using_supabase():
        return

    conn = get_sqlite_conn(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS validacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            data_validacao TEXT,
            hora_validacao TEXT,
            operador TEXT,
            serial TEXT NOT NULL,
            tecnico TEXT,
            tecnico_vicky TEXT,
            codigo_produto TEXT,
            sku TEXT,
            descricao TEXT,
            classificacao_sgm TEXT,
            estado_vicky TEXT,
            tipo_material TEXT,
            status TEXT NOT NULL,
            categoria TEXT NOT NULL,
            motivo TEXT,
            alerta_classificacao TEXT,
            origem_planilha TEXT
        );
    """)

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_serial ON validacoes(serial);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_status ON validacoes(status);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_timestamp ON validacoes(timestamp);")

    conn.commit()
    conn.close()


# --- OPERAÇÕES DE GRAVAÇÃO ---

def save_validation(result, operador="Almoxarife", origem_planilha="dados.xlsx", db_path=DB_FILE):
    """
    Insere o registro de validação no Supabase (nuvem) ou SQLite (local).
    Retorna (id_inserido, destino_str). Em caso de erro no Supabase, exibe erro visível.
    """
    now = _now_brasil()
    data_str = result.get('data', now.strftime("%d/%m/%Y"))
    hora_str = result.get('hora', now.strftime("%H:%M:%S"))

    # Campos na ordem exata das colunas do banco
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
    serial = result.get('serial', '')
    tecnico = result.get('tecnico', result.get('tecnico_sgm', ''))
    tecnico_vicky = result.get('tecnico_vicky', '')
    codigo_produto = result.get('codigo_produto', '')
    sku = result.get('sku', '')
    descricao = result.get('descricao', '')
    classificacao_sgm = result.get('classificacao', '')
    estado_vicky = result.get('estado', '')
    tipo_material = result.get('tipo_material', '')
    status = result.get('status', 'BLOQUEADO')
    categoria = result.get('categoria', 'Desconhecido')
    motivo = result.get('motivo', '')
    alerta_classificacao = result.get('alerta_classificacao', '')

    record = {
        "timestamp": timestamp_str,
        "data_validacao": data_str,
        "hora_validacao": hora_str,
        "operador": operador,
        "serial": serial,
        "tecnico": tecnico,
        "tecnico_vicky": tecnico_vicky,
        "codigo_produto": codigo_produto,
        "sku": sku,
        "descricao": descricao,
        "classificacao_sgm": classificacao_sgm,
        "estado_vicky": estado_vicky,
        "tipo_material": tipo_material,
        "status": status,
        "categoria": categoria,
        "motivo": motivo,
        "alerta_classificacao": alerta_classificacao,
        "origem_planilha": origem_planilha
    }

    client = get_supabase_client()
    if client:
        try:
            res = client.table("validacoes").insert(record).execute()
            if res.data and len(res.data) > 0:
                return res.data[0].get('id', 'Nuvem'), 'Supabase Nuvem'
            return 'Nuvem', 'Supabase Nuvem'
        except Exception as e:
            # Erro visível ao operador — NÃO falhar silenciosamente
            try:
                import streamlit as st
                st.error(f"⚠️ Falha ao gravar no Supabase: {e}. Gravando localmente como backup.")
            except Exception:
                pass

    # Fallback para SQLite local (desenvolvimento ou erro temporário)
    init_db(db_path)
    conn = get_sqlite_conn(db_path)
    cursor = conn.cursor()

    # Tupla explícita na ordem exata das colunas SQL
    cursor.execute("""
        INSERT INTO validacoes (
            timestamp, data_validacao, hora_validacao, operador, serial,
            tecnico, tecnico_vicky, codigo_produto, sku, descricao,
            classificacao_sgm, estado_vicky, tipo_material, status,
            categoria, motivo, alerta_classificacao, origem_planilha
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        timestamp_str, data_str, hora_str, operador, serial,
        tecnico, tecnico_vicky, codigo_produto, sku, descricao,
        classificacao_sgm, estado_vicky, tipo_material, status,
        categoria, motivo, alerta_classificacao, origem_planilha
    ))

    inserted_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return inserted_id, 'SQLite Local'


# --- CONSULTA DE HISTÓRICO ---

def load_validations(limit=1000, operador_filter=None, status_filter=None, db_path=DB_FILE):
    """
    Carrega histórico do Supabase Cloud ou do SQLite local em formato DataFrame.
    """
    client = get_supabase_client()
    if client:
        try:
            query = client.table("validacoes").select("*").order("id", desc=True).limit(limit)
            if operador_filter and operador_filter != "Todos":
                query = query.eq("operador", operador_filter)
            if status_filter and status_filter != "Todos":
                query = query.eq("status", status_filter)

            res = query.execute()
            if res.data:
                return pd.DataFrame(res.data)
            return pd.DataFrame()
        except Exception as e:
            try:
                import streamlit as st
                st.warning(f"⚠️ Erro ao consultar Supabase: {e}. Exibindo dados locais.")
            except Exception:
                pass

    # Fallback para SQLite
    init_db(db_path)
    conn = get_sqlite_conn(db_path)
    query = "SELECT * FROM validacoes WHERE 1=1"
    params = []

    if operador_filter and operador_filter != "Todos":
        query += " AND operador = ?"
        params.append(operador_filter)
    if status_filter and status_filter != "Todos":
        query += " AND status = ?"
        params.append(status_filter)

    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


# --- MÉTRICAS DO BANCO (eficientes) ---

def get_db_metrics(db_path=DB_FILE):
    """
    Retorna métricas consolidadas usando COUNT em vez de carregar todas as linhas.
    """
    client = get_supabase_client()
    if client:
        try:
            # Usar select com count para evitar transferir dados inteiros
            total_res = client.table("validacoes").select("id", count="exact").execute()
            total = total_res.count if total_res.count is not None else 0

            aprovados_res = client.table("validacoes").select("id", count="exact").eq("status", "LIBERADO").execute()
            aprovados = aprovados_res.count if aprovados_res.count is not None else 0

            bloqueados_res = client.table("validacoes").select("id", count="exact").eq("status", "BLOQUEADO").execute()
            bloqueados = bloqueados_res.count if bloqueados_res.count is not None else 0

            return {
                'total': total,
                'aprovados': aprovados,
                'bloqueados': bloqueados,
                'is_cloud': True
            }
        except Exception:
            pass

    # Fallback SQLite — usar COUNT() nativo
    init_db(db_path)
    conn = get_sqlite_conn(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM validacoes")
    total = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM validacoes WHERE status = 'LIBERADO'")
    aprovados = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM validacoes WHERE status = 'BLOQUEADO'")
    bloqueados = cursor.fetchone()[0]

    conn.close()
    return {
        'total': total,
        'aprovados': aprovados,
        'bloqueados': bloqueados,
        'is_cloud': False
    }


# --- EXPORTAÇÃO ---

def export_db_to_excel(limit=50000, db_path=DB_FILE):
    """
    Gera um arquivo Excel a partir dos dados do banco (Supabase/SQLite).
    Limite elevado para exportações completas.
    """
    df = load_validations(limit=limit, db_path=db_path)
    output = io.BytesIO()

    if not df.empty:
        df_export = df.rename(columns={
            'id': 'ID',
            'timestamp': 'Data/Hora Registro',
            'data_validacao': 'Data',
            'hora_validacao': 'Hora',
            'operador': 'Operador',
            'serial': 'Número de Série',
            'tecnico': 'Técnico SGM',
            'tecnico_vicky': 'Técnico Vicky',
            'codigo_produto': 'Código Produto SGM',
            'sku': 'SKU Vicky',
            'descricao': 'Descrição',
            'classificacao_sgm': 'Classificação SGM',
            'estado_vicky': 'Estado Vicky',
            'tipo_material': 'Tipo de Material',
            'status': 'Status (Situação)',
            'categoria': 'Categoria Divergência',
            'motivo': 'Motivo do Bloqueio/Observação',
            'alerta_classificacao': 'Alerta de Classificação',
            'origem_planilha': 'Planilha Origem'
        })
    else:
        df_export = df

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Histórico Validações')

    output.seek(0)
    return output.getvalue()
