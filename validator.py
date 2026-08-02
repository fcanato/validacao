import pandas as pd
import unicodedata

def normalize_str(val):
    """
    Remove acentos, espaços extras e converte para maiúsculo.
    """
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    return unicodedata.normalize('NFKD', val_str).encode('ASCII', 'ignore').decode('utf-8').upper().strip()

def find_column(df, possible_names):
    """
    Procura uma coluna no DataFrame priorizando correspondência exata antes de busca parcial.
    """
    # 1. Busca por correspondência exata
    for name in possible_names:
        norm_name = normalize_str(name)
        for col in df.columns:
            if normalize_str(col) == norm_name:
                return col
                
    # 2. Busca por substring
    for name in possible_names:
        norm_name = normalize_str(name)
        for col in df.columns:
            if norm_name in normalize_str(col):
                return col
    return None

def load_excel_data(file_source):
    """
    Carrega e normaliza as abas SGM e Vicky da planilha Excel.
    """
    try:
        excel_file = pd.ExcelFile(file_source)
        sheet_names = excel_file.sheet_names
        
        sgm_sheet = None
        vicky_sheet = None
        
        for sheet in sheet_names:
            norm_sheet = normalize_str(sheet)
            if 'SGM' in norm_sheet:
                sgm_sheet = sheet
            elif 'VICKY' in norm_sheet:
                vicky_sheet = sheet
                
        if sgm_sheet is None:
            sgm_sheet = sheet_names[0]
        if vicky_sheet is None:
            vicky_sheet = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
            
        df_sgm = pd.read_excel(excel_file, sheet_name=sgm_sheet)
        df_vicky = pd.read_excel(excel_file, sheet_name=vicky_sheet)
        
        # Mapeamento e limpeza de colunas SGM
        sgm_col_map = {
            'serial': find_column(df_sgm, ['COMPLEMENTAR', 'SERIAL', 'NUMERO_SERIE']),
            'tecnico': find_column(df_sgm, ['NOME_VOLANTE', 'VOLANTE', 'TECNICO']),
            'codigo_produto': find_column(df_sgm, ['CODIGO_PRODUTO', 'COD_PRODUTO', 'PRODUTO']),
            'descricao_produto': find_column(df_sgm, ['DESCRICAO_PRODUTO', 'DESC_PRODUTO']),
            'quantidade': find_column(df_sgm, ['QT_ESTOQUE', 'QUANTIDADE', 'QTD']),
            'classificacao': find_column(df_sgm, ['DESCRICAO_CLASSIFICACAO', 'CLASSIFICACAO'])
        }
        
        # Mapeamento e limpeza de colunas Vicky
        vicky_col_map = {
            'serial': find_column(df_vicky, ['NUMERO_DE_SERIE', 'NUMERO_SERIE', 'SERIAL']),
            'tecnico': find_column(df_vicky, ['NOME_DO_TECNICO', 'NOME_TECNICO', 'TECNICO']),
            'sku': find_column(df_vicky, ['SKU']),
            'descricao_sku': find_column(df_vicky, ['DESCRICAO_SKU', 'DESC_SKU']),
            'quantidade': find_column(df_vicky, ['QUANTIDADE', 'QTD']),
            'estado': find_column(df_vicky, ['DESCRICAO_ESTADO', 'ESTADO'])
        }
        
        # Criar séries normalizadas para pesquisa rápida
        sgm_serials = df_sgm[sgm_col_map['serial']].apply(normalize_str) if sgm_col_map['serial'] else pd.Series(dtype=str)
        vicky_serials = df_vicky[vicky_col_map['serial']].apply(normalize_str) if vicky_col_map['serial'] else pd.Series(dtype=str)
        
        sgm_counts = sgm_serials[sgm_serials != ''].value_counts()
        vicky_counts = vicky_serials[vicky_serials != ''].value_counts()
        
        return {
            'df_sgm': df_sgm,
            'df_vicky': df_vicky,
            'sgm_col_map': sgm_col_map,
            'vicky_col_map': vicky_col_map,
            'sgm_serials': sgm_serials,
            'vicky_serials': vicky_serials,
            'sgm_counts': sgm_counts,
            'vicky_counts': vicky_counts,
            'loaded': True
        }
    except Exception as e:
        return {
            'loaded': False,
            'error': str(e)
        }

def identify_material_type(sgm_class, vicky_state):
    """
    Identifica o tipo de material e verifica divergência de classificação entre SGM e Vicky.
    Categorias: Material Novo, Material Retirado, Recuperado, Outro.
    """
    norm_sgm = normalize_str(sgm_class)
    norm_vicky = normalize_str(vicky_state)
    
    # Determinar tipo SGM
    if 'NOVO' in norm_sgm:
        tipo_sgm = 'Material Novo'
    elif 'RETIRAD' in norm_sgm:
        tipo_sgm = 'Material Retirado'
    elif 'CLIENTE' in norm_sgm or 'RECUPERAD' in norm_sgm:
        tipo_sgm = 'Recuperado'
    else:
        tipo_sgm = 'Outro'
        
    # Determinar tipo Vicky
    if 'NOVO' in norm_vicky or 'DISPONIVEL' in norm_vicky:
        tipo_vicky = 'Material Novo'
    elif 'RETIRAD' in norm_vicky:
        tipo_vicky = 'Material Retirado'
    elif 'DEVOLUC' in norm_vicky or 'DEVOLVER' in norm_vicky or 'WO' in norm_vicky:
        tipo_vicky = 'Recuperado'
    else:
        tipo_vicky = 'Outro'
        
    # Tipo final unificado
    tipo_final = tipo_sgm if tipo_sgm != 'Outro' else tipo_vicky
    
    alerta_divergencia = None
    if tipo_sgm != 'Outro' and tipo_vicky != 'Outro' and tipo_sgm != tipo_vicky:
        alerta_divergencia = f"Divergência de classificação: SGM indica '{sgm_class}' ({tipo_sgm}) e Vicky indica '{vicky_state}' ({tipo_vicky})."
        
    return tipo_final, alerta_divergencia

def validate_equipment_serial(serial_input, data_store):
    """
    Valida um número de série aplicando as Regras de 1 a 7 e Regra de Prioridade SGM -> Vicky.
    """
    if not data_store.get('loaded', False):
        return {
            'status': 'BLOQUEADO',
            'motivo': 'Planilha Excel não carregada.',
            'categoria': 'Planilha Não Carregada',
            'serial': serial_input
        }
        
    serial_norm = normalize_str(serial_input)
    if not serial_norm:
        return {
            'status': 'BLOQUEADO',
            'motivo': 'Número de série em branco.',
            'categoria': 'Serial em Branco',
            'serial': ''
        }
        
    df_sgm = data_store['df_sgm']
    df_vicky = data_store['df_vicky']
    sgm_col_map = data_store['sgm_col_map']
    vicky_col_map = data_store['vicky_col_map']
    sgm_serials = data_store['sgm_serials']
    vicky_serials = data_store['vicky_serials']
    sgm_counts = data_store['sgm_counts']
    vicky_counts = data_store['vicky_counts']
    
    in_sgm = serial_norm in sgm_counts and sgm_counts[serial_norm] > 0
    in_vicky = serial_norm in vicky_counts and vicky_counts[serial_norm] > 0
    
    # Regra 1: Inexistente no SGM
    if not in_sgm and not in_vicky:
        return {
            'status': 'BLOQUEADO',
            'motivo': 'Serial não localizado no SGM. Recolhimento não permitido.',
            'categoria': 'Divergência SGM',
            'serial': serial_norm
        }
        
    # Regra 8 (Leitura manual apenas no Vicky):
    if not in_sgm and in_vicky:
        vicky_matches = df_vicky[vicky_serials == serial_norm]
        v_row = vicky_matches.iloc[0]
        v_tech = v_row[vicky_col_map['tecnico']] if vicky_col_map['tecnico'] else ''
        v_sku = v_row[vicky_col_map['sku']] if vicky_col_map['sku'] else ''
        return {
            'status': 'BLOQUEADO',
            'motivo': 'Serial localizado apenas no Vicky. Regularizar cadastro no SGM antes do recolhimento.',
            'categoria': 'Divergência SGM',
            'serial': serial_norm,
            'tecnico': str(v_tech),
            'sku': str(v_sku)
        }
        
    # Regra 3: Duplicidade no SGM
    if in_sgm and sgm_counts[serial_norm] > 1:
        return {
            'status': 'BLOQUEADO',
            'motivo': f'Serial duplicado na base SGM ({sgm_counts[serial_norm]} ocorrências).',
            'categoria': 'Serial Duplicado',
            'serial': serial_norm
        }
        
    # Regra 2: Inexistente no Vicky
    if not in_vicky:
        sgm_matches = df_sgm[sgm_serials == serial_norm]
        s_row = sgm_matches.iloc[0]
        s_tech = s_row[sgm_col_map['tecnico']] if sgm_col_map['tecnico'] else ''
        s_prod = s_row[sgm_col_map['codigo_produto']] if sgm_col_map['codigo_produto'] else ''
        return {
            'status': 'BLOQUEADO',
            'motivo': 'Serial localizado no SGM, porém inexistente no Vicky. Recolhimento bloqueado.',
            'categoria': 'Divergência Vicky',
            'serial': serial_norm,
            'tecnico': str(s_tech),
            'codigo_produto': str(s_prod)
        }
        
    # Regra 3: Duplicidade no Vicky
    if in_vicky and vicky_counts[serial_norm] > 1:
        return {
            'status': 'BLOQUEADO',
            'motivo': f'Serial duplicado na base Vicky ({vicky_counts[serial_norm]} ocorrências).',
            'categoria': 'Serial Duplicado',
            'serial': serial_norm
        }
        
    # Localizado em ambas as bases - Extração dos dados
    sgm_row = df_sgm[sgm_serials == serial_norm].iloc[0]
    vicky_row = df_vicky[vicky_serials == serial_norm].iloc[0]
    
    sgm_tech_orig = str(sgm_row[sgm_col_map['tecnico']]) if sgm_col_map['tecnico'] else ''
    vicky_tech_orig = str(vicky_row[vicky_col_map['tecnico']]) if vicky_col_map['tecnico'] else ''
    
    sgm_prod_orig = str(sgm_row[sgm_col_map['codigo_produto']]) if sgm_col_map['codigo_produto'] else ''
    vicky_sku_orig = str(vicky_row[vicky_col_map['sku']]) if vicky_col_map['sku'] else ''
    
    sgm_desc = str(sgm_row[sgm_col_map['descricao_produto']]) if sgm_col_map['descricao_produto'] else ''
    vicky_desc = str(vicky_row[vicky_col_map['descricao_sku']]) if vicky_col_map['descricao_sku'] else ''
    
    sgm_class = str(sgm_row[sgm_col_map['classificacao']]) if sgm_col_map['classificacao'] else ''
    vicky_state = str(vicky_row[vicky_col_map['estado']]) if vicky_col_map['estado'] else ''
    
    tipo_material, alerta_classificacao = identify_material_type(sgm_class, vicky_state)
    
    # Regra 4: Divergência entre Produto (SGM) e SKU (Vicky)
    if normalize_str(sgm_prod_orig) != normalize_str(vicky_sku_orig):
        return {
            'status': 'BLOQUEADO',
            'motivo': f'Divergência entre Código de Produto SGM ({sgm_prod_orig}) e SKU Vicky ({vicky_sku_orig}).',
            'categoria': 'Divergência de Produto',
            'serial': serial_norm,
            'tecnico': sgm_tech_orig,
            'codigo_produto': sgm_prod_orig,
            'sku': vicky_sku_orig,
            'descricao': sgm_desc or vicky_desc,
            'classificacao': sgm_class,
            'estado': vicky_state,
            'tipo_material': tipo_material,
            'alerta_classificacao': alerta_classificacao
        }
        
    # Regra 5: Divergência de Técnico responsável
    if normalize_str(sgm_tech_orig) != normalize_str(vicky_tech_orig):
        return {
            'status': 'BLOQUEADO',
            'motivo': f'Divergência de Técnico: SGM ({sgm_tech_orig}) vs Vicky ({vicky_tech_orig}).',
            'categoria': 'Divergência de Técnico',
            'serial': serial_norm,
            'tecnico_sgm': sgm_tech_orig,
            'tecnico_vicky': vicky_tech_orig,
            'codigo_produto': sgm_prod_orig,
            'sku': vicky_sku_orig,
            'descricao': sgm_desc or vicky_desc,
            'classificacao': sgm_class,
            'estado': vicky_state,
            'tipo_material': tipo_material,
            'alerta_classificacao': alerta_classificacao
        }
        
    # Tudo aprovado -> Liberação
    return {
        'status': 'LIBERADO',
        'motivo': 'Equipamento validado com sucesso e liberado para recolhimento.',
        'categoria': 'Aprovado',
        'serial': serial_norm,
        'tecnico': sgm_tech_orig,
        'codigo_produto': sgm_prod_orig,
        'sku': vicky_sku_orig,
        'descricao': sgm_desc or vicky_desc,
        'classificacao': sgm_class,
        'estado': vicky_state,
        'tipo_material': tipo_material,
        'alerta_classificacao': alerta_classificacao
    }
