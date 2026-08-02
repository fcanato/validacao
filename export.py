import pandas as pd
import io

def generate_excel_report(history_list):
    """
    Gera um relatório em Excel (.xlsx) contendo todas as validações realizadas.
    Retorna os bytes do arquivo Excel para download no Streamlit.
    """
    if not history_list:
        df = pd.DataFrame(columns=[
            'Serial', 'Técnico', 'Produto', 'Sistema SGM', 'Sistema Vicky',
            'Situação', 'Motivo do Bloqueio', 'Data', 'Hora', 'Operador'
        ])
    else:
        rows = []
        for item in history_list:
            rows.append({
                'Serial': item.get('serial', ''),
                'Técnico': item.get('tecnico', item.get('tecnico_sgm', '')),
                'Produto': item.get('codigo_produto', item.get('sku', '')),
                'Sistema SGM': 'Sim' if item.get('categoria') not in ['Divergência SGM', 'Planilha Não Carregada', 'Serial em Branco'] else 'Não',
                'Sistema Vicky': 'Sim' if item.get('categoria') not in ['Divergência Vicky', 'Planilha Não Carregada', 'Serial em Branco'] else 'Não',
                'Situação': item.get('status', ''),
                'Motivo do Bloqueio': item.get('motivo', '') if item.get('status') == 'BLOQUEADO' else 'Nenhum (Liberado)',
                'Data': item.get('data', ''),
                'Hora': item.get('hora', ''),
                'Operador': item.get('operador', '')
            })
        df = pd.DataFrame(rows)
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Validações')
    
    output.seek(0)
    return output.getvalue()
