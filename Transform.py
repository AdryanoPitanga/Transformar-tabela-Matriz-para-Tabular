import pandas as pd
import numpy as np
import os

# ============================================================================
# CONFIGURAÇÕES
# ============================================================================
INPUT_FILE_PATH = r"C:\Users\adryano.pitanga\Downloads\RAW DATA\Palmsnov11.xlsx"
OUTPUT_FILE_PATH = r"C:\Users\adryano.pitanga\Downloads\transformadas\tabela_transformada_final_POWERBI.xlsx"

print("=" * 80)
print("TRANSFORMAÇÃO DE DADOS PARA POWER BI")
print("=" * 80)
print(f"Entrada: {INPUT_FILE_PATH}")
print(f"Saída: {OUTPUT_FILE_PATH}")
print("=" * 80)

# ============================================================================
# 1. VERIFICAR ARQUIVO
# ============================================================================
if not os.path.exists(INPUT_FILE_PATH):
    print(f"❌ ERRO: Arquivo não encontrado: {INPUT_FILE_PATH}")
    exit()

print(f"✅ Arquivo encontrado")

# ============================================================================
# 2. FUNÇÃO PARA CONVERTER VALORES BRASILEIROS
# ============================================================================
def converter_valor_brasileiro(valor):
    """
    Converte valores no formato brasileiro para float.
    Exemplos:
    - "1.234,56" → 1234.56
    - "1.234" → 1234.0
    - "1234,56" → 1234.56
    """
    if pd.isna(valor):
        return 0.0
    
    if isinstance(valor, (int, float, np.number)):
        return float(valor)
    
    valor_str = str(valor).strip()
    
    # Limpar
    valor_str = valor_str.replace('R$', '').replace(' ', '').replace('\xa0', '')
    
    if valor_str == '' or valor_str == '-':
        return 0.0
    
    # Se tem vírgula (decimal brasileiro)
    if ',' in valor_str:
        # Se tem ponto (separador de milhar)
        if '.' in valor_str:
            # Formato: 1.234,56
            valor_str = valor_str.replace('.', '').replace(',', '.')
        else:
            # Formato: 1234,56
            valor_str = valor_str.replace(',', '.')
    
    try:
        return float(valor_str)
    except:
        return 0.0

# ============================================================================
# 3. CARREGAR ARQUIVO
# ============================================================================
try:
    df = pd.read_excel(INPUT_FILE_PATH, sheet_name='Planilha1', header=None)
    print(f"✅ Arquivo carregado: {df.shape[0]} linhas × {df.shape[1]} colunas")
except Exception as e:
    print(f"❌ ERRO ao carregar: {e}")
    exit()

# ============================================================================
# 4. ENCONTRAR DATAS
# ============================================================================
dates_info = []
for col_idx in range(min(300, df.shape[1])):
    try:
        cell_value = df.iat[0, col_idx]
        if isinstance(cell_value, str) and ('2025' in cell_value or '/' in cell_value):
            date_part = cell_value.split()[0] if ' ' in cell_value else str(cell_value)
            dates_info.append((col_idx, date_part))
    except:
        continue

print(f"📅 Datas encontradas: {len(dates_info)}")

# ============================================================================
# 5. ENCONTRAR CLIENTES
# ============================================================================
clientes = []
for row_idx in range(3, min(200, df.shape[0])):
    try:
        cliente = df.iat[row_idx, 0]
        if pd.notna(cliente) and str(cliente).strip() != "":
            clientes.append(str(cliente).strip())
    except:
        continue

print(f"👥 Clientes encontrados: {len(clientes)}")

# ============================================================================
# 6. TRANSFORMAÇÃO
# ============================================================================
print("\nTransformando dados...")

all_data = []

for cliente_idx, cliente in enumerate(clientes):
    linha_original = cliente_idx + 3
    
    for col_start, data_str in dates_info:
        if col_start + 6 >= df.shape[1]:
            continue
        
        try:
            # Extrair valores
            registro = {
                'CLIENTE': cliente,
                'DATA': data_str,
                'QTDE_PERNOITES': converter_valor_brasileiro(df.iat[linha_original, col_start]),
                'QTDE_CANCELAMENTOS': converter_valor_brasileiro(df.iat[linha_original, col_start + 1]),
                'QTDE_RESERVAS': converter_valor_brasileiro(df.iat[linha_original, col_start + 2]),
                'QTDE_VAGOS': converter_valor_brasileiro(df.iat[linha_original, col_start + 3]),
                'QTDE_HOSPEDES': converter_valor_brasileiro(df.iat[linha_original, col_start + 4]),
                'QTDE_OCUPADAS': converter_valor_brasileiro(df.iat[linha_original, col_start + 5]),
                'TOTAL_DIARIAS': converter_valor_brasileiro(df.iat[linha_original, col_start + 6])
            }
            
            all_data.append(registro)
            
        except:
            # Se der erro, cria registro vazio
            all_data.append({
                'CLIENTE': cliente,
                'DATA': data_str,
                'QTDE_PERNOITES': 0,
                'QTDE_CANCELAMENTOS': 0,
                'QTDE_RESERVAS': 0,
                'QTDE_VAGOS': 0,
                'QTDE_HOSPEDES': 0,
                'QTDE_OCUPADAS': 0,
                'TOTAL_DIARIAS': 0.0
            })

# Criar DataFrame
df_final = pd.DataFrame(all_data)

print(f"✅ Dados transformados: {len(df_final)} registros")

# ============================================================================
# 7. LIMPEZA FINAL
# ============================================================================
# Converter DATA para datetime
df_final['DATA'] = pd.to_datetime(df_final['DATA'], dayfirst=True, errors='coerce')

# Remover datas inválidas
df_final = df_final.dropna(subset=['DATA'])

# Remover duplicatas
df_final = df_final.drop_duplicates()

# Ordenar por Data e Cliente
df_final = df_final.sort_values(['DATA', 'CLIENTE']).reset_index(drop=True)

# Formatos consistentes para colunas numéricas
colunas_numericas = ['QTDE_PERNOITES', 'QTDE_CANCELAMENTOS', 'QTDE_RESERVAS', 
                     'QTDE_VAGOS', 'QTDE_HOSPEDES', 'QTDE_OCUPADAS', 'TOTAL_DIARIAS']

for col in colunas_numericas:
    df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

# ============================================================================
# 8. SALVAR APENAS EXCEL
# ============================================================================
print(f"\nSalvando arquivo Excel...")

try:
    with pd.ExcelWriter(OUTPUT_FILE_PATH, engine='openpyxl') as writer:
        # Salvar dados principais
        df_final.to_excel(writer, sheet_name='DADOS', index=False)
        
        # Ajustar largura das colunas
        worksheet = writer.sheets['DADOS']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"✅ Excel salvo com sucesso: {OUTPUT_FILE_PATH}")
    print(f"   • {len(df_final)} registros")
    print(f"   • {df_final.shape[1]} colunas")
    
except Exception as e:
    print(f"❌ ERRO ao salvar Excel: {e}")
    
    # Tentar método simples
    try:
        df_final.to_excel(OUTPUT_FILE_PATH, index=False)
        print(f"✅ Salvo em formato simples")
    except:
        print("❌ Falha total ao salvar")

# ============================================================================
# 9. RESUMO FINAL
# ============================================================================
print("\n" + "=" * 80)
print("RESUMO")
print("=" * 80)

print(f"\n📊 ESTATÍSTICAS:")
print(f"   • Total de registros: {len(df_final):,}")
print(f"   • Clientes únicos: {df_final['CLIENTE'].nunique()}")
print(f"   • Período: {df_final['DATA'].min().strftime('%d/%m/%Y')} a {df_final['DATA'].max().strftime('%d/%m/%Y')}")
print(f"   • Dias únicos: {df_final['DATA'].nunique()}")

# Calcular totais
try:
    soma_diarias = df_final['TOTAL_DIARIAS'].sum()
    print(f"   • Total diárias: R$ {soma_diarias:,.2f}")
except:
    print(f"   • Total diárias: R$ 0,00")

print(f"\n📍 Arquivo gerado em:")
print(f"   {os.path.abspath(OUTPUT_FILE_PATH)}")

print("\n" + "=" * 80)
print("CONCLUÍDO")
print("=" * 80)