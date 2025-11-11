import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

RAW_DATA_PATH = "data_raw"
TRANSFORMED_DATA_PATH = "data_transformed"
OUTPUT_FILE = "data_costumer_care.xlsx"
ACTIVE_SALES_STATUSES = ['OPEN ORDER'] 
TRACKED_PICKING_STATUSES = ['ACTIVATED', 'COMPLETED'] 
COVERAGE_STATUS = 'AVAILABLE'

def transform_data():
    
    print("-" * 50)
    print("Iniciando Transformação de Dados...")
    print("Carregando arquivos de origem...")

    try:
        df_sales = pd.read_excel(os.path.join(RAW_DATA_PATH, "CHINTSalesDetail.xlsx"))
        df_picking = pd.read_excel(os.path.join(RAW_DATA_PATH, "SalesPickingList.xlsx"))
        df_stock = pd.read_excel(os.path.join(RAW_DATA_PATH, "OnHandInventory.xlsx"))
        df_customer = pd.read_excel(os.path.join(RAW_DATA_PATH, "AllCostumers.xlsx"))
        df_po = pd.read_excel(os.path.join(RAW_DATA_PATH, "OpenPurchaseOrderLines.xlsx"))

    except FileNotFoundError as e:
        print(f"Erro: arquivo não encontrado: {e.filename}. Verifique a pasta '{RAW_DATA_PATH}'.")
        return
    except Exception as e:
        print(f"Erro ao ler arquivos: {e}")
        return

    print("Pré-processando e garantindo tipos de dados...")

    def clean_merge_keys(df, cols):
        for col in cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
        return df
    
   
    sales_cols_map = {
        'SalesId': 'salesid',
        'Item Id': 'itemid',
        'Cust Account': 'cust_account_id',
        'Sales Amount': 'sales_amount',
        'Open Qty': 'open_qty_order',
        'Create Date': 'order_date'
    }
    
    df_sales = df_sales.rename(columns={k: v for k, v in sales_cols_map.items() if k in df_sales.columns})
    df_sales = clean_merge_keys(df_sales, ['salesid', 'itemid', 'cust_account_id'])

    if 'salesid' not in df_sales.columns or 'itemid' not in df_sales.columns:
        print(f"ERRO CRÍTICO DE CHAVE DE MERGE: As colunas 'salesid' (esperada de 'SalesId') ou 'itemid' (esperada de 'Item Id') não foram encontradas em 'CHINTSalesDetail.xlsx'.")
        print(f"Colunas disponíveis em df_sales: {df_sales.columns.tolist()}")
        return

    if 'Sales Status' in df_sales.columns:
        df_sales = df_sales.rename(columns={'Sales Status': 'sales_status'})
    elif 'Status Venda' in df_sales.columns:
         df_sales = df_sales.rename(columns={'Status Venda': 'sales_status'})
    
    if 'sales_status' not in df_sales.columns:
        print(f"ERRO CRÍTICO DE COLUNA: Não foi possível encontrar 'sales_status' em 'CHINTSalesDetail.xlsx'.")
        print(f"Colunas disponíveis: {df_sales.columns.tolist()}")
        return

    df_sales['sales_status'] = df_sales['sales_status'].astype(str).str.strip().str.upper()

    picking_cols_map = {
        'Number': 'salesid',
        'Item number': 'itemid',
        'Route': 'picking_route',                 
        'Handling status': 'picking_status',      
        'Created date and time': 'picking_date',  
        'Quantity': 'picking_qty',                
    }
    
    df_picking = df_picking.rename(columns={k: v for k, v in picking_cols_map.items() if k in df_picking.columns})
    df_picking = clean_merge_keys(df_picking, ['salesid', 'itemid'])
    
    if 'salesid' not in df_picking.columns or 'itemid' not in df_picking.columns:
        print(f"ERRO CRÍTICO DE CHAVE DE MERGE: As colunas 'salesid' (esperada de 'Number') ou 'itemid' não foram encontradas em 'SalesPickingList.xlsx'.")
        print(f"Colunas disponíveis em df_picking: {df_picking.columns.tolist()}")
        return

    required_picking_cols = ['picking_status', 'picking_route', 'picking_qty', 'picking_date', 'salesid', 'itemid']
    missing_picking_cols = [col for col in required_picking_cols if col not in df_picking.columns]

    if missing_picking_cols:
        print(f"ERRO CRÍTICO DE COLUNA: Ainda faltam colunas essenciais no Picking List ('SalesPickingList.xlsx'): {missing_picking_cols}")
        print("Por favor, verifique se os nomes das colunas estão EXATAMENTE corretos (incluindo Capitalização).")
        return

    if 'picking_status' in df_picking.columns:
        df_picking['picking_status'] = df_picking['picking_status'].astype(str).str.strip().str.upper()
    else:
        df_picking['picking_status'] = pd.NA

    stock_cols_map = {
        'Item number': 'itemid', 
    }
    df_stock = df_stock.rename(columns={k: v for k, v in stock_cols_map.items() if k in df_stock.columns})
    df_stock = clean_merge_keys(df_stock, ['itemid'])
    
    if 'Total available' in df_stock.columns:
        df_stock = df_stock.rename(columns={'Total available': 'stock_available'})
    elif 'Total available' not in df_stock.columns and 'stock_available' not in df_stock.columns:
        df_stock = df_stock.rename(columns={'Quantity': 'stock_available'})

    customer_cols_map = {
        'Account': 'cust_account_id',
        'Name': 'customer_name',
        'Customer group': 'customer_group',
    }
    df_customer = df_customer.rename(columns={k: v for k, v in customer_cols_map.items() if k in df_customer.columns})
    df_customer = clean_merge_keys(df_customer, ['cust_account_id'])
    
    if 'Employee responsible' in df_customer.columns:
        df_customer = df_customer.rename(columns={'Employee responsible': 'sales_responsible'})

    po_cols_map = {
        'Item number': 'itemid',
        'Requested receipt date': 'requested_receipt_date'
    }

    if 'Item number' in df_po.columns or 'itemid' in df_po.columns:
        df_po = df_po.rename(columns={k: v for k, v in po_cols_map.items() if k in df_po.columns})
        df_po = clean_merge_keys(df_po, ['itemid'])
    else:
        print("Aviso: Não foi possível encontrar a coluna de item (Item number/itemid) em OpenPurchaseOrderLines. A importação será ignorada.")
        df_po = pd.DataFrame(columns=['itemid', 'requested_receipt_date', 'Quantity'])

    print(f"Filtrando linhas ativas ({', '.join(ACTIVE_SALES_STATUSES)})...")
    df_filtered = df_sales[df_sales['sales_status'].isin(ACTIVE_SALES_STATUSES)].copy()
    
    if df_filtered.empty:
        print("Aviso: Nenhum registro de vendas ativo encontrado após a filtragem inicial.")
        return

    print(f"Realizando merge com Picking List (status {', '.join(TRACKED_PICKING_STATUSES)})...")

    df_picking_tracked = df_picking[df_picking['picking_status'].isin(TRACKED_PICKING_STATUSES)].copy()
    
    picking_cols = [col for col in ['salesid', 'itemid', 'picking_route', 'picking_status', 'picking_qty', 'picking_date'] if col in df_picking_tracked.columns]
    
    df_picking_tracked = df_picking_tracked[picking_cols]
    
    df_merged = pd.merge(
        df_filtered, 
        df_picking_tracked, 
        on=['salesid', 'itemid'], 
        how='left', 
        suffixes=('_order', '_picking')
    )

    print("Adicionando informações de Estoque...")

    df_stock_filtered = df_stock.rename(columns={'coverage_status': 'coverage_status_stock'})
    
    cols_to_select = ['itemid']
    if 'stock_available' in df_stock_filtered.columns:
        cols_to_select.append('stock_available')
    if 'coverage_status_stock' in df_stock_filtered.columns:
        cols_to_select.append('coverage_status_stock')

    if 'itemid' not in df_stock_filtered.columns:
        df_stock_filtered = df_stock_filtered.assign(itemid=pd.NA)

    df_stock_filtered = df_stock_filtered[cols_to_select]

    df_merged = pd.merge(
        df_merged, 
        df_stock_filtered, 
        on='itemid', 
        how='left'
    )
    
    if 'stock_available' in df_merged.columns:
        df_merged['stock_available'] = df_merged['stock_available'].fillna(0)
    else:
        df_merged['stock_available'] = 0

    if 'coverage_status_stock' in df_merged.columns:
        df_merged['coverage_status'] = df_merged['coverage_status_stock'].fillna('NO COVERAGE').astype(str).str.strip().str.upper()
        df_merged = df_merged.drop(columns=['coverage_status_stock'])
    else:
        df_merged['coverage_status'] = 'NO COVERAGE'

    print("Adicionando dados do Cliente (Nome, Vendedor/Grupo)...")
    
    df_customer_filtered = df_customer[['cust_account_id', 'customer_name', 'sales_responsible', 'customer_group']]

    df_merged = pd.merge(
        df_merged, 
        df_customer_filtered, 
        on='cust_account_id', 
        how='left'
    )
    
    print("Adicionando previsão de chegada de Importação (PO)...")

    df_po_filtered = df_po.dropna(subset=['requested_receipt_date']).copy()
    
    if not df_po_filtered.empty:
        df_po_filtered['requested_receipt_date'] = pd.to_datetime(df_po_filtered['requested_receipt_date'], errors='coerce')
        
        df_po_grouped = df_po_filtered.groupby('itemid')['requested_receipt_date'].min().reset_index()
        df_po_grouped.rename(columns={'requested_receipt_date': 'Chegada Importação'}, inplace=True)
        
        df_merged = pd.merge(
            df_merged, 
            df_po_grouped, 
            on='itemid', 
            how='left'
        )
    else:
        df_merged['Chegada Importação'] = pd.NaT

    print("Calculando status logístico e datas de faturamento...")

    if 'picking_status' not in df_merged.columns:
        df_merged['picking_status'] = pd.NA

    df_merged['status_logistica'] = df_merged.apply(
        lambda row: 'Em Picking (ATIVO)' if row['picking_status'] == 'ACTIVATED'
        else ('Picking Concluído' if row['picking_status'] == 'COMPLETED' else row['sales_status']), 
        axis=1
    )
    
    hoje = datetime.now()
    data_120_dias = (hoje.replace(day=1) + timedelta(days=120)).replace(day=1)

    def calcular_data_fatura(row):
        data_chegada = row['Chegada Importação']

        if pd.notna(data_chegada) and data_chegada != 'Sem Cobertura':
            try:
                return pd.to_datetime(data_chegada, errors='coerce') + timedelta(days=4)
            except:
                return data_120_dias
        else:
            return data_120_dias

    df_merged['Data prevista para fatura'] = df_merged.apply(calcular_data_fatura, axis=1)

    df_merged['Chegada Importação'] = df_merged['Chegada Importação'].apply(
        lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) and isinstance(x, (datetime, pd.Timestamp)) else 'Sem Cobertura'
    )
    
    if not os.path.exists(TRANSFORMED_DATA_PATH):
        os.makedirs(TRANSFORMED_DATA_PATH)
        
    output_path = os.path.join(TRANSFORMED_DATA_PATH, OUTPUT_FILE)
    df_merged.to_excel(output_path, index=False)
    
    print(f"Transformação concluída! Arquivo salvo em '{output_path}'.")
    print(f"Total de linhas na base final: {len(df_merged)}")
    print("-" * 50)
    
if __name__ == "__main__":
    transform_data()
