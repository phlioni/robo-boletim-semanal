# -*- coding: utf-8 -*-
import pandas as pd

def processar_relatorio_geral(caminho_arquivo):
    """Lê o relatório baixado e retorna um DataFrame limpo."""
    print("Processando a planilha geral...")
    try:
        df = pd.read_excel(caminho_arquivo, skiprows=18, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['Data'], inplace=True)
        
        # --- CORREÇÃO APLICADA AQUI ---
        # A linha que filtrava apenas por 'ZRC' foi REMOVIDA.
        # Agora, os filtros de 'Situação' e 'Comentários' são aplicados a todos os dados.
        
        df_filtrado = df[df['Situação'] == 'Aprovado'].copy()
        df_filtrado['Comentários'] = df_filtrado['Comentários'].astype(str)
        df_filtrado = df_filtrado[df_filtrado['Comentários'].str.len() >= 5]
        
        print("Planilha geral processada com sucesso.")
        return df_filtrado

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar a planilha: {e}")
        return None