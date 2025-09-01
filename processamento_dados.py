# -*- coding: utf-8 -*-
import pandas as pd
import re

def extrair_id_do_comentario(comentario):
    """Extrai o primeiro ID numérico de um comentário (ex: #12345, WI12345)."""
    if not isinstance(comentario, str): return None
    match = re.search(r'#(\d+)|(?:WI|wi|workitem|item|task|bug|card)\s*(\d+)', comentario, re.IGNORECASE)
    if match: return int(next(g for g in match.groups() if g is not None))
    return None

def processar_relatorio_geral(caminho_arquivo):
    """Lê o relatório baixado e retorna um DataFrame limpo."""
    print("Processando a planilha geral...")
    try:
        df = pd.read_excel(caminho_arquivo, skiprows=18, engine='openpyxl')
        df.columns = df.columns.str.strip()
        if 'Projeto' in df.columns:
            df['Projeto'] = df['Projeto'].str.strip()
        else:
            raise KeyError("A coluna 'Projeto' não foi encontrada no relatório.")
        
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['Data'], inplace=True)
        df_filtrado = df[df['Situação'] == 'Aprovado'].copy()
        df_filtrado['Comentários'] = df_filtrado['Comentários'].astype(str)
        df_filtrado['ClienteCod'] = df_filtrado['Projeto'].str[:3]
        print("Planilha geral processada com sucesso.")
        return df_filtrado
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar a planilha: {e}")
        return None

def calcular_kpis_de_projeto(df_projeto):
    """Calcula os KPIs específicos para um único projeto (para o e-mail do líder)."""
    if df_projeto.empty: return {}
    
    total_horas = df_projeto['Horas'].sum()
    if total_horas == 0: return {}

    if 'Categoria' not in df_projeto.columns: df_projeto['Categoria'] = 'Outros'

    horas_bug = df_projeto[df_projeto['Categoria'] == 'Correção de Bug']['Horas'].sum()
    pct_bug = (horas_bug / total_horas * 100) if total_horas > 0 else 0
    
    try:
        top_contribuidor = df_projeto.groupby('Profissional')['Horas'].sum().idxmax()
    except ValueError:
        top_contribuidor = "N/A"
    
    return {
        "total_horas": f"{total_horas:.1f}h",
        "percentual_bugs": f"{pct_bug:.1f}%",
        "principal_contribuidor": top_contribuidor
    }

def calcular_kpis_avancados(df_classificado):
    """Calcula as métricas avançadas consolidadas (para o e-mail executivo)."""
    print("Calculando KPIs avançados de portfólio e alocação de equipe...")
    if 'Categoria' not in df_classificado.columns: df_classificado['Categoria'] = 'Outros'
    
    portfolio_clientes = []
    codigos_cliente = df_classificado['ClienteCod'].unique()
    
    for cliente in codigos_cliente:
        df_cliente = df_classificado[df_classificado['ClienteCod'] == cliente]
        total_horas = df_cliente['Horas'].sum()
        if total_horas == 0: continue
        
        horas_dev = df_cliente[df_cliente['Categoria'] == 'Desenvolvimento de Feature']['Horas'].sum()
        horas_bug = df_cliente[df_cliente['Categoria'] == 'Correção de Bug']['Horas'].sum()
        horas_reuniao = df_cliente[df_cliente['Categoria'] == 'Reunião']['Horas'].sum()
        pct_dev = (horas_dev / total_horas * 100)
        pct_bug = (horas_bug / total_horas * 100)
        pct_reuniao = (horas_reuniao / total_horas * 100)
        saude = "🔴 Crítico" if pct_bug > 35 else ("🟡 Atenção" if pct_bug > 20 else "🟢 Saudável")
        
        portfolio_clientes.append({ "cliente": cliente, "total_horas": f"{total_horas:.1f}h", "pct_dev": f"{pct_dev:.1f}%", "pct_bug": f"{pct_bug:.1f}%", "pct_reuniao": f"{pct_reuniao:.1f}%", "saude": saude })

    proj_por_profissional = df_classificado.groupby('Profissional')['Projeto'].nunique()
    fragmentados_df = proj_por_profissional[proj_por_profissional >= 3].reset_index().rename(columns={'Profissional': 'profissional', 'Projeto': 'num_projetos'})
    fragmentados_list = fragmentados_df.to_dict('records')
    
    dependencia_critica = []
    projetos = df_classificado['Projeto'].unique()
    for proj in projetos:
        df_proj = df_classificado[df_classificado['Projeto'] == proj]
        horas_totais_proj = df_proj['Horas'].sum()
        if horas_totais_proj == 0: continue
        horas_por_profissional = df_proj.groupby('Profissional')['Horas'].sum()
        if not horas_por_profissional.empty:
            profissional_principal = horas_por_profissional.idxmax()
            horas_principal = horas_por_profissional.max()
            concentracao = (horas_principal / horas_totais_proj * 100)
            if concentracao > 70:
                dependencia_critica.append({ "projeto": proj, "profissional": profissional_principal, "concentracao": f"{concentracao:.1f}%" })

    return { "portfolio_clientes": sorted(portfolio_clientes, key=lambda x: float(x['total_horas'][:-1]), reverse=True), "foco_equipe": { "fragmentados": fragmentados_list, "dependencia_critica": dependencia_critica } }