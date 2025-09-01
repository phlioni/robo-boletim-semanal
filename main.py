# -*- coding: utf-8 -*-
import os
import shutil
import pandas as pd
import time

import config
import automacao_web
import excel_handler
import processamento_dados
import ia_handler
import envio_email
import azure_devops_handler

def run():
    """Função principal que orquestra todo o processo do boletim semanal."""
    print("--- INICIANDO ROBÔ DE BOLETIM DE INTELIGÊNCIA SEMANAL ---")
    caminho_relatorio_bruto = None
    
    try:
        # 1. COLETA E PROCESSAMENTO INICIAL (igual)
        caminho_relatorio_bruto = automacao_web.login_e_download_semanal()
        excel_handler.unprotect_and_save(caminho_relatorio_bruto)
        df_completo = processamento_dados.processar_relatorio_geral(caminho_relatorio_bruto)
        
        if df_completo is None or df_completo.empty:
            print("Nenhum dado encontrado no relatório para processar.")
            return

        df_completo['ID_DevOps'] = df_completo['Comentários'].apply(processamento_dados.extrair_id_do_comentario)

        # 2. ENRIQUECIMENTO GLOBAL DE DADOS (NOVA LÓGICA)
        print("\n--- ENRIQUECENDO DADOS DE TODOS OS PROJETOS ---")
        dfs_enriquecidos_globais = []
        
        # Enriquece projetos com integração Azure
        for org_url, projetos_na_org in config.MAPEAMENTO_PROJETOS_AZURE.items():
            for prefixo, azure_project_name in projetos_na_org.items():
                df_filtrado = df_completo[df_completo['ClienteCod'] == prefixo].copy()
                if df_filtrado.empty: continue
                
                ids_para_buscar = df_filtrado['ID_DevOps'].dropna().unique().tolist()
                dados_devops = azure_devops_handler.get_work_items_details(org_url, azure_project_name, ids_para_buscar)
                
                if dados_devops:
                    df_devops = pd.DataFrame.from_dict(dados_devops, orient='index').reset_index().rename(columns={'index': 'ID_DevOps'})
                    df_enriquecido = pd.merge(df_filtrado, df_devops, on='ID_DevOps', how='left')
                    dfs_enriquecidos_globais.append(df_enriquecido)
                else:
                    dfs_enriquecidos_globais.append(df_filtrado)

        # Adiciona projetos sem integração Azure
        df_sem_azure = df_completo[df_completo['ClienteCod'].isin(config.PROJETOS_SEM_AZURE)]
        dfs_enriquecidos_globais.append(df_sem_azure)

        if not dfs_enriquecidos_globais:
            print("Nenhum projeto configurado foi encontrado no relatório.")
            return
            
        # Consolida TODOS os dados em um único DataFrame
        df_global_enriquecido = pd.concat(dfs_enriquecidos_globais, ignore_index=True)

        # Classificação por IA de TODOS os comentários de uma só vez (mais eficiente)
        comentarios_unicos = df_global_enriquecido['Comentários'].unique().tolist()
        mapa_categorias = ia_handler.classificar_comentarios_em_lote(comentarios_unicos)
        if mapa_categorias:
            df_global_enriquecido['Categoria'] = df_global_enriquecido['Comentários'].map(mapa_categorias).fillna('Outros')

        # --- ESTEIRA 1: GERAÇÃO DO BOLETIM EXECUTIVO GERAL ---
        print("\n--- INICIANDO ESTEIRA DO BOLETIM EXECUTIVO ---")
        # AGORA, os dados são completos e refletem TODOS os projetos do relatório.
        kpis_avancados_geral = processamento_dados.calcular_kpis_avancados(df_global_enriquecido)
        kpis = {"total_horas": df_global_enriquecido['Horas'].sum()}
        alocacao_horas = df_global_enriquecido.groupby('Categoria')['Horas'].sum()
        alocacao_esforco = [{"categoria": cat, "horas": float(f"{horas:.1f}"), "percentual": f"{(horas / kpis['total_horas'] * 100):.1f}%"} for cat, horas in alocacao_horas.items()]
        dados_fatos = { "kpis": kpis, "alocacao_esforco": alocacao_esforco }
        
        amostra_comentarios = "\n".join(df_global_enriquecido.sample(min(len(df_global_enriquecido), 20))['Comentários'].tolist())
        narrativas_ia = ia_handler.gerar_boletim_geral(dados_fatos, kpis_avancados_geral, amostra_comentarios)
        envio_email.enviar_boletim_geral(dados_fatos, narrativas_ia, kpis_avancados_geral)

        # --- ESTEIRA 2: GERAÇÃO DOS BOLETINS INDIVIDUAIS PARA LÍDERES ---
        print("\n--- INICIANDO ESTEIRA DE BOLETINS PARA LÍDERES ---")
        for lider_email, prefixos_lider in config.LIDERES_E_SEUS_PROJETOS.items():
            # Filtra o DataFrame GERAL para obter apenas os dados deste líder
            df_lider = df_global_enriquecido[df_global_enriquecido['ClienteCod'].isin(prefixos_lider)].copy()
            if df_lider.empty:
                print(f"Nenhum dado encontrado para o líder {lider_email}.")
                continue
            
            projetos_do_lider = df_lider['Projeto'].unique()
            for projeto_nome in projetos_do_lider:
                df_projeto_especifico = df_lider[df_lider['Projeto'] == projeto_nome]
                if df_projeto_especifico.empty: continue

                print(f"--- Processando Boletim para Projeto: {projeto_nome} ---")
                # (A chamada para as funções de boletim de líder será melhorada no Problema 2)
                kpis_projeto = processamento_dados.calcular_kpis_de_projeto(df_projeto_especifico)
                analise_ia_projeto = ia_handler.gerar_boletim_para_projeto(projeto_nome, df_projeto_especifico)
                envio_email.enviar_boletim(lider_email, projeto_nome, kpis_projeto, analise_ia_projeto)

    except Exception as e:
        print(f"\n--- ERRO CRÍTICO NO PROCESSO ---\n{e}")
    finally:
        if caminho_relatorio_bruto and os.path.exists(caminho_relatorio_bruto):
            os.remove(caminho_relatorio_bruto)
        print("\n--- EXECUÇÃO DO ROBÔ FINALIZADA ---")

if __name__ == "__main__":
    run()