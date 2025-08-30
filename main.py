# -*- coding: utf-8 -*-
import os
import shutil

import config
import automacao_web
import excel_handler
import processamento_dados
import visualizacoes
import ia_handler
import envio_email

def run():
    """Função principal que orquestra todo o processo do boletim semanal."""
    print("--- INICIANDO ROBÔ DE BOLETIM DE INTELIGÊNCIA SEMANAL ---")
    caminho_relatorio_bruto = None
    
    try:
        caminho_relatorio_bruto = automacao_web.login_e_download_semanal()
        excel_handler.unprotect_and_save(caminho_relatorio_bruto)
        
        df_completo = processamento_dados.processar_relatorio_geral(caminho_relatorio_bruto)
        
        if df_completo is not None and not df_completo.empty:
            # Etapa 1: Envio dos boletins individuais para os líderes
            projetos_encontrados = df_completo['Projeto'].unique()
            
            for lider_email, prefixos_projetos in config.LIDERES_PROJETOS.items():
                print(f"\nProcessando projetos para o líder: {lider_email}")
                for prefixo in prefixos_projetos:
                    projetos_do_lider = [p for p in projetos_encontrados if p.startswith(prefixo)]
                    
                    for projeto_nome in projetos_do_lider:
                        df_projeto = df_completo[df_completo['Projeto'] == projeto_nome]
                        if df_projeto.empty: continue

                        print(f"--- Processando Projeto: {projeto_nome} ---")
                        
                        comentarios = " ".join(df_projeto['Comentários'].tolist())
                        caminho_nuvem = visualizacoes.gerar_nuvem_de_palavras(comentarios, projeto_nome)
                        boletim_html = ia_handler.gerar_boletim_para_projeto(projeto_nome, df_projeto)
                        envio_email.enviar_boletim(lider_email, projeto_nome, boletim_html, caminho_nuvem)

            # --- (NOVA ETAPA) Etapa 2: Envio do boletim geral para a diretoria ---
            boletim_geral_html = ia_handler.gerar_boletim_geral(df_completo)
            envio_email.enviar_boletim_geral(boletim_geral_html)
            
        else:
            print("Nenhum dado encontrado no relatório para processar.")

    except Exception as e:
        print(f"\n--- ERRO CRÍTICO NO PROCESSO ---")
        print(e)
    finally:
        if os.path.exists(config.PASTA_TEMP):
            shutil.rmtree(config.PASTA_TEMP)
        if caminho_relatorio_bruto and os.path.exists(caminho_relatorio_bruto):
            os.remove(caminho_relatorio_bruto)
            
        print("\n--- EXECUÇÃO DO ROBÔ FINALIZADA ---")

if __name__ == "__main__":
    run()