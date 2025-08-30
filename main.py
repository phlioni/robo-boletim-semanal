# -*- coding: utf-8 -*-
import os
import shutil
import time

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
        # Etapa 1: Baixar e desbloquear o relatório
        caminho_relatorio_bruto = automacao_web.login_e_download_semanal()
        excel_handler.unprotect_and_save(caminho_relatorio_bruto)
        
        # Etapa 2: Processar a planilha geral
        df_completo = processamento_dados.processar_relatorio_geral(caminho_relatorio_bruto)
        
        if df_completo is not None and not df_completo.empty:
            
            # --- CORREÇÃO: Bloco para envio de e-mails individuais aos líderes RESTAURADO ---
            # Etapa 3: Envio dos boletins individuais para os líderes
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
            # --- FIM DO BLOCO RESTAURADO ---

            # Etapa 4: Geração do Boletim Executivo Geral
            print("\n--- INICIANDO GERAÇÃO DO BOLETIM EXECUTIVO COM DADOS PRECISOS ---")

            comentarios_unicos = df_completo['Comentários'].unique().tolist()
            mapa_categorias = ia_handler.classificar_comentarios_em_lote(comentarios_unicos)
            
            if mapa_categorias:
                df_completo['Categoria'] = df_completo['Comentários'].map(mapa_categorias).fillna('Outros')
            else:
                print("AVISO: Classificação por IA falhou. Usando classificação por palavras-chave como fallback.")
                def classificar_fallback(c):
                    c = c.lower()
                    if any(w in c for w in ["bug", "erro", "correção", "ajuste", "hotfix"]): return "Correção de Bug"
                    if any(w in c for w in ["reunião", "alinhamento", "call", "planning"]): return "Reunião"
                    return "Desenvolvimento"
                df_completo['Categoria'] = df_completo['Comentários'].apply(classificar_fallback)
            
            kpis = {"total_horas": df_completo['Horas'].sum(), "projetos_ativos": df_completo['Projeto'].nunique(), "profissionais_ativos": df_completo['Profissional'].nunique()}
            alocacao_horas = df_completo.groupby('Categoria')['Horas'].sum()
            alocacao_esforco = [{"categoria": cat, "horas": float(f"{horas:.1f}"), "percentual": f"{(horas / kpis['total_horas'] * 100):.1f}%"} for cat, horas in alocacao_horas.items()]
            
            projetos_selecionados = {}
            try:
                bugs_df = df_completo[df_completo['Categoria'] == 'Correção de Bug']
                projetos_selecionados['critico'] = bugs_df.groupby('Projeto')['Horas'].sum().idxmax() if not bugs_df.empty else "Nenhum"
            except Exception: projetos_selecionados['critico'] = "Nenhum"
            try:
                dev_df = df_completo[df_completo['Categoria'] == 'Desenvolvimento de Feature']
                projetos_selecionados['destaque'] = dev_df.groupby('Projeto')['Horas'].sum().idxmax() if not dev_df.empty else "Nenhum"
            except Exception: projetos_selecionados['destaque'] = "Nenhum"

            dados_fatos = { "kpis": kpis, "alocacao_esforco": alocacao_esforco, "projetos_selecionados": projetos_selecionados }
            comentarios_relevantes = df_completo[df_completo['Categoria'].isin(['Correção de Bug', 'Suporte'])]['Comentários'].tolist()
            amostra_comentarios = "\n".join(comentarios_relevantes[:20])

            narrativas_ia = ia_handler.gerar_boletim_geral(dados_fatos, amostra_comentarios)
            caminho_nuvem_geral = visualizacoes.gerar_nuvem_global(df_completo)
            envio_email.enviar_boletim_geral(dados_fatos, narrativas_ia, caminho_nuvem_geral)
            
        else:
            print("Nenhum dado encontrado no relatório para processar.")

    except Exception as e:
        print(f"\n--- ERRO CRÍTICO NO PROCESSO ---\n{e}")
    finally:
        if os.path.exists(config.PASTA_TEMP):
            shutil.rmtree(config.PASTA_TEMP)
        if caminho_relatorio_bruto and os.path.exists(caminho_relatorio_bruto):
            os.remove(caminho_relatorio_bruto)
        print("\n--- EXECUÇÃO DO ROBÔ FINALIZADA ---")

if __name__ == "__main__":
    run()