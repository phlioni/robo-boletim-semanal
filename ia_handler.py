# -*- coding: utf-8 -*-
from openai import OpenAI
import config
import html
import json
import time

try:
    client = OpenAI(api_key=config.OPENAI_API_KEY)
except Exception as e:
    print(f"ERRO: Não foi possível inicializar o cliente da OpenAI. Verifique sua API Key no config.py. Erro: {e}")
    client = None

def classificar_comentarios_em_lote(comentarios):
    """
    Etapa 1: Envia uma lista de comentários para a IA em lotes para classificação precisa.
    """
    if not client: return {}
    print(f"Classificando {len(comentarios)} comentários únicos com IA em lotes...")
    
    mapa_geral_categorias = {}
    tamanho_lote = 100 # Envia 100 comentários por vez

    for i in range(0, len(comentarios), tamanho_lote):
        lote_comentarios = comentarios[i:i + tamanho_lote]
        print(f" -> Processando lote {i//tamanho_lote + 1}...")

        prompt = f"""
        Sua tarefa é classificar cada comentário de apontamento de horas em uma das seguintes 8 categorias: "Desenvolvimento de Feature", "Correção de Bug", "Reunião", "Suporte/Atendimento", "Análise/Planejamento", "Testes/QA", "Documentação", "Outros".
        Analise a lista de comentários a seguir e retorne sua resposta SOMENTE como um objeto JSON, onde cada chave é o comentário original e o valor é a categoria.
        Comentários para classificar:
        {json.dumps(lote_comentarios)}
        """
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "Você é um assistente de categorização de dados que responde em JSON."},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"}
            )
            resultado_lote = json.loads(response.choices[0].message.content)
            mapa_geral_categorias.update(resultado_lote)
            time.sleep(1) # Pausa entre as chamadas para não sobrecarregar a API
        except Exception as e:
            print(f"ERRO ao classificar o lote {i//tamanho_lote + 1} com a IA: {e}")
            # Se um lote falhar, preenche com 'Outros' para não parar o processo
            for comentario in lote_comentarios:
                if comentario not in mapa_geral_categorias:
                    mapa_geral_categorias[comentario] = 'Outros'

    print(" -> Comentários classificados com sucesso.")
    return mapa_geral_categorias


def gerar_boletim_geral(dados_fatos, amostra_comentarios_relevantes):
    """Etapa 2: Usa fatos pré-calculados para gerar a narrativa executiva."""
    if not client: return None
    print("\nIniciando análise ESTRATÉGICA com IA para o boletim geral...")
    
    proj_critico = dados_fatos.get("projetos_selecionados", {}).get("critico", "Nenhum")
    proj_destaque = dados_fatos.get("projetos_selecionados", {}).get("destaque", "Nenhum")
    
    prompt = f"""
    Aja como um Diretor de Tecnologia (CTO) pragmático e direto, criando um boletim executivo.
    Sua prioridade é identificar riscos e ineficiências. Use os FATOS NUMÉRICOS PRECISOS que fornecerei.
    Sua resposta DEVE ser um objeto JSON bem formatado.

    FATOS (Fonte única da verdade):
    - KPIs da Semana: {json.dumps(dados_fatos['kpis'])}
    - Alocação de Esforço Calculada: {json.dumps(dados_fatos['alocacao_esforco'])}
    - Projetos Chave Identificados: Projeto Crítico = {proj_critico}, Projeto Destaque = {proj_destaque}
    - Amostra de Comentários Relevantes para Contexto: "{amostra_comentarios_relevantes}"

    ESTRUTURA JSON DE SAÍDA (preencha com sua análise detalhada e focada em problemas):
    {{
      "kpis": {{
        "ponto_critico": "Com base nos comentários do '{proj_critico}', descreva em uma frase o maior e mais urgente PONTO DE ATENÇÃO da semana."
      }},
      "analise_estrategica": "Escreva um parágrafo (4-5 frases) com uma análise estratégica. Comece pelo maior DESAFIO ou ineficiência da semana (ex: alto % de bugs no projeto '{proj_critico}'). Depois, mencione uma conquista importante para balancear. Finalize com uma RECOMENDAÇÃO DE AÇÃO clara e direta para a gestão.",
      "radar_projetos": [
        {{ "projeto": "{proj_critico}", "saude": "Vermelho", "tendencia": "⬇️", "justificativa": "Analise os comentários do '{proj_critico}' e escreva uma justificativa DETALHADA para seu status crítico, mencionando o problema específico." }},
        {{ "projeto": "{proj_destaque}", "saude": "Verde", "tendencia": "⬆️", "justificativa": "Analise os comentários do '{proj_destaque}' e escreva uma justificativa DETALHADA para seu sucesso, mencionando a entrega." }}
      ],
      "destaques_pessoas": {{
        "destaque_profissional": "Identifique o 'Profissional em Destaque' da semana. Descreva o feito em DETALHES, citando o NOME DO PROFISSIONAL e, se possível, parte do seu comentário entre aspas para dar contexto.",
        "ponto_atencao_equipe": "Identifique o desafio mais significativo que parece afetar a EQUIPE. Descreva o problema DETALHADAMENTE, incluindo o NOME DO PROFISSIONAL que o reportou e o PROJETO impactado."
      }}
    }}
    """
    try:
        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "system", "content": "Você é um CTO que gera relatórios executivos em JSON a partir de fatos pré-calculados."}, {"role": "user", "content": prompt}], response_format={"type": "json_object"}, temperature=0.5)
        dados_analisados = json.loads(response.choices[0].message.content)
        print(" -> Análise GERAL da IA concluída com sucesso.")
        return dados_analisados
    except Exception as e:
        print(f"ERRO: Falha ao chamar a API de IA para o boletim geral. Erro: {e}")
        return None

def gerar_boletim_para_projeto(nome_projeto, df_projeto):
    """Envia os dados de UM projeto para a IA e retorna o boletim em HTML."""
    # (Esta função não foi alterada e permanece aqui para os e-mails individuais)
    if not client: return "<p>Cliente da API da OpenAI não inicializado.</p>"
    print(f"Iniciando análise com IA para o projeto {nome_projeto}...")
    comentarios_texto = "\n".join(f"- {row['Profissional']}: {row['Comentários']} ({row['Horas']:.2f}h)" for index, row in df_projeto.iterrows())
    prompt = f"""
    Aja como um analista de projetos sênior. Crie um boletim de inteligência semanal em HTML para o projeto "{nome_projeto}".
    Sua resposta DEVE conter 3 seções com títulos <h3>: "Resumo Executivo", "Alocação de Esforço" e "Destaques e Pontos de Atenção".
    - Na "Alocação de Esforço", classifique as tarefas (ex: Desenvolvimento, Correção de Bug, Reunião, Suporte) e mostre o total de horas e a porcentagem de cada.
    - Nos "Destaques", crie sub-seções <h4> "Destaques Positivos" e "Pontos de Atenção". OBRIGATORIAMENTE mencione o nome do profissional ao citar um comentário.
    Use <b> para destacar pontos importantes.
    Dados da semana:
    ---
    {comentarios_texto}
    ---
    """
    try:
        response = client.chat.completions.create(model="gpt-3.5-turbo", messages=[{"role": "system", "content": "Você é um analista de projetos sênior que cria relatórios HTML."}, {"role": "user", "content": prompt}], temperature=0.5, max_tokens=1000)
        conteudo_ia_bruto = response.choices[0].message.content
        return html.unescape(conteudo_ia_bruto)
    except Exception as e:
        print(f"ERRO: Falha ao chamar a API de IA para o projeto {nome_projeto}. Erro: {e}")
        return "<p>Ocorreu um erro ao gerar a análise da IA.</p>"