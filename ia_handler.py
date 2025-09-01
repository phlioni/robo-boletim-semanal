# -*- coding: utf-8 -*-
from openai import OpenAI
import config
import html
import json
import time
import pandas as pd

try:
    client = OpenAI(api_key=config.OPENAI_API_KEY)
except Exception as e:
    print(f"ERRO: Não foi possível inicializar o cliente da OpenAI. Verifique sua API Key. Erro: {e}")
    client = None

def classificar_comentarios_em_lote(comentarios):
    if not client: return {}
    print(f"Classificando {len(comentarios)} comentários únicos com IA em lotes...")
    mapa_geral_categorias = {}
    tamanho_lote = 100
    for i in range(0, len(comentarios), tamanho_lote):
        lote_comentarios = comentarios[i:i + tamanho_lote]
        print(f" -> Processando lote {i//tamanho_lote + 1}...")
        prompt = f"""
        Classifique cada comentário em uma das 8 categorias: "Desenvolvimento de Feature", "Correção de Bug", "Reunião", "Suporte/Atendimento", "Análise/Planejamento", "Testes/QA", "Documentação", "Outros".
        Responda SOMENTE com um objeto JSON onde a chave é o comentário e o valor é a categoria.
        Comentários: {json.dumps(lote_comentarios)}
        """
        try:
            response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "system", "content": "Você é um assistente de categorização de dados que responde em JSON."}, {"role": "user", "content": prompt}], response_format={"type": "json_object"})
            resultado_lote = json.loads(response.choices[0].message.content)
            mapa_geral_categorias.update(resultado_lote)
            time.sleep(1)
        except Exception as e:
            print(f"ERRO ao classificar o lote {i//tamanho_lote + 1}: {e}")
            for comentario in lote_comentarios:
                if comentario not in mapa_geral_categorias: mapa_geral_categorias[comentario] = 'Outros'
    print(" -> Comentários classificados com sucesso.")
    return mapa_geral_categorias

def gerar_boletim_para_projeto(nome_projeto, df_projeto_enriquecido):
    """Gera uma análise tática e rica para o boletim do líder de projeto."""
    if not client: return {"resumo_tatico": "N/A", "conquistas": [], "riscos_e_atencao": []}
    print(f"Gerando análise tática via IA para o projeto {nome_projeto}...")

    # Converte o dataframe em uma lista de dicionários para um contexto mais estruturado
    # Inclui campos importantes como status, tipo, esforço, etc.
    contexto_detalhado = df_projeto_enriquecido[['Profissional', 'Comentários', 'titulo', 'status', 'tipo', 'esforco_concluido', 'estimativa_original', 'Categoria']].to_dict(orient='records')
    
    prompt = f"""
    Aja como um Analista de Projetos Sênior auxiliando o Tech Lead do projeto "{nome_projeto}".
    Analise os seguintes dados de apontamentos da última semana e forneça uma análise tática e objetiva.
    Sua resposta DEVE ser um objeto JSON com três chaves: "resumo_tatico", "conquistas" e "riscos_e_atencao".

    - "resumo_tatico": Um parágrafo curto (2-3 frases) descrevendo o foco da semana. Ex: "O foco da semana foi na conclusão de novas funcionalidades, com um esforço significativo em desenvolvimento, enquanto a correção de bugs se manteve controlada."
    - "conquistas": Uma lista em markdown com 2 a 3 pontos positivos e entregas de valor. Ex: "- Conclusão da task crítica 'XYZ' (Card #12345) por [Nome].". Seja específico.
    - "riscos_e_atencao": Uma lista em markdown com 2 a 3 pontos de atenção ou riscos que o líder precisa investigar. Use os dados. Ex: "- O card #54321, com {_x_}h de esforço, já ultrapassou em {y_}% a estimativa original. Recomenda-se reavaliação.", "- Alta concentração de horas de suporte ({z_}h) com o profissional [Nome], indicando possível gargalo."

    Baseie TODA a sua análise exclusivamente nos dados fornecidos em formato JSON. Não invente informações. Se não houver pontos claros, retorne uma lista vazia.

    DADOS DA SEMANA:
    ---
    {json.dumps(contexto_detalhado, indent=2)}
    ---
    """
    try:
        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "system", "content": "Você é um analista de projetos sênior que responde em JSON."}, {"role": "user", "content": prompt}], response_format={"type": "json_object"}, temperature=0.3)
        resultado = json.loads(response.choices[0].message.content)
        return resultado
    except Exception as e:
        print(f"ERRO: Falha na API para o projeto {nome_projeto}: {e}")
        return {"resumo_tatico": "Falha ao gerar análise.", "conquistas": [], "riscos_e_atencao": []}

def gerar_boletim_geral(dados_fatos, kpis_avancados, amostra_comentarios_relevantes):
    if not client: return None
    print("\nIniciando análise ESTRATÉGICA com IA para o boletim geral...")
    prompt = f"""
    Aja como um CTO pragmático. Sua prioridade é identificar riscos e ineficiências usando os FATOS precisos que fornecerei. Não invente números. Sua resposta DEVE ser um objeto JSON.

    FATOS:
    - KPIs: {json.dumps(dados_fatos['kpis'])}
    - Alocação de Esforço: {json.dumps(dados_fatos['alocacao_esforco'])}
    - Portfólio de Clientes: {json.dumps(kpis_avancados['portfolio_clientes'])}
    - Foco da Equipe: {json.dumps(kpis_avancados['foco_equipe'])}
    - Amostra de Comentários: "{amostra_comentarios_relevantes}"

    ESTRUTURA JSON DE SAÍDA:
    {{
      "kpis": {{ "ponto_critico": "Com base em TODAS as análises, descreva em uma frase o maior PONTO DE ATENÇÃO da semana." }},
      "analise_estrategica": "Escreva um parágrafo (4-5 frases) com uma análise estratégica. Conecte os dados: comente sobre a alocação de esforço, aponte qual cliente está mais crítico e mencione o risco de foco. Finalize com uma recomendação de ação.",
      "destaques_pessoas": {{ "destaque_profissional": "Identifique o 'Profissional em Destaque' e descreva o feito em DETALHES, citando NOME e comentário.", "ponto_atencao_equipe": "Identifique o desafio mais significativo para a EQUIPE, descrevendo o problema DETALHADAMENTE com NOMES e projetos." }}
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