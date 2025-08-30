# -*- coding: utf-8 -*-
from openai import OpenAI
import config
import html
import json

try:
    client = OpenAI(api_key=config.OPENAI_API_KEY)
except Exception as e:
    print(f"ERRO: Não foi possível inicializar o cliente da OpenAI. Verifique sua API Key. Erro: {e}")
    client = None

# A função gerar_boletim_para_projeto (individual) permanece a mesma...

def gerar_boletim_geral(df_completo):
    """Envia os dados de TODOS os projetos para a IA e retorna um dicionário estruturado com a análise."""
    if not client: return None
    print("\nIniciando análise ESTRATÉGICA com IA para todos os projetos...")
    
    comentarios_texto = "\n".join(f"- Projeto {row['Projeto']} / {row['Profissional']}: {row['Comentários']} ({row['Horas']:.2f}h)" for index, row in df_completo.iterrows())
    total_horas = df_completo['Horas'].sum()
    num_projetos = df_completo['Projeto'].nunique()
    num_profissionais = df_completo['Profissional'].nunique()

    prompt = f"""
    Aja como um Diretor de Tecnologia (CTO) analisando os apontamentos de horas da semana.
    Os dados brutos contêm {num_projetos} projetos ativos, {num_profissionais} profissionais e um total de {total_horas:.2f} horas.
    Sua tarefa é gerar um relatório de inteligência para outros executivos.
    Sua resposta DEVE ser um objeto JSON bem formatado com a seguinte estrutura:
    {{
      "kpis": {{
        "total_horas": {total_horas:.2f},
        "projetos_ativos": {num_projetos},
        "profissionais_ativos": {num_profissionais},
        "ponto_critico": "Identifique e descreva em uma frase o maior ponto de atenção da semana"
      }},
      "analise_estrategica": "Escreva um parágrafo (4-5 frases) com uma análise estratégica da semana, conectando os dados de esforço, uma grande conquista e o principal risco.",
      "alocacao_esforco": [
        {{ "categoria": "Desenvolvimento", "horas": "some as horas e coloque aqui", "percentual": "calcule o percentual e coloque aqui" }},
        {{ "categoria": "Correção de Bug", "horas": "some as horas e coloque aqui", "percentual": "calcule o percentual e coloque aqui" }},
        {{ "categoria": "Reunião", "horas": "some as horas e coloque aqui", "percentual": "calcule o percentual e coloque aqui" }},
        {{ "categoria": "Suporte", "horas": "some as horas e coloque aqui", "percentual": "calcule o percentual e coloque aqui" }}
      ],
      "radar_projetos": [
        {{ "projeto": "Nome do Projeto 1", "saude": "Verde, Amarelo ou Vermelho", "tendencia": "Use um emoji de seta para cima, para o lado ou para baixo", "justificativa": "Justificativa concisa" }},
        {{ "projeto": "Nome do Projeto 2", "saude": "Verde, Amarelo ou Vermelho", "tendencia": "Use um emoji de seta para cima, para o lado ou para baixo", "justificativa": "Justificativa concisa" }}
      ],
      "destaques_pessoas": {{
        "destaque_profissional": "Identifique um 'Profissional em Destaque' e descreva o feito com base em um comentário específico.",
        "ponto_atencao_equipe": "Identifique um desafio ou problema que parece afetar a equipe como um todo."
      }}
    }}
    Use os comentários a seguir como base para sua análise:
    ---
    {comentarios_texto}
    ---
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o", # Usar um modelo mais avançado para melhor análise
            messages=[
                {"role": "system", "content": "Você é um CTO que gera relatórios executivos em formato JSON."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.5
        )
        conteudo_ia_bruto = response.choices[0].message.content
        dados_analisados = json.loads(conteudo_ia_bruto)
        print(" -> Análise GERAL da IA concluída com sucesso.")
        return dados_analisados
    except Exception as e:
        print(f"ERRO: Falha ao chamar a API de IA para o boletim geral. Erro: {e}")
        return None