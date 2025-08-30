# -*- coding: utf-8 -*-
import os
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
import config

def gerar_nuvem_de_palavras(texto_comentarios, nome_projeto):
    """Gera e salva uma imagem de nuvem de palavras."""
    print(f"Gerando nuvem de palavras para o projeto {nome_projeto}...")
    
    # Palavras comuns em português para serem ignoradas
    stopwords_pt = set(STOPWORDS)
    stopwords_pt.update([
        "de", "a", "o", "que", "e", "do", "da", "em", "um", "para", "com", "não",
        "uma", "os", "no", "na", "por", "mais", "as", "dos", "como", "mas", "ao",
        "ele", "das", "à", "seu", "sua", "ou", "quando", "muito", "nos", "já",
        "eu", "também", "só", "pelo", "pela", "até", "isso", "ela", "entre",
        "depois", "sem", "mesmo", "aos", "seus", "quem", "nas", "me", "esse",
        "eles", "você", "essa", "num", "nem", "suas", "meu", "às", "minha",
        "numa", "pelos", "elas", "qual", "nós", "lhes", "esses", "essas", "pelas",
        "este", "aquele", "aquela", "isto", "aquilo", "estou", "está", "estamos",
        "estão", "estive", "esteve", "estivemos", "estiveram", "estava", "estávamos",
        "estavam", "estivera", "estivéramos", "esteja", "estejamos", "estejam",
        "estivesse", "estivéssemos", "estivessem", "estiver", "estivermos",
        "estiverem", "hei", "há", "havemos", "hão", "houve", "houvemos", "houveram",
        "houvera", "houvéramos", "haja", "hajamos", "hajam", "houvesse",
        "houvéssemos", "houvessem", "houver", "houvermos", "houverem", "houverei",
        "houverá", "houveremos", "houverão", "houveria", "houveríamos", "houveriam",
        "sou", "somos", "são", "era", "éramos", "eram", "fui", "foi", "fomos",
        "foram", "fora", "fôramos", "seja", "sejamos", "sejam", "fosse", "fôssemos",
        "fossem", "for", "formos", "forem", "serei", "será", "seremos", "serão",
        "seria", "seríamos", "seriam", "tenho", "tem", "temos", "têm", "tinha",
        "tínhamos", "tinham", "tive", "teve", "tivemos", "tiveram", "tivera",
        "tivéramos", "tenha", "tenhamos", "tenham", "tivesse", "tivéssemos",
        "tivessem", "tiver", "tivermos", "tiverem", "terei", "terá", "teremos",
        "terão", "teria", "teríamos", "teriam", "foi", "atividade", "projeto",
        "realizado", "desenvolvimento", "tarefa"
    ])

    if not texto_comentarios.strip():
        print("AVISO: Não há comentários para gerar a nuvem de palavras.")
        return None

    wordcloud = WordCloud(
        width=800,
        height=400,
        background_color='white',
        stopwords=stopwords_pt,
        max_words=50,
        collocations=False
    ).generate(texto_comentarios)

    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    
    # Cria a pasta temp se não existir
    os.makedirs(config.PASTA_TEMP, exist_ok=True)
    
    caminho_arquivo = os.path.join(config.PASTA_TEMP, f"nuvem_{nome_projeto}.png")
    plt.savefig(caminho_arquivo)
    plt.close()
    
    print(f" -> Nuvem de palavras salva em: {caminho_arquivo}")
    return caminho_arquivo

def gerar_nuvem_global(df_completo):
    """Gera e salva uma imagem de nuvem de palavras com todos os comentários."""
    print("Gerando nuvem de palavras GLOBAL...")
    
    # Concatena todos os comentários de todos os projetos
    texto_comentarios = " ".join(df_completo['Comentários'].tolist())
    
    caminho_arquivo = gerar_nuvem_de_palavras(texto_comentarios, "GLOBAL")
    return caminho_arquivo