# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime
import os

import config

def enviar_boletim(destinatario, nome_projeto, conteudo_ia_html, caminho_nuvem_palavras):
    # (Esta função para líderes individuais permanece a mesma)
    print(f"Preparando boletim para {destinatario} sobre o projeto {nome_projeto}...")
    msg = MIMEMultipart('related')
    msg['Subject'] = f"Boletim de Inteligência Semanal - Projeto {nome_projeto}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = destinatario
    corpo_html = f"""<html><head></head><body style="font-family: Calibri, sans-serif; font-size: 12pt;">{conteudo_ia_html}<hr><h3>Nuvem de Palavras dos Comentários</h3><p>Os termos mais frequentes mencionados pela equipe no projeto esta semana:</p><img src="cid:wordcloud"><br><br><p style="font-size: 9pt; color: #888;">Este e-mail foi gerado automaticamente pelo Robô Analista de Projetos.</p><p style="font-size: 9pt; color: #888;">Data da execução: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p></body></html>"""
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))

    if caminho_nuvem_palavras and os.path.exists(caminho_nuvem_palavras):
        try:
            with open(caminho_nuvem_palavras, 'rb') as f:
                img_nuvem = MIMEImage(f.read()); img_nuvem.add_header('Content-ID', '<wordcloud>'); msg.attach(img_nuvem)
        except Exception as e:
            print(f"AVISO: Falha ao anexar a nuvem de palavras. Erro: {e}")
    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, destinatario, msg.as_string())
        server.quit()
        print(f" -> Boletim para {nome_projeto} enviado com sucesso para {destinatario}.")
    except Exception as e:
        print(f"ERRO ao enviar boletim para {destinatario}: {e}")

def enviar_boletim_geral(dados_fatos, narrativas_ia, caminho_nuvem_global):
    """Envia o boletim executivo combinando fatos do Python e narrativas da IA."""
    if not dados_fatos or not narrativas_ia:
        print("AVISO: Faltam dados ou análise da IA para gerar o boletim geral. Pulando envio.")
        return
        
    print("\nPreparando boletim GERAL EXECUTIVO para a gestão...")
    # ... (cabeçalho do e-mail permanece o mesmo) ...
    msg = MIMEMultipart('related')
    msg['Subject'] = f"Boletim Executivo Semanal de Projetos | {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = ", ".join(config.GERAL_DESTINATARIOS)
    
    kpis = dados_fatos.get('kpis', {})
    alocacao = dados_fatos.get('alocacao_esforco', [])
    # As narrativas agora vêm de uma fonte diferente
    analise_estrategica = narrativas_ia.get('analise_estrategica', 'N/A')
    radar = narrativas_ia.get('radar_projetos', [])
    destaques = narrativas_ia.get('destaques_pessoas', {})
    ponto_critico = narrativas_ia.get('kpis', {}).get('ponto_critico', 'N/A') # Pega o ponto crítico da narrativa

    html_alocacao = ""
    cores_alocacao = {'Desenvolvimento': '#4CAF50', 'Correção de Bug': '#f44336', 'Reunião': '#2196F3', 'Suporte': '#FFC107', 'Outros': '#888'}
    for item in alocacao:
        cor = cores_alocacao.get(item.get('categoria', 'Outros'), '#888')
        html_alocacao += f"<p style='margin: 5px 0; font-size: 11pt;'><strong>{item.get('categoria', '')}:</strong> {item.get('horas')}h ({item.get('percentual', '0%')})<br><div style='background-color: {cor}; width: {item.get('percentual', '0%')}; height: 20px; border-radius: 5px;'></div></p>"

    html_radar = ""
    for proj in radar:
        cor_saude = "green" if proj.get('saude') == 'Verde' else ('#FFC107' if proj.get('saude') == 'Amarelo' else 'red')
        html_radar += f"<tr><td style='padding: 8px; border-bottom: 1px solid #ddd;'><strong>{proj.get('projeto','')}</strong></td><td style='padding: 8px; border-bottom: 1px solid #ddd; color: {cor_saude}; font-weight:bold;'>{proj.get('saude','')}</td><td style='padding: 8px; border-bottom: 1px solid #ddd; text-align:center; font-size:16pt;'>{proj.get('tendencia','')}</td><td style='padding: 8px; border-bottom: 1px solid #ddd;'><em>{proj.get('justificativa','')}</em></td></tr>"

    corpo_html = f"""
    <html><body style="font-family: Calibri, sans-serif; font-size: 12pt; color: #333;">
        <h2 style="color: #041C2F;">Boletim Executivo Semanal de Projetos</h2>
        <table style="width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 20px; font-size:11pt;">
            <tr style="background-color: #f2f2f2;">
                <td style="padding: 10px;">📈 <strong>{kpis.get('total_horas', 0):.2f}</strong> Horas Apontadas</td>
                <td style="padding: 10px;">📂 <strong>{kpis.get('projetos_ativos', 0)}</strong> Projetos Ativos</td>
                <td style="padding: 10px;">👥 <strong>{kpis.get('profissionais_ativos', 0)}</strong> Profissionais Ativos</td>
                <td style="padding: 10px; color:red;">🔴 <strong>Ponto Crítico:</strong> {narrativas_ia.get('kpis', {}).get('ponto_critico', 'N/A')}</td>
            </tr>
        </table>
        <hr>
        <h3 style="color: #041C2F;">Análise Estratégica da Semana</h3>
        <p><em>{narrativas_ia.get('analise_estrategica', 'N/A')}</em></p>
        <hr>
        <h3 style="color: #041C2F;">Alocação de Esforço Consolidada</h3>
        {html_alocacao}
        <hr>
        <h3 style="color: #041C2F;">Radar de Projetos</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 11pt;">
            <tr style="text-align: left; background-color: #f2f2f2;"><th style="padding: 8px;">Projeto</th><th style="padding: 8px;">Saúde</th><th style="padding: 8px; text-align:center;">Tendência</th><th style="padding: 8px;">Justificativa da IA</th></tr>
            {html_radar}
        </table>
        <hr>
        <h3 style="color: #041C2F;">Nuvem de Palavras Global</h3>
        <p>Os temas mais comentados pela equipe em todos os projetos esta semana:</p>
        <img src="cid:wordcloud" style="width:100%; max-width:600px;">
        <hr>
        <h3 style="color: #041C2F;">Destaques e Pessoas</h3>
        <h4>🏆 Profissional em Destaque</h4>
        <p><em>{destaques.get('destaque_profissional', 'N/A')}</em></p>
        <h4>⚠️ Ponto de Atenção (Equipe)</h4>
        <p><em>{destaques.get('ponto_atencao_equipe', 'N/A')}</em></p>
        <br>
        <p style="font-size: 9pt; color: #888;">Este e-mail foi gerado automaticamente pelo Robô Analista de Projetos.</p>
    </body></html>
    """
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))
    
    if caminho_nuvem_global and os.path.exists(caminho_nuvem_global):
        with open(caminho_nuvem_global, 'rb') as f:
            img = MIMEImage(f.read()); img.add_header('Content-ID', '<wordcloud>'); msg.attach(img)

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, config.GERAL_DESTINATARIOS, msg.as_string())
        server.quit()
        print(f" -> Boletim GERAL EXECUTIVO enviado com sucesso.")
    except Exception as e:
        print(f"ERRO ao enviar boletim GERAL: {e}")