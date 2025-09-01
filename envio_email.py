# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import os

import config

def enviar_boletim(destinatario, nome_projeto, kpis_projeto, analise_ia_projeto):
    """Envia o boletim data-driven e com resumo da IA para líderes de projeto."""
    print(f"Preparando boletim para {destinatario} sobre o projeto {nome_projeto}...")
    msg = MIMEMultipart('related')
    msg['Subject'] = f"Boletim Tático Semanal - Projeto {nome_projeto}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = destinatario
    
    # Extrai as novas informações da análise da IA
    resumo_ia = analise_ia_projeto.get('resumo_tatico', 'N/A')
    conquistas_ia = "".join([f"<li>{item}</li>" for item in analise_ia_projeto.get('conquistas', [])])
    riscos_ia = "".join([f"<li>{item}</li>" for item in analise_ia_projeto.get('riscos_e_atencao', [])])

    corpo_html = f"""
    <html><body style="font-family: Calibri, sans-serif; font-size: 12pt;">
        <h2 style="color: #041C2F;">Boletim Tático Semanal: {nome_projeto}</h2>
        <p><i>{resumo_ia}</i></p>
        
        <h3 style="color: #041C2F;">Métricas Chave da Semana</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 11pt; text-align: left;">
            <tr style="background-color: #f2f2f2;">
                <th style="padding: 8px;">Horas Totais</th>
                <th style="padding: 8px;">% em Correção de Bugs</th>
                <th style="padding: 8px;">Principal Contribuidor</th>
            </tr>
            <tr>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{kpis_projeto.get('total_horas', 'N/A')}</td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{kpis_projeto.get('percentual_bugs', 'N/A')}</td>
                <td style="padding: 8px; border-bottom: 1px solid #ddd;">{kpis_projeto.get('principal_contribuidor', 'N/A')}</td>
            </tr>
        </table>
        <br>

        <h3 style="color: #28a745;">✅ Conquistas e Entregas</h3>
        <ul>{conquistas_ia or "<li>Nenhuma conquista específica destacada pela IA.</li>"}</ul>

        <h3 style="color: #dc3545;">⚠️ Riscos e Pontos de Atenção</h3>
        <ul>{riscos_ia or "<li>Nenhum risco ou ponto de atenção identificado pela IA.</li>"}</ul>

        <br>
        <p style="font-size: 9pt; color: #888;">Este e-mail foi gerado automaticamente pelo Robô Analista de Projetos.</p>
    </body></html>
    """
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))
    
    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, destinatario, msg.as_string())
        server.quit()
        print(f" -> Boletim para {nome_projeto} enviado com sucesso para {destinatario}.")
    except Exception as e:
        print(f"ERRO ao enviar boletim para {destinatario}: {e}")

def _gerar_html_kpis_avancados(kpis_avancados):
    if not kpis_avancados: return ""
    html_gerado = ""
    portfolio = kpis_avancados.get('portfolio_clientes', [])
    if portfolio:
        html_gerado += "<h3 style='color: #041C2F;'>Visão por Portfólio de Cliente</h3><table style='width: 100%; border-collapse: collapse; font-size: 11pt;'><thead><tr style='text-align: left; background-color: #f2f2f2;'><th style='padding: 8px;'>Cliente</th><th>Horas Totais</th><th>% Dev</th><th>% Bug</th><th>% Reunião</th><th>Saúde</th></tr></thead><tbody>"
        for item in portfolio:
            html_gerado += f"<tr><td style='padding: 8px; border-bottom: 1px solid #ddd;'><strong>{item['cliente']}</strong></td><td>{item['total_horas']}</td><td>{item['pct_dev']}</td><td>{item['pct_bug']}</td><td>{item['pct_reuniao']}</td><td>{item['saude']}</td></tr>"
        html_gerado += "</tbody></table>"
    foco = kpis_avancados.get('foco_equipe', {})
    fragmentados = foco.get('fragmentados', [])
    dependentes = foco.get('dependencia_critica', [])
    if fragmentados or dependentes:
        html_gerado += "<hr><h3 style='color: #041C2F;'>Mapeamento de Foco e Risco Humano</h3>"
        if fragmentados:
            html_gerado += "<h4>⚠️ Profissionais Fragmentados (&gt;3 projetos)</h4><ul>"
            for item in fragmentados:
                html_gerado += f"<li>{item['profissional']} ({item['num_projetos']} projetos)</li>"
            html_gerado += "</ul><p style='font-size:10pt; color:#888;'><em>Risco: Perda de eficiência por troca de contexto.</em></p>"
        if dependentes:
            html_gerado += "<h4>🔴 Projetos com Dependência Crítica (&gt;70% em 1 prof.)</h4><ul>"
            for item in dependentes:
                html_gerado += f"<li>{item['projeto']} (<strong>{item['concentracao']}</strong> com {item['profissional']})</li>"
            html_gerado += "</ul><p style='font-size:10pt; color:#888;'><em>Risco: Gargalo e ponto único de falha.</em></p>"
    return html_gerado

def enviar_boletim_geral(dados_fatos, narrativas_ia, kpis_avancados):
    if not dados_fatos or not narrativas_ia:
        print("AVISO: Faltam dados ou análise da IA para gerar o boletim geral. Pulando envio.")
        return
        
    print("\nPreparando boletim GERAL EXECUTIVO para a gestão...")
    msg = MIMEMultipart('related')
    msg['Subject'] = f"Boletim Executivo Semanal de Projetos | {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = ", ".join(config.GERAL_DESTINATARIOS)
    kpis = dados_fatos.get('kpis', {})
    alocacao = dados_fatos.get('alocacao_esforco', [])
    analise_estrategica = narrativas_ia.get('analise_estrategica', 'N/A')
    destaques = narrativas_ia.get('destaques_pessoas', {})
    ponto_critico = narrativas_ia.get('kpis', {}).get('ponto_critico', 'N/A')
    html_alocacao = ""
    cores_alocacao = {'Desenvolvimento de Feature': '#4CAF50', 'Correção de Bug': '#f44336', 'Reunião': '#2196F3', 'Suporte/Atendimento': '#FFC107', 'Análise/Planejamento': '#00BCD4', 'Testes/QA': '#9C27B0', 'Documentação': '#607D8B', 'Outros': '#888'}
    for item in alocacao:
        categoria = item.get('categoria', 'Outros')
        percent_str = str(item.get('percentual', '0%'))
        cor = cores_alocacao.get(categoria, '#888')
        html_alocacao += f"<p style='margin: 5px 0; font-size: 11pt;'><strong>{categoria}:</strong> {item.get('horas')}h ({percent_str})<br><div style='background-color: {cor}; width: {percent_str}; height: 20px; border-radius: 5px;'></div></p>"
    
    html_kpis_avancados = _gerar_html_kpis_avancados(kpis_avancados)
    corpo_html = f"""
    <html><body style="font-family: Calibri, sans-serif; font-size: 12pt; color: #333;">
        <h2 style="color: #041C2F;">Boletim Executivo Semanal de Projetos</h2>
        <table style="width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 20px; font-size:11pt;">
            <tr style="background-color: #f2f2f2;">
        <td style="padding: 10px; width: 30%;">📈 <strong>{kpis.get('total_horas', 0):.1f}</strong> Horas Apontadas na Semana</td>
        <td style="padding: 10px; color:red; text-align: left;">🔴 <strong>Ponto Crítico Principal:</strong> {ponto_critico}</td>
            </tr>
        </table>
        <hr>
        <h3 style="color: #041C2F;">Análise Estratégica da Semana</h3>
        <p><em>{analise_estrategica}</em></p>
        <hr>
        <h3 style="color: #041C2F;">Alocação de Esforço Consolidada</h3>
        {html_alocacao}
        <hr>
        {html_kpis_avancados}
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
    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, config.GERAL_DESTINATARIOS, msg.as_string())
        server.quit()
        print(f" -> Boletim GERAL EXECUTIVO enviado com sucesso.")
    except Exception as e:
        print(f"ERRO ao enviar boletim GERAL: {e}")