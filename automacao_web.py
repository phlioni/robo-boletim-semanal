# -*- coding: utf-8 -*-
import os
import time
import shutil
# A biblioteca de cálculo de datas não é mais necessária aqui
from selenium import webdriver
from selenium.webdriver.common.by import By
# A importação de 'Keys' não é mais necessária
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config

def login_e_download_semanal():
    """Baixa o relatório de horas clicando na opção 'Semana passada'."""
    options = webdriver.ChromeOptions()
    #options.add_argument("--headless") # Roda em modo invisível por padrão
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    
    try:
        wait = WebDriverWait(driver, 20)
        print("Acessando o site...")
        driver.get(config.SITE_URL)

        try:
            print("Realizando login...")
            wait.until(EC.presence_of_element_located((By.NAME, "LoginName"))).send_keys(config.SITE_LOGIN)
            driver.find_element(By.NAME, "Password").send_keys(config.SITE_SENHA)
            driver.find_element(By.ID, "button_processLogin").click()
        except TimeoutException:
            print("Campos de login não encontrados. Presumindo que já está logado.")
        
        print("Navegando para 'Resumo de Horas por Profissional'...")
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Resumo de Horas por Profissional"))).click()

        print("Aguardando a página de filtros carregar...")
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        # --- LÓGICA DE DATA TOTALMENTE SIMPLIFICADA ---
        print("Aplicando filtro para 'Semana passada'...")
        
        # 1. Abre o seletor de data
        # Usamos 'P_DATA_show' pois é o campo visível onde o usuário clica
        wait.until(EC.element_to_be_clickable((By.ID, "P_DATA_show"))).click()
        
        # 2. Clica diretamente na opção "Semana passada"
        wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Semana passada']"))).click()
        
        time.sleep(1) # Pausa para o filtro aplicar
        # --- FIM DA NOVA LÓGICA ---

        print("Executando relatório...")
        driver.find_element(By.ID, "button_Execute").click()

        print("Aguardando o carregamento do relatório...")
        time.sleep(10)
        
        arquivos_antes = set(os.listdir(config.PASTA_DOWNLOADS))
        print("Exportando para Excel...")
        driver.find_element(By.ID, "button_ExecuteXSL").click()
        print("Aguardando 10 segundos para o início do download...")
        time.sleep(10)

        print(f"Monitorando a pasta {config.PASTA_DOWNLOADS} por um novo arquivo...")
        start_time = time.time()
        caminho_arquivo_baixado = None
        
        while time.time() - start_time < 60:
            arquivos_depois = set(os.listdir(config.PASTA_DOWNLOADS))
            novos_arquivos = arquivos_depois - arquivos_antes
            if novos_arquivos and not list(novos_arquivos)[0].endswith('.crdownload'):
                nome_do_arquivo = novos_arquivos.pop()
                caminho_arquivo_baixado = os.path.join(config.PASTA_DOWNLOADS, nome_do_arquivo)
                print(f"Novo arquivo detectado: {nome_do_arquivo}")
                break
            time.sleep(1)

        if not caminho_arquivo_baixado:
            raise TimeoutException("O download do arquivo demorou mais de 60 segundos ou não foi detectado.")
        
        time.sleep(2)
        print("Download concluído.")
        
        caminho_destino = os.path.join(os.getcwd(), os.path.basename(caminho_arquivo_baixado))
        shutil.move(caminho_arquivo_baixado, caminho_destino)
        print(f"Arquivo movido para a pasta do projeto: {caminho_destino}")
        
        return caminho_destino

    finally:
        if driver:
            for handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.close()