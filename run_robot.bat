@echo off
rem Este comando garante que o script sempre execute a partir de sua própria pasta.
cd /d "%~dp0"

echo Ativando o ambiente virtual...
rem Ativa o ambiente virtual (venv).
call venv\Scripts\activate

echo Ambiente ativado. Executando o script Python...
rem Executa o seu robô. Todo o progresso aparecerá nesta janela.
python main.py

echo Processo finalizado.

rem Se você quiser que esta janela preta permaneça aberta no final para ler
rem todas as mensagens, remova o "rem" da linha abaixo.
rem pause