# -*- coding: utf-8 -*-
import os
import win32com.client as win32

def unprotect_and_save(filepath):
    """
    Abre uma instância do Excel para 'Habilitar Edição' de um arquivo
    e salvá-lo, removendo o Modo de Exibição Protegido.
    """
    # A biblioteca pywin32 requer um caminho absoluto.
    absolute_path = os.path.abspath(filepath)
    print(f"Desbloqueando o arquivo: {absolute_path}")

    excel = None
    wb = None
    try:
        excel = win32.Dispatch('Excel.Application')
        # Manter o Excel invisível durante o processo
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Abre o workbook
        wb = excel.Workbooks.Open(absolute_path)
        
        # Simplesmente salvar o arquivo e fechar já remove o Modo de Exibição Protegido.
        wb.Save()
        wb.Close()
        print("Arquivo salvo e desbloqueado com sucesso.")
        
    except Exception as e:
        print(f"Ocorreu um erro ao tentar desbloquear o arquivo Excel: {e}")
        # Garante que o processo do Excel seja encerrado em caso de erro
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        # Lança o erro novamente para que o main.py saiba que algo deu errado
        raise e
    finally:
        # Garante que o processo do Excel seja sempre encerrado
        if excel:
            excel.Quit()