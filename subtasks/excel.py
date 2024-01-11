import pandas as pd
import datetime, os, sys

def readExcel(maestro, strLabelLog, strFolderPath, strFileName, numIndexHeader):
    try:
        maestro.new_log_entry(activity_label = strLabelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "readExcel", "Description": "Iniciando Tarefa"})
        strFilePath = os.path.join(strFolderPath, strFileName)
        if os.path.exists(strFilePath):
            worksheet = pd.read_excel(strFilePath, header=0, skiprows=range(numIndexHeader))
        else:
            raise Exception("Planilha "+strFilePath+" nao encontrada")
        return worksheet
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = strLabelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "readExcel", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[readExcel] Erro "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))

