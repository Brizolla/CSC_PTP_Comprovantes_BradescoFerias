import xml.etree.ElementTree as ET
import os
import datetime

def getConfig(strEnviroment, strConfigPath):
    dictConfig = {}
    nos = [strEnviroment, "COMMON"]
    # open xml file
    try:
        root = ET.parse(strConfigPath).getroot()
    except:
        raise Exception("XML nao foi aberto, verifique se ele esta correto ou no caminho correto")
    # get root nodes
    for no in nos:
        # get child nodes from root nodes
        for child in root.findall("./enviroment[@name='"+no+"']/"):
            tag = child.tag
            # get text
            for grandchild in root.findall("./enviroment[@name='"+no+"']"+tag+"/"):
                # insert in dictonary
                dictConfig[grandchild.tag] = grandchild.text
    return dictConfig

def closeApps(maestro, listApplications, labelLog):
    maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "closeApps", "Description": "Iniciando tarefa"})
    for app in listApplications:
        os.system("taskkill /f /im  "+app)
    maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "closeApps", "Description": "Tarefa finalizada"})
