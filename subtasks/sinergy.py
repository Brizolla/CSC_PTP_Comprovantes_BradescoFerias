from botcity.web import By
from botcity.web.util import element_as_select
from datetime import datetime, timedelta
import os, sys, time

def get_days_of_previous_month():
    today = datetime.today()
    last_day_of_previous_month = today.replace(day=1) - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
    days_of_previous_month = []
    current_day = first_day_of_previous_month
    while current_day <= last_day_of_previous_month:
        days_of_previous_month.append(current_day.strftime("%d/%m/%Y"))
        current_day += timedelta(days=1)

    return days_of_previous_month

def loginSinergy(maestro, labelLog, bot, strURLLogin, strUsuario, strOperador, strSenha):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Iniciando tarefa"})
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Abrir URL "+ strURLLogin})
        bot.browse(strURLLogin)
        bot.maximize_window()

        # Usuario
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Inserir usuario"})
        username_field = bot.find_element(selector='UsuarioCheck', by=By.ID, waiting_time=60000)
        if username_field:
            username_field.click()
        else:
            raise Exception("Campo do usuario nao encontrado")
        username_field.send_keys(strUsuario)
        # Prosseguir
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Clicar em prosseguir"})
        bot.find_element(selector='btnLoginPasso1', by=By.ID).click()
        # Verificar botao login
        button_login = bot.find_element(selector='btnLoginPasso2', by=By.ID, waiting_time=60000)
        if button_login:
            perfilNum = -1
            boolExist = False
        else:
            raise Exception("Campo de senha nao encontrado")
        for count in range(5):
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Selecionar perfil tentativa "+str(count)})
            tabelaPerfil = bot.find_element("#perfilContainer > tbody:nth-child(1)", By.CSS_SELECTOR)
            for linha in tabelaPerfil.find_elements_by_tag_name("tr"):
                texto_perfil = linha.find_element_by_class_name("item-ltext").text
                perfilNum += 1
                if texto_perfil == strOperador:
                    perfil = "Perfil_"+str(perfilNum)
                    bot.execute_javascript(f"document.getElementById('{perfil}').click()")
                    boolExist = True
                    break
            if boolExist:
                break
            else:
                perfilNum = 0
        if not boolExist:
            raise Exception("Perfil nao encontrado")
        # Senha
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Colocar senha"})
        bot.wait_for_element_visibility(bot.find_element("Senha", By.ID), True, 60000)
        bot.find_element("Senha", By.ID).send_keys(strSenha)
        # Login
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Clicar em Login"})
        bot.wait_for_element_visibility(bot.find_element(selector='btnLoginPasso2', by=By.ID), True, 60000)
        bot.find_element(selector='btnLoginPasso2', by=By.ID).click()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "loginSinergy", "Description": "Tarefa finalizada"})
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "loginSinergy", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[loginSinergy] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))        

def selectCompany(maestro, labelLog, bot, strEmpresa):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "selectCompany", "Description": "Iniciando tarefa"})
        bot.find_element(selector='NomeEmpresa', by=By.ID, waiting_time=60000).send_keys(strEmpresa)
        boolFiltrar = False
        for count in range(10):
            try:
                bot.wait_for_element_visibility(bot.find_element(selector='btnFiltrar', by=By.ID), True, 60000)
                bot.wait(1000)
                bot.find_element(selector='btnFiltrar', by=By.ID).click()
                boolFiltrar = True
                break
            except Exception as e:
                print(str(e))
                if str(e) == "'NoneType' object has no attribute 'is_displayed'":
                    boolFiltrar = True
                    break
                else:
                    bot.wait(2000)
        if not boolFiltrar:
            raise Exception("Falha ao filtrar empresa")
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "selectCompany", "Description": "Tarefa finalizada"})
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "selectCompany", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[selectCompany] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))
        

def navMenu(maestro, labelLog, bot, strURLMenu, strTipoRelatorio):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "navMenu", "Description": "Iniciando tarefa"})
        bot.wait(10000)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "navMenu", "Description": "Acessar URL " + strURLMenu})
        bot.navigate_to(strURLMenu)
        bot.wait(2000)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "navMenu", "Description": "Selecionar opção " + strTipoRelatorio})
        bot.find_element('#tblRelatorio_filter > label > input', By.CSS_SELECTOR).send_keys(strTipoRelatorio)
        bot.wait(2000)
        bot.find_element('#tblRelatorio > tbody > tr:nth-child(1) > td:nth-child(2) > a > span', By.CSS_SELECTOR).click()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "navMenu", "Description": "Tarefa finalizada"})
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "navMenu", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[navMenu] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))
        

def extractReport(maestro, labelLog, bot, days_of_previous_month, strFiltro, strUrlHome, strNomeArquivo, strFolderPath):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Iniciando tarefa"})
        # Aguardar
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Aguardar pagina carregar"})
        while bot.find_element('.modal-loading', By.CSS_SELECTOR, ensure_visible=True):
            bot.wait(2000)
        # Limpar Filtro
        while bot.find_element('#tblFiltros > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1) > span:nth-child(1)', By.CSS_SELECTOR):
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Limpar filtro"})
            bot.find_element('#tblFiltros > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1) > span:nth-child(1)', By.CSS_SELECTOR).click()
        for data in days_of_previous_month:
            # Filtrar
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Selecionar "+strFiltro})
            filtroOptions = bot.find_element('#Filtro', By.CSS_SELECTOR)
            filtroOptions = element_as_select(filtroOptions)
            filtroOptions.select_by_visible_text(strFiltro)
            # Operador
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Selecionar Operador 'IGUAL'"})
            operadorOptions = bot.find_element('#OperadorTodos', By.CSS_SELECTOR)
            operadorOptions = element_as_select(operadorOptions)
            operadorOptions.select_by_value('IGUAL')
            # Condicional
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Selecionar Condicional 'OU'"})
            condicionalOptions = bot.find_element('#Condicional', By.CSS_SELECTOR)
            condicionalOptions = element_as_select(condicionalOptions)
            condicionalOptions.select_by_visible_text('OU')
            # Datas
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Informar data: "+data})
            bot.find_element('#Selecao_Texto', By.CSS_SELECTOR).send_keys(data)
            bot.find_element('#btnAdicionarFiltro', By.CSS_SELECTOR).click()
            
        # Gerar PDF
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Gerar PDF"})
        bot.find_element('#btnGerarPDFRel', By.CSS_SELECTOR).click()
        # Aguardar
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Aguardar processamento concluir"})
        while bot.find_element('.modal-loading', By.CSS_SELECTOR, ensure_visible=True):
            bot.wait(2000)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Processamento realizado"})
        # Voltar na navegacao
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Acessar página "+strUrlHome})
        bot.navigate_to(strUrlHome)
        # Pesquisar Recibo
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Pesquisar recibo "+strNomeArquivo})
        bot.find_element('input.form-control', By.CSS_SELECTOR, 30000).send_keys(strNomeArquivo)
        bot.wait(2000)
        strNomeCompleto = bot.find_element("tr.odd:nth-child(1) > td:nth-child(1)", By.CSS_SELECTOR).text
        strNomeCompleto = strNomeCompleto.replace('001 - Arquivo - ','')
        bot.find_element(selector='.odd > td:nth-child(3) > a:nth-child(1)', by=By.CSS_SELECTOR, waiting_time=30000).click()
        # Confirmar download
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Aguardar janela de confirmação"})
        dialog = bot.get_js_dialog()
        bot.wait(2000)
        dialog.accept()
        bot.wait(2000)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Aguardar download do arquivo"})
        strNomeArquivo = f'{strFolderPath}\{strNomeCompleto}'
        countArquivo = 0
        while not os.path.exists(strNomeArquivo):
            if countArquivo == 450:
                raise Exception('Tempo de espera excedido')
            else:
                countArquivo += 1
                time.sleep(2)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "extractReport", "Description": "Arquivo baixado: "+strNomeArquivo})
        return strNomeArquivo
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "extractReport", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[extractReport] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))
        

