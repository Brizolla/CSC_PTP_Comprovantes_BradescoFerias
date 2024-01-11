from botcity.web import WebBot, Browser
from botcity.maestro import *
from subtasks import config, onedrive, sinergy, teste, pdf
import datetime, os, sys, shutil, glob

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    boolExcptGlobal = True
    # Definir ambiente de exec: "DEV", "QA"ou "PRD"
    if execution.parameters:
        strEnvironment = execution.parameters.get("ambiente", "PRD")
        strConfigPath = execution.parameters.get("caminhoConfig", r"E:\TH\GESTÃO DE DOCUMENTOS\AA161. Separação de comprovantes Bradesco férias\Config\config.xml")
    else:
        maestro.login("https://cscalgar.botcity.dev", "cscalgar", "CSC_V97ZNSLVXX8VH6PR5MNG") # São os mesmos dados do arquivo conf.bcf
        strEnvironment = "DEV"
        strConfigPath = r"E:\TH\GESTÃO DE DOCUMENTOS\AA161. Separação de comprovantes Bradesco férias\Config\config.xml"
    # Declarar variaveis
    dictConfig = {}
    meses = {'01':'JANEIRO','02':'FEVEREIRO','03':'MARÇO','04':'ABRIL','05':'MAIO','06':'JUNHO','07':'JULHO','08':'AGOSTO','09':'SETEMBRO','10':'OUTUBRO','11':'NOVEMBRO','12':'DEZEMBRO'}
    try:
        # *** Inicio ler arquivo de configuracao ***
        dictConfig  = config.getConfig(strEnvironment, strConfigPath)
        # *** Fim ler arquivo de configuracao ***
        maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "Main", "Description": "Iniciando o bot"})
        # *** Inicio capturar credenciais ***
        credUserSinergy = maestro.get_credential(label=dictConfig["credSinergy"], key="username")
        credPasswordSinergy = maestro.get_credential(label=dictConfig["credSinergy"], key="password")
        credClientID = maestro.get_credential(label=dictConfig["credOnedrive"], key="clientID")
        credClientKey = maestro.get_credential(label=dictConfig["credOnedrive"], key="clientKey")
        # *** Fim capturar credenciais ***
        # *** Inicio encerrar aplicacoes ***
        listApplications = ['EXCEL.EXE', 'firefox.exe']
        config.closeApps(maestro, listApplications, dictConfig["labelLog"])
        # *** Fim encerrar aplicacoes ***
        # *** Inicio imprimir arquivos ***
        listEmpresas = []
        listEmpresas = dictConfig['listEmpresas'].split(sep=', ')
        listTipoRecibos = []
        listTipoRecibos = dictConfig['listTipoRecibos'].split(sep=', ')
        strRecibo = []
        # Chamada da função e obtenção da lista de dias do mês anterior
        listDiasMesAnterior = sinergy.get_days_of_previous_month()
        
        for strEmpresa in listEmpresas:
            # Obter caminho de execucao desse script
            strCaminhoScript = os.path.abspath(sys.argv[0])
            strCaminhoScript = strCaminhoScript.replace(r'\bot.py','')
            for strItemTipoRecibo in listTipoRecibos:
                for numRetry in range(2):
                    try:
                        strRecibo = strItemTipoRecibo.split(sep='/')
                        strNovoNomeArquivo = strEmpresa[6:]+'_'+strRecibo[1]+'.pdf'
                        strCaminhoCompletoNovo = f'{strCaminhoScript}\{strNovoNomeArquivo}'
                        vAno, vMes = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").getData(listDiasMesAnterior[0])
                        vBoolConsolidado = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").verifyItem(maestro=maestro, labelLog=dictConfig["labelLog"], iStrCaminhoItem=f"{dictConfig['rootFolder']}\{vAno}\{vMes}\{dictConfig['companyGroupRoot']}\{dictConfig['relFolder']}\{dictConfig['vacationFolder']}\{strNovoNomeArquivo}")
                        if not vBoolConsolidado:
                            bot = WebBot()
                            # Configure whether or not to run on headless mode
                            bot.headless = False
                            maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "Main", "Description": "Instanciar firefox"})
                            bot.browser = Browser.FIREFOX
                            print(strCaminhoScript+dictConfig["resourcesFolder"]+dictConfig["driverName"])
                            bot.driver_path = strCaminhoScript+dictConfig["resourcesFolder"]+dictConfig["driverName"]
                            # Deletar arquivo renomeado caso ja exista na pasta script
                            if os.path.exists(strCaminhoCompletoNovo):
                                os.remove(strCaminhoCompletoNovo)
                            # Inicio login sinergy
                            sinergy.loginSinergy(maestro=maestro, labelLog=dictConfig["labelLog"], bot=bot, strURLLogin=dictConfig["urlLogin"], strUsuario=credUserSinergy, strOperador=dictConfig["operadorSinergy"], strSenha=credPasswordSinergy)
                            # Fim login sinergy
                            # Inicio selecionar empresa
                            sinergy.selectCompany(maestro=maestro, labelLog=dictConfig["labelLog"], bot=bot, strEmpresa=strEmpresa)
                            # Fim selecionar empresa
                            # Inicio navegar menu
                            sinergy.navMenu(maestro=maestro, labelLog=dictConfig["labelLog"], bot=bot, strURLMenu=dictConfig["urlRelatorio"], strTipoRelatorio=strRecibo[0])
                            # Fim navegar menu
                            # Inicio baixar relatorios
                            strCaminhoArquivo = sinergy.extractReport(maestro=maestro, labelLog=dictConfig["labelLog"], bot=bot, days_of_previous_month=listDiasMesAnterior, strFiltro=dictConfig["tipoFiltro"], strUrlHome=dictConfig["urlLogin"], strNomeArquivo=strRecibo[1], strFolderPath=strCaminhoScript)
                            # Fim baixar relatorios
                            bot.wait(3000) 
                            #bot.close_page()
                            try: bot.stop_browser()
                            except: pass
                            os.rename(strCaminhoArquivo, strCaminhoCompletoNovo)
                            # Inicio subir no sharepoint
                            onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").uploadFile(maestro, dictConfig["labelLog"], dictConfig['rootFolder'], listDiasMesAnterior[0], dictConfig['companyGroupRoot'], dictConfig['relFolder'], dictConfig['vacationFolder'], strCaminhoCompletoNovo)
                            # Fim subir no sharepoint
                            # Mover para Downloads
                            strCaminhoArquivoDownload = f'{dictConfig["logicPath"]}\{strNovoNomeArquivo}'
                            shutil.move(strCaminhoCompletoNovo, strCaminhoArquivoDownload)
                            # Aguardar upload do arquivo
                            for count in range(240):
                                vBoolConsolidado = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").verifyItem(maestro=maestro, labelLog=dictConfig["labelLog"], iStrCaminhoItem=f"{dictConfig['rootFolder']}\{vAno}\{vMes}\{dictConfig['companyGroupRoot']}\{dictConfig['relFolder']}\{dictConfig['vacationFolder']}\{strNovoNomeArquivo}")
                                if vBoolConsolidado:
                                    os.remove(strCaminhoCompletoNovo)
                                    break
                                else:
                                    bot.wait(1000)
                        else:
                            # Download arquivo
                            onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").downloadFile(maestro=maestro, labelLog=dictConfig["labelLog"], iStrRootFolder=dictConfig['rootFolder'], iStrVacationDate=listDiasMesAnterior[0], iStrCompanyPath=dictConfig['companyGroupRoot'], iStrVacationFolder=f"{dictConfig['relFolder']}\{dictConfig['vacationFolder']}", iStrFileName=strNovoNomeArquivo, iStrPathDownload=dictConfig["logicPath"])
                        # Parar retry
                        break
                    except Exception as erroSinergy:
                        exc_type, exc_obj, exc_tb  = sys.exc_info()
                        if numRetry == 2:
                            raise Exception("Numero de tentativas excedida. Mensagem de erro: " + str(erroSinergy) + " Na linha: " + str(exc_tb.tb_lineno))
                        else:
                            bot.wait(2000)
                            print(str(erroSinergy))
                            maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "WARN", "Task": "Main", "Description": "Mensagem de erro: " + str(erroSinergy) + " Na linha: " + str(exc_tb.tb_lineno)})
                            maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "WARN", "Task": "Main", "Description": "Retentando..."})
                    finally:
                        config.closeApps(maestro, listApplications, dictConfig["labelLog"])
        # *** Fim imprimir arquivos ***
        
        # *** Inicio ler planilha ***
        dfInput = teste.readExcel(maestro=maestro, strLabelLog=dictConfig["labelLog"], strFolderPath=dictConfig["inputPath"], strFileName=dictConfig["inputFileName"], numIndexHeader=2)
        # *** Fim ler planilha ***
        for numIndexInput, dflinhaInput in dfInput.iterrows():
            # Verificar se esta no escopo.
            if dflinhaInput['COD EMPRESA'] == 29 or dflinhaInput['COD EMPRESA'] == 33:
                print('entrou')
                for strItemTipoRecibo in listTipoRecibos:
                    strRecibo = strItemTipoRecibo.split(sep='/')
                    strCodEmpresa = str(dflinhaInput['COD EMPRESA']).strip()
                    strCodEmpresa = strCodEmpresa.replace('.0','')
                    strNovoNomeArquivo = strCodEmpresa+'_'+strRecibo[1]+'.pdf'
                    strCaminhoArquivoDownload = f'{dictConfig["logicPath"]}\{strNovoNomeArquivo}'
                    if not os.path.exists(strCaminhoArquivoDownload):
                        raise Exception(f'Arquivo {strCaminhoArquivoDownload} nao baixado')
                try:
                    # *** Inicio baixar planilha Mapa ***
                    if dflinhaInput['COD EMPRESA'] == 29:
                        strEmpresaArquivoPgto = 'TECH'
                        strCompanyPath = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").getFolderName(maestro, dictConfig["labelLog"], dictConfig['rootFolder'], listDiasMesAnterior[0], dictConfig['companyGroupRoot'], 'ALGAR TECH')
                    else:
                        strEmpresaArquivoPgto = 'TI'
                        strCompanyPath = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").getFolderName(maestro, dictConfig["labelLog"], dictConfig['rootFolder'], listDiasMesAnterior[0], dictConfig['companyGroupRoot'], 'ALGAR TI')
                    strEspecificaName = str(dflinhaInput['NOME CLIENTE/TOMADOR']).strip() + '.xlsx'
                    onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").downloadFile(maestro, dictConfig["labelLog"], dictConfig['rootFolder'], listDiasMesAnterior[0], dictConfig['companyGroupRoot']+'/'+strCompanyPath, str(dflinhaInput['NOME CLIENTE/TOMADOR']).strip(), strEspecificaName, dictConfig["logicPath"])
                    # *** Fim baixar planilha Mapa ***
                    dfEspecifica = teste.readExcel(maestro=maestro, strLabelLog=dictConfig["labelLog"], strFolderPath=dictConfig["logicPath"], strFileName=strEspecificaName, numIndexHeader=0)
                    for numIndexEspecifica, dflinhaEspecifica in dfEspecifica.iterrows():
                        try:
                            # Filtrar coluna M por diferente de vazio e coluna Q por bradesco
                            if (str(dflinhaEspecifica['INICIO_FERIAS']).strip() != '' and str(dflinhaEspecifica['INICIO_FERIAS']).strip() != 'nan') and str(dflinhaEspecifica['DESC_BANCO']).strip().upper().find("BRADESCO") != -1:
                                # Verificar se arquivo de pgto existe
                                vAno, vMes = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").getData(listDiasMesAnterior[0])
                                # Verificar se arquivo consolidado ja existe
                                vBoolConsolidado = onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").verifyItem(maestro=maestro, labelLog=dictConfig["labelLog"], iStrCaminhoItem=f"{dictConfig['rootFolder']}\{vAno}\{vMes}\{dictConfig['companyGroupRoot']}\{strCompanyPath}\{str(dflinhaInput['NOME CLIENTE/TOMADOR']).strip()}\{dictConfig['comprovanteFolder']}\{str(dflinhaEspecifica['NOME_ASSOCIADO']).strip()}.pdf")
                                if vBoolConsolidado:
                                    continue
                                    
                                vStrDataPgto = str(dflinhaEspecifica['DATA_PG_FERIAS'])
                                maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "WARN", "Task": "Main", "Description": "Data de pagamento: "+vStrDataPgto})
                                vDiaAtual = vStrDataPgto[:2]
                                vMesAtual = f'{vStrDataPgto[3:5]} - {meses[vStrDataPgto[3:5]]}'
                                vAnoAtual = vStrDataPgto[6:10]
                                vBoolEncontrouPg = False
                                listCaminhoarquivos = []
                                os.chdir(f'{dictConfig["paymentPath"]}\{vAnoAtual}\{vMesAtual}\{vDiaAtual}')
                                for arquivosPgto in glob.glob("*.pdf"):
                                    if f'BRADESCO {strEmpresaArquivoPgto}' in arquivosPgto:
                                        # Pegar comprovante do funcionario
                                        try:
                                            strCaminhoArquivo = pdf.savePDF(maestro=maestro, labelLog=dictConfig["labelLog"], iStrCaminhoPDF=f'{dictConfig["paymentPath"]}\{vAnoAtual}\{vMesAtual}\{vDiaAtual}\{arquivosPgto}', iStrNomeFuncionario=str(dflinhaEspecifica['CPF']).strip(), iStrIdentificacao='CPF:', iStrTipoDoc='ComprovantedePagamento', iStrPastaDownload=dictConfig["logicPath"])
                                            vBoolEncontrouPg = True
                                            listCaminhoarquivos.append(strCaminhoArquivo)
                                            break
                                        except Exception as erroPagamento:
                                            pass
                                if not vBoolEncontrouPg:
                                    raise Exception('Nao foi encontrado para o comprovante de pagamento')
                                for strItemTipoRecibo in listTipoRecibos:
                                    strRecibo = strItemTipoRecibo.split(sep='/')
                                    strNovoNomeArquivo = strCodEmpresa+'_'+strRecibo[1]+'.pdf'
                                    strCaminhoArquivoDownload = f'{dictConfig["logicPath"]}\{strNovoNomeArquivo}'
                                    strCaminhoArquivo = pdf.savePDF(maestro, dictConfig["labelLog"], strCaminhoArquivoDownload, str(dflinhaEspecifica['NOME_ASSOCIADO']).strip(), strRecibo[2], strRecibo[1], dictConfig["logicPath"])
                                    listCaminhoarquivos.append(strCaminhoArquivo)
                                pdf.mergePDF(maestro=maestro, labelLog=dictConfig["labelLog"], listaArquivos=listCaminhoarquivos, iStrNomeArquivoConsolidado=f'{dictConfig["logicPath"]}\{str(dflinhaEspecifica["NOME_ASSOCIADO"]).strip()}.pdf')
                                for remover in listCaminhoarquivos:
                                    os.remove(remover)
                                onedrive.Sharepoint(dictConfig["siteHostname"],dictConfig["siteName"], credClientID, credClientKey, f"{dictConfig['configPath']}\{dictConfig['authPath']}").uploadFile(maestro=maestro, labelLog=dictConfig["labelLog"], iStrRootFolder=dictConfig['rootFolder'], iStrVacationDate=listDiasMesAnterior[0], iStrCompanyGroup=dictConfig['companyGroupRoot']+'/'+strCompanyPath, iStrCompanyName=str(dflinhaInput['NOME CLIENTE/TOMADOR']).strip(), iStrVacationFolder=dictConfig['comprovanteFolder'], iStrFile=f'{dictConfig["logicPath"]}\{str(dflinhaEspecifica["NOME_ASSOCIADO"]).strip()}.pdf')
                                os.remove(f'{dictConfig["logicPath"]}\{str(dflinhaEspecifica["NOME_ASSOCIADO"]).strip()}.pdf')
                                return
                        except Exception as erroChild:
                            exc_type, exc_obj, exc_tb  = sys.exc_info()
                            maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "Main", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(erroChild) + " Linha: " + str(exc_tb.tb_lineno)})
                            print(str(erroChild)+ " Linha: " + str(exc_tb.tb_lineno))
                    # Deletar especifica
                    os.remove(f'{dictConfig["logicPath"]}\{strEspecificaName}')
                except Exception as erroParent:
                    exc_type, exc_obj, exc_tb  = sys.exc_info()
                    maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "Main", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(erroParent) + " Linha: " + str(exc_tb.tb_lineno)})
                    print(str(erroParent)+ " Linha: " + str(exc_tb.tb_lineno))

    except Exception as erroPrincipal:
        erroFinal = ""
        exc_type, exc_obj, exc_tb  = sys.exc_info()
        maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "Main", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(erroPrincipal) + " Linha: " + str(exc_tb.tb_lineno)})
        erroFinal = str(erroPrincipal)
        boolExcptGlobal = False
        print(str(exc_type) + " Mensagem: " + str(erroPrincipal) + " Linha: " + str(exc_tb.tb_lineno))
    finally:
        maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "Main", "Description": "Limpar pasta"})
        os.chdir(f'{dictConfig["logicPath"]}')
        for arquivosLogic in glob.glob("*.*"):
            maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "Main", "Description": f"Excluir arquivo: {arquivosLogic}"})
            os.remove(arquivosLogic)
        maestro.new_log_entry(activity_label = dictConfig["labelLog"],values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "Main", "Description": "BOT FINALIZADO"})
        if execution.parameters:
            if boolExcptGlobal:
                # Uncomment to mark this task as finished on BotMaestro
                maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.SUCCESS,
                    message="Task Finished OK."
                )
            else:
                maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.FAILED,
                    message="Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(erroFinal) + " Linha: " + str(exc_tb.tb_lineno)
                )




def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
