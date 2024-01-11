#from O365 import Account
import datetime, sys
from botcity.plugins.ms365.credentials import MS365CredentialsPlugin, Scopes
from botcity.plugins.ms365.onedrive import MS365OneDrivePlugin

class Sharepoint():
	def __init__(self,iStrSiteHostname,iStrSiteName, iStrClientID, iStrClientKey, iStrTokenPath):
		self.service = MS365CredentialsPlugin(
            client_id=iStrClientID,
            client_secret=iStrClientKey,
            token_path=iStrTokenPath
		)
		self.service.authenticate(scopes=[Scopes.BASIC, Scopes.FILES_READ_WRITE_ALL, Scopes.SITES_READ_WRITE_ALL])
		self.sharepoint = MS365OneDrivePlugin(
    	    service_account=self.service,
    	    use_sharepoint=True,
    	    host_name=iStrSiteHostname,
    	    path_to_site=iStrSiteName
		)
		self.defaultDrive = self.sharepoint.default_drive
		
	def getData(self, iStrVacationDate):
		oStrMonth = iStrVacationDate[3:5]
		oStrYear = iStrVacationDate[6:10]
		return oStrYear, oStrMonth
	
	def verifyItem(self, maestro, labelLog, iStrCaminhoItem):
		maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "verifyItem", "Description": "Iniciando tarefa"})
		maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "verifyItem", "Description": "Verificar se item existe: "+iStrCaminhoItem})
		try:
			vItem = self.defaultDrive.get_item_by_path(iStrCaminhoItem)
			oBoolExist = True
		except Exception as e:
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "WARN": "verifyItem", "Description": str(e)})
			oBoolExist = False
		maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "verifyItem", "Description": "Tarefa finalizada"})
		return oBoolExist

	def getFolderName(self, maestro, labelLog, iStrRootFolder,iStrVacationDate, iStrCompanyRoot, iStrFolder):
		try:
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "getFolderName", "Description": "Iniciando tarefa"})
			oFolderName = ''
			vStrYear, vStrMonth = self.getData(iStrVacationDate)
			iStrRootFolder = f'{iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyRoot}'
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "getFolderName", "Description": "Verificar pastas filho da pasta: "+iStrRootFolder})
			try:
				vItens = self.defaultDrive.get_item_by_path(iStrRootFolder).get_child_folders()
			except:
				raise Exception(f"Pasta {iStrRootFolder} não encontrado.")
			for item in vItens:
				if str(item).find(iStrFolder) != -1:
					oFolderName = str(item).replace('Folder: ','')
			if oFolderName == '':
				raise Exception(f'Pasta com o nome {iStrFolder} nao foi encontrada no caminho {iStrRootFolder}')
			else:
				maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "getFolderName", "Description": "Tarefa finalizada"})
				return oFolderName
		except Exception as e:
			exc_type, exc_obj, exc_tb = sys.exc_info()
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "getFolderName", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
			raise Exception("[getFolderName] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))	

	def downloadFile(self, maestro, labelLog, iStrRootFolder, iStrVacationDate, iStrCompanyPath, iStrVacationFolder, iStrFileName, iStrPathDownload):
		try:
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "downloadFile", "Description": "Iniciando tarefa"})
			vStrYear, vStrMonth = self.getData(iStrVacationDate)
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "downloadFile", "Description": "Verificar se arquivo existe: "+f"{iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyPath}/{iStrVacationFolder}/{iStrFileName}"})
			try:
				vFile = self.defaultDrive.get_item_by_path(f"{iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyPath}/{iStrVacationFolder}/{iStrFileName}")
			except:
				raise Exception(f"Arquivo {iStrFileName} não encontrado no diretorio {iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyPath}/{iStrVacationFolder}")
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "downloadFile", "Description": "Baixar arquivo"})
			vFile.download(iStrPathDownload)
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "downloadFile", "Description": "Tarefa finalizada"})
		except Exception as e:
			exc_type, exc_obj, exc_tb = sys.exc_info()
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "downloadFile", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
			raise Exception("[downloadFile] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))

	def uploadFile(self, maestro, labelLog, iStrRootFolder, iStrVacationDate, iStrCompanyGroup, iStrCompanyName, iStrVacationFolder, iStrFile):
		vListPathSharepoint = []
		vStrFolderVerify = ""
		try:
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "uploadFile", "Description": "Iniciando tarefa"})
			vStrYear, vStrMonth = self.getData(iStrVacationDate)
			try:
				maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "uploadFile", "Description": "Verificar se pasta existe: "+f"{iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyGroup}/{iStrCompanyName}/{iStrVacationFolder}"})
				vFolderName = self.defaultDrive.get_item_by_path(f"{iStrRootFolder}/{vStrYear}/{vStrMonth}/{iStrCompanyGroup}/{iStrCompanyName}/{iStrVacationFolder}")
			except:
				maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "uploadFile", "Description": "Pasta nao existe. Criando..."})
				vListPathSharepoint = [iStrRootFolder,vStrYear,vStrMonth,iStrCompanyGroup, iStrCompanyName,iStrVacationFolder]
				for vStrFolderSharepoint in vListPathSharepoint:
					vStrFolderVerify = vStrFolderVerify + vStrFolderSharepoint
					try:
						vFolderName = self.defaultDrive.get_item_by_path(vStrFolderVerify)
					except:
						vFolderName = vFolderName.create_child_folder(vStrFolderSharepoint)
					finally:
						vStrFolderVerify = vStrFolderVerify + "/"
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "uploadFile", "Description": "Upload do arquivo "+iStrFile})
			vFolderName.upload_file(iStrFile)
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "uploadFile", "Description": "Tarefa finalizada"})
		except Exception as e:
			exc_type, exc_obj, exc_tb = sys.exc_info()
			maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "validateDatas", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
			raise Exception("[uploadFile] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))

