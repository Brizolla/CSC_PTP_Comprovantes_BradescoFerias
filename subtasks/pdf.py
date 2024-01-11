import PyPDF2, sys, datetime

def savePDF(maestro, labelLog, iStrCaminhoPDF, iStrNomeFuncionario, iStrIdentificacao, iStrTipoDoc, iStrPastaDownload):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Iniciando tarefa"})
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Abrir PDF " + iStrCaminhoPDF})
        vBoolEncontrado = False
        with open(iStrCaminhoPDF, 'rb') as vObjArquivoPDF:
            # Criar um objeto PDF Reader
            vObjLeitorPDF = PyPDF2.PdfReader(vObjArquivoPDF)
            nNumPaginasPDF = len(vObjLeitorPDF.pages)
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Quantidade de paginas no PDF: "+str(nNumPaginasPDF)})
            # Percorrer todas as páginas do arquivo PDF
            for vNumPagina in range(nNumPaginasPDF):
                # Obter o texto da página
                vObjPagina = vObjLeitorPDF.pages[vNumPagina]
                vStrTextoPagina = vObjPagina.extract_text()
                # Verificar se a string está presente na página
                if iStrNomeFuncionario in vStrTextoPagina:
                    maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "O funcionario foi encontrado na pagina: "+str(vNumPagina+1)})
                    vBoolEncontrado = True
                    break
            if vBoolEncontrado:
                # Criar um objeto PDF Writer
                vObjEscritorPDF = PyPDF2.PdfWriter()
                # Verificar se existe mais de um na pagina
                vNumOcorrencias = vStrTextoPagina.count(iStrIdentificacao)
                if vNumOcorrencias > 1:
                    maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Existe mais de um funcionario na pagina"})
                    # Tamanhos da pagina
                    vListPosicao = vObjPagina.mediabox
                    # Encontrar a posição da string na página
                    vNumPosicao = vStrTextoPagina.index(iStrNomeFuncionario)
                    if vNumPosicao < 500:
                        #Superior
                        x0 = vListPosicao[2]
                        y0 = vListPosicao[3]
                        x1 = vListPosicao[0]
                        y1 = (vListPosicao[3]/2)+19
                    else:
                        #Inferior
                        x0 = vListPosicao[2]
                        y0 = (vListPosicao[3]/2)+19
                        x1 = vListPosicao[0]
                        y1 = vListPosicao[1]
                    maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Cortar pagina"})
                    vObjCaixaCorte = (x0, y0, x1, y1) # (x0, y0, x1, y1)
                    vObjPagina.cropbox.lower_left = vObjCaixaCorte[:2]
                    vObjPagina.cropbox.upper_right = vObjCaixaCorte[2:]
                # Adicionar a página selecionada ao objeto PDF Writer
                vObjEscritorPDF.add_page(vObjPagina)
                maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Salvar comprovante"})
                with open('{}{}.pdf'.format(iStrPastaDownload,"\\"+iStrNomeFuncionario+"_"+iStrTipoDoc), 'wb') as vObjNovoPDF:
                    vObjEscritorPDF.write(vObjNovoPDF)
                    oStrCaminhoNovoArquivo = iStrPastaDownload+'\\'+iStrNomeFuncionario+"_"+iStrTipoDoc+'.pdf'
            else:
                raise Exception("Funcionario nao encontrado")
            maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "savePDF", "Description": "Tarefa finalizada"})
            return oStrCaminhoNovoArquivo
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "savePDF", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[savePDF] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))
        
def mergePDF(maestro, labelLog, listaArquivos, iStrNomeArquivoConsolidado):
    try:
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "mergePDF", "Description": "Iniciando tarefa"})
        pdf_merger = PyPDF2.PdfMerger()
    
        for arquivo in listaArquivos:
            with open(arquivo, 'rb') as pdf_file:
                pdf_merger.append(pdf_file)

        with open(iStrNomeArquivoConsolidado, 'wb') as merged_pdf:
            pdf_merger.write(merged_pdf)
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "INFO", "Task": "mergePDF", "Description": "Tarefa finalizada"})
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        maestro.new_log_entry(activity_label = labelLog,values = {"Datetime": datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"),"Status": "ERRO", "Task": "mergePDF", "Description": "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno)})
        raise Exception("[mergePDF] Erro inesperado "+ "Tipo: "+ str(exc_type) + " Mensagem: " + str(e) + " Linha: " + str(exc_tb.tb_lineno))
        
