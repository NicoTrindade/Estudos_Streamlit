from PyPDF2 import PdfReader
import csv
from funcoes import DadosRetornoCSV

from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo

import customtkinter as ctk
import time
import threading

import os.path

import datetime
# import os
import pandas as pd

from pathlib import Path
# PASTA_RAIZ = Path(__file__).parent
PASTA_RAIZ = os.getcwd() 

PDF = Path(PASTA_RAIZ + '/PDFS_ORIGINAIS')
CSV = Path(PASTA_RAIZ + '/CSV')
EXCEL = Path(PASTA_RAIZ + '/EXCEL')

PDF.mkdir(exist_ok=True)
CSV.mkdir(exist_ok=True)
EXCEL.mkdir(exist_ok=True)

totalRegistros = 0
versao = 'v.1.0.0' 
arquivoCSV = ""
arquivoExcel = ""
nomeArquivoCSV = ""
mensaGeracao = ""

janela = ctk.CTk()
janela.title("CONVERTER ARQUIVOS PDF EM CSV "+versao)
janela.geometry("700x400")
janela.resizable(False, False)

'''janela.resizable(0, 0)
# dimensoes
largura = 550
altura = 500

# resulução do sistema
larguta_tela = janela.winfo_screenwidth()
altura_tela =  janela.winfo_screenmmheight()

# posicao da tela
posix = larguta_tela/2 - largura/2
posiy = altura_tela/2 - altura/2

# definir a geometry
janela.geometry("%dx%d+%d+%d" % (largura,altura,posix,posiy))'''

barraDeStatusFixa = ctk.CTkButton(janela, text="")

# Testes de criação de pastas.
# pastaRAIZ = "Pasta raiz: " + str(PASTA_RAIZ)
# ctk.CTkLabel(janela, text=pastaRAIZ, font=("arial bold", 15)).pack(pady=10)
# ctk.CTkLabel(janela, text="Pasta PDF: " + str(PDF), font=("arial bold", 15)).pack(pady=10)
# ctk.CTkLabel(janela, text="Pasta CSV: " + str(CSV), font=("arial bold", 15)).pack(pady=10)
# ctk.CTkLabel(janela, text="Pasta EXCEL: " + str(EXCEL), font=("arial bold", 15)).pack(pady=10)

# janela._set_appearance_mode("Dark")

dataExpirar = datetime.date(2024,7,25)
dataAtual = datetime.date.today()

def gerarExcel():   
   buscarCaminhoArquivoCSV = escolherArqCSV()

   listaCaminhoArquivoCSV = buscarCaminhoArquivoCSV
   listNomeArquivoCSV = listaCaminhoArquivoCSV[buscarCaminhoArquivoCSV.rfind('/')+1 : len(buscarCaminhoArquivoCSV)]

   csv_COELBA = pd.read_csv(buscarCaminhoArquivoCSV, encoding="ISO-8859-1")

   # # print(csv_COELBA['Dados do cliente'])
   # # print(csv_COELBA)
   # for x in range(csv_COELBA):      
   #    print(csv_COELBA['Dados do cliente'])
   #    print(csv_COELBA['Endereço da Unidade Consumidora'])  
   
   # arquivoExcel = listNomeArquivoCSV.replace(".csv","")+'.xlsx'   
   # nomeArquivoXLSX = EXCEL / arquivoExcel   
   # # with pd.ExcelWriter(nomeArquivoXLSX, engine='xlsxwriter') as writer:
   #    csv_COELBA.to_excel(writer, index=False)   
   #    wb = writer.book  # get workbook
   #    ws = writer.sheets['Sheet1']  # and worksheet

   #    # add a new format then apply it on right columns
   #    currency_fmt = wb.add_format({'num_format': 'R$ #,###.##'})
   #    ws.set_column('H:P', None, currency_fmt)

   arquivoExcel = listNomeArquivoCSV.replace(".csv","")+'.xlsx'   
   nomeArquivoXLSX = EXCEL / arquivoExcel   
   csv_COELBA.to_excel(nomeArquivoXLSX, index=False)   

   arquivoExcel = listNomeArquivoCSV.replace(".csv","")+'.xls'
   nomeArquivoXLS = EXCEL / arquivoExcel   
   csv_COELBA.to_excel(nomeArquivoXLS, index=False)

   try:
      mensagem = showinfo(title="FINALIZADO", message="ARQUIVO: " + "\"" + arquivoExcel + "\"" +  " GERADO COM SUCESSO!") 
   except:
      mensagem = showinfo(title="FINALIZADO", message="ERRO NA TENTATIVA DE EXPORTAR PARA EXCEL.") 
   stop_progressBar()

# def gerarExcelAutomatico(caminho):  

#    caminhoTratado = str(caminho)
#    quantidade = caminhoTratado.count("\\")    
#    for x in range(quantidade):
#       caminhoTratado = caminhoTratado.replace("\\","/")      

#    buscarCaminhoArquivoCSV = caminhoTratado

#    # print(buscarCaminhoArquivoCSV)
#    # print(Path(buscarCaminhoArquivoCSV))   

#    listaCaminhoArquivoCSV = buscarCaminhoArquivoCSV
#    listNomeArquivoCSV = listaCaminhoArquivoCSV[buscarCaminhoArquivoCSV.rfind('/')+1 : len(buscarCaminhoArquivoCSV)]
   
#    buscarCaminhoArquivoCSVAux = Path(buscarCaminhoArquivoCSV)
   
#    csv_COELBA = pd.read_csv(buscarCaminhoArquivoCSVAux, encoding="ISO-8859-1")
   
#    print(csv_COELBA)

#    arquivoExcel = listNomeArquivoCSV.replace(".csv","")+'.xlsx'   
#    nomeArquivoXLSX = EXCEL / arquivoExcel   
#    csv_COELBA.to_excel(nomeArquivoXLSX, index=False)

#    arquivoExcel = listNomeArquivoCSV.replace(".csv","")+'.xls'
#    nomeArquivoXLS = EXCEL / arquivoExcel   
#    csv_COELBA.to_excel(nomeArquivoXLS, index=False)

def start_progressBar(porcent):   
   progressBar.start()   
   janela.update_idletasks()
   janela.title("CONVERSÃO ESTÁ EM: " + porcent )
   time.sleep(1)
   
def stop_progressBar():
   progressBar.stop()
   janela.title("CONVERTER ARQUIVOS PDF EM CSV "+versao)

def escolherArqPDF():
   if dataAtual > dataExpirar:
      mensagem = showinfo(title="Exprirou", message="A versão " + versao + " da aplicação expirou. Entre em contato com o desenvolvedor. (71)98149-8002")   
   else:
      return askopenfilename(title="Selecione um arquivo PDF", filetypes = (("Arquivos PDF", "*.PDF"), ("Todos os arquivos", "*.*")))  

def escolherArqCSV():   
   return askopenfilename(title="Selecione um arquivo PDF", filetypes = (("Arquivos csv", "*.csv"), ("Arquivos xls", "*.xls"), ("Arquivos xlsx", "*.xlsx"), ("Todos os arquivos", "*.*")))  

def gerarCSV2022_SemTravamento():

   def gerarCSV2022():         
      
      buscarCaminhoArquivoPDF = escolherArqPDF()

      listaCaminhoArquivoPDF = buscarCaminhoArquivoPDF
      listNomeArquivoPDF = listaCaminhoArquivoPDF[buscarCaminhoArquivoPDF.rfind('/')+1 : len(buscarCaminhoArquivoPDF)]

      RELATORIO_COELBA = listNomeArquivoPDF               

      arquivoCSV = RELATORIO_COELBA.replace(".pdf","") + '.csv'

      CAMINHO_CSV = CSV / arquivoCSV 

      if buscarCaminhoArquivoPDF == "":
         exit()

      reader = PdfReader(buscarCaminhoArquivoPDF) 
      page = reader.pages

      totalRegistros = len(reader.pages)
      
      lista_cabecalho = ['Dados do cliente', 
                     'Endereço da Unidade Consumidora', 
                     'Número da Nota Fiscal', 
                     'N da Instalação', 
                     'Classificação', 
                     'Descrição da Nota Fiscal', 
                     'Tarifas Aplicadas', 
                     'ICMS Base de Cálculo', 
                     'ICMS Base 2', 
                     'ICMS Base 3', 
                     'ICMS Base 4', 
                     'ICMS Base 5', 
                     'ICMS Base 6', 
                     'ICMS Base 7', 
                     'ICMS Base 8', 
                     'ICMS Base 9', 
                     'Número do medidor', 
                     'Conta Contrato', 
                     'Mês Ano', 
                     'Total a pagar']

      with open(CAMINHO_CSV, 'w', newline='') as csvfile:   
         csv.DictWriter(csvfile, fieldnames=lista_cabecalho, quoting=csv.QUOTE_ALL, delimiter=',').writeheader()
      
         contPag = 0  
      
         statusbar2022 = ctk.CTkButton(janela, text="COELBA: ANO 2022 - Arquivo: " + RELATORIO_COELBA)
         statusbar2022.pack(side=ctk.BOTTOM, fill=ctk.X)
         
         # labelGerarCSV2022 = ctk.CTkLabel(janela, text="COELBA: ANO 2022 - Arquivo: " + RELATORIO_COELBA, font=("arial bold", 15), anchor="center")
         # labelGerarCSV2022.pack(pady=15)             
         # tamanhoDoTexto = len("COELBA: Ano 2022 - Arquivo: " + RELATORIO_COELBA)
         # inicioDoTexto = (300 - tamanhoDoTexto) / 2
         # labelGerarCSV2022.place(x=inicioDoTexto,y=210)                   

         for page in reader.pages:                
         
            if contPag > 0:                                                                        
      
               TEXTO_COMPLETO = page.extract_text()            
               
               # Dados do cliente
               lista_dados_cliente = DadosRetornoCSV(len('DADOS DO CLIENTE'), page.extract_text().find('DADOS DO CLIENTE'), page.extract_text().find('DATA DE VENCIMENTO'), TEXTO_COMPLETO)            

               # Endereço Unidade Consumidora
               lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO DA UNIDADE CONSUMIDORA'), page.extract_text().find('ENDEREÇO DA UNIDADE CONSUMIDORA'), page.extract_text().find('RESERVADO AO FISCO'), TEXTO_COMPLETO)                      

               # Número da Nota Fiscal
               lista_num_nota_fiscal = DadosRetornoCSV(len('NÚMERO DA NOTA FISCAL'), page.extract_text().find('NÚMERO DA NOTA FISCAL'), page.extract_text().find('CONTA CONTRATO'), TEXTO_COMPLETO)          
            
               # Nº da Instlação
               lista_num_Instalacao = DadosRetornoCSV(len('Nº DA INSTALAÇÃO'), page.extract_text().find('Nº DA INSTALAÇÃO'), page.extract_text().find('CLASSIFICAÇÃO'), TEXTO_COMPLETO)                      
               
               # Classificação
               lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO'), page.extract_text().find('CLASSIFICAÇÃO'), page.extract_text().find('ENDEREÇO DA UNIDADE CONSUMIDORA'), TEXTO_COMPLETO)          

               # Descrição da Nota Fiscal
               lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('DESCRIÇÃO DA NOTA FISCAL'), TEXTO_COMPLETO)
               lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split())

               # Tarifas Aplicadas
               if page.extract_text().find('DATA PREVISTA DA PRÓXIMA LEITURA:') > 0:
                  lista_tarifas_aplicadas_tratada = DadosRetornoCSV(len('DATA PREVISTA DA PRÓXIMA LEITURA:')+11, page.extract_text().find('DATA PREVISTA DA PRÓXIMA LEITURA:'), page.extract_text().find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
               else:
                  lista_tarifas_aplicadas = DadosRetornoCSV(len('AJUSTECONSUMO'), page.extract_text().find('AJUSTECONSUMO'), page.extract_text().find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
                  lista_tarifas_aplicadas_tratada = lista_tarifas_aplicadas.replace('(kWh)','').strip()

               # Informações de Tributos
               lista_inform_tributos = DadosRetornoCSV(len('INFORMAÇÕES DE TRIBUTOS'), page.extract_text().find('INFORMAÇÕES DE TRIBUTOS'), page.extract_text().find('AUTENTICAÇÃO MECÂNICA'), TEXTO_COMPLETO)          
               lista_inform_tributos_tratado = " ".join(lista_inform_tributos.split()) # Retrar os espaços entre as palavras      
               lista_inform_tributos_list = lista_inform_tributos_tratado.split(" ") # Converter em lista
               if len(lista_inform_tributos_list) > 8: # Definir se existe o ICMS Base de Cálculo 
                  lista_inform_tributos_list.insert(0,lista_inform_tributos_list[0]) 
               else:
                  lista_inform_tributos_list.insert(0,' ')     

               # Número do Medidor
               if page.extract_text().find('CAT') > 0 and page.extract_text().find('AJUSTECONSUMO') > 0:
                  lista_num_medidor = DadosRetornoCSV(len('AJUSTECONSUMO'), page.extract_text().find('AJUSTECONSUMO'), page.extract_text().find('CAT'), TEXTO_COMPLETO)
                  lista_num_medidor_tratado = lista_num_medidor.replace('(kWh)','').strip()
               else:
                  lista_num_medidor_tratado = ""

               # Conta Contrato
               lista_conta_contato = DadosRetornoCSV(len('CONTA CONTRATO'), page.extract_text().find('CONTA CONTRATO'), page.extract_text().find('Nº DO CLIENTE'), TEXTO_COMPLETO)
         
               # Mês Ano
               lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), page.extract_text().find('MÊS/ANO'), page.extract_text().find('TOTAL A PAGAR(R$)')+1, TEXTO_COMPLETO) 
               
               # Total a pagar
               lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR (R$)'), page.extract_text().find('TOTAL A PAGAR (R$)'), page.extract_text().find('DATA DA EMISSÃO DA NOTA FISCAL'), TEXTO_COMPLETO)                      
               
               csv.writer(csvfile, quoting=csv.QUOTE_ALL, delimiter=',').writerow([lista_dados_cliente, 
                                                                                 lista_end_unid_consum, 
                                                                                 lista_num_nota_fiscal, 
                                                                                 lista_num_Instalacao, 
                                                                                 lista_classificacao, 
                                                                                 lista_desc_nota_fiscal_tratado, 
                                                                                 lista_tarifas_aplicadas_tratada, 
                                                                                 lista_inform_tributos_list[0], 
                                                                                 lista_inform_tributos_list[1],
                                                                                 lista_inform_tributos_list[2],
                                                                                 lista_inform_tributos_list[3],
                                                                                 lista_inform_tributos_list[4],
                                                                                 lista_inform_tributos_list[5],
                                                                                 lista_inform_tributos_list[6],
                                                                                 lista_inform_tributos_list[7],
                                                                                 lista_inform_tributos_list[8],
                                                                                 lista_num_medidor_tratado,
                                                                                 lista_conta_contato,
                                                                                 lista_mes_ano,
                                                                                 lista_total_pagar])                   

            contPag += 1          
            enviarProgessoPercent = f'{(contPag/totalRegistros)*100:_.2f}'+'%'
            start_progressBar(enviarProgessoPercent)
            progressBar.step()
               
         mensagem = showinfo(title="Finalizado", message="Arquivo: " + "\"" + arquivoCSV + "\"" +  " gerado com sucesso!")          
         statusbar2022.destroy() 
            
         # nomeArquivoCSV = CSV / arquivoCSV
         # gerarExcelAutomatico(nomeArquivoCSV)
                     
         # gerarExcel()

         stop_progressBar()                
   
   threading.Thread(target=gerarCSV2022).start()

def gerarCSV2024_SemTravamento():

   def gerarCSV2024():   
      
      buscarCaminhoArquivoPDF = escolherArqPDF()

      listaCaminhoArquivoPDF = buscarCaminhoArquivoPDF
      listNomeArquivoPDF = listaCaminhoArquivoPDF[buscarCaminhoArquivoPDF.rfind('/')+1 : len(buscarCaminhoArquivoPDF)]

      RELATORIO_COELBA = listNomeArquivoPDF    
         
      arquivoCSV = RELATORIO_COELBA.replace(".pdf","") + '.csv'

      CAMINHO_CSV = CSV / arquivoCSV 
      
      if buscarCaminhoArquivoPDF == "":
         exit()
         
      reader = PdfReader(buscarCaminhoArquivoPDF) 
      page = reader.pages

      totalRegistros = len(reader.pages)
      
      lista_cabecalho = ['Dados do cliente', 
                     'Endereço da Unidade Consumidora', 
                     'Número da Nota Fiscal', 
                     'N da Instalação', 
                     'Classificação', 
                     'Descrição da Nota Fiscal', 
                     'Tarifas Aplicadas', 
                     'ICMS Base de Cálculo', 
                     'ICMS Base 2', 
                     'ICMS Base 3', 
                     'ICMS Base 4', 
                     'ICMS Base 5', 
                     'ICMS Base 6', 
                     'ICMS Base 7', 
                     'ICMS Base 8', 
                     'ICMS Base 9', 
                     'Número do medidor', 
                     'Conta Contrato', 
                     'Mês Ano', 
                     'Total a pagar']

      with open(CAMINHO_CSV, 'w', newline='') as csvfile:   
         csv.DictWriter(csvfile, fieldnames=lista_cabecalho, quoting=csv.QUOTE_ALL, delimiter=',').writeheader()         
      
         contPag = 0   
         
         statusbar2024 = ctk.CTkButton(janela, text="COELBA: ANO 2024 - Arquivo: " + RELATORIO_COELBA)
         statusbar2024.pack(side=ctk.BOTTOM, fill=ctk.X)
                  
         # labelGerarCSV2024 = ctk.CTkLabel(janela, text="COELBA: Ano 2024 - Arquivo: " + RELATORIO_COELBA, font=("arial bold", 15), anchor="center")
         # labelGerarCSV2024.pack(pady=15)      
         # tamanhoDoTexto = len("COELBA: Ano 2024 - Arquivo: " + RELATORIO_COELBA)
         # inicioDoTexto = (400 - tamanhoDoTexto) / 2
         # labelGerarCSV2024.place(x=inicioDoTexto,y=210)             
         
         for page in reader.pages:                
         
            if contPag > 0 and contPag % 2 != 0: 
      
               TEXTO_COMPLETO = page.extract_text()   

               #   print(TEXTO_COMPLETO)         
               
               # Dados do cliente
               lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), page.extract_text().find('NOME DO CLIENTE:'), page.extract_text().find('ENDEREÇO:'), TEXTO_COMPLETO)            

               # Endereço Unidade Consumidora
               lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), page.extract_text().find('ENDEREÇO:'), page.extract_text().find('CÓDIGO DA')+1, TEXTO_COMPLETO)  # +1 por conta do caractér especial, pois, não está considerando ara contagem      

               # Número da Nota Fiscal
               lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), page.extract_text().find('NOTA FISCAL N°'), page.extract_text().find('- SÉRIE'), TEXTO_COMPLETO)          
            
               # Nº da Instlação
               lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), page.extract_text().find('INSTALAÇÃO'), page.extract_text().find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)                      
               
               # Classificação
               lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), page.extract_text().find('CLASSIFICAÇÃO:'), page.extract_text().find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)          

               # Descrição da Nota Fiscal
               # lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), page.extract_text().find('CLASSIFICAÇÃO:'), page.extract_text().find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)          

               lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
               lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split()) # Retrar os espaços entre as palavras
               lista_desc_nota_fiscal_tratado_list = lista_desc_nota_fiscal_tratado.split(" ") # Converter em lista                            
                          
               # for listar in lista_desc_nota_fiscal_tratado_list:
               #    print(listar)

               lista_desc_nota_fiscal_separados = []
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[0])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[2])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[3])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[4])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[10])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[12])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[13])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[14])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[20] + " " +
                                                       lista_desc_nota_fiscal_tratado_list[21] + " " +
                                                       lista_desc_nota_fiscal_tratado_list[22] )
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[23])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[24])
               lista_desc_nota_fiscal_separados.append(lista_desc_nota_fiscal_tratado_list[25])
               
               lista_desc_nota_fiscal_gerar = ""
               for lista_separados in lista_desc_nota_fiscal_separados:
                  lista_desc_nota_fiscal_gerar = lista_desc_nota_fiscal_gerar + " " + lista_separados
               
               # Tarifas Aplicadas          
               lista_desc_tarifa_separados = []
               lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[0])
               lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[9])
               lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[10])
               lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[19])
               
               lista_desc_tarifa_gerar = ""
               for lista_separados in lista_desc_tarifa_separados:
                  lista_desc_tarifa_gerar = lista_desc_tarifa_gerar + " " + lista_separados
               
               # Informações de Tributos            
               lista_inform_tributos_ICMS = DadosRetornoCSV(len('ICMS'), page.extract_text().find('ICMS'), page.extract_text().find('CONSUMO / kWh'), TEXTO_COMPLETO)          
               lista_inform_tributos_ICMS_tratado = " ".join(lista_inform_tributos_ICMS.split()) # Retrar os espaços entre as palavras      
               lista_inform_tributos_list_ICMS = lista_inform_tributos_ICMS_tratado.split(" ") # Converter em lista                            

               lista_inform_tributos_PIS = DadosRetornoCSV(len('(%) PIS'), page.extract_text().find('(%) PIS'), page.extract_text().find('COFINS'), TEXTO_COMPLETO)          
               lista_inform_tributos_PIS_tratado = " ".join(lista_inform_tributos_PIS.split()) # Retrar os espaços entre as palavras      
               lista_inform_tributos_list_PIS = lista_inform_tributos_PIS_tratado.split(" ") # Converter em lista   

               lista_inform_tributos_COFINS = DadosRetornoCSV(len('COFINS'), page.extract_text().find('COFINS'), page.extract_text().find('ICMS'), TEXTO_COMPLETO)          
               lista_inform_tributos_COFINS_tratado = " ".join(lista_inform_tributos_COFINS.split()) # Retrar os espaços entre as palavras      
               lista_inform_tributos_list_COFINS = lista_inform_tributos_COFINS_tratado.split(" ") # Converter em lista                         

               # Número do Medidor              
               if page.extract_text().find('MEDIDOR kWh') > 0 and page.extract_text().find('Energia Ativa') > 0:
                  lista_num_medidor_tratado = DadosRetornoCSV(len('MEDIDOR kWh'), page.extract_text().find('MEDIDOR kWh'), page.extract_text().find('Energia Ativa'), TEXTO_COMPLETO)
                  #   lista_num_medidor_tratado = lista_num_medidor.replace('(kWh)','').strip()
               else:
                  lista_num_medidor_tratado = ""

               # Conta Contrato
               lista_conta_contato = DadosRetornoCSV(len('Conta Contrato Coletiva nº'), page.extract_text().find('Conta Contrato Coletiva nº'), page.extract_text().find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO)
         
               # Mês Ano
               lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), page.extract_text().find('MÊS/ANO'), page.extract_text().find('VENCIMENTO'), TEXTO_COMPLETO) 
               
               # Total a pagar
               lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), page.extract_text().find('TOTAL A PAGAR R$'), page.extract_text().find('Cadastra-se e receba'), TEXTO_COMPLETO)                      
               
               csv.writer(csvfile, quoting=csv.QUOTE_ALL, delimiter=',').writerow([lista_dados_cliente, 
                                                                                 lista_end_unid_consum, 
                                                                                 lista_num_nota_fiscal, 
                                                                                 lista_num_Instalacao, 
                                                                                 lista_classificacao, 
                                                                                 lista_desc_nota_fiscal_gerar, 
                                                                                 lista_desc_tarifa_gerar, 
                                                                                 lista_inform_tributos_list_ICMS[0], 
                                                                                 lista_inform_tributos_list_ICMS[1],
                                                                                 lista_inform_tributos_list_ICMS[2],
                                                                                 lista_inform_tributos_list_PIS[0],
                                                                                 lista_inform_tributos_list_PIS[1],
                                                                                 lista_inform_tributos_list_PIS[2],
                                                                                 lista_inform_tributos_list_COFINS[0],
                                                                                 lista_inform_tributos_list_COFINS[1],
                                                                                 lista_inform_tributos_list_COFINS[2],
                                                                                 lista_num_medidor_tratado,
                                                                                 lista_conta_contato,
                                                                                 lista_mes_ano,
                                                                                 lista_total_pagar])                   

            contPag += 1          
            enviarProgessoPercent = f'{(contPag/totalRegistros)*100:_.2f}'+'%'
            start_progressBar(enviarProgessoPercent)
            progressBar.step()
               
         mensagem = showinfo(title="Finalizado", message="Arquivo: " + "\"" + arquivoCSV + "\"" +  " gerado com sucesso!") 
         statusbar2024.destroy()
               
         # nomeArquivoCSV = CSV / arquivoCSV
         # gerarExcelAutomatico(nomeArquivoCSV)
                     
         # gerarExcel()

         stop_progressBar()                
   threading.Thread(target=gerarCSV2024).start()

# ctk.CTkLabel(janela, text="CONVERSÃO DOS DADOS DE ARQUIVOS: PDF >> CSV >> EXCEL.", font=("arial bold", 15)).pack(pady=10)
barraDeStatusFixa = ctk.CTkButton(janela, text="App Conversão dos dados de arquivos: PDF >> CSV >> EXCEL. Versão: 1.0.01")
barraDeStatusFixa.pack(side=ctk.TOP, fill=ctk.X)   

ctk.CTkLabel(janela, text="Instruções: ", font=("arial bold", 18)).pack(pady=1)
ctk.CTkLabel(janela, text="(DE PDF PARA CSV) BOTÃO: GERAR CSV 2022 / GERAR CSV 2024 - PASTA: PDFS_ORIGINAIS", font=("arial bold", 12)).pack(pady=1)
ctk.CTkLabel(janela, text="(DE CSV PARA EXCEL) BOTÃO: EXPORTAR EXCEL - PASTA: CSV", font=("arial bold", 12)).pack(pady=1)

progressBar = ctk.CTkProgressBar(janela, width=400, height=20, corner_radius=20, fg_color="#003", progress_color="#060", mode='determinate')
progressBar.pack(pady=50)

btn2022 = ctk.CTkButton(janela, text="GERAR CSV 2022", command=gerarCSV2022_SemTravamento)
btn2022.pack(padx=20, pady=10)
btn2022.place(x=110, y=300)

btn2024 = ctk.CTkButton(janela, text="GERAR CSV 2024", command=gerarCSV2024_SemTravamento)
btn2024.pack(padx=20, pady=10)
btn2024.place(x=280, y=300)

btn_xlsx = ctk.CTkButton(janela, text="EXPORTAR EXCEL", command=gerarExcel)
btn_xlsx.pack(padx=20, pady=10)
btn_xlsx.place(x=450, y=300)  

janela.mainloop()
