
import ROTINA as rt
import PROMAX as pro
import MOVIMENTAR as mv
from datetime import datetime
import time
import os
import pyautogui as auto
import subprocess
start_time = datetime.now()



time.sleep(30)

#Abrir o promax
pro.abri_promax()
pro.inserir_dados(1)
pro.definir_unidade(1,1)
pro.seguir(1)

tempo_padrao=3

#BAIXAR AS ROTINAS DO REVEBI
#ATENDIMENTO/BEES DELIVERY




rt.rotina_03_02_24(tempo_padrao,'Matriz',3,r'Z:\ATENDIMENTO\BEES DELIVERY\03.02.24\MATRIZ')
rt.rotina_03_11_20(tempo_padrao,"Matriz",r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\MATRIZ')
#ATENDIMENTO
rt.rotina_01_11(tempo_padrao,r'Z:\ATENDIMENTO\FEFO\01.11')
rt.rotina_03_02_37(tempo_padrao,'Matriz',14,5,24,1,r'Z:\ATENDIMENTO\FEFO\03.02.37\MATRIZ',0,1,0)#Operação/Vendedor/Motorista
rt.rotina_01_20_01_24(tempo_padrao,'Matriz',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\01.20.01.24\MATRIZ')
rt.rotina_03_01_47_01(tempo_padrao,1,'Revenda',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\03.01.47.01\REVENDAS',1)
#ATENDIMENTO/POLITICA DE ESTOQUE
rt.rotina_02_05_02(tempo_padrao,'Matriz',1,r'Z:\ATENDIMENTO\POLITICA DE ESTOQUE\02.05.02\MATRIZ')
#ATENDIMENTO/RATING
rt.rotina_01_05_07_04_02(tempo_padrao,'Matriz',1,r'Z:\ATENDIMENTO\RATING\01.05.07.04.02')
rt.rotina_03_02_37(tempo_padrao,'Matriz',25,35,36,0,r'Z:\ATENDIMENTO\RATING\03.02.37\MATRIZ',0,1,0)#Motorista/Ajudante1/Ajudante2
#COMERCIAL/BEES
rt.rotina_23_01_01_02(tempo_padrao,'Matriz',r'Z:\COMERCIAL\BEES\23.01.01.02\MATRIZ')
#COMERCIAL/VOLUME
rt.rotina_05_10(tempo_padrao,'Matriz',r'Z:\COMERCIAL\VOLUME\05.10\MATRIZ')
#CONTROLE/VOLUME
rt.rotina_03_05_26(tempo_padrao,'Matriz',r'Z:\CONTROLE\GESTÃO DE VALES\03.05.26\MATRIZ')
rt.rotina_03_07_13(tempo_padrao,'Matriz',r'Z:\CONTROLE\GESTÃO DE VALES\03.07.13\MATRIZ')
rt.rotina_03_11_40(tempo_padrao,'Matriz',r'Z:\CONTROLE\GESTÃO DE VALES\03.11.40\MATRIZ')

#MOVIMENTAR ARQUIVOS
#CRITICA DE PEDIDO
rt.rotina_01_05_07_04_02(tempo_padrao,'Matriz',1,r'Z:\CRITICA PEDIDOS\01.05.07.04.02')
rt.rotina_01_09(tempo_padrao,r'Z:\CRITICA PEDIDOS\01.09')
rt.rotina_01_12(tempo_padrao,r'Z:\CRITICA PEDIDOS\01.12')
rt.rotina_02_05_02(tempo_padrao,'Matriz',2,r'Z:\CRITICA PEDIDOS\02.05.02')
rt.rotina_03_01_11(tempo_padrao,'Matriz',r'Z:\CRITICA PEDIDOS\03.01.11')
rt.rotina_03_01_36_04(tempo_padrao,'Matriz',r'Z:\CRITICA PEDIDOS\03.01.36.04',0)
rt.rotina_03_02_24(tempo_padrao,'Matriz',4,r'Z:\CRITICA PEDIDOS\03.02.24',2)

#FINANCEIRO COMODATO
rt.rotina_02_02_20(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\COMODATO\02.02.20')
rt.rotina_02_02_53(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\COMODATO\02.02.53\MATRIZ')
rt.rotina_02_02_29(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\COMODATO\02.02.29')
#FINANCEIRO OBZ
#rt.rotina_15_05_01(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\OBZ\15.05.01\MATRIZ')
#rt.rotina_15_05_04(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\OBZ\15.05.04\MATRIZ')
#FINANCEIRO RECEBIMENTO
rt.rotina_01_20_01_27(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.27\MATRIZ')
rt.rotina_01_20_01_47(tempo_padrao,'Matriz',r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\MATRIZ')
rt.rotina_03_02_37(tempo_padrao,'Matriz',14,15,24,0,r'Z:\FINANCEIRO\RECEBIMENTO\03.02.37\MATRIZ',0,1,0)#Operação/Condição/Motorista

#FROTA GESTRAN
rt.rotina_03_11_49_02(tempo_padrao,'Matriz',r'Z:\FROTA\GESTRAN\03.11.49.02\MATRIZ')
#PRODUTIVIDADE CURVA ABC
rt.rotina_03_02_36_08(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\CURVA ABC\03.02.36.08\MATRIZ')#
#PRODUTIVIDADE GESTAO DE PRODUTIVIDADE
rt.rotina_03_02_37(tempo_padrao,'Matriz',36,36,0,0,r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.02.37\MATRIZ',0,1,0)#Ajudante1 e Ajduante2
rt.rotina_03_11_35_04(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\MATRIZ')
#PRODUTIVIDADE ID AJUDANTE
rt.rotina_03_02_37(tempo_padrao,'Matriz',14,35,36,1,r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.02.37\MATRIZ',0,1,0)#Operacação/Ajudante1 e Ajduante2
#PRODUTIVIDADE PLANNER
rt.rotina_01_05_07_04_02(tempo_padrao,'Matriz',2,r'Z:\PRODUTIVIDADE\PLANNER\01.05.07.04.02\MATRIZ')
rt.rotina_01_20_11(tempo_padrao,'Matriz',4,r'Z:\PRODUTIVIDADE\PLANNER\01.20.11\MATRIZ')
rt.rotina_03_01_47_01(tempo_padrao,0,'Matriz',r'Z:\PRODUTIVIDADE\PLANNER\03.01.47.01\MATRIZ',0)
#PRODUTIVIDADE PLANNER
rt.rotina_01_20_01_24(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\PREJUÍZO\01.20.01.24\MATRIZ')
rt.rotina_03_11_34_03(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.34.03\MATRIZ')
rt.rotina_03_11_34_05(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.34.05\MATRIZ')
rt.rotina_03_11_35_06(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.35.06\MATRIZ')
rt.rotina_03_18_05(tempo_padrao,'Matriz',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.18.05\MATRIZ')
#
rt.fechar_promax()

time.sleep(60)
auto.click(1,0)

#Abrir o promax
pro.abri_promax()
pro.inserir_dados(2)
pro.definir_unidade(1,2)
pro.seguir(1)
tempo_padrao=3

#BAIXAR AS ROTINAS DO REVEBI
#ATENDIMENTO/BEES DELIVERY
#rt.rotina_01_20_01_47(tempo_padrao,'Filial',r'Z:\ATENDIMENTO\BEES DELIVERY\01.20.01.47\FILIAL')#NÃO RODAR PORQUE NICOLE DISSE QUE ESTÁ ERRADO
rt.rotina_03_02_24(tempo_padrao,'Filial',3,r'Z:\ATENDIMENTO\BEES DELIVERY\03.02.24\FILIAL')
rt.rotina_03_11_20(tempo_padrao,"Filial",r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\FILIAL')
#ATENDIMENTO/FEFO
rt.rotina_03_02_37(tempo_padrao,'Filial',14,5,24,1,r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',0,1,0)#Operação/Vendedor/Motorista
rt.rotina_01_20_01_24(tempo_padrao,'Filial',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\01.20.01.24\FILIAL')
#ATENDIMENTO/POLITICA DE ESTOQUE
rt.rotina_02_05_02(tempo_padrao,'Filial',1,r'Z:\ATENDIMENTO\POLITICA DE ESTOQUE\02.05.02\FILIAL')
#ATENDIMENTO/RATING
rt.rotina_01_05_07_04_02(tempo_padrao,'Filial',1,r'Z:\ATENDIMENTO\RATING\01.05.07.04.02')
rt.rotina_03_02_37(tempo_padrao,'Filial',25,35,36,0,r'Z:\ATENDIMENTO\RATING\03.02.37\FILIAL',0,1,0)#Motorista/Ajudante1/Ajudante2
#COMERCIAL/BEES
rt.rotina_23_01_01_02(tempo_padrao,'Filial',r'Z:\COMERCIAL\BEES\23.01.01.02\FILIAL')
#COMERCIAL/VOLUME
rt.rotina_05_10(tempo_padrao,'Filial',r'Z:\COMERCIAL\VOLUME\05.10\FILIAL')
#CONTROLE/VOLUME
rt.rotina_03_05_26(tempo_padrao,'Filial',r'Z:\CONTROLE\GESTÃO DE VALES\03.05.26\FILIAL')
rt.rotina_03_07_13(tempo_padrao,'Filial',r'Z:\CONTROLE\GESTÃO DE VALES\03.07.13\FILIAL')
rt.rotina_03_11_40(tempo_padrao,'Filial',r'Z:\CONTROLE\GESTÃO DE VALES\03.11.40\FILIAL')
#MOVIMENTAR ARQUIVOS
#CRITICA DE PEDIDO
rt.rotina_01_05_07_04_02(tempo_padrao,'Filial',1,r'Z:\CRITICA PEDIDOS\01.05.07.04.02')
rt.rotina_02_05_02(tempo_padrao,'Filial',2,r'Z:\CRITICA PEDIDOS\02.05.02')
rt.rotina_03_01_11(tempo_padrao,'Filial',r'Z:\CRITICA PEDIDOS\03.01.11')
rt.rotina_03_01_36_04(tempo_padrao,'Filial',r'Z:\CRITICA PEDIDOS\03.01.36.04',0)
rt.rotina_03_02_24(tempo_padrao,'Filial',4,r'Z:\CRITICA PEDIDOS\03.02.24',2)
#FINANCEIRO COMODATO
rt.rotina_02_02_20(tempo_padrao,'Filial',r'Z:\FINANCEIRO\COMODATO\02.02.20')
rt.rotina_02_02_53(tempo_padrao,'Filial',r'Z:\FINANCEIRO\COMODATO\02.02.53\FILIAL')
rt.rotina_02_02_29(tempo_padrao,'Filial',r'Z:\FINANCEIRO\COMODATO\02.02.29')
#FINANCEIRO OBZ
#rt.rotina_15_05_01(tempo_padrao,'Filial',r'Z:\FINANCEIRO\OBZ\15.05.01\FILIAL')
#rt.rotina_15_05_04(tempo_padrao,'Filial',r'Z:\FINANCEIRO\OBZ\15.05.04\FILIAL')
#FINANCEIRO RECEBIMENTO
rt.rotina_01_20_01_27(tempo_padrao,'Filial',r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.27\FILIAL')
rt.rotina_01_20_01_47(tempo_padrao,'Filial',r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\FILIAL')
rt.rotina_03_02_37(tempo_padrao,'Filial',14,15,24,0,r'Z:\FINANCEIRO\RECEBIMENTO\03.02.37\FILIAL',0,1,0)#Operação/Condição/Motorista

#FROTA GESTRAN
rt.rotina_03_11_49_02(tempo_padrao,'Filial',r'Z:\FROTA\GESTRAN\03.11.49.02\FILIAL')
#PRODUTIVIDADE CURVA ABC
rt.rotina_03_02_36_08(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\CURVA ABC\03.02.36.08\FILIAL')
#PRODUTIVIDADE GESTAO DE PRODUTIVIDADE
rt.rotina_03_02_37(tempo_padrao,'Filial',36,36,0,0,r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.02.37\FILIAL',0,1,0)#Ajudante1 e Ajduante2
rt.rotina_03_11_35_04(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\FILIAL')
#PRODUTIVIDADE ID AJUDANTE
rt.rotina_03_02_37(tempo_padrao,'Filial',14,35,36,1,r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.02.37\FILIAL',0,1,0)#Operacação/Ajudante1 e Ajduante2
#PRODUTIVIDADE PLANNER
rt.rotina_01_05_07_04_02(tempo_padrao,'Filial',2,r'Z:\PRODUTIVIDADE\PLANNER\01.05.07.04.02\FILIAL')
rt.rotina_01_20_11(tempo_padrao,'Filial',4,r'Z:\PRODUTIVIDADE\PLANNER\01.20.11\FILIAL')
rt.rotina_03_01_47_01(tempo_padrao,0,'Filial',r'Z:\PRODUTIVIDADE\PLANNER\03.01.47.01\FILIAL',0)
#PRODUTIVIDADE PLANNER
rt.rotina_01_20_01_24(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\PREJUÍZO\01.20.01.24\FILIAL')
rt.rotina_03_11_34_03(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.34.03\FILIAL')
rt.rotina_03_11_34_05(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.34.05\FILIAL')
rt.rotina_03_11_35_06(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.35.06\FILIAL')
rt.rotina_03_18_05(tempo_padrao,'Filial',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.18.05\FILIAL')



# Fechar o promax
rt.fechar_promax()


intervalo_verificacao=10



#MOVIMENTAÇÃO MATRIZ
#01.11
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\ATENDIMENTO\PNC\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\ATENDIMENTO\POLITICA DE ESTOQUE\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\ATENDIMENTO\PUXADA\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\COMERCIAL\BEES\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\COMERCIAL\VOLUME\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\CONTROLE\GESTÃO DE VALES\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\CRITICA PEDIDOS\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\CURVA ABC\Configuração\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\FAST PICKING\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\ID AJUDANTE\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\PICKING\01.11\01.11.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\FEFO\01.11\01.11.csv',r'Z:\PRODUTIVIDADE\PREJUÍZO\01.11\01.11.csv')


#01.05.07.04.02
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\ATENDIMENTO\SOLICITAÇÕES\01.05.07.04.02\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\ATENDIMENTO\TOP LINE\01.05.07.04.02\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\COMERCIAL\BEES\01.05.07.04.02\MATRIZ\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\COMERCIAL\VOLUME\01.05.07.04.02\MATRIZ\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\FINANCEIRO\COMODATO\01.05.07.04.02\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\FINANCEIRO\RECEBIMENTO\01.05.07.04.02\MATRIZ\Matriz.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Matriz.csv',r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\01.05.07.04.02\MATRIZ\Matriz.csv')
#01.20.01.24
mv.substituir_arquivo(r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\01.20.01.24\MATRIZ\Matriz.csv',r'Z:\COMERCIAL\VOLUME\01.20.01.24\MATRIZ\Matriz.csv')
#03.02.37#Operação/Vendedor/Motorista
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\MATRIZ',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\03.02.37\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\MATRIZ',r'Z:\COMERCIAL\VOLUME\03.02.37\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\MATRIZ',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.02.37\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\MATRIZ',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.02.37\MATRIZ','Matriz')
#Roterizador Cliente
#mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\Clientes\MATRIZ',r'Z:\CRITICA PEDIDOS\Clientes','Matriz')
#01.20.01.47
mv.substituir_arquivo_com_nome_especifico(r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\MATRIZ',r'Z:\FROTA\GESTRAN\01.20.01.47\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\MATRIZ',r'Z:\PRODUTIVIDADE\PREJUÍZO\01.20.01.47\MATRIZ','Matriz')
#3.11.20
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\MATRIZ',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.20\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\MATRIZ',r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.11.20\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\MATRIZ',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.11.20\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\MATRIZ',r'Z:\PRODUTIVIDADE\JUSTIFICATIVA MPD\03.11.20\MATRIZ','Matriz')
#03.11.35.04
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\MATRIZ',r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.11.35.04\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\MATRIZ',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.35.04\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\MATRIZ',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.11.34.05\MATRIZ','Matriz')
#3.11.49.02
mv.substituir_arquivo_com_nome_especifico(r'Z:\FROTA\GESTRAN\03.11.49.02\MATRIZ',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.49.02\MATRIZ','Matriz')
mv.substituir_arquivo_com_nome_especifico(r'Z:\FROTA\GESTRAN\03.11.49.02\MATRIZ',r'Z:\PRODUTIVIDADE\JUSTIFICATIVA MPD\03.11.49.02\MATRIZ','Matriz')
#3.11.40
mv.substituir_arquivo_com_nome_especifico(r'Z:\CONTROLE\GESTÃO DE VALES\03.11.40\MATRIZ',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.40\MATRIZ','Matriz')
#MOVIMENTAÇÃO FILIAL
#01.05.07.04.02
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\ATENDIMENTO\SOLICITAÇÕES\01.05.07.04.02\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\ATENDIMENTO\TOP LINE\01.05.07.04.02\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\COMERCIAL\BEES\01.05.07.04.02\FILIAL\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\COMERCIAL\VOLUME\01.05.07.04.02\FILIAL\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\FINANCEIRO\COMODATO\01.05.07.04.02\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\FINANCEIRO\RECEBIMENTO\01.05.07.04.02\FILIAL\Filial.csv')
mv.substituir_arquivo(r'Z:\ATENDIMENTO\RATING\01.05.07.04.02\Filial.csv',r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\01.05.07.04.02\FILIAL\Filial.csv')
#01.20.01.24
mv.substituir_arquivo(r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\01.20.01.24\FILIAL\Filial.csv',r'Z:\COMERCIAL\VOLUME\01.20.01.24\FILIAL\Filial.csv')
#03.02.37#Operação/Vendedor/Motorista
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\03.02.37\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',r'Z:\COMERCIAL\VOLUME\03.02.37\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.02.37\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.02.37\FILIAL','Filial')
#03.02.37#Motorista/Ajudante1/Ajudante2
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\FEFO\03.02.37\FILIAL',r'Z:\ATENDIMENTO\NÍVEL DE SERVIÇO\03.02.37\FILIAL','Filial')#
#Roterizador Cliente
#mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\ROTEIRIZADOR\Clientes\FILIAL',r'Z:\CRITICA PEDIDOS\Clientes','Filial')
#01.20.01.47
mv.substituir_arquivo_com_nome_especifico(r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\FILIAL',r'Z:\FROTA\GESTRAN\01.20.01.47\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\FINANCEIRO\RECEBIMENTO\01.20.01.47\FILIAL',r'Z:\PRODUTIVIDADE\PREJUÍZO\01.20.01.47\FILIAL','Filial')
#3.11.20
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\FILIAL',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.20\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\FILIAL',r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.11.20\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\FILIAL',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.11.20\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\ATENDIMENTO\BEES DELIVERY\03.11.20\FILIAL',r'Z:\PRODUTIVIDADE\JUSTIFICATIVA MPD\03.11.20\FILIAL','Filial')
#03.11.35.04
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\FILIAL',r'Z:\PRODUTIVIDADE\ID AJUDANTE\03.11.35.04\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\FILIAL',r'Z:\PRODUTIVIDADE\PREJUÍZO\03.11.35.04\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.35.04\FILIAL',r'Z:\PRODUTIVIDADE\INPUT ARMAZÉM\03.11.34.05\FILIAL','Filial')
#3.11.49.02
mv.substituir_arquivo_com_nome_especifico(r'Z:\FROTA\GESTRAN\03.11.49.02\FILIAL',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.49.02\FILIAL','Filial')
mv.substituir_arquivo_com_nome_especifico(r'Z:\FROTA\GESTRAN\03.11.49.02\FILIAL',r'Z:\PRODUTIVIDADE\JUSTIFICATIVA MPD\03.11.49.02\FILIAL','Filial')
#3.11.40
mv.substituir_arquivo_com_nome_especifico(r'Z:\CONTROLE\GESTÃO DE VALES\03.11.40\FILIAL',r'Z:\PRODUTIVIDADE\GESTÃO MPD\03.11.40\FILIAL','Filial')


#VALIDAR SE O ARQUIVO FOI BAIXADO


caminho_executavel = r"C:\Users\pedro.alves\Desktop\BASE\ROTERIZADOR.exe"
# Executar o programa
subprocess.run(caminho_executavel)

end_time = datetime.now()

file_path = "log.txt"
with open(file_path, "a") as file:
    file.write(f'\n\n Processo inicio a {start_time} e finalizou as {end_time}.')

while True:
    agora = datetime.now()
    if agora.hour > 6:
        # Abrir o promax
        pro.abri_promax()
        pro.inserir_dados(2)
        pro.definir_unidade(1, 1)
        pro.seguir(1)
        rt.rotina_12_06_01(tempo_padrao, 'Matriz', r'Z:\CRITICA PEDIDOS\12.06.01', "5", 5)
        rt.rotina_12_06_01(tempo_padrao, 'Matriz', r'Z:\CONTROLE\GESTÃO DE VALES\12.06.01\MATRIZ', '99', 6)
        rt.fechar_promax()
        pro.abri_promax()
        pro.inserir_dados(2)
        pro.definir_unidade(1, 2)
        pro.seguir(1)
        rt.rotina_12_06_01(tempo_padrao, 'Filial', r'Z:\CRITICA PEDIDOS\12.06.01', "5", 5)
        rt.rotina_12_06_01(tempo_padrao, 'Filial', r'Z:\CONTROLE\GESTÃO DE VALES\12.06.01\FILIAL', '99', 6)
        time.sleep(60)
        rt.fechar_promax()
        break
    else:
        # Esperar por 1 minuto antes de verificar novamente
        print("Aguardando")
        auto.click(1,0)
        time.sleep(60)
# Caminho para o executável
caminho_executavel = r"C:\Users\pedro.alves\Desktop\MIGUELITOFINAL.exe"

# Executar o programa
subprocess.run(caminho_executavel)


os.system("shutdown /s /t 1800")