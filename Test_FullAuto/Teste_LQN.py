import      pyautogui
import      pyscreeze
import      pandas
from      faker     import       Faker
import      pandas as pd
import      time
import      os


class Dados():
    def Abrir_planilha():
        
        pasta_destino=r'C:\Users\nsn102241\OneDrive - nansen.com.br\Documents\Test_FullAuto\Resultado'
        pasta_planilhas=r'C:\Users\nsn102241\OneDrive - nansen.com.br\Documents\Test_FullAuto\Planilhas'
        planilhas_dict = {}
        
        print("\nIDs falhados:  \n") 

        for arquivo in os.listdir(pasta_planilhas):
            if arquivo.endswith('.xls') or arquivo.endswith('.xlsx'):
                caminho_arq = os.path.join(pasta_planilhas, arquivo)
                df = pd.read_excel(caminho_arq)
                planilhas_dict[arquivo] = df

                ids_falhados = []
                
                for l in range(6, 17, +2):  ## Range para identificar fail
                     result = df.iloc[66, l]
                     ## print(result)
                     if result == ('Failed') and result != ('null'):
                        valor_ID = df.iloc[65, l-1]  ## Coletar ID na linha - PANILHA
                        print(valor_ID)  
                        ids_falhados.append(valor_ID)
                df_ids = pd.DataFrame(ids_falhados, columns=['ID_Falhado'])
                caminho_arquivo = os.path.join(pasta_destino, f'{arquivo}.xlsx')
                df_ids.to_excel(caminho_arquivo, index=False)       
            ## else:
                   ## print('\nNão falhado\n')
    Abrir_planilha()
Dados()