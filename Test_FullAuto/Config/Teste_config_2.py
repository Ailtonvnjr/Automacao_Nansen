import      pyautogui
import      pyscreeze
import      pandas
from      faker     import       Faker
import      pandas as pd
import      time
import      os


class Dados():
    def Abrir_planilha():
        
        pasta_destino=r'C:\Users\nsn102241\Documents\Test_FullAuto\Resultado'
        pasta_planilhas=r'C:\Users\nsn102241\Documents\Test_FullAuto\Planilhas'
        lista_dfs = []
        
        for arquivo in os.listdir(pasta_planilhas):
            if arquivo.endswith('.xls') or arquivo.endswith('.xlsx'):
                caminho_arq = os.path.join(pasta_planilhas, arquivo)
                df = pd.read_excel(caminho_arq)
                lista_dfs.append(df)

                df_final = pd.concat(lista_dfs, ignore_index=True)
                print(df_final.head())

                 
                caminho_planilha = os.path.join(pasta_planilhas, 'teste_fullauto.xls')
                df = pd.read_excel(caminho_planilha)
                print(df)
                ids_falhados = []