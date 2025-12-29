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

        caminho_planilha = os.path.join(pasta_planilhas, 'teste_fullauto.xls')
        df = pd.read_excel(caminho_planilha)
        print(df)
        ids_falhados = []
        for l in range(6, 17, +2):  ## Range para identificar fail
            result = df.iloc[15, l]
            ## print(result)
            if result == ('Failed'):
                valor_ID = df.iloc[14, l-1]  ## Coletar ID na linha - PANILHA
                print(valor_ID)  
                ids_falhados.append(valor_ID)
        fake=Faker() 
        df_ids = pd.DataFrame(ids_falhados, columns=['ID_Falhado'])
        caminho_arquivo = os.path.join(pasta_destino, f'{fake.random_number(digits=2)}.IDs_falhados.xlsx')
        df_ids.to_excel(caminho_arquivo, index=False)
        print("Arquivo 'IDs_falhados.xlsx' gerado com sucesso!")        
            ## else:
                   ## print('\nNão falhado\n')

    Abrir_planilha()
Dados()

