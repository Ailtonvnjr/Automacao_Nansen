
import os
import pandas as pd

class Dados:
    @staticmethod
    def Abrir_planilha():
        pasta_destino = r'C:\Users\nsn102241\OneDrive - nansen.com.br\Documents\Test_FullAuto\Resultado'
        pasta_planilhas = r'C:\Users\nsn102241\OneDrive - nansen.com.br\Documents\Test_FullAuto\Planilhas'

        # Garante que a pasta de destino exista
        os.makedirs(pasta_destino, exist_ok=True)
        
        print("\nIDs falhados:\n")

        # Lista acumuladora para todos os IDs de todas as planilhas
        todos_ids = []

        for arquivo in os.listdir(pasta_planilhas):
            if arquivo.endswith('.xls') or arquivo.endswith('.xlsx'):
                caminho_arq = os.path.join(pasta_planilhas, arquivo)

                # Lê de acordo com a extensão
                try:
                    if arquivo.endswith('.xlsx'):
                        df = pd.read_excel(caminho_arq, engine='openpyxl')
                    else:
                        df = pd.read_excel(caminho_arq, engine='xlrd')
                except Exception as e:
                    print(f"Aviso: não foi possível ler '{arquivo}': {e}")
                    continue

                # Coleta IDs falhados
                ids_falhados = []

                # Range para identificar colunas (6 até 16, de 2 em 2): 6, 8, 10, 12, 14, 16
                for l in range(6, 17, 2):
                    try:
                        result = df.iloc[66, l]  # célula com status
                    except Exception:
                        result = None

                    # Checa 'Failed' com segurança
                    if isinstance(result, str) and result.strip().lower() == 'failed':
                        try:
                            valor_ID = df.iloc[65, l - 1]  # célula com ID
                        except Exception:
                            valor_ID = None

                        if pd.notna(valor_ID) and str(valor_ID).strip().lower() != 'null':
                            print(valor_ID)
                            ids_falhados.append(valor_ID)

                # Acumula com o nome do arquivo (para rastreabilidade)
                if ids_falhados:
                    for _id in ids_falhados:
                        todos_ids.append({'Arquivo': arquivo, 'ID_Falhado': _id})

        # Cria DataFrame final e salva em um único Excel
        if todos_ids:
            df_final = pd.DataFrame(todos_ids, columns=['Arquivo', 'ID_Falhado'])
            caminho_unico = os.path.join(pasta_destino, 'IDs_Falhados.xlsx')
            df_final.to_excel(caminho_unico, index=False)
            print(f"\n✔ IDs consolidados salvos em: {caminho_unico}\n")
        else:
            print("\nNenhum ID falhado encontrado.\n")

# Executa
if __name__ == '__main__':
    Dados.Abrir_planilha()
