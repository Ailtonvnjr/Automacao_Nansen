
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

        # Acumuladores globais
        todos_ids_falhados = []  # para aba 'IDs_Falhados'
        resumo_por_arquivo = []  # para aba 'Resumo'

        # Totais gerais
        total_verificados_geral = 0
        total_falhados_geral = 0

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

                # Contadores por arquivo
                verificados = 0
                falhados = 0

                # Range das colunas: 6, 8, 10, 12, 14, 16
                for l in range(6, 17, 2):
                    # Lê status na linha 66, coluna l
                    try:
                        status = df.iloc[66, l]
                    except Exception:
                        status = None

                    # Lê ID na linha 65, coluna l-1
                    try:
                        id_val = df.iloc[65, l - 1]
                    except Exception:
                        id_val = None

                    # Normaliza valores para checagens
                    id_str = str(id_val).strip() if pd.notna(id_val) else ""
                    status_str = str(status).strip().lower() if isinstance(status, str) else (str(status).strip().lower() if pd.notna(status) else "")

                    # É considerado verificado se existe um ID válido (não vazio/nulo/'null')
                    if id_str and id_str.lower() != 'null':
                        verificados += 1
                        # Considera falha se status == 'failed'
                        if status_str == 'failed':
                            falhados += 1
                            todos_ids_falhados.append({'Arquivo': arquivo, 'ID_Falhado': id_str})
                            print(id_str)

                # Atualiza totais gerais
                total_verificados_geral += verificados
                total_falhados_geral += falhados

                # Percentual por arquivo
                perc = (falhados / verificados * 100) if verificados > 0 else 0.0

                # Guarda o resumo deste arquivo
                resumo_por_arquivo.append({
                    'Arquivo': arquivo,
                    'IDs_Verificados': verificados,
                    'IDs_Falhados': falhados,
                    'Percentual_Falha(%)': round(perc, 2)
                })

        # Monta DataFrames
        df_ids = pd.DataFrame(todos_ids_falhados, columns=['Arquivo', 'ID_Falhado'])
        df_resumo = pd.DataFrame(resumo_por_arquivo, columns=['Arquivo', 'IDs_Verificados', 'IDs_Falhados', 'Percentual_Falha(%)'])

        # Adiciona linha de total geral no resumo
        if total_verificados_geral > 0:
            perc_geral = round((total_falhados_geral / total_verificados_geral) * 100, 2)
        else:
            perc_geral = 0.0

        # Linha de total geral
        linha_total = {
            'Arquivo': 'TOTAL_GERAL',
            'IDs_Verificados': total_verificados_geral,
            'IDs_Falhados': total_falhados_geral,
            'Percentual_Falha(%)': perc_geral
        }

        # Concatena total geral ao final do resumo
        if not df_resumo.empty:
            df_resumo = pd.concat([df_resumo, pd.DataFrame([linha_total])], ignore_index=True)
        else:
            df_resumo = pd.DataFrame([linha_total])

        # Salva tudo em um único arquivo Excel com duas abas
        caminho_unico = os.path.join(pasta_destino, 'IDs_Falhados_Consolidado.xlsx')
        with pd.ExcelWriter(caminho_unico, engine='openpyxl') as writer:
            df_ids.to_excel(writer, sheet_name='IDs_Falhados', index=False)
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)

        # Mensagens finais
        if not df_ids.empty:
            print(f"\n✔ IDs consolidados salvos em: {caminho_unico}")
            print("   Abas: 'IDs_Falhados' e 'Resumo'.\n")
        else:
            print(f"\nNenhum ID falhado encontrado. Mesmo assim, o resumo foi salvo em: {caminho_unico}\n")

# Executa
if __name__ == '__main__':
    Dados.Abrir_planilha()
