1 ° PASSO CRIAR UM AMBIENTE VIRTUAL UTILIZANDO O TERMINAL (LEMBRAR DE BAIXAR O PYTHON COM O PATH ATIVO NA HORA DA INSTAÇÃO)
- python -m venv nome_do_ambiente
- nome_do_ambiente\Scripts\activate
- pip install openpyxl
- pip install pyinstaller

2° PASSO TRANSFORMANDO O ARQUIVO EM EXECUTÁVEL
- pyinstaller --onefile seu_script.py



LEMBRANDO QUE AS PLANILHAS DEVEM ESTAR DENTRO DA PASTA DO EXECUTÁVEL, QUANDO FAZ TODOS OS PASSOS, O EXECUTÁVEL FICA NA PASTA DIST.



===================================== LEMBRAR DE SEMPRE SEGUIR O PASSO A PASSO ====================================================




import openpyxl

def carregar_planilha(nome_arquivo):
    try:
        workbook = openpyxl.load_workbook(nome_arquivo)
        return workbook
    except FileNotFoundError:
        print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

def procurar_palavra_em_aba(sheet, palavra):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and palavra.lower() == str(cell.value).lower():
                return row
    return None

def procurar_palavra_em_planilha(workbook, palavra):
    for sheet in workbook.worksheets:
        resultado = procurar_palavra_em_aba(sheet, palavra)
        if resultado:
            return sheet.title, resultado
    return None, None

def main():
    # Lista de arquivos de planilhas que você deseja verificar
    arquivos_planilhas = ["nome da planilha.xlsx", "nome da planilha", "nome da planilha.xlsx"] # AQUI É O NOME DAS PLANILHAS QUE QUER PESQUISAR, PODE SER VÁRIAS PLANILHAS

    while True:
        palavra = input("Coloque o serial que deseja buscar (ou 'sair' para encerrar): ")
        if palavra.lower() == 'sair':
            print("Encerrando o programa.")
            break

        encontrado = False
        for nome_arquivo in arquivos_planilhas:
            workbook = carregar_planilha(nome_arquivo)
            if workbook:
                aba, resultado = procurar_palavra_em_planilha(workbook, palavra)
                if resultado:
                    print(f"Serial '{palavra}' encontrado no arquivo '{nome_arquivo}', aba '{aba}':")
                    for cell in resultado:
                        print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")
                        print(cell.value)
                    encontrado = True
                    break
            else:
                print(f"Não foi possível carregar o arquivo '{nome_arquivo}'.")

        if not encontrado:
            print(f"Serial '{palavra}' não encontrado em nenhuma planilha.")

if __name__ == "__main__":
    main()
