import os
import win32com.client as win32

PASTA_ARQUIVOS = r"C:\Users\eduar\OneDrive\Documentos\JR EMPREENDIMENTO\Pasta que armazena os arquivos"
PASTA_MACROS = r"C:\Users\eduar\OneDrive\Documentos\JR EMPREENDIMENTO\Macros\EMPRESA 3"

SEQUENCIA_TXT = [
    "1 - Macro Movimentos.txt",
    "2 - Macro Matriculas.txt",
    "3 - Macro Verbas.txt",
    "4 - Macro Ajustar Unidade.txt",
    "5 - Macro Centro de Custo.txt",
    "6 - Macro Modelo.txt",
    "7 - Macro Ajustar Numeros Padrao Protheus.txt",
    "8 - Macro Ajustar Filiais.txt",
    "9 - Macro preencher coluna J e E.txt",
    "10 - Macro Ajustar Horas para Decimal.txt",
    "11 - Atualizar Sequencia.txt",
]

MACROS_PARA_RODAR = [
    "Extrair_Movimentos_Novo_Layout",
    "AtualizarMatriculasPorNome_Embutido",
    "AtualizarCodigosPorNumero_Protheus",
    "DefinirUnidadeVH",
    "Preencher_CC_Por_Matricula",
    "Transformar_Para_Modelo",
    "InserirPontoAntesDos2Ultimos_CORRETO",
    "AtualizarFilialPorMatricula",
    "PreencherColunas",
    "ConverterHorasParaDecimal_ColunaE",
    "AtualizarSequencia",

]


def importar_txt_como_modulo(vb_project, caminho_txt, nome_modulo):
    componente = vb_project.VBComponents.Add(1)
    componente.Name = nome_modulo

    with open(caminho_txt, "r", encoding="utf-8") as f:
        codigo = f.read()

    componente.CodeModule.AddFromString(codigo)
    return componente


def listar_arquivos_excel(pasta):
    extensoes = (".xlsx", ".xlsm", ".xls")
    arquivos = []

    for nome in os.listdir(pasta):
        if nome.startswith("~$"):
            continue

        if nome.lower().endswith(extensoes):
            arquivos.append(os.path.join(pasta, nome))

    return arquivos


def processar_arquivo(excel, caminho_excel):
    print(f"\nProcessando: {caminho_excel}")

    wb = None
    modulos_importados = []

    try:
        wb = excel.Workbooks.Open(caminho_excel, ReadOnly=False)

        if wb.ReadOnly:
            raise PermissionError(f"O arquivo abriu como somente leitura: {caminho_excel}")

        vb_project = wb.VBProject

        for i, nome_txt in enumerate(SEQUENCIA_TXT, start=1):
            caminho_txt = os.path.join(PASTA_MACROS, nome_txt)

            if not os.path.exists(caminho_txt):
                raise FileNotFoundError(f"Macro TXT não encontrada: {caminho_txt}")

            modulo = importar_txt_como_modulo(
                vb_project,
                caminho_txt,
                f"modMacroTemp{i}"
            )

            modulos_importados.append(modulo)

        for nome_macro in MACROS_PARA_RODAR:
            print(f"Rodando macro: {nome_macro}")
            excel.Run(f"'{wb.Name}'!{nome_macro}")

        # Remove os módulos temporários ANTES de salvar
        for modulo in modulos_importados:
            vb_project.VBComponents.Remove(modulo)

        modulos_importados.clear()

        # Salva por cima do arquivo original
        wb.Save()
        wb.Close(SaveChanges=True)

        print(f"Concluído e salvo por cima: {caminho_excel}")

    except Exception as e:
        print(f"Erro no arquivo {caminho_excel}: {e}")

        try:
            if wb is not None:
                vb_project = wb.VBProject

                for modulo in modulos_importados:
                    try:
                        vb_project.VBComponents.Remove(modulo)
                    except:
                        pass

                wb.Close(SaveChanges=False)
        except:
            pass


def main():
    arquivos_excel = listar_arquivos_excel(PASTA_ARQUIVOS)

    if not arquivos_excel:
        print("Nenhum arquivo Excel encontrado na pasta.")
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    try:
        for caminho_excel in arquivos_excel:
            processar_arquivo(excel, caminho_excel)

    finally:
        excel.EnableEvents = True
        excel.DisplayAlerts = True
        excel.Quit()

    print("\nTodos os arquivos foram processados.")


if __name__ == "__main__":
    main()