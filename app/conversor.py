import pandas as pd
import re
# from time import sleep

# Constantes para índices e tamanhos de campo
INDICE_CNPJ_EMPRESA = 2
INDICE_PERIODO_ESCRITURACAO = 2
INDICE_CNPJ_FONTE = 0
INDICE_CODIGO_RECEITA = 9
INDICE_IR_RETIDO_NA_FONTE_1708 = 15
INDICE_CSLL_RETIDA_FONTE_5987 = 16
TAMANHO_CNPJ = 14
TAMANHO_PERIODO = 16
TAMANHO_CODIGO_RECEITA = 4
TAMANHO_RENDIMENTO = 14
TAMANHO_ESPACAMENTO = 14
VALOR_PADRAO_PERIODO = "0000000000000000"

indice_rendimento = {
    '1708': INDICE_IR_RETIDO_NA_FONTE_1708,
    '5987': INDICE_CSLL_RETIDA_FONTE_5987
}


def limpar_cnpj(cnpj):
    return re.sub(r'\D', '', str(cnpj))


def formatar_periodo(periodo):
    datas = re.findall(r'\d{2}/\d{2}/\d{4}', str(periodo))
    if len(datas) == 2:
        return datas[0].replace('/', '') + datas[1].replace('/', '')
    return VALOR_PADRAO_PERIODO


def formatar_rendimento(valor):
    if pd.isna(valor) or valor == '':
        valor = 0
    else:
        valor = f"{float(valor):.2f}"
        valor = str(valor).replace('.', '')
        valor = valor.zfill(TAMANHO_RENDIMENTO)
    return valor


def processar_planilha(arquivo_entrada: str, nome_planilha: str, arquivo_saida: str):
    if arquivo_entrada.endswith('.xlsx'):
        engine = "openpyxl"
    else:
        engine = "xlrd"

    try:
        df = pd.read_excel(arquivo_entrada, sheet_name=nome_planilha, engine=engine)

        cnpj_empresa = limpar_cnpj(df.iloc[0, INDICE_CNPJ_EMPRESA])
        periodo_escrituracao = formatar_periodo(df.iloc[1, INDICE_PERIODO_ESCRITURACAO])

        inicio_dados = df[df.iloc[:, 0] == "CNPJ da fonte pagadora"].index[0] + 1
        df_dados = df.iloc[inicio_dados:].reset_index(drop=True)

        linhas_formatadas_1708 = []
        linhas_formatadas_5987 = []

        for _, row in df_dados.iterrows():
            cnpj_fonte = limpar_cnpj(row.iloc[INDICE_CNPJ_FONTE])
            codigo_receita = str(row.iloc[INDICE_CODIGO_RECEITA]).zfill(TAMANHO_CODIGO_RECEITA)

            indice = indice_rendimento.get(codigo_receita)

            if indice is not None:
                rendimento = formatar_rendimento(row.iloc[indice])

                if codigo_receita == '1708':
                    linha = f"R29{cnpj_empresa}{' ' * TAMANHO_ESPACAMENTO}{periodo_escrituracao}{cnpj_fonte}{codigo_receita}00{rendimento}\n "
                    linhas_formatadas_1708.append(linha)
                elif codigo_receita == '5987':
                    linha = f"R36{cnpj_empresa}{' ' * TAMANHO_ESPACAMENTO}{periodo_escrituracao}{cnpj_fonte}{codigo_receita}10{rendimento}\n "
                    linhas_formatadas_5987.append(linha)

        saida_ir_retido_na_fonte = arquivo_saida + '_1708.txt'
        saida_csll_retido_na_fonte = arquivo_saida + '_5987.txt'

        with open(saida_ir_retido_na_fonte, "w", newline="\r\n") as f:
            f.write("".join(linhas_formatadas_1708))

        with open(saida_csll_retido_na_fonte, "w", newline="\r\n") as f:
            f.write("".join(linhas_formatadas_5987))

    except Exception as e:
        raise e
