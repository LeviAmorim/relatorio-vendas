#!/usr/bin/env python3
"""Gera um relatório de vendas em Excel a partir de uma planilha (.xlsx ou .csv).

Uso:
    python gerar_relatorio.py exemplos/vendas_exemplo.xlsx
    python gerar_relatorio.py minhas_vendas.csv -o relatorio_outubro.xlsx

Colunas esperadas na planilha de entrada:
    Data, Produto, Categoria, Vendedor, Quantidade, Preco_Unitario
"""

import argparse
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COLUNAS_OBRIGATORIAS = [
    "data",
    "produto",
    "categoria",
    "vendedor",
    "quantidade",
    "preco_unitario",
]

# Nomes alternativos aceitos nos cabeçalhos (já normalizados).
APELIDOS = {
    "qtd": "quantidade",
    "qtde": "quantidade",
    "preco": "preco_unitario",
    "valor_unitario": "preco_unitario",
}

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

FORMATO_MOEDA = '"R$" #,##0.00'
FORMATO_PERCENTUAL = "0.0%"
FORMATO_INTEIRO = "#,##0"

COR_CABECALHO = "1F3864"
COR_DESTAQUE = "DDEBF7"
BORDA = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


# --------------------------------------------------------------------------
# Leitura e limpeza
# --------------------------------------------------------------------------
def normalizar_nome(nome: str) -> str:
    """'Preço Unitário ' -> 'preco_unitario'."""
    texto = unicodedata.normalize("NFKD", str(nome))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.strip().lower().replace(" ", "_").replace("-", "_")
    return APELIDOS.get(texto, texto)


def para_numero(serie: pd.Series) -> pd.Series:
    """Converte para número, aceitando textos como 'R$ 1.234,50'."""
    if not pd.api.types.is_numeric_dtype(serie):
        texto = serie.astype(str).str.replace(r"[R$\s]", "", regex=True)
        if texto.str.contains(",").any():
            texto = texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        serie = texto
    return pd.to_numeric(serie, errors="coerce")


def ler_planilha(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        sys.exit(f"Erro: arquivo não encontrado: {caminho}")
    sufixo = caminho.suffix.lower()
    if sufixo in (".xlsx", ".xlsm"):
        return pd.read_excel(caminho)
    if sufixo in (".csv", ".txt"):
        return pd.read_csv(caminho, sep=None, engine="python")
    sys.exit(f"Erro: formato '{sufixo}' não suportado. Use .xlsx ou .csv.")


def preparar_dados(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Valida colunas, converte tipos e descarta linhas inválidas.

    Retorna o DataFrame limpo e quantas linhas foram descartadas.
    """
    df = df.rename(columns=normalizar_nome)

    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        sys.exit(
            "Erro: faltam colunas na planilha: "
            + ", ".join(faltando)
            + f"\nColunas encontradas: {', '.join(map(str, df.columns))}"
        )

    df = df[COLUNAS_OBRIGATORIAS].copy()
    total_lido = len(df)

    if not pd.api.types.is_datetime64_any_dtype(df["data"]):
        df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    df["quantidade"] = para_numero(df["quantidade"])
    df["preco_unitario"] = para_numero(df["preco_unitario"])

    for coluna in ("produto", "categoria", "vendedor"):
        df[coluna] = df[coluna].fillna("Não informado").astype(str).str.strip()

    df = df.dropna(subset=["data", "quantidade", "preco_unitario"])
    df = df[(df["quantidade"] > 0) & (df["preco_unitario"] >= 0)]
    descartadas = total_lido - len(df)

    if df.empty:
        sys.exit("Erro: nenhuma linha válida encontrada na planilha.")

    df["total"] = df["quantidade"] * df["preco_unitario"]
    return df.sort_values("data").reset_index(drop=True), descartadas


# --------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------
def resumir_por(df: pd.DataFrame, chave: str) -> pd.DataFrame:
    """Faturamento, unidades, nº de vendas, ticket médio e participação por chave."""
    resumo = (
        df.groupby(chave)
        .agg(faturamento=("total", "sum"), unidades=("quantidade", "sum"), vendas=("total", "size"))
        .reset_index()
        .sort_values("faturamento", ascending=False)
        .reset_index(drop=True)
    )
    resumo["ticket_medio"] = resumo["faturamento"] / resumo["vendas"]
    resumo["participacao"] = resumo["faturamento"] / resumo["faturamento"].sum()
    return resumo


def resumir_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    periodo = df["data"].dt.to_period("M")
    resumo = (
        df.groupby(periodo)
        .agg(faturamento=("total", "sum"), unidades=("quantidade", "sum"), vendas=("total", "size"))
        .reset_index()
        .rename(columns={"data": "periodo"})
        .sort_values("periodo")
        .reset_index(drop=True)
    )
    resumo["mes"] = resumo["periodo"].apply(lambda p: f"{MESES[p.month - 1]}/{p.year}")
    resumo["variacao"] = resumo["faturamento"].pct_change()
    return resumo


# --------------------------------------------------------------------------
# Escrita do Excel
# --------------------------------------------------------------------------
def escrever_tabela(ws, df, colunas, titulos, formatos=None, linha_inicial=1):
    """Escreve um DataFrame como tabela formatada. Retorna a última linha usada."""
    formatos = formatos or {}

    for j, titulo in enumerate(titulos, start=1):
        celula = ws.cell(row=linha_inicial, column=j, value=titulo)
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = BORDA

    for i, linha in enumerate(df[colunas].itertuples(index=False), start=linha_inicial + 1):
        for j, (coluna, valor) in enumerate(zip(colunas, linha), start=1):
            if pd.isna(valor):
                valor = None
            elif isinstance(valor, pd.Timestamp):
                valor = valor.to_pydatetime()
            celula = ws.cell(row=i, column=j, value=valor)
            celula.border = BORDA
            if coluna in formatos:
                celula.number_format = formatos[coluna]

    ultima = linha_inicial + len(df)
    ajustar_larguras(ws, colunas, titulos, df, linha_inicial)
    return ultima


def ajustar_larguras(ws, colunas, titulos, df, linha_inicial):
    for j, (coluna, titulo) in enumerate(zip(colunas, titulos), start=1):
        valores = df[coluna].astype(str).str.len()
        largura = max(len(titulo), int(valores.max()) if len(valores) else 0) + 4
        letra = get_column_letter(j)
        ws.column_dimensions[letra].width = max(ws.column_dimensions[letra].width or 0, min(largura, 40))


def aba_resumo(wb, df, por_produto, por_vendedor, por_mes):
    ws = wb.active
    ws.title = "Resumo"
    ws.sheet_view.showGridLines = False

    inicio, fim = df["data"].min(), df["data"].max()
    ws["A1"] = "Relatório de Vendas"
    ws["A1"].font = Font(size=18, bold=True, color=COR_CABECALHO)
    ws["A2"] = f"Período: {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"
    ws["A3"] = f"Gerado em {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(size=11)
    ws["A3"].font = Font(size=9, italic=True, color="7F7F7F")

    faturamento = df["total"].sum()
    vendas = len(df)
    melhor_mes = por_mes.loc[por_mes["faturamento"].idxmax()]
    pior_mes = por_mes.loc[por_mes["faturamento"].idxmin()]
    top_produto = por_produto.iloc[0]
    top_vendedor = por_vendedor.iloc[0]

    indicadores = [
        ("Faturamento total", faturamento, FORMATO_MOEDA),
        ("Número de vendas", vendas, FORMATO_INTEIRO),
        ("Unidades vendidas", df["quantidade"].sum(), FORMATO_INTEIRO),
        ("Ticket médio", faturamento / vendas, FORMATO_MOEDA),
        ("Produto campeão", top_produto["produto"], None),
        ("  faturamento do produto", top_produto["faturamento"], FORMATO_MOEDA),
        ("Vendedor destaque", top_vendedor["vendedor"], None),
        ("  faturamento do vendedor", top_vendedor["faturamento"], FORMATO_MOEDA),
        ("Melhor mês", melhor_mes["mes"], None),
        ("  faturamento do mês", melhor_mes["faturamento"], FORMATO_MOEDA),
        ("Mês mais fraco", pior_mes["mes"], None),
        ("  faturamento do mês", pior_mes["faturamento"], FORMATO_MOEDA),
    ]

    for i, (rotulo, valor, formato) in enumerate(indicadores, start=5):
        a = ws.cell(row=i, column=1, value=rotulo)
        b = ws.cell(row=i, column=2, value=valor)
        a.border = b.border = BORDA
        b.alignment = Alignment(horizontal="right")
        if not rotulo.startswith(" "):
            a.font = Font(bold=True)
            a.fill = PatternFill("solid", fgColor=COR_DESTAQUE)
            b.font = Font(bold=True)
        if formato:
            b.number_format = formato

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 24


def adicionar_grafico_barras(ws, titulo, rotulos_col, valores_col, linha_cab, linha_fim, ancora, horizontal=False):
    grafico = BarChart()
    grafico.type = "bar" if horizontal else "col"
    grafico.title = titulo
    grafico.legend = None
    grafico.height, grafico.width = 9, 18
    dados = Reference(ws, min_col=valores_col, min_row=linha_cab, max_row=linha_fim)
    categorias = Reference(ws, min_col=rotulos_col, min_row=linha_cab + 1, max_row=linha_fim)
    grafico.add_data(dados, titles_from_data=True)
    grafico.set_categories(categorias)
    grafico.y_axis.numFmt = '"R$" #,##0'
    grafico.x_axis.delete = False
    grafico.y_axis.delete = False
    if horizontal:
        grafico.x_axis.scaling.orientation = "maxMin"  # maior no topo
    ws.add_chart(grafico, ancora)


def aba_por_mes(wb, por_mes):
    ws = wb.create_sheet("Por Mês")
    colunas = ["mes", "faturamento", "variacao", "unidades", "vendas"]
    titulos = ["Mês", "Faturamento", "Variação vs. mês anterior", "Unidades", "Vendas"]
    formatos = {"faturamento": FORMATO_MOEDA, "variacao": FORMATO_PERCENTUAL,
                "unidades": FORMATO_INTEIRO, "vendas": FORMATO_INTEIRO}
    ultima = escrever_tabela(ws, por_mes, colunas, titulos, formatos)
    ws.freeze_panes = "A2"

    grafico = LineChart()
    grafico.title = "Faturamento por mês"
    grafico.legend = None
    grafico.height, grafico.width = 9, 18
    grafico.add_data(Reference(ws, min_col=2, min_row=1, max_row=ultima), titles_from_data=True)
    grafico.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ultima))
    serie = grafico.series[0]
    serie.smooth = False  # sem suavizar: a curva não pode inventar picos
    serie.marker.symbol = "circle"
    serie.marker.size = 6
    grafico.y_axis.numFmt = '"R$" #,##0'
    grafico.x_axis.delete = False
    grafico.y_axis.delete = False
    ws.add_chart(grafico, "H2")


def aba_resumo_por(wb, nome_aba, resumo, chave, rotulo_chave, com_grafico=True):
    ws = wb.create_sheet(nome_aba)
    colunas = [chave, "faturamento", "participacao", "unidades", "vendas", "ticket_medio"]
    titulos = [rotulo_chave, "Faturamento", "Participação", "Unidades", "Vendas", "Ticket médio"]
    formatos = {"faturamento": FORMATO_MOEDA, "participacao": FORMATO_PERCENTUAL,
                "unidades": FORMATO_INTEIRO, "vendas": FORMATO_INTEIRO,
                "ticket_medio": FORMATO_MOEDA}
    ultima = escrever_tabela(ws, resumo, colunas, titulos, formatos)
    ws.freeze_panes = "A2"
    if com_grafico:
        ultima_grafico = min(ultima, 11)  # no máximo os 10 primeiros no gráfico
        adicionar_grafico_barras(
            ws, f"Faturamento por {rotulo_chave.lower()}", 1, 2, 1, ultima_grafico, "I2", horizontal=True
        )


def aba_dados(wb, df):
    ws = wb.create_sheet("Dados")
    colunas = ["data", "produto", "categoria", "vendedor", "quantidade", "preco_unitario", "total"]
    titulos = ["Data", "Produto", "Categoria", "Vendedor", "Quantidade", "Preço unitário", "Total"]
    formatos = {"data": "DD/MM/YYYY", "preco_unitario": FORMATO_MOEDA,
                "total": FORMATO_MOEDA, "quantidade": FORMATO_INTEIRO}
    ultima = escrever_tabela(ws, df, colunas, titulos, formatos)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(colunas))}{ultima}"


def gerar_excel(df, saida: Path):
    por_produto = resumir_por(df, "produto")
    por_categoria = resumir_por(df, "categoria")
    por_vendedor = resumir_por(df, "vendedor")
    por_mes = resumir_por_mes(df)

    wb = Workbook()
    aba_resumo(wb, df, por_produto, por_vendedor, por_mes)
    aba_por_mes(wb, por_mes)
    aba_resumo_por(wb, "Por Produto", por_produto, "produto", "Produto")
    aba_resumo_por(wb, "Por Categoria", por_categoria, "categoria", "Categoria")
    aba_resumo_por(wb, "Por Vendedor", por_vendedor, "vendedor", "Vendedor")
    aba_dados(wb, df)
    wb.save(saida)


# --------------------------------------------------------------------------
# Linha de comando
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Gera um relatório de vendas em Excel a partir de uma planilha.")
    parser.add_argument("entrada", type=Path, help="planilha de vendas (.xlsx ou .csv)")
    parser.add_argument("-o", "--saida", type=Path, default=Path("relatorio_vendas.xlsx"),
                        help="arquivo Excel de saída (padrão: relatorio_vendas.xlsx)")
    args = parser.parse_args()

    bruto = ler_planilha(args.entrada)
    df, descartadas = preparar_dados(bruto)
    gerar_excel(df, args.saida)

    print(f"Relatório gerado: {args.saida}")
    print(f"  Vendas processadas: {len(df)}")
    total_br = f"{df['total'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    print(f"  Faturamento total:  R$ {total_br}")
    if descartadas:
        print(f"  Atenção: {descartadas} linha(s) inválida(s) foram ignoradas.")


if __name__ == "__main__":
    main()
