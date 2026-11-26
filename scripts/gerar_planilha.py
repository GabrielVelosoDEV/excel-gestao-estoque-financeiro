"""Gera a planilha de gestão de estoque e controle financeiro com dados de exemplo.

A empresa "Loja Exemplo" é FICTÍCIA. Os dados de exemplo (produtos, movimentações e lançamentos)
são gerados por este script com semente fixa, só para demonstrar as fórmulas, validações e o painel.
Toda a lógica da planilha (estoque, margens, fluxo de caixa, projeção, alertas) está em fórmulas do
Excel, então ao apagar os exemplos e digitar dados reais tudo recalcula.

Uso:
    python scripts/gerar_planilha.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parents[1] / "planilha" / "Gestao_Estoque_Financeiro.xlsx"
ANO = 2025
LINHAS = 5000                    # capacidade da tabela de movimentações (linhas com fórmulas prontas)
random.seed(42)

# ----------------------------------------------------------------- estilos
FONTE = "Arial"
AZUL = "0000FF"                  # entradas digitadas pelo usuário
TITULO = Font(name=FONTE, size=16, bold=True, color="1F2937")
SUB = Font(name=FONTE, size=10, italic=True, color="6B7280")
CAB = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
NORMAL = Font(name=FONTE, size=10, color="111827")
ENTRADA = Font(name=FONTE, size=10, color=AZUL)
NEGRITO = Font(name=FONTE, size=10, bold=True, color="111827")
FILL_CAB = PatternFill("solid", fgColor="1F3A5F")
FILL_ENTRADA = PatternFill("solid", fgColor="FFF9C4")
FILL_TOTAL = PatternFill("solid", fgColor="E5E7EB")
FILL_KPI = PatternFill("solid", fgColor="EEF2FF")
LINHA = Side(style="thin", color="D1D5DB")
BORDA = Border(bottom=LINHA)
BRL = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
BRL0 = 'R$ #,##0;[Red]-R$ #,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
INT = '#,##0;[Red]-#,##0;"-"'
DATA = "dd/mm/yyyy"
MES = "mmm/yyyy"


def cabecalho(ws, titulo: str, subtitulo: str, colunas: list[tuple[str, int]], linha: int = 4) -> None:
    ws["A1"] = titulo
    ws["A1"].font = TITULO
    ws["A2"] = subtitulo
    ws["A2"].font = SUB
    for i, (nome, largura) in enumerate(colunas, start=1):
        c = ws.cell(row=linha, column=i, value=nome)
        c.font, c.fill = CAB, FILL_CAB
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.row_dimensions[linha].height = 30
    ws.freeze_panes = ws.cell(row=linha + 1, column=1)
    ws.sheet_view.showGridLines = False


def fmt(cell, numero: str | None = None, fonte: Font = NORMAL) -> None:
    cell.font = fonte
    cell.border = BORDA
    if numero:
        cell.number_format = numero


# ----------------------------------------------------------------- dados de exemplo (fictícios)
PRODUTOS = [  # código, descrição, categoria, custo, preço, estoque mínimo, venda média/mês
    ("P001", "Caderno universitário 200 fls", "Papelaria", 12.50, 24.90, 40, 60),
    ("P002", "Caneta esferográfica azul (cx 50)", "Papelaria", 28.00, 49.90, 15, 18),
    ("P003", "Papel A4 75g (resma 500 fls)", "Papelaria", 21.00, 34.90, 60, 90),
    ("P004", "Marca-texto (kit 6 cores)", "Papelaria", 9.80, 19.90, 25, 30),
    ("P005", "Agenda executiva", "Papelaria", 18.00, 39.90, 10, 12),
    ("P006", "Mouse sem fio", "Informática", 32.00, 69.90, 15, 20),
    ("P007", "Teclado USB ABNT2", "Informática", 45.00, 89.90, 10, 12),
    ("P008", "Pendrive 64 GB", "Informática", 24.00, 49.90, 20, 28),
    ("P009", "Headset com microfone", "Informática", 68.00, 139.90, 8, 9),
    ("P010", "Cabo HDMI 2 m", "Informática", 11.00, 29.90, 15, 22),
    ("P011", "Monitor 24 pol.", "Informática", 520.00, 849.90, 3, 4),
    ("P012", "Webcam Full HD", "Informática", 95.00, 179.90, 5, 6),
    ("P013", "Cadeira de escritório", "Escritório", 310.00, 599.90, 3, 3),
    ("P014", "Organizador de mesa", "Escritório", 16.00, 34.90, 10, 11),
    ("P015", "Grampeador de mesa", "Escritório", 14.00, 29.90, 10, 9),
    ("P016", "Quadro branco 90x60", "Escritório", 85.00, 169.90, 3, 4),
    ("P017", "Garrafa térmica 1 L", "Utilidades", 27.00, 59.90, 8, 10),
    ("P018", "Luminária de mesa LED", "Utilidades", 48.00, 99.90, 5, 7),
    ("P019", "Filtro de linha 6 tomadas", "Utilidades", 22.00, 44.90, 10, 14),
    ("P020", "Mochila para notebook", "Utilidades", 75.00, 159.90, 5, 8),
]
SAZONAL = {1: 1.6, 2: 1.4, 3: 1.0, 4: 0.9, 5: 0.9, 6: 0.8, 7: 1.1, 8: 1.0, 9: 0.9, 10: 1.0, 11: 1.3, 12: 1.5}
CATEGORIAS = ["Papelaria", "Informática", "Escritório", "Utilidades"]
CAT_DESPESA = ["Aluguel", "Salários e encargos", "Energia e internet", "Marketing", "Contabilidade", "Manutenção", "Outras despesas"]
CAT_RECEITA = ["Serviços", "Outras receitas"]


def gerar_movimentos() -> list[tuple]:
    mov, estoque = [], {}
    for cod, _, _, custo, preco, minimo, _ in PRODUTOS:
        q = minimo * 4
        mov.append((date(ANO, 1, 2), cod, "Entrada", q, custo))
        estoque[cod] = q
    d = date(ANO, 1, 3)
    while d <= date(ANO, 12, 30):
        if d.weekday() < 6:  # segunda a sábado
            for cod, _, _, custo, preco, minimo, media in PRODUTOS:
                lam = media * 1.8 * SAZONAL[d.month] / 26
                q = sum(1 for _ in range(int(lam * 4) + 1) if random.random() < lam / (int(lam * 4) + 1))
                if q and estoque[cod] >= q:
                    mov.append((d, cod, "Saída", q, preco))
                    estoque[cod] -= q
                if estoque[cod] < minimo and random.random() < 0.35:   # reposição
                    q = minimo * 3 - estoque[cod]
                    mov.append((d, cod, "Entrada", q, round(custo * random.uniform(0.97, 1.05), 2)))
                    estoque[cod] += q
        d += timedelta(days=1)
    return sorted(mov, key=lambda m: (m[0], m[1]))


def gerar_financeiro() -> list[tuple]:
    fin = []
    for m in range(1, 13):
        v = date(ANO, m, 5)
        fin += [
            (date(ANO, m, 1), "Aluguel da loja", "Aluguel", "Despesa", 3800.00, v),
            (date(ANO, m, 1), "Folha de pagamento", "Salários e encargos", "Despesa", 9800.00 + (4900.0 if m == 12 else 0), v),
            (date(ANO, m, 10), "Conta de energia e internet", "Energia e internet", "Despesa", round(random.uniform(780, 1150), 2), date(ANO, m, 15)),
            (date(ANO, m, 12), "Anúncios e redes sociais", "Marketing", "Despesa", round(random.uniform(700, 1600) * SAZONAL[m], 2), date(ANO, m, 20)),
            (date(ANO, m, 15), "Honorários contábeis", "Contabilidade", "Despesa", 650.00, date(ANO, m, 20)),
            (date(ANO, m, 18), "Configuração e instalação de equipamentos", "Serviços", "Receita", round(random.uniform(900, 2600), 2), date(ANO, m, 25)),
        ]
        if m in (3, 7, 10):
            fin.append((date(ANO, m, 22), "Manutenção do ar-condicionado", "Manutenção", "Despesa", round(random.uniform(350, 800), 2), date(ANO, m, 28)))
    fin.sort(key=lambda r: r[0])
    out = []
    for r in fin:
        pendente = r[5] >= date(ANO, 12, 20)
        out.append(r + ("Pendente" if pendente else "Pago",))
    return out


# ----------------------------------------------------------------- planilhas
def aba_leiame(wb):
    ws = wb.active
    ws.title = "LEIA-ME"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 95
    ws["B2"] = "Gestão de Estoque e Controle Financeiro"
    ws["B2"].font = TITULO
    ws["B3"] = "Loja Exemplo — empresa FICTÍCIA. Todos os dados de exemplo foram gerados para demonstração."
    ws["B3"].font = Font(name=FONTE, size=10, bold=True, color="B91C1C")
    linhas = [
        ("Como usar", ""),
        ("1. Parametros", "Ajuste o ano, o saldo inicial de caixa, a alíquota de impostos e as listas de categorias."),
        ("2. Produtos", "Cadastre os produtos (código, custo, preço, estoque mínimo). Margem, estoque atual e status são calculados."),
        ("3. Lancamento", "Preencha a área amarela e rode a macro RegistrarMovimentacao (Alt+F8) para gravar entradas e saídas com validação."),
        ("4. Movimentacoes", "Histórico de entradas e saídas. Pode ser digitado direto, respeitando as listas de validação."),
        ("5. Financeiro", "Despesas e receitas que não são venda de mercadoria (aluguel, folha, serviços...)."),
        ("6. Fluxo_Caixa", "Calculado: vendas, compras, impostos e despesas por mês, com saldo acumulado."),
        ("7. Projecao", "Calculado: projeção dos próximos 6 meses a partir da média recente e da taxa de crescimento."),
        ("8. Dashboard", "Indicadores, alertas de reposição e gráficos."),
        ("", ""),
        ("Legenda", ""),
        ("Texto azul", "Valor digitado (entrada). Pode ser alterado."),
        ("Fundo amarelo", "Premissa ou campo de preenchimento."),
        ("Texto preto", "Fórmula. Não edite."),
        ("", ""),
        ("Macros (VBA)", "Os módulos estão na pasta vba/ do repositório. Importe-os no Editor do VBA (Alt+F11 > Arquivo > Importar) "
                         "e salve como .xlsm. Veja o README para a lista de macros."),
    ]
    for i, (a, b) in enumerate(linhas, start=5):
        ws.cell(row=i, column=2, value=a).font = NEGRITO if b == "" or a.startswith(("Como", "Legenda", "Macros")) else NORMAL
        ws.cell(row=i, column=3, value=b).font = NORMAL
        ws.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    ws["B16"].font = ENTRADA
    ws["B17"].fill = FILL_ENTRADA


def aba_parametros(wb):
    ws = wb.create_sheet("Parametros")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Parâmetros"
    ws["A1"].font = TITULO
    ws["A2"] = "Premissas usadas nas fórmulas. Altere os valores em azul."
    ws["A2"].font = SUB
    for col, w in zip("ABCDEFG", (34, 18, 4, 24, 4, 26, 4)):
        ws.column_dimensions[col].width = w
    premissas = [
        ("Empresa", "Loja Exemplo (fictícia)", None, None),
        ("Ano de referência", ANO, "0", "Ano exibido no fluxo de caixa e no painel."),
        ("Data de referência", date(ANO, 12, 22), DATA, "Usada para identificar contas vencidas."),
        ("Saldo inicial de caixa", 25000.00, BRL, "Saldo em 1º de janeiro. Premissa do exemplo."),
        ("Alíquota de impostos sobre vendas", 0.06, PCT, "Premissa do exemplo (faixa inicial do Simples Nacional para comércio). Ajuste à sua realidade."),
        ("Crescimento mensal projetado", 0.02, PCT, "Premissa usada na aba Projecao."),
        ("Meses para média da projeção", 3, "0", "Quantos meses recentes compõem a base da projeção."),
        ("Cobertura de reposição (x mínimo)", 3, "0", "Sugestão de compra repõe o estoque até N vezes o mínimo."),
    ]
    ws["A4"], ws["B4"] = "Premissa", "Valor"
    for c in (ws["A4"], ws["B4"]):
        c.font, c.fill = CAB, FILL_CAB
    for i, (nome, valor, formato, nota) in enumerate(premissas, start=5):
        ws.cell(row=i, column=1, value=nome).font = NORMAL
        c = ws.cell(row=i, column=2, value=valor)
        c.font, c.fill = ENTRADA, FILL_ENTRADA
        if formato:
            c.number_format = formato
        if nota:
            c.comment = Comment(nota, "Gabriel Veloso")
    listas = {"D": ("Categorias de produto", CATEGORIAS), "F": ("Categorias financeiras", CAT_DESPESA + CAT_RECEITA)}
    for col, (titulo, itens) in listas.items():
        ws[f"{col}4"] = titulo
        ws[f"{col}4"].font, ws[f"{col}4"].fill = CAB, FILL_CAB
        for j, it in enumerate(itens, start=5):
            ws[f"{col}{j}"] = it
            ws[f"{col}{j}"].font = ENTRADA
    ws["D15"], ws["D16"], ws["D17"] = "Tipos de movimento", "Entrada", "Saída"
    ws["D18"], ws["D19"], ws["D20"] = "Tipos financeiros", "Receita", "Despesa"
    ws["D21"], ws["D22"], ws["D23"] = "Status", "Pago", "Pendente"
    for r in (15, 18, 21):
        ws[f"D{r}"].font, ws[f"D{r}"].fill = CAB, FILL_CAB
    for r in (16, 17, 19, 20, 22, 23):
        ws[f"D{r}"].font = NORMAL
    # nomes usados nas fórmulas
    from openpyxl.workbook.defined_name import DefinedName
    nomes = {
        "pAno": "$B$6", "pDataRef": "$B$7", "pSaldoInicial": "$B$8", "pAliquota": "$B$9",
        "pCrescimento": "$B$10", "pMesesMedia": "$B$11", "pCobertura": "$B$12",
        "lstCategorias": "$D$5:$D$12", "lstCatFinanceiras": "$F$5:$F$14",
        "lstTipoMov": "$D$16:$D$17", "lstTipoFin": "$D$19:$D$20", "lstStatus": "$D$22:$D$23",
    }
    for n, ref in nomes.items():
        wb.defined_names[n] = DefinedName(n, attr_text=f"Parametros!{ref}")


def aba_produtos(wb):
    ws = wb.create_sheet("Produtos")
    cols = [("Código", 9), ("Descrição", 34), ("Categoria", 14), ("Custo unit. (R$)", 13), ("Preço venda (R$)", 13),
            ("Margem bruta", 11), ("Markup", 10), ("Estoque mínimo", 11), ("Entradas", 10), ("Saídas", 10),
            ("Estoque atual", 11), ("Valor em estoque (R$)", 15), ("Status", 14), ("Sugestão de compra (un.)", 13),
            ("Vendas no ano (R$)", 15), ("Nº do alerta", 9)]
    cabecalho(ws, "Cadastro de produtos", "Colunas em azul são digitadas; as demais são calculadas.", cols)
    mov = "Movimentacoes"
    for i, (cod, desc, cat, custo, preco, minimo, _) in enumerate(PRODUTOS, start=5):
        for col, v in zip("ABCDEH", (cod, desc, cat, custo, preco, minimo)):
            fmt(ws[f"{col}{i}"], BRL if col in "DE" else (INT if col == "H" else None), ENTRADA)
            ws[f"{col}{i}"] = v
        f = {
            "F": (f"=IFERROR((E{i}-D{i})/E{i},0)", PCT),
            "G": (f"=IFERROR(E{i}/D{i}-1,0)", PCT),
            "I": (f'=SUMIFS({mov}!$E$5:$E${LINHAS},{mov}!$B$5:$B${LINHAS},A{i},{mov}!$D$5:$D${LINHAS},"Entrada")', INT),
            "J": (f'=SUMIFS({mov}!$E$5:$E${LINHAS},{mov}!$B$5:$B${LINHAS},A{i},{mov}!$D$5:$D${LINHAS},"Saída")', INT),
            "K": (f"=I{i}-J{i}", INT),
            "L": (f"=K{i}*D{i}", BRL),
            "M": (f'=IF(K{i}<=0,"SEM ESTOQUE",IF(K{i}<H{i},"REPOR","OK"))', None),
            "N": (f"=IF(K{i}<H{i},H{i}*pCobertura-K{i},0)", INT),
            "O": (f'=SUMIFS({mov}!$G$5:$G${LINHAS},{mov}!$B$5:$B${LINHAS},A{i},{mov}!$D$5:$D${LINHAS},"Saída",{mov}!$H$5:$H${LINHAS},pAno)', BRL),
            "P": (f'=IF(M{i}="OK","",COUNTIF(M$5:M{i},"<>OK"))', "0"),
        }
        for col, (formula, numero) in f.items():
            ws[f"{col}{i}"] = formula
            fmt(ws[f"{col}{i}"], numero)
        ws[f"M{i}"].alignment = Alignment(horizontal="center")
    fim = 4 + len(PRODUTOS)
    t = fim + 1
    ws[f"A{t}"] = "Total"
    for col, numero in (("I", INT), ("J", INT), ("K", INT), ("L", BRL), ("O", BRL)):
        ws[f"{col}{t}"] = f"=SUM({col}5:{col}{fim})"
        fmt(ws[f"{col}{t}"], numero, NEGRITO)
    ws[f"F{t}"] = f"=IFERROR(SUMPRODUCT(J5:J{fim},E5:E{fim}-D5:D{fim})/SUMPRODUCT(J5:J{fim},E5:E{fim}),0)"
    fmt(ws[f"F{t}"], PCT, NEGRITO)
    ws[f"F{t}"].comment = Comment("Margem média ponderada pelas unidades vendidas.", "Gabriel Veloso")
    for c in ws[t]:
        c.fill = FILL_TOTAL
    ws[f"A{t}"].font = NEGRITO
    dv = DataValidation(type="list", formula1="=lstCategorias", allow_blank=True)
    dv.error, dv.errorTitle = "Escolha uma categoria da lista em Parametros.", "Categoria inválida"
    ws.add_data_validation(dv)
    dv.add(f"C5:C{fim + 50}")
    dvn = DataValidation(type="decimal", operator="greaterThan", formula1="0")
    dvn.error, dvn.errorTitle = "Informe um valor maior que zero.", "Valor inválido"
    ws.add_data_validation(dvn)
    dvn.add(f"D5:E{fim + 50}")
    red = PatternFill("solid", fgColor="FEE2E2")
    amber = PatternFill("solid", fgColor="FEF3C7")
    ws.conditional_formatting.add(f"M5:M{fim}", CellIsRule(operator="equal", formula=['"SEM ESTOQUE"'], fill=red, font=Font(name=FONTE, bold=True, color="B91C1C")))
    ws.conditional_formatting.add(f"M5:M{fim}", CellIsRule(operator="equal", formula=['"REPOR"'], fill=amber, font=Font(name=FONTE, bold=True, color="92400E")))
    ws.conditional_formatting.add(f"F5:F{fim}", CellIsRule(operator="lessThan", formula=["0.3"], font=Font(name=FONTE, color="B91C1C")))
    return fim


def aba_movimentacoes(wb, movs):
    ws = wb.create_sheet("Movimentacoes")
    cols = [("Data", 12), ("Código", 9), ("Descrição", 34), ("Tipo", 10), ("Quantidade", 11),
            ("Valor unit. (R$)", 13), ("Valor total (R$)", 14), ("Ano", 7), ("Mês", 11)]
    cabecalho(ws, "Movimentações de estoque", "Entradas (compras) e saídas (vendas). Valor unitário = custo na entrada e preço na saída.", cols)
    for i, (d, cod, tipo, q, vu) in enumerate(movs, start=5):
        for col, v, numero in (("A", d, DATA), ("B", cod, None), ("D", tipo, None), ("E", q, INT), ("F", vu, BRL)):
            ws[f"{col}{i}"] = v
            fmt(ws[f"{col}{i}"], numero, ENTRADA)
    # fórmulas pré-preenchidas em toda a capacidade da tabela
    for i in range(5, LINHAS + 1):
        ws[f"C{i}"] = f'=IF(B{i}="","",IFERROR(INDEX(Produtos!$B:$B,MATCH(B{i},Produtos!$A:$A,0)),"CÓDIGO NÃO CADASTRADO"))'
        ws[f"G{i}"] = f'=IF(E{i}="","",E{i}*F{i})'
        ws[f"H{i}"] = f'=IF(A{i}="","",YEAR(A{i}))'
        ws[f"I{i}"] = f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))'
        for col, numero in (("C", None), ("G", BRL), ("H", "0"), ("I", MES)):
            fmt(ws[f"{col}{i}"], numero)
    dv = DataValidation(type="list", formula1="=lstTipoMov")
    dv.error, dv.errorTitle = "Use Entrada ou Saída.", "Tipo inválido"
    ws.add_data_validation(dv)
    dv.add(f"D5:D{LINHAS}")
    dvq = DataValidation(type="whole", operator="greaterThan", formula1="0")
    dvq.error, dvq.errorTitle = "A quantidade deve ser um número inteiro maior que zero.", "Quantidade inválida"
    ws.add_data_validation(dvq)
    dvq.add(f"E5:E{LINHAS}")
    dvc = DataValidation(type="list", formula1=f"=Produtos!$A$5:$A$100")
    dvc.error, dvc.errorTitle = "Código não encontrado no cadastro de produtos.", "Código inválido"
    ws.add_data_validation(dvc)
    dvc.add(f"B5:B{LINHAS}")
    ws.conditional_formatting.add(f"C5:C{LINHAS}", CellIsRule(operator="equal", formula=['"CÓDIGO NÃO CADASTRADO"'], font=Font(name=FONTE, bold=True, color="B91C1C")))
    ws.auto_filter.ref = f"A4:I{4 + len(movs)}"


def aba_financeiro(wb, fin):
    ws = wb.create_sheet("Financeiro")
    cols = [("Data", 12), ("Descrição", 38), ("Categoria", 20), ("Tipo", 10), ("Valor (R$)", 13),
            ("Vencimento", 12), ("Status", 11), ("Mês", 11), ("Situação", 12)]
    cabecalho(ws, "Lançamentos financeiros", "Receitas e despesas que não são compra ou venda de mercadoria.", cols)
    for i, (d, desc, cat, tipo, valor, venc, status) in enumerate(fin, start=5):
        for col, v, numero in (("A", d, DATA), ("B", desc, None), ("C", cat, None), ("D", tipo, None),
                               ("E", valor, BRL), ("F", venc, DATA), ("G", status, None)):
            ws[f"{col}{i}"] = v
            fmt(ws[f"{col}{i}"], numero, ENTRADA)
    for i in range(5, 601):
        ws[f"H{i}"] = f'=IF(A{i}="","",DATE(YEAR(A{i}),MONTH(A{i}),1))'
        ws[f"I{i}"] = f'=IF(G{i}="","",IF(AND(G{i}="Pendente",F{i}<pDataRef),"VENCIDA",IF(G{i}="Pendente","A vencer","Quitada")))'
        fmt(ws[f"H{i}"], MES)
        fmt(ws[f"I{i}"])
    for rng, nome in (("C5:C600", "lstCatFinanceiras"), ("D5:D600", "lstTipoFin"), ("G5:G600", "lstStatus")):
        dv = DataValidation(type="list", formula1=f"={nome}")
        dv.error, dv.errorTitle = "Escolha um valor da lista.", "Valor inválido"
        ws.add_data_validation(dv)
        dv.add(rng)
    ws.conditional_formatting.add("I5:I600", CellIsRule(operator="equal", formula=['"VENCIDA"'], fill=PatternFill("solid", fgColor="FEE2E2"), font=Font(name=FONTE, bold=True, color="B91C1C")))
    ws.auto_filter.ref = f"A4:I{4 + len(fin)}"


def aba_fluxo(wb):
    ws = wb.create_sheet("Fluxo_Caixa")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Fluxo de caixa mensal"
    ws["A1"].font = TITULO
    ws["A2"] = "Calculado a partir de Movimentacoes e Financeiro. Ano definido em Parametros."
    ws["A2"].font = SUB
    ws.column_dimensions["A"].width = 30
    ws["A4"] = "R$"
    ws["A4"].font, ws["A4"].fill = CAB, FILL_CAB
    for m in range(1, 13):
        col = get_column_letter(m + 1)
        ws.column_dimensions[col].width = 12
        c = ws[f"{col}4"]
        c.value = f"=DATE(pAno,{m},1)"
        c.number_format = "mmm"
        c.font, c.fill = CAB, FILL_CAB
        c.alignment = Alignment(horizontal="center")
    ws["N4"] = "Total do ano"
    ws["N4"].font, ws["N4"].fill = CAB, FILL_CAB
    ws.column_dimensions["N"].width = 14
    mv, fi = "Movimentacoes", "Financeiro"
    linhas = [
        (5, "Vendas de mercadorias", lambda c: f'=SUMIFS({mv}!$G$5:$G${LINHAS},{mv}!$D$5:$D${LINHAS},"Saída",{mv}!$I$5:$I${LINHAS},{c}$4)'),
        (6, "Outras receitas", lambda c: f'=SUMIFS({fi}!$E$5:$E$600,{fi}!$D$5:$D$600,"Receita",{fi}!$H$5:$H$600,{c}$4)'),
        (7, "Total de entradas", lambda c: f"={c}5+{c}6"),
        (9, "Compras de mercadorias", lambda c: f'=SUMIFS({mv}!$G$5:$G${LINHAS},{mv}!$D$5:$D${LINHAS},"Entrada",{mv}!$I$5:$I${LINHAS},{c}$4)'),
        (10, "Impostos sobre vendas", lambda c: f"={c}5*pAliquota"),
        (11, "Despesas operacionais", lambda c: f'=SUMIFS({fi}!$E$5:$E$600,{fi}!$D$5:$D$600,"Despesa",{fi}!$H$5:$H$600,{c}$4)'),
        (12, "Total de saídas", lambda c: f"={c}9+{c}10+{c}11"),
        (14, "Saldo do mês", lambda c: f"={c}7-{c}12"),
        (15, "Saldo acumulado", lambda c: "=pSaldoInicial+B14" if c == "B" else f"={get_column_letter(ord(c) - 65)}15+{c}14"),
        (17, "Resultado operacional (sem compras)", lambda c: f"={c}5*(1-pAliquota)-{c}5*(1-Produtos!$F${4 + len(PRODUTOS) + 1})+{c}6-{c}11"),
    ]
    for r, nome, f in linhas:
        ws[f"A{r}"] = nome
        ws[f"A{r}"].font = NEGRITO if r in (7, 12, 14, 15, 17) else NORMAL
        for m in range(1, 13):
            col = get_column_letter(m + 1)
            ws[f"{col}{r}"] = f(col)
            fmt(ws[f"{col}{r}"], BRL0, NEGRITO if r in (7, 12, 14, 15, 17) else NORMAL)
        if r != 15:
            ws[f"N{r}"] = f"=SUM(B{r}:M{r})"
        else:
            ws[f"N{r}"] = "=M15"
        fmt(ws[f"N{r}"], BRL0, NEGRITO)
        if r in (7, 12, 14, 15):
            for c in ws[r][:14]:
                c.fill = FILL_TOTAL
    ws["A17"].comment = Comment("Vendas líquidas de impostos menos o custo das mercadorias vendidas (pela margem média "
                                "ponderada da aba Produtos), mais outras receitas, menos despesas operacionais.", "Gabriel Veloso")
    ws["A8"], ws["A13"], ws["A16"] = "", "", ""
    ws.conditional_formatting.add("B14:N15", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONTE, bold=True, color="B91C1C")))
    ws.freeze_panes = "B5"
    # gráfico entradas x saídas
    ch = BarChart()
    ch.type = "col"
    ch.title = "Entradas x saídas de caixa"
    ch.y_axis.title = "R$"
    ch.add_data(Reference(ws, min_col=1, max_col=13, min_row=7, max_row=7), from_rows=True, titles_from_data=True)
    ch.add_data(Reference(ws, min_col=1, max_col=13, min_row=12, max_row=12), from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=13, min_row=4, max_row=4))
    ch.height, ch.width = 8, 22
    _ajustar(ch)
    ws.add_chart(ch, "A20")
    return ws


def _ajustar(ch):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.overlay = False
        ch.legend.position = "b"
    if ch.title is not None:
        ch.title.overlay = False
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
        if ax.title is not None:
            ax.title.overlay = False


def aba_projecao(wb):
    ws = wb.create_sheet("Projecao")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Projeção dos próximos 6 meses"
    ws["A1"].font = TITULO
    ws["A2"] = "Base: média dos últimos N meses do fluxo de caixa (Parametros) crescendo à taxa mensal projetada."
    ws["A2"].font = SUB
    ws.column_dimensions["A"].width = 34
    ws["A4"], ws["B4"] = "Base de cálculo", "Valor"
    for c in (ws["A4"], ws["B4"]):
        c.font, c.fill = CAB, FILL_CAB
    ws.column_dimensions["B"].width = 16
    base = [
        ("Média mensal de vendas", "=AVERAGE(OFFSET(Fluxo_Caixa!$M$5,0,1-pMesesMedia,1,pMesesMedia))", BRL),
        ("Média mensal de outras receitas", "=AVERAGE(OFFSET(Fluxo_Caixa!$M$6,0,1-pMesesMedia,1,pMesesMedia))", BRL),
        ("Média mensal de despesas operacionais", "=AVERAGE(OFFSET(Fluxo_Caixa!$M$11,0,1-pMesesMedia,1,pMesesMedia))", BRL),
        ("Margem bruta média ponderada", f"=Produtos!$F${4 + len(PRODUTOS) + 1}", PCT),
        ("Crescimento mensal projetado", "=pCrescimento", PCT),
        ("Alíquota de impostos", "=pAliquota", PCT),
    ]
    for i, (nome, f, numero) in enumerate(base, start=5):
        ws[f"A{i}"] = nome
        ws[f"A{i}"].font = NORMAL
        ws[f"B{i}"] = f
        fmt(ws[f"B{i}"], numero, Font(name=FONTE, size=10, color="008000") if i >= 8 else NORMAL)
    ws["A13"] = "R$"
    ws["A13"].font, ws["A13"].fill = CAB, FILL_CAB
    for n in range(1, 7):
        col = get_column_letter(n + 1)
        ws.column_dimensions[col].width = 13
        ws[f"{col}13"] = f"=DATE(pAno+1,{n},1)"
        ws[f"{col}13"].number_format = MES
        ws[f"{col}13"].font, ws[f"{col}13"].fill = CAB, FILL_CAB
    ws["H13"] = "Total"
    ws["H13"].font, ws["H13"].fill = CAB, FILL_CAB
    ws.column_dimensions["H"].width = 14
    linhas = [
        (14, "Vendas projetadas", lambda c, n: f"=$B$5*(1+$B$9)^{n}"),
        (15, "(-) Custo das mercadorias vendidas", lambda c, n: f"=-{c}14*(1-$B$8)"),
        (16, "(-) Impostos sobre vendas", lambda c, n: f"=-{c}14*$B$10"),
        (17, "(+) Outras receitas", lambda c, n: "=$B$6"),
        (18, "(-) Despesas operacionais", lambda c, n: "=-$B$7"),
        (19, "Resultado projetado", lambda c, n: f"=SUM({c}14:{c}18)"),
        (20, "Margem líquida projetada", lambda c, n: f"=IFERROR({c}19/{c}14,0)"),
    ]
    for r, nome, f in linhas:
        ws[f"A{r}"] = nome
        ws[f"A{r}"].font = NEGRITO if r in (19, 20) else NORMAL
        for n in range(1, 7):
            col = get_column_letter(n + 1)
            ws[f"{col}{r}"] = f(col, n)
            fmt(ws[f"{col}{r}"], PCT if r == 20 else BRL0, NEGRITO if r in (19, 20) else NORMAL)
        ws[f"H{r}"] = "=IFERROR(H19/H14,0)" if r == 20 else f"=SUM(B{r}:G{r})"
        fmt(ws[f"H{r}"], PCT if r == 20 else BRL0, NEGRITO)
    for c in ws[19][:8]:
        c.fill = FILL_TOTAL
    ws["A22"] = ("Observação: projeção simples para planejamento. Não considera sazonalidade; os meses de janeiro e "
                 "dezembro costumam vender acima da média.")
    ws["A22"].font = SUB
    ws.conditional_formatting.add("B19:H20", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONTE, bold=True, color="B91C1C")))


def aba_dashboard(wb, fim_prod):
    ws = wb.create_sheet("Dashboard", 1)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", (2, 26, 16, 3, 26, 16, 3, 26, 16, 3)):
        ws.column_dimensions[col].width = w
    ws["B1"] = "Painel de gestão"
    ws["B1"].font = TITULO
    ws["B2"] = '="Loja Exemplo (fictícia) — ano "&pAno&" — atualizado automaticamente pelas fórmulas"'
    ws["B2"].font = SUB
    tp = 4 + len(PRODUTOS) + 1
    kpis = [
        ("B4", "Vendas no ano", "=Fluxo_Caixa!N5", BRL0),
        ("E4", "Resultado de caixa no ano", "=Fluxo_Caixa!N14", BRL0),
        ("H4", "Saldo de caixa final", "=Fluxo_Caixa!N15", BRL0),
        ("B7", "Margem bruta média", f"=Produtos!F{tp}", PCT),
        ("E7", "Valor em estoque (custo)", f"=Produtos!L{tp}", BRL0),
        ("H7", "Produtos para repor", f'=COUNTIF(Produtos!M5:M{fim_prod},"REPOR")+COUNTIF(Produtos!M5:M{fim_prod},"SEM ESTOQUE")', INT),
        ("B10", "Contas pendentes (R$)", '=SUMIFS(Financeiro!E5:E600,Financeiro!G5:G600,"Pendente",Financeiro!D5:D600,"Despesa")', BRL0),
        ("E10", "Contas vencidas", '=COUNTIF(Financeiro!I5:I600,"VENCIDA")', INT),
        ("H10", "Giro de estoque no ano (x)", f'=IFERROR(SUMPRODUCT(Produtos!J5:J{fim_prod},Produtos!D5:D{fim_prod})/Produtos!L{tp},0)', '0.0"x"'),
    ]
    for ref, nome, f, numero in kpis:
        col, row = ref[0], int(ref[1:])
        nxt = chr(ord(col) + 1)
        ws[ref] = nome
        ws[ref].font = Font(name=FONTE, size=9, color="4B5563")
        ws[f"{col}{row + 1}"] = f
        ws[f"{col}{row + 1}"].font = Font(name=FONTE, size=16, bold=True, color="1F3A5F")
        ws[f"{col}{row + 1}"].number_format = numero
        for r in (row, row + 1):
            for c in (col, nxt):
                ws[f"{c}{r}"].fill = FILL_KPI
    ws["H10"].comment = Comment("Custo das mercadorias vendidas no ano dividido pelo valor atual em estoque.", "Gabriel Veloso")
    # tabela: estoque e vendas por categoria
    ws["B13"] = "Por categoria"
    ws["B13"].font = NEGRITO
    for col, t in zip("BCEF", ("Categoria", "Valor em estoque", "Vendas no ano", "% das vendas")):
        ws[f"{col}14"] = t
        ws[f"{col}14"].font, ws[f"{col}14"].fill = CAB, FILL_CAB
    for i, cat in enumerate(CATEGORIAS, start=15):
        ws[f"B{i}"] = f"=Parametros!D{i - 10}"
        ws[f"C{i}"] = f"=SUMIFS(Produtos!$L$5:$L${fim_prod},Produtos!$C$5:$C${fim_prod},B{i})"
        ws[f"E{i}"] = f"=SUMIFS(Produtos!$O$5:$O${fim_prod},Produtos!$C$5:$C${fim_prod},B{i})"
        ws[f"F{i}"] = f"=IFERROR(E{i}/SUM($E$15:$E$18),0)"
        for col, numero in (("B", None), ("C", BRL0), ("E", BRL0), ("F", PCT)):
            fmt(ws[f"{col}{i}"], numero)
    # alertas de reposição (até 10 produtos)
    ws["H13"] = "Alertas de reposição"
    ws["H13"].font = NEGRITO
    ws["H14"], ws["I14"] = "Produto", "Comprar (un.)"
    for c in (ws["H14"], ws["I14"]):
        c.font, c.fill = CAB, FILL_CAB
    for k in range(1, 11):
        r = 14 + k
        # k-ésimo produto em alerta, pela coluna auxiliar "Nº do alerta" da aba Produtos
        ws[f"H{r}"] = f'=IFERROR(INDEX(Produtos!$B$5:$B${fim_prod},MATCH({k},Produtos!$P$5:$P${fim_prod},0)),"")'
        ws[f"I{r}"] = f'=IF(H{r}="","",INDEX(Produtos!$N$5:$N${fim_prod},MATCH(H{r},Produtos!$B$5:$B${fim_prod},0)))'
        fmt(ws[f"H{r}"])
        fmt(ws[f"I{r}"], INT)
    # gráfico de vendas mensais
    ln = LineChart()
    ln.title = "Vendas e despesas operacionais por mês"
    ln.y_axis.title = "R$"
    fc = wb["Fluxo_Caixa"]
    ln.add_data(Reference(fc, min_col=1, max_col=13, min_row=5, max_row=5), from_rows=True, titles_from_data=True)
    ln.add_data(Reference(fc, min_col=1, max_col=13, min_row=11, max_row=11), from_rows=True, titles_from_data=True)
    ln.set_categories(Reference(fc, min_col=2, max_col=13, min_row=4, max_row=4))
    ln.height, ln.width = 8, 17
    for serie in ln.series:
        serie.smooth = False
    _ajustar(ln)
    ws.add_chart(ln, "B27")
    bc = BarChart()
    bc.type = "bar"
    bc.title = "Vendas no ano por categoria"
    bc.add_data(Reference(ws, min_col=5, min_row=14, max_row=18), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=2, min_row=15, max_row=18))
    bc.height, bc.width = 8, 13
    bc.legend = None
    _ajustar(bc)
    ws.add_chart(bc, "H27")


def aba_lancamento(wb):
    ws = wb.create_sheet("Lancamento")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 30
    ws["B1"] = "Lançamento rápido de movimentação"
    ws["B1"].font = TITULO
    ws["B2"] = "Preencha os campos amarelos e rode a macro RegistrarMovimentacao (Alt+F8)."
    ws["B2"].font = SUB
    campos = [("Data", date(ANO, 12, 31), DATA), ("Código do produto", "P001", None), ("Tipo", "Saída", None), ("Quantidade", 1, INT)]
    for i, (nome, ex, numero) in enumerate(campos, start=4):
        ws[f"B{i}"] = nome
        ws[f"B{i}"].font = NEGRITO
        c = ws[f"C{i}"]
        c.value = ex
        c.font, c.fill = ENTRADA, FILL_ENTRADA
        if numero:
            c.number_format = numero
    ws["B9"], ws["C9"] = "Produto", '=IFERROR(INDEX(Produtos!B:B,MATCH(C5,Produtos!A:A,0)),"Código não cadastrado")'
    ws["B10"], ws["C10"] = "Estoque atual", "=IFERROR(INDEX(Produtos!K:K,MATCH(C5,Produtos!A:A,0)),0)"
    ws["B11"], ws["C11"] = "Valor unitário sugerido", '=IFERROR(IF(C6="Entrada",INDEX(Produtos!D:D,MATCH(C5,Produtos!A:A,0)),INDEX(Produtos!E:E,MATCH(C5,Produtos!A:A,0))),0)'
    ws["B12"], ws["C12"] = "Estoque após o lançamento", '=IF(C6="Entrada",C10+C7,C10-C7)'
    for r, numero in ((9, None), (10, INT), (11, BRL), (12, INT)):
        ws[f"B{r}"].font = NORMAL
        fmt(ws[f"C{r}"], numero)
    ws.conditional_formatting.add("C12", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONTE, bold=True, color="B91C1C")))
    for rng, f in (("C5", "=Produtos!$A$5:$A$100"), ("C6", "=lstTipoMov")):
        dv = DataValidation(type="list", formula1=f)
        ws.add_data_validation(dv)
        dv.add(rng)


def main():
    wb = Workbook()
    aba_leiame(wb)
    aba_parametros(wb)
    fim = aba_produtos(wb)
    movs = gerar_movimentos()
    aba_movimentacoes(wb, movs)
    aba_financeiro(wb, gerar_financeiro())
    aba_fluxo(wb)
    aba_projecao(wb)
    aba_dashboard(wb, fim)
    aba_lancamento(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"{OUT.name}: {len(movs)} movimentações")


if __name__ == "__main__":
    main()
