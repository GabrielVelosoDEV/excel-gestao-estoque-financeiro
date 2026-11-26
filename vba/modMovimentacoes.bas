Attribute VB_Name = "modMovimentacoes"
'=============================================================================
' modMovimentacoes
' Registra entradas e saídas de estoque a partir da aba Lancamento,
' com validação de código, tipo, quantidade e estoque disponível.
'=============================================================================
Option Explicit

Private Const PRIMEIRA_LINHA As Long = 5
Private Const ULTIMA_LINHA As Long = 5000     ' linhas com fórmulas prontas na aba Movimentacoes

Public Sub RegistrarMovimentacao()
    Dim wsL As Worksheet, wsM As Worksheet, wsP As Worksheet
    Set wsL = ThisWorkbook.Worksheets("Lancamento")
    Set wsM = ThisWorkbook.Worksheets("Movimentacoes")
    Set wsP = ThisWorkbook.Worksheets("Produtos")

    Dim dataMov As Variant, codigo As String, tipo As String, qtd As Variant
    dataMov = wsL.Range("C4").Value
    codigo = UCase$(Trim$(CStr(wsL.Range("C5").Value)))
    tipo = Trim$(CStr(wsL.Range("C6").Value))
    qtd = wsL.Range("C7").Value

    ' ---- validações
    If Not IsDate(dataMov) Then
        Avisar "Informe uma data válida.": Exit Sub
    End If
    If codigo = "" Then
        Avisar "Informe o código do produto.": Exit Sub
    End If

    Dim linhaProduto As Variant
    linhaProduto = Application.Match(codigo, wsP.Range("A:A"), 0)
    If IsError(linhaProduto) Then
        Avisar "O código " & codigo & " não está cadastrado na aba Produtos.": Exit Sub
    End If

    If tipo <> "Entrada" And tipo <> "Saída" Then
        Avisar "O tipo deve ser Entrada ou Saída.": Exit Sub
    End If
    If Not IsNumeric(qtd) Then
        Avisar "A quantidade deve ser um número.": Exit Sub
    End If
    If qtd <= 0 Or qtd <> Int(qtd) Then
        Avisar "A quantidade deve ser um número inteiro maior que zero.": Exit Sub
    End If

    Dim estoque As Double
    estoque = wsP.Cells(linhaProduto, "K").Value
    If tipo = "Saída" And qtd > estoque Then
        Avisar "Estoque insuficiente: há " & estoque & " unidade(s) de " & _
               wsP.Cells(linhaProduto, "B").Value & "."
        Exit Sub
    End If

    ' ---- grava na próxima linha livre
    Dim linha As Long
    linha = ProximaLinhaLivre(wsM)
    If linha > ULTIMA_LINHA Then
        Avisar "A aba Movimentacoes atingiu a capacidade de " & (ULTIMA_LINHA - PRIMEIRA_LINHA + 1) & _
               " lançamentos. Copie as fórmulas das colunas C, G, H e I para mais linhas."
        Exit Sub
    End If

    Dim valorUnit As Double
    If tipo = "Entrada" Then
        valorUnit = wsP.Cells(linhaProduto, "D").Value      ' custo
    Else
        valorUnit = wsP.Cells(linhaProduto, "E").Value      ' preço de venda
    End If

    wsM.Cells(linha, "A").Value = CDate(dataMov)
    wsM.Cells(linha, "B").Value = codigo
    wsM.Cells(linha, "D").Value = tipo
    wsM.Cells(linha, "E").Value = CLng(qtd)
    wsM.Cells(linha, "F").Value = valorUnit
    Application.Calculate

    MsgBox tipo & " de " & qtd & " un. de " & wsP.Cells(linhaProduto, "B").Value & _
           " registrada na linha " & linha & "." & vbCrLf & _
           "Estoque atual: " & wsP.Cells(linhaProduto, "K").Value & " un.", _
           vbInformation, "Movimentação registrada"
End Sub

Private Function ProximaLinhaLivre(ByVal ws As Worksheet) As Long
    Dim ultima As Long
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultima < PRIMEIRA_LINHA Then ultima = PRIMEIRA_LINHA - 1
    ProximaLinhaLivre = ultima + 1
End Function

Private Sub Avisar(ByVal texto As String)
    MsgBox texto, vbExclamation, "Lançamento não registrado"
End Sub
