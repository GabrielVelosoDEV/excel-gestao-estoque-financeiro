Attribute VB_Name = "modValidacao"
'=============================================================================
' modValidacao
' Confere a consistência dos dados digitados e destaca as linhas com problema.
'=============================================================================
Option Explicit

Private Const COR_ERRO As Long = 13421823     ' RGB(255, 204, 204)

Public Sub VerificarConsistencia()
    Dim wsM As Worksheet, wsP As Worksheet, wsF As Worksheet
    Set wsM = ThisWorkbook.Worksheets("Movimentacoes")
    Set wsP = ThisWorkbook.Worksheets("Produtos")
    Set wsF = ThisWorkbook.Worksheets("Financeiro")
    Application.ScreenUpdating = False
    Application.CalculateFull
    LimparDestaques

    Dim ano As Long
    ano = ThisWorkbook.Names("pAno").RefersToRange.Value

    Dim r As Long, ultima As Long
    Dim semCadastro As Long, qtdInvalida As Long, tipoInvalido As Long, foraDoAno As Long
    ultima = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row
    For r = 5 To ultima
        Dim problema As Boolean
        problema = False
        If wsM.Cells(r, "C").Value = "CÓDIGO NÃO CADASTRADO" Then semCadastro = semCadastro + 1: problema = True
        If Not IsNumeric(wsM.Cells(r, "E").Value) Then
            qtdInvalida = qtdInvalida + 1: problema = True
        ElseIf wsM.Cells(r, "E").Value <= 0 Then
            qtdInvalida = qtdInvalida + 1: problema = True
        End If
        If wsM.Cells(r, "D").Value <> "Entrada" And wsM.Cells(r, "D").Value <> "Saída" Then
            tipoInvalido = tipoInvalido + 1: problema = True
        End If
        If IsDate(wsM.Cells(r, "A").Value) Then
            If Year(wsM.Cells(r, "A").Value) <> ano Then foraDoAno = foraDoAno + 1
        End If
        If problema Then wsM.Range("A" & r & ":I" & r).Interior.Color = COR_ERRO
    Next r

    Dim precoAbaixoCusto As Long, semEstoque As Long
    r = 5
    Do While wsP.Cells(r, "A").Value <> "" And wsP.Cells(r, "A").Value <> "Total"
        If wsP.Cells(r, "E").Value <= wsP.Cells(r, "D").Value Then
            precoAbaixoCusto = precoAbaixoCusto + 1
            wsP.Range("A" & r & ":E" & r).Interior.Color = COR_ERRO
        End If
        If wsP.Cells(r, "K").Value < 0 Then semEstoque = semEstoque + 1
        r = r + 1
    Loop

    Dim vencidas As Long
    vencidas = Application.WorksheetFunction.CountIf(wsF.Range("I5:I600"), "VENCIDA")

    Application.ScreenUpdating = True
    Dim msg As String
    msg = "Movimentações:" & vbCrLf & _
          "  • código não cadastrado: " & semCadastro & vbCrLf & _
          "  • quantidade inválida: " & qtdInvalida & vbCrLf & _
          "  • tipo inválido: " & tipoInvalido & vbCrLf & _
          "  • fora do ano de referência (" & ano & "): " & foraDoAno & vbCrLf & vbCrLf & _
          "Produtos:" & vbCrLf & _
          "  • preço menor ou igual ao custo: " & precoAbaixoCusto & vbCrLf & _
          "  • estoque negativo: " & semEstoque & vbCrLf & vbCrLf & _
          "Financeiro:" & vbCrLf & _
          "  • contas vencidas: " & vencidas
    MsgBox msg, IIf(semCadastro + qtdInvalida + tipoInvalido + precoAbaixoCusto + semEstoque > 0, _
                    vbExclamation, vbInformation), "Verificação de consistência"
End Sub

Public Sub LimparDestaques()
    With ThisWorkbook
        .Worksheets("Movimentacoes").Range("A5:I5000").Interior.Pattern = xlNone
        .Worksheets("Produtos").Range("A5:E200").Interior.Pattern = xlNone
    End With
End Sub
