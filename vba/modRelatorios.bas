Attribute VB_Name = "modRelatorios"
'=============================================================================
' modRelatorios
' Gera a lista de compras para reposição e exporta o painel em PDF.
'=============================================================================
Option Explicit

Public Sub GerarRelatorioReposicao()
    Dim wsP As Worksheet, wsR As Worksheet
    Set wsP = ThisWorkbook.Worksheets("Produtos")
    Application.ScreenUpdating = False
    Application.CalculateFull

    Set wsR = ObterAbaLimpa("Reposicao")
    wsR.Range("A1").Value = "Lista de compras para reposição"
    wsR.Range("A1").Font.Bold = True
    wsR.Range("A1").Font.Size = 14
    wsR.Range("A2").Value = "Gerada em " & Format(Now, "dd/mm/yyyy hh:nn")
    wsR.Range("A4:G4").Value = Array("Código", "Produto", "Estoque atual", "Estoque mínimo", _
                                     "Comprar (un.)", "Custo unit. (R$)", "Custo da compra (R$)")
    With wsR.Range("A4:G4")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 58, 95)
    End With

    Dim r As Long, linha As Long, total As Double
    linha = 5
    r = 5
    Do While wsP.Cells(r, "A").Value <> "" And wsP.Cells(r, "A").Value <> "Total"
        If wsP.Cells(r, "M").Value <> "OK" Then
            wsR.Cells(linha, "A").Value = wsP.Cells(r, "A").Value
            wsR.Cells(linha, "B").Value = wsP.Cells(r, "B").Value
            wsR.Cells(linha, "C").Value = wsP.Cells(r, "K").Value
            wsR.Cells(linha, "D").Value = wsP.Cells(r, "H").Value
            wsR.Cells(linha, "E").Value = wsP.Cells(r, "N").Value
            wsR.Cells(linha, "F").Value = wsP.Cells(r, "D").Value
            wsR.Cells(linha, "G").Formula = "=E" & linha & "*F" & linha
            total = total + wsP.Cells(r, "N").Value * wsP.Cells(r, "D").Value
            linha = linha + 1
        End If
        r = r + 1
    Loop

    If linha = 5 Then
        wsR.Range("A5").Value = "Nenhum produto abaixo do estoque mínimo."
    Else
        wsR.Cells(linha, "A").Value = "Total"
        wsR.Cells(linha, "G").Formula = "=SUM(G5:G" & linha - 1 & ")"
        wsR.Rows(linha).Font.Bold = True
        wsR.Range("F5:G" & linha).NumberFormat = "R$ #,##0.00"
    End If
    wsR.Columns("A:G").AutoFit
    Application.ScreenUpdating = True
    wsR.Activate
    MsgBox IIf(linha = 5, "Nenhum produto precisa de reposição.", _
               (linha - 5) & " produto(s) para repor. Custo estimado: " & Format(total, "R$ #,##0.00")), _
           vbInformation, "Reposição"
End Sub

Public Sub ExportarDashboardPDF()
    Dim caminho As String
    If ThisWorkbook.Path = "" Then
        MsgBox "Salve a planilha antes de exportar.", vbExclamation
        Exit Sub
    End If
    caminho = ThisWorkbook.Path & Application.PathSeparator & _
              "Painel_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    Application.CalculateFull
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=caminho, Quality:=xlQualityStandard, OpenAfterPublish:=True
End Sub

Private Function ObterAbaLimpa(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nome
    Else
        ws.Cells.Clear
    End If
    ActiveWindow.DisplayGridlines = False
    Set ObterAbaLimpa = ws
End Function
