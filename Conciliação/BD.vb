Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data
Imports Conciliação_Rateio.MainWindow
Imports System.Windows.Threading
Imports System.IO

Public Class BD
    Dim DS As New DataSet
    Dim DT_BF As New DataTable
    Dim DT_BC As New DataTable
    Dim DV_BC As New DataView
    Public DT_RESULTADO As New DataTable
    Dim DT_BF_Back As New DataTable
    Dim DT_BC_Back As New DataTable
    Dim DV_Excel As New DataView

    Dim Nome_Coluna As String

    Dim n_Rodada As Integer
    'Stores the value of the ProgressBar
    Public value As Double = 0
    'Dim MW As New MainWindow


    Public Sub Exportacao_SF_SC(Txt As TextBox, Menu As MenuItem)

        Try

            Dim Soma_BC As Single
            Dim Soma_BF As Single
            Dim dv_bf As New DataView
            Dim dv_bc As New DataView

            Menu.IsEnabled = False
            Menu.Header = "Aguarde a Exportação do Arquivo"
            DoEvents()

            If DT_BC.Rows.Count > 0 Then
                dv_bc = DT_BC.DefaultView
                For Each R_BC In dv_bc
                    Soma_BC += R_BC.item(11)
                Next
            Else
                Soma_BC = 0
            End If

            If DT_BF.Rows.Count > 0 Then
                dv_bf = DT_BF.DefaultView
                For Each R_BF In dv_bf
                    Soma_BF += R_BF.item(11)
                Next
            Else
                Soma_BF = 0
            End If

            Txt.Text += vbCrLf & " | RODADA FINAL | SOBRA FÍSICA: " & Math.Round(Soma_BF, 2) & " | SOBRA CONTÁBIL " & Math.Round(Soma_BC, 2)
        Catch
            MsgBox("Erro na Rodada Final. Verifique os valores das quantidades.", vbCritical)
        End Try
    End Sub

    Public Sub Modelo_Excel()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T_BF As Excel.Worksheet
        Dim Sh_T_BC As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkBook.Sheets.Add()
            Sh_T_BF = xlWorkBook.Sheets(1)
            Sh_T_BF.Name = "Base Física"
            Sh_T_BF.Range("a1").Value = "CHAVE"
            Sh_T_BF.Range("b1").Value = "CAMPO1"
            Sh_T_BF.Range("c1").Value = "CAMPO2"
            Sh_T_BF.Range("d1").Value = "CAMPO3"
            Sh_T_BF.Range("e1").Value = "CAMPO4"
            Sh_T_BF.Range("f1").Value = "CAMPO5"
            Sh_T_BF.Range("g1").Value = "CAMPO6"
            Sh_T_BF.Range("h1").Value = "CAMPO7"
            Sh_T_BF.Range("i1").Value = "CAMPO8"
            Sh_T_BF.Range("j1").Value = "CAMPO9"
            Sh_T_BF.Range("k1").Value = "CAMPO10"
            Sh_T_BF.Range("l1").Value = "QUANTIDADE"
            Sh_T_BF.Range("m1").Value = "PRIORIDADE"
            Sh_T_BF.Columns.AutoFit()
            Sh_T_BF.Range("a1:m1").Font.Bold = True
            Sh_T_BF.Range("a1:m1").Font.ColorIndex = 2
            Sh_T_BF.Range("a1:m1").Interior.ColorIndex = 51

            Sh_T_BC = xlWorkBook.Sheets(2)
            Sh_T_BC.Name = "Base Contábil"
            Sh_T_BC.Range("a1").Value = "CHAVE"
            Sh_T_BC.Range("b1").Value = "CAMPO1"
            Sh_T_BC.Range("c1").Value = "CAMPO2"
            Sh_T_BC.Range("d1").Value = "CAMPO3"
            Sh_T_BC.Range("e1").Value = "CAMPO4"
            Sh_T_BC.Range("f1").Value = "CAMPO5"
            Sh_T_BC.Range("g1").Value = "CAMPO6"
            Sh_T_BC.Range("h1").Value = "CAMPO7"
            Sh_T_BC.Range("i1").Value = "CAMPO8"
            Sh_T_BC.Range("j1").Value = "CAMPO9"
            Sh_T_BC.Range("k1").Value = "CAMPO10"
            Sh_T_BC.Range("l1").Value = "QUANTIDADE"
            Sh_T_BC.Range("m1").Value = "DATA"
            Sh_T_BC.Range("n1").Value = "VOC"
            Sh_T_BC.Range("o1").Value = "DAC"
            Sh_T_BC.Columns.AutoFit()
            Sh_T_BC.Range("a1:o1").Font.Bold = True
            Sh_T_BC.Range("a1:o1").Font.ColorIndex = 2
            Sh_T_BC.Range("a1:o1").Interior.ColorIndex = 56

            xlApp.Visible = True

        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Juntar_DT()
        Try
            If DT_BF.Rows.Count > 0 Then
                Dim dv_bf As New DataView
                dv_bf = DT_BF.DefaultView
                DT_RESULTADO.BeginLoadData()
                For Each R_BF In dv_bf
                    DT_RESULTADO.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", 0, 0, 0,
                                          "SOBRA FÍSICA", R_BF.Item(0), R_BF.Item(1), R_BF.Item(2), R_BF.Item(3), R_BF.Item(4),
                                              R_BF.Item(5), R_BF.Item(6), R_BF.Item(7), R_BF.Item(8), R_BF.Item(9),
                                              R_BF.Item(10), R_BF.Item(11))
                Next
                DT_RESULTADO.EndLoadData()
                DT_BF.Clear()
            End If
        Catch
        End Try

        Try
            If DT_BC.Rows.Count > 0 Then
                Dim dv_bc As New DataView
                dv_bc = DT_BC.DefaultView
                DT_RESULTADO.BeginLoadData()
                For Each R_BC In dv_bc
                    DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), R_BC.Item(12), Math.Round(R_BC.Item(13), 4).ToString, R_BC.Item(14), R_BC.Item(11),
                                              "SOBRA CONTÁBIL", "", "", "", "", "", "", "", "", "", "", "", 0)
                Next
                DT_RESULTADO.EndLoadData()
                DT_BC.Clear()
            End If
        Catch
        End Try
    End Sub

    Public Sub Exportar_Excel(Txt As TextBox, Casa_Decimal_Qtde As Integer, Casa_Decimal_Valor As Integer, DgResult As DataGrid,
                              DgBC As DataGrid, DgBF As DataGrid, Tbcontrole As TabControl, TiResultado As TabItem,
                              TiBC As TabItem, TiBF As TabItem)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim StResultado As Excel.Worksheet
        'Dim StSF As Excel.Worksheet
        'Dim StSC As Excel.Worksheet
        Dim StRodadas As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim Formato_Qtde As String = ""
        Dim Formato_Valor As String = ""

        Select Case Casa_Decimal_Qtde
            Case 0
                Formato_Qtde = "#,###"
            Case 1
                Formato_Qtde = "#,###.0"
            Case 2
                Formato_Qtde = "#,###.00"
            Case 3
                Formato_Qtde = "#,###.000"
            Case 4
                Formato_Qtde = "#,###.0000"
        End Select

        Select Case Casa_Decimal_Valor
            Case 0
                Formato_Valor = "#,###"
            Case 1
                Formato_Valor = "#,###.0"
            Case 2
                Formato_Valor = "#,###.00"
            Case 3
                Formato_Valor = "#,###.000"
            Case 4
                Formato_Valor = "#,###.0000"
        End Select
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)


        'xlWorkBook.Sheets.Add(Before:=StResultado)
        StRodadas = xlWorkBook.Sheets(1)
        StRodadas.Cells(1, 1).value = Txt.Text
        StRodadas.Cells(1, 1).VerticalAlignment = Excel.Constants.xlTop
        StRodadas.Columns("A:A").ColumnWidth = 65
        'StRodadas.Rows.AutoFit()
        StRodadas.Name = "Rodadas"
        xlWorkBook.Sheets.Add()
        StResultado = xlWorkBook.Sheets(1)
        StResultado.Name = "Resultado"
        Dim Contar_DT_Resultado As Integer = DT_RESULTADO.Rows.Count

        'Dim colIndex As Integer
        'Dim dc As System.Data.DataColumn

        Dim dv As New DataView
        Dim Linha As Integer = 1
        dv = DT_RESULTADO.DefaultView

        xlApp.Visible = False

        'CONCILIADO
        Dim colIndex As Integer
        For Each dc In DT_RESULTADO.Columns
            colIndex = colIndex + 1
            'Column headers
            StResultado.Cells(1, colIndex) = dc.ColumnName
        Next

        If DT_RESULTADO.Rows.Count > 0 Then
            Try
                Dim A_1(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_2(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_3(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_4(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_5(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_6(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_7(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_8(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_9(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_10(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_11(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_12(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_13(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_14(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_15(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_16(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_17(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_18(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_19(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_20(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_21(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_22(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_23(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_24(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_25(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_26(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_27(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String
                Dim A_28(0 To DT_RESULTADO.Rows.Count - 1, 0 To 0) As String

                For i = 0 To DT_RESULTADO.Rows.Count - 1
                    A_1(i, 0) = DT_RESULTADO.Rows(i)(0)
                    A_2(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(1)), "", DT_RESULTADO.Rows(i)(1))
                    A_3(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(2)), "", DT_RESULTADO.Rows(i)(2))
                    A_4(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(3)), "", DT_RESULTADO.Rows(i)(3))
                    A_5(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(4)), "", DT_RESULTADO.Rows(i)(4))
                    A_6(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(5)), "", DT_RESULTADO.Rows(i)(5))
                    A_7(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(6)), "", DT_RESULTADO.Rows(i)(6))
                    A_8(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(7)), "", DT_RESULTADO.Rows(i)(7))
                    A_9(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(8)), "", DT_RESULTADO.Rows(i)(8))
                    A_10(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(9)), "", DT_RESULTADO.Rows(i)(9))
                    A_11(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(10)), "", DT_RESULTADO.Rows(i)(10))
                    A_12(i, 0) = DT_RESULTADO.Rows(i)(11)
                    A_13(i, 0) = DT_RESULTADO.Rows(i)(12)
                    A_14(i, 0) = DT_RESULTADO.Rows(i)(13)
                    A_15(i, 0) = DT_RESULTADO.Rows(i)(14)
                    A_16(i, 0) = DT_RESULTADO.Rows(i)(15)
                    A_17(i, 0) = DT_RESULTADO.Rows(i)(16)
                    A_18(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(17)), "", DT_RESULTADO.Rows(i)(17))
                    A_19(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(18)), "", DT_RESULTADO.Rows(i)(18))
                    A_20(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(19)), "", DT_RESULTADO.Rows(i)(19))
                    A_21(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(20)), "", DT_RESULTADO.Rows(i)(20))
                    A_22(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(21)), "", DT_RESULTADO.Rows(i)(21))
                    A_23(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(22)), "", DT_RESULTADO.Rows(i)(22))
                    A_24(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(23)), "", DT_RESULTADO.Rows(i)(23))
                    A_25(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(24)), "", DT_RESULTADO.Rows(i)(24))
                    A_26(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(25)), "", DT_RESULTADO.Rows(i)(25))
                    A_27(i, 0) = IIf(IsDBNull(DT_RESULTADO.Rows(i)(26)), "", DT_RESULTADO.Rows(i)(26))
                    A_28(i, 0) = DT_RESULTADO.Rows(i)(27)
                Next
                StResultado.Cells(2, 1).resize(DT_RESULTADO.Rows.Count).value = A_1
                StResultado.Cells(2, 2).resize(DT_RESULTADO.Rows.Count).value = A_2
                StResultado.Cells(2, 3).resize(DT_RESULTADO.Rows.Count).value = A_3
                StResultado.Cells(2, 4).resize(DT_RESULTADO.Rows.Count).value = A_4
                StResultado.Cells(2, 5).resize(DT_RESULTADO.Rows.Count).value = A_5
                StResultado.Cells(2, 6).resize(DT_RESULTADO.Rows.Count).value = A_6
                StResultado.Cells(2, 7).resize(DT_RESULTADO.Rows.Count).value = A_7
                StResultado.Cells(2, 8).resize(DT_RESULTADO.Rows.Count).value = A_8
                StResultado.Cells(2, 9).resize(DT_RESULTADO.Rows.Count).value = A_9
                StResultado.Cells(2, 10).resize(DT_RESULTADO.Rows.Count).value = A_10
                StResultado.Cells(2, 11).resize(DT_RESULTADO.Rows.Count).value = A_11
                StResultado.Cells(2, 12).resize(DT_RESULTADO.Rows.Count).value = A_12
                StResultado.Cells(2, 13).resize(DT_RESULTADO.Rows.Count).value = A_13
                StResultado.Cells(2, 14).resize(DT_RESULTADO.Rows.Count).value = A_14
                StResultado.Cells(2, 15).resize(DT_RESULTADO.Rows.Count).value = A_15
                StResultado.Cells(2, 16).resize(DT_RESULTADO.Rows.Count).value = A_16
                StResultado.Cells(2, 17).resize(DT_RESULTADO.Rows.Count).value = A_17
                StResultado.Cells(2, 18).resize(DT_RESULTADO.Rows.Count).value = A_18
                StResultado.Cells(2, 19).resize(DT_RESULTADO.Rows.Count).value = A_19
                StResultado.Cells(2, 20).resize(DT_RESULTADO.Rows.Count).value = A_20
                StResultado.Cells(2, 21).resize(DT_RESULTADO.Rows.Count).value = A_21
                StResultado.Cells(2, 22).resize(DT_RESULTADO.Rows.Count).value = A_22
                StResultado.Cells(2, 23).resize(DT_RESULTADO.Rows.Count).value = A_23
                StResultado.Cells(2, 24).resize(DT_RESULTADO.Rows.Count).value = A_24
                StResultado.Cells(2, 25).resize(DT_RESULTADO.Rows.Count).value = A_25
                StResultado.Cells(2, 26).resize(DT_RESULTADO.Rows.Count).value = A_26
                StResultado.Cells(2, 27).resize(DT_RESULTADO.Rows.Count).value = A_27
                StResultado.Cells(2, 28).resize(DT_RESULTADO.Rows.Count).value = A_28
            Catch
                MsgBox("Erro na Extração Resultado")
            End Try
        End If

        'SOBRA CONTÁBIL
        If DT_BC.Rows.Count > 0 Then
            Try
                Dim A_1(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_2(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_3(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_4(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_5(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_6(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_7(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_8(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_9(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_10(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_11(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_12(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_13(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_14(0 To DT_BC.Rows.Count - 1, 0 To 0) As String
                Dim A_15(0 To DT_BC.Rows.Count - 1, 0 To 0) As String

                For i = 0 To DT_BC.Rows.Count - 1
                    A_1(i, 0) = DT_BC.Rows(i)(0)
                    A_2(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(1)), "", DT_BC.Rows(i)(1))
                    A_3(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(2)), "", DT_BC.Rows(i)(2))
                    A_4(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(3)), "", DT_BC.Rows(i)(3))
                    A_5(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(4)), "", DT_BC.Rows(i)(4))
                    A_6(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(5)), "", DT_BC.Rows(i)(5))
                    A_7(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(6)), "", DT_BC.Rows(i)(6))
                    A_8(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(7)), "", DT_BC.Rows(i)(7))
                    A_9(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(8)), "", DT_BC.Rows(i)(8))
                    A_10(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(9)), "", DT_BC.Rows(i)(9))
                    A_11(i, 0) = IIf(IsDBNull(DT_BC.Rows(i)(10)), "", DT_BC.Rows(i)(10))
                    A_12(i, 0) = DT_BC.Rows(i)(11)
                    A_13(i, 0) = DT_BC.Rows(i)(12)
                    A_14(i, 0) = DT_BC.Rows(i)(13)
                    A_15(i, 0) = DT_BC.Rows(i)(14)
                Next
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 1).resize(DT_BC.Rows.Count).value = A_1
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 2).resize(DT_BC.Rows.Count).value = A_2
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 3).resize(DT_BC.Rows.Count).value = A_3
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 4).resize(DT_BC.Rows.Count).value = A_4
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 5).resize(DT_BC.Rows.Count).value = A_5
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 6).resize(DT_BC.Rows.Count).value = A_6
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 7).resize(DT_BC.Rows.Count).value = A_7
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 8).resize(DT_BC.Rows.Count).value = A_8
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 9).resize(DT_BC.Rows.Count).value = A_9
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 10).resize(DT_BC.Rows.Count).value = A_10
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 11).resize(DT_BC.Rows.Count).value = A_11
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 12).resize(DT_BC.Rows.Count).value = A_13
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 13).resize(DT_BC.Rows.Count).value = A_14
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 14).resize(DT_BC.Rows.Count).value = A_15
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 15).resize(DT_BC.Rows.Count).value = A_12
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count, 16).resize(DT_BC.Rows.Count).value = "SOBRA CONTÁBIL"
            Catch
                MsgBox("Erro na Extração Sobra Contábil")
            End Try
        End If

        'SOBRA FÍSICA
        If DT_BF.Rows.Count > 0 Then
            Try
                Dim A_17(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_18(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_19(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_20(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_21(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_22(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_23(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_24(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_25(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_26(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_27(0 To DT_BF.Rows.Count - 1, 0 To 0) As String
                Dim A_28(0 To DT_BF.Rows.Count - 1, 0 To 0) As String

                For i = 0 To DT_BF.Rows.Count - 1
                    A_17(i, 0) = DT_BF.Rows(i)(0)
                    A_18(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(1)), "", DT_BF.Rows(i)(1))
                    A_19(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(2)), "", DT_BF.Rows(i)(2))
                    A_20(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(3)), "", DT_BF.Rows(i)(3))
                    A_21(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(4)), "", DT_BF.Rows(i)(4))
                    A_22(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(5)), "", DT_BF.Rows(i)(5))
                    A_23(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(6)), "", DT_BF.Rows(i)(6))
                    A_24(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(7)), "", DT_BF.Rows(i)(7))
                    A_25(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(8)), "", DT_BF.Rows(i)(8))
                    A_26(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(9)), "", DT_BF.Rows(i)(9))
                    A_27(i, 0) = IIf(IsDBNull(DT_BF.Rows(i)(10)), "", DT_BF.Rows(i)(10))
                    A_28(i, 0) = DT_BF.Rows(i)(11)
                Next
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 16).resize(DT_BF.Rows.Count).value = "SOBRA FÍSICA"
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 17).resize(DT_BF.Rows.Count).value = A_17
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 18).resize(DT_BF.Rows.Count).value = A_18
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 19).resize(DT_BF.Rows.Count).value = A_19
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 20).resize(DT_BF.Rows.Count).value = A_20
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 21).resize(DT_BF.Rows.Count).value = A_21
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 22).resize(DT_BF.Rows.Count).value = A_22
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 23).resize(DT_BF.Rows.Count).value = A_23
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 24).resize(DT_BF.Rows.Count).value = A_24
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 25).resize(DT_BF.Rows.Count).value = A_25
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 26).resize(DT_BF.Rows.Count).value = A_26
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 27).resize(DT_BF.Rows.Count).value = A_27
                StResultado.Cells(2 + DT_RESULTADO.Rows.Count + DT_BC.Rows.Count, 28).resize(DT_BF.Rows.Count).value = A_28
            Catch
                MsgBox("Erro na Extração Sobra Física")
            End Try
        End If

        DT_RESULTADO.Clear()
        DT_BC.Clear()
        DT_BF.Clear()

        StResultado.Columns("M:M").TextToColumns(Destination:=StResultado.Range("M1"), DataType:=Excel.XlTextParsingType.xlDelimited,
    TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True,
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True)
        StResultado.Columns("N:N").TextToColumns(Destination:=StResultado.Range("N1"), DataType:=Excel.XlTextParsingType.xlDelimited,
    TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True,
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True)
        StResultado.Columns("O:O").TextToColumns(Destination:=StResultado.Range("O1"), DataType:=Excel.XlTextParsingType.xlDelimited,
    TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True,
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True)
        StResultado.Columns("AB:AB").TextToColumns(Destination:=StResultado.Range("AB1"), DataType:=Excel.XlTextParsingType.xlDelimited,
    TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True,
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True)

        'Qtde
        StResultado.Range("O:O").NumberFormat = Formato_Qtde
        StResultado.Range("AB:AB").NumberFormat = Formato_Qtde
        'Valor
        StResultado.Range("M:N").NumberFormat = Formato_Valor

        StResultado.Range("a1:ab1").Font.Bold = True
        StResultado.Range("a1:ab1").Font.ColorIndex = 2
        StResultado.Range("p1").Font.ColorIndex = 1
        StResultado.Range("a1:o1").Interior.ColorIndex = 56
        StResultado.Range("p1").Interior.ColorIndex = 2
        StResultado.Range("q1:ab1").Interior.ColorIndex = 51

        StResultado.Columns.AutoFit()
        MsgBox("Dados Exportados com Sucesso!")
        xlApp.Visible = True

    End Sub

    Public Sub Importar_Excel(ArquivoExcel As String, DGV_BF As DataGrid, DGV_BC As DataGrid)
        'Try
        'Limpa dados
        n_Rodada = 0
        DT_RESULTADO.Clear()
        DS.Clear()
        DT_BC.Clear()
        DT_BF.Clear()
        DGV_BF.ItemsSource = ""
        DGV_BC.ItemsSource = ""

        GC.Collect()

        Dim con As New OleDbConnection
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ArquivoExcel & "';Extended Properties= 'Excel 12.0';"
        Dim cmd As New OleDbCommand
        Dim DA_BF As New OleDbDataAdapter
        Dim DA_BC As New OleDbDataAdapter
        con.Open()
        DA_BF.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE],[PRIORIDADE] FROM [Base Física$];", con)
        'DA_BC.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE],FORMAT ([DATA],'dd/MM/yyyy') as DATA,[VOC],[DAC] FROM [Base Contábil$];", con)
        DA_BC.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE],[DATA],[VOC],[DAC] FROM [Base Contábil$];", con)

        DA_BF.Fill(DS, "TB_BF")
        DA_BC.Fill(DS, "TB_BC")
        con.Close()
        'Dividir BF e BC em DataTables
        DT_BF = DS.Tables("TB_BF")
        DT_BC = DS.Tables("TB_BC")

        'Colocando ID na BC
        If DT_BC.Columns.Count = 15 Then
            DT_BC.Columns.Add("ID")
            DT_BC.BeginLoadData()
            For j = 0 To DT_BC.Rows.Count - 1
                DT_BC.Rows(j).Item("ID") = j + 1
            Next
            DT_BC.EndLoadData()
        End If
        'Analisar VOC e QTD
        If Analise_VOC() = False Then
            MsgBox("Quantidade e/ou VOC menor ou igual a zero localizado." & Chr(13) & " Ajuste a base antes de Importar.", vbInformation)
            Exit Sub
        End If

        DT_BF_Back = DT_BF
        DT_BC_Back = DT_BC

        DGV_BF.ItemsSource = DS.Tables("TB_BF").DefaultView
        DGV_BC.ItemsSource = DS.Tables("TB_BC").DefaultView

        MsgBox("Dados carregados com sucesso", MsgBoxStyle.Information)
        'Catch
        'MsgBox("Erro ao Carregar os Dados", MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub Classificar_BD(Prioridade As String, Ordem As String)
        Try

            'Ordenar tabelas
            If Prioridade = "Valor" And Ordem = "Crescente" Then
                DT_BC = DT_BC.Select("", "[VOC] asc").CopyToDataTable
            End If
            If Prioridade = "Valor" And Ordem = "Decrescente" Then
                DT_BC = DT_BC.Select("", "[VOC] desc").CopyToDataTable
            End If
            If Prioridade = "Data" And Ordem = "Crescente" Then
                DT_BC = DT_BC.Select("", "[DATA] asc").CopyToDataTable
            End If
            If Prioridade = "Data" And Ordem = "Decrescente" Then
                DT_BC = DT_BC.Select("", "[DATA] desc").CopyToDataTable
            End If
            DT_BF = DT_BF.Select("", "[PRIORIDADE] asc").CopyToDataTable
        Catch
        End Try
    End Sub

    Public Sub Criar_DT_Resultado()
        DT_RESULTADO.Columns.Add("ID_C")
        DT_RESULTADO.Columns.Add("CAMPO1_C")
        DT_RESULTADO.Columns.Add("CAMPO2_C")
        DT_RESULTADO.Columns.Add("CAMPO3_C")
        DT_RESULTADO.Columns.Add("CAMPO4_C")
        DT_RESULTADO.Columns.Add("CAMPO5_C")
        DT_RESULTADO.Columns.Add("CAMPO6_C")
        DT_RESULTADO.Columns.Add("CAMPO7_C")
        DT_RESULTADO.Columns.Add("CAMPO8_C")
        DT_RESULTADO.Columns.Add("CAMPO9_C")
        DT_RESULTADO.Columns.Add("CAMPO10_C")
        DT_RESULTADO.Columns.Add("DATA")
        DT_RESULTADO.Columns.Add("VOC")
        DT_RESULTADO.Columns.Add("DAC")
        DT_RESULTADO.Columns.Add("QUANTIDADE_C")
        DT_RESULTADO.Columns.Add("STATUS")
        DT_RESULTADO.Columns.Add("ID_F")
        DT_RESULTADO.Columns.Add("CAMPO1_F")
        DT_RESULTADO.Columns.Add("CAMPO2_F")
        DT_RESULTADO.Columns.Add("CAMPO3_F")
        DT_RESULTADO.Columns.Add("CAMPO4_F")
        DT_RESULTADO.Columns.Add("CAMPO5_F")
        DT_RESULTADO.Columns.Add("CAMPO6_F")
        DT_RESULTADO.Columns.Add("CAMPO7_F")
        DT_RESULTADO.Columns.Add("CAMPO8_F")
        DT_RESULTADO.Columns.Add("CAMPO9_F")
        DT_RESULTADO.Columns.Add("CAMPO10_F")
        DT_RESULTADO.Columns.Add("QUANTIDADE_F")
    End Sub

    'Public Sub Criar_DT_N_CONTABIL()
    '    DT_N_BC.Columns.Add("CHAVE")
    '    DT_N_BC.Columns.Add("CAMPO1")
    '    DT_N_BC.Columns.Add("CAMPO2")
    '    DT_N_BC.Columns.Add("CAMPO3")
    '    DT_N_BC.Columns.Add("CAMPO4")
    '    DT_N_BC.Columns.Add("CAMPO5")
    '    DT_N_BC.Columns.Add("CAMPO6")
    '    DT_N_BC.Columns.Add("CAMPO7")
    '    DT_N_BC.Columns.Add("CAMPO8")
    '    DT_N_BC.Columns.Add("CAMPO9")
    '    DT_N_BC.Columns.Add("CAMPO10")
    '    DT_N_BC.Columns.Add("QUANTIDADE")
    '    DT_N_BC.Columns.Add("DATA")
    '    DT_N_BC.Columns.Add("VOC")
    '    DT_N_BC.Columns.Add("DAC")
    'End Sub

    Public Sub Limpar_Limite(Limite_F As Single, Limite_C As Single, Txt As TextBox)
        On Error GoTo Err
        'Limpar DT
        Dim Limp_BF As New ArrayList
        Dim Limp_BC As New ArrayList
        Dim Loop_n_BF As Integer = 0
        Dim Loop_n_BC As Integer = 0
        Dim n_limpos_BF As Integer = 0
        Dim n_limpos_BC As Integer = 0
        Dim Q_BC As Decimal = 0
        Dim Q_BF As Decimal = 0

        For Each R_BF In DT_BF.Rows
            Loop_n_BF += 1
            If R_BF.Item(11) <= Limite_F Then
                Limp_BF.Add(Loop_n_BF - 1)
                Q_BF += R_BF.item(11)
            End If
        Next

        For Each R_BC In DT_BC.Rows
            Loop_n_BC += 1
            If R_BC.Item(11) <= Limite_C Then
                Limp_BC.Add(Loop_n_BC - 1)
                Q_BC += R_BC.item(11)
            End If
        Next

        If Limp_BF.Count > 0 Then
            For k = 0 To Limp_BF.Count - 1
                'Adicionar na Tabela Resultado
                DT_RESULTADO.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", 0, 0, 0,
                                      "SOBRA FÍSICA", DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(0), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(1), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(2),
                                      DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(3), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(4),
                                      DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(5), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(6), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(7),
                                      DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(8), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(9),
                                      DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(10), DT_BF.Rows(Limp_BF(k) - n_limpos_BF)(11))
                'Limpar Linha
                DT_BF.Rows.RemoveAt(Limp_BF(k) - n_limpos_BF)
                n_limpos_BF += 1
            Next
        End If

        If Limp_BC.Count > 0 Then
            For k = 0 To Limp_BC.Count - 1
                'Adicionar na Tabela Resultado
                DT_RESULTADO.BeginLoadData()
                DT_RESULTADO.Rows.Add(DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(0), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(1), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(2),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(3), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(4),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(5), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(6), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(7),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(8), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(9), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(10),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(12), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(13), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(14),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(11), "SOBRA CONTÁBIL", "", "", "", "", "", "", "", "", "", "", "", 0)
                DT_RESULTADO.EndLoadData()
                'Limpar Linha
                DT_BC.Rows.RemoveAt(Limp_BC(k) - n_limpos_BC)
                n_limpos_BC += 1
            Next
        End If
        Txt.Text = " | RODADA MÍN. QTDE | SOBRA FÍSICA: " & Math.Round(Q_BF, 2) & " | SOBRA CONTÁBIL: " & Math.Round(Q_BC, 2)
Err:
    End Sub

    Public Sub Conciliar(DGV_BF As DataGrid, DGV_BC As DataGrid, DGV_RESULTADO As DataGrid, CAMPO1 As Boolean,
                         CAMPO2 As Boolean, CAMPO3 As Boolean, CAMPO4 As Boolean, CAMPO5 As Boolean, CAMPO6 As Boolean,
                         CAMPO7 As Boolean, CAMPO8 As Boolean, CAMPO9 As Boolean, CAMPO10 As Boolean,
                         Txt As TextBox, Campos As String, Menu As MenuItem)
        'Limpar Bases
        Dim Limp_BF As New ArrayList
        Dim Limp_BC As New ArrayList
        Dim Loop_n_BF As Integer
        Dim Loop_n_BC As Integer
        Dim n_limpos_BF As Integer = 0
        Dim n_limpos_BC As Integer = 0

        'Var BF e BC
        Dim CHAVE_BF As String
        Dim CAMPO1_BF As String
        Dim CAMPO2_BF As String
        Dim CAMPO3_BF As String
        Dim CAMPO4_BF As String
        Dim CAMPO5_BF As String
        Dim CAMPO6_BF As String
        Dim CAMPO7_BF As String
        Dim CAMPO8_BF As String
        Dim CAMPO9_BF As String
        Dim CAMPO10_BF As String

        Dim CAMPO1_BF_Loop As String
        Dim CAMPO2_BF_Loop As String
        Dim CAMPO3_BF_Loop As String
        Dim CAMPO4_BF_Loop As String
        Dim CAMPO5_BF_Loop As String
        Dim CAMPO6_BF_Loop As String
        Dim CAMPO7_BF_Loop As String
        Dim CAMPO8_BF_Loop As String
        Dim CAMPO9_BF_Loop As String
        Dim CAMPO10_BF_Loop As String

        Dim QUANTIDADE_BF As Decimal
        Dim CHAVE_BC As String
        Dim CAMPO1_BC As String
        Dim CAMPO2_BC As String
        Dim CAMPO3_BC As String
        Dim CAMPO4_BC As String
        Dim CAMPO5_BC As String
        Dim CAMPO6_BC As String
        Dim CAMPO7_BC As String
        Dim CAMPO8_BC As String
        Dim CAMPO9_BC As String
        Dim CAMPO10_BC As String
        Dim QUANTIDADE_BC As Decimal
        Dim VOC_BC As Decimal
        Dim DAC_BC As Decimal

        Dim TEXTOFIM As String = ""

        Dim VOC_UNIT As Decimal
        Dim DAC_UNIT As Decimal
        Dim BF_BC_CONCIL As Decimal
        Dim Resultado_BF As Decimal
        Dim Resultado_BC As Decimal
        Dim Status As String

        Dim n_BF As Decimal = 0
        Dim n_BC As Decimal = 0

        Dim n_CO As Decimal = 0

        Dim Linhas_Totais As Integer = DT_BF.Rows.Count
        n_Rodada += 1

        Menu.IsEnabled = False
        For Each R_BF In DT_BF.Rows
            CHAVE_BF = R_BF.item(0)
            CAMPO1_BF = IIf(IsDBNull(R_BF.item(1)), "", R_BF.item(1))
            CAMPO2_BF = IIf(IsDBNull(R_BF.item(2)), "", R_BF.item(2))
            CAMPO3_BF = IIf(IsDBNull(R_BF.item(3)), "", R_BF.item(3))
            CAMPO4_BF = IIf(IsDBNull(R_BF.item(4)), "", R_BF.item(4))
            CAMPO5_BF = IIf(IsDBNull(R_BF.item(5)), "", R_BF.item(5))
            CAMPO6_BF = IIf(IsDBNull(R_BF.item(6)), "", R_BF.item(6))
            CAMPO7_BF = IIf(IsDBNull(R_BF.item(7)), "", R_BF.item(7))
            CAMPO8_BF = IIf(IsDBNull(R_BF.item(8)), "", R_BF.item(8))
            CAMPO9_BF = IIf(IsDBNull(R_BF.item(9)), "", R_BF.item(9))
            CAMPO10_BF = IIf(IsDBNull(R_BF.item(10)), "", R_BF.item(10))

            CAMPO1_BF_Loop = IIf(IsDBNull(R_BF.item(1)), "IS NULL", "='" & R_BF.item(1) & "'")
            CAMPO2_BF_Loop = IIf(IsDBNull(R_BF.item(2)), "IS NULL", "='" & R_BF.item(2) & "'")
            CAMPO3_BF_Loop = IIf(IsDBNull(R_BF.item(3)), "IS NULL", "='" & R_BF.item(3) & "'")
            CAMPO4_BF_Loop = IIf(IsDBNull(R_BF.item(4)), "IS NULL", "='" & R_BF.item(4) & "'")
            CAMPO5_BF_Loop = IIf(IsDBNull(R_BF.item(5)), "IS NULL", "='" & R_BF.item(5) & "'")
            CAMPO6_BF_Loop = IIf(IsDBNull(R_BF.item(6)), "IS NULL", "='" & R_BF.item(6) & "'")
            CAMPO7_BF_Loop = IIf(IsDBNull(R_BF.item(7)), "IS NULL", "='" & R_BF.item(7) & "'")
            CAMPO8_BF_Loop = IIf(IsDBNull(R_BF.item(8)), "IS NULL", "='" & R_BF.item(8) & "'")
            CAMPO9_BF_Loop = IIf(IsDBNull(R_BF.item(9)), "IS NULL", "='" & R_BF.item(9) & "'")
            CAMPO10_BF_Loop = IIf(IsDBNull(R_BF.item(10)), "IS NULL", "='" & R_BF.item(10) & "'")

            QUANTIDADE_BF = Decimal.Round(R_BF.Item(11), 4)

            n_BF = n_BF + 1
            n_BC = 0
            'Process(n_BF, Linhas_Totais)
            Menu.Header = n_BF & " de " & Linhas_Totais & " (" & Math.Round((n_BF / Linhas_Totais), 2) * 100 & " %)"
            DoEvents()
            'Loop BC
            If CAMPO1 = True Then
                TEXTOFIM = "CAMPO1 " & CAMPO1_BF_Loop
            End If
            If CAMPO2 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO2 " & CAMPO2_BF_Loop
                Else
                    TEXTOFIM = "CAMPO2 " & CAMPO2_BF_Loop
                End If
            End If
            If CAMPO3 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO3 " & CAMPO3_BF_Loop
                Else
                    TEXTOFIM = "CAMPO3 " & CAMPO3_BF_Loop
                End If
            End If
            If CAMPO4 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO4 " & CAMPO4_BF_Loop
                Else
                    TEXTOFIM = "CAMPO4 " & CAMPO4_BF_Loop
                End If
            End If
            If CAMPO5 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO5 " & CAMPO5_BF_Loop
                Else
                    TEXTOFIM = "CAMPO5 " & CAMPO5_BF_Loop
                End If
            End If
            If CAMPO6 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO6 " & CAMPO6_BF_Loop
                Else
                    TEXTOFIM = "CAMPO6 " & CAMPO6_BF_Loop
                End If
            End If
            If CAMPO7 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO7 " & CAMPO7_BF_Loop
                Else
                    TEXTOFIM = "CAMPO7 " & CAMPO7_BF_Loop
                End If
            End If
            If CAMPO8 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO8 " & CAMPO8_BF_Loop
                Else
                    TEXTOFIM = "CAMPO8 " & CAMPO8_BF_Loop
                End If
            End If
            If CAMPO9 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO9 " & CAMPO9_BF_Loop
                Else
                    TEXTOFIM = "CAMPO9 " & CAMPO9_BF_Loop
                End If
            End If
            If CAMPO10 = True Then
                If TEXTOFIM <> "" Then
                    TEXTOFIM = TEXTOFIM & " AND CAMPO10 " & CAMPO10_BF_Loop
                Else
                    TEXTOFIM = "CAMPO10 " & CAMPO10_BF_Loop
                End If
            End If

            DV_BC = DT_BC.DefaultView
            DV_BC.RowFilter = TEXTOFIM
            TEXTOFIM = ""
            DT_RESULTADO.BeginLoadData()
            For Each R_BC In DV_BC
                n_BC = n_BC + 1
                On Error GoTo Err

                If R_BF.Item(11) <= 0 Then
                    GoTo Prox_BF
                End If
                If R_BC.Item(11) <= 0 Then
                    GoTo Prox_BC
                End If

                '-----------------------------------------------------------------
                'Subtração e colocar dados na DT resultado

                QUANTIDADE_BC = Math.Round(R_BC.Item(11), 4)
                VOC_BC = Math.Round(R_BC.Item(13), 4)
                DAC_BC = Math.Round(R_BC.Item(14), 4)
                'Valores unit.
                VOC_UNIT = VOC_BC / QUANTIDADE_BC
                DAC_UNIT = DAC_BC / QUANTIDADE_BC
                'Diminui BC
                If QUANTIDADE_BF >= QUANTIDADE_BC Then
                    'Var Conciliado
                    BF_BC_CONCIL = QUANTIDADE_BC
                    'Zera BC
                    Resultado_BC = 0
                Else
                    Resultado_BC = QUANTIDADE_BC - QUANTIDADE_BF
                    'Var Conciliado
                    BF_BC_CONCIL = QUANTIDADE_BF
                End If

                'Diminui BF
                Resultado_BF = QUANTIDADE_BF - QUANTIDADE_BC
                If Resultado_BF < 0 Then
                    Resultado_BF = 0
                End If

                R_BF.Item(11) = Resultado_BF
                QUANTIDADE_BF = Resultado_BF
                R_BC.Item(11) = Resultado_BC

                'Arrumar DT_BC - VOC e DAC
                R_BC.Item(13) = Resultado_BC * VOC_UNIT
                R_BC.Item(14) = Resultado_BC * DAC_UNIT
                'Preencher DT resultado
                Status = "CONCILIADO"
                n_CO += BF_BC_CONCIL
                DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), Format(R_BC.Item(12), "dd/MM/yyyy"), Math.Round(VOC_UNIT * BF_BC_CONCIL, 4), Math.Round(DAC_UNIT * BF_BC_CONCIL, 4), BF_BC_CONCIL,
                                          Status, CHAVE_BF, CAMPO1_BF, CAMPO2_BF, CAMPO3_BF, CAMPO4_BF,
                                          CAMPO5_BF, CAMPO6_BF, CAMPO7_BF, CAMPO8_BF, CAMPO9_BF,
                                          CAMPO10_BF, BF_BC_CONCIL)

                '------------------------------------------------------------------
                'Limpar BC
                If R_BC.Item(11) <= 0 Then
                    R_BC.Item(11) = 0
                End If
                'Limpar BF
                If R_BF.Item(11) <= 0 Then
                    R_BF.Item(11) = 0
                    QUANTIDADE_BF = 0
                End If
Prox_BC:
            Next
            DT_RESULTADO.EndLoadData()
Prox_BF:

        Next
Err:
        'Limpar DT
        For Each R_BC In DT_BC.Rows
            If R_BC.Item(11) <= 0 Then
                R_BC.delete
            End If
        Next
        DT_BC.AcceptChanges()

        For Each R_BF In DT_BF.Rows
            If R_BF.Item(11) <= 0 Then
                R_BF.delete
            End If
        Next
        DT_BF.AcceptChanges()

        Txt.Text = Txt.Text & IIf(Txt.Text = "", "", vbCrLf) & " | RODADA: " & n_Rodada & " | CONCILIADO: " &
            Math.Round(n_CO, 2) & " | CAMPOS: " & Campos
        'Devolver resultado para DGV
        DGV_RESULTADO.ItemsSource = DT_RESULTADO.DefaultView

        DT_BC.DefaultView.RowFilter = ""
        DGV_BC.ItemsSource = DT_BC.DefaultView
        DGV_BF.ItemsSource = DT_BF.DefaultView
        Menu.IsEnabled = True
        Menu.Header = "Arquivo"

        '-----------------------------------------------------------
        'Alternativa
        '        DT_RESULTADO.BeginLoadData()

        '        For Each R_BC In DT_BC.Rows
        '            n_BC = n_BC + 1
        '            On Error GoTo Err

        '            If R_BF.Item(11) <= 0 Then
        '                GoTo Prox_BF
        '            End If
        '            If R_BC.Item(11) <= 0 Then
        '                GoTo Prox_BC
        '            End If

        '            'Dados selecionados
        '            If CAMPO1 = True Then
        '                TEXTO1 = CAMPO1_BF = IIf(IsDBNull(R_BC.item(1)), "", R_BC.item(1))
        '            Else
        '                TEXTO1 = True
        '            End If
        '            If CAMPO2 = True Then
        '                TEXTO2 = CAMPO2_BF = IIf(IsDBNull(R_BC.item(2)), "", R_BC.item(2))
        '            Else
        '                TEXTO2 = True
        '            End If
        '            If CAMPO3 = True Then
        '                TEXTO3 = CAMPO3_BF = IIf(IsDBNull(R_BC.item(3)), "", R_BC.item(3))
        '            Else
        '                TEXTO3 = True
        '            End If
        '            If CAMPO4 = True Then
        '                TEXTO4 = CAMPO4_BF = IIf(IsDBNull(R_BC.item(4)), "", R_BC.item(4))
        '            Else
        '                TEXTO4 = True
        '            End If
        '            If CAMPO5 = True Then
        '                TEXTO5 = CAMPO5_BF = IIf(IsDBNull(R_BC.item(5)), "", R_BC.item(5))
        '            Else
        '                TEXTO5 = True
        '            End If
        '            If CAMPO6 = True Then
        '                TEXTO6 = CAMPO6_BF = IIf(IsDBNull(R_BC.item(6)), "", R_BC.item(6))
        '            Else
        '                TEXTO6 = True
        '            End If
        '            If CAMPO7 = True Then
        '                TEXTO7 = CAMPO7_BF = IIf(IsDBNull(R_BC.item(7)), "", R_BC.item(7))
        '            Else
        '                TEXTO7 = True
        '            End If
        '            If CAMPO8 = True Then
        '                TEXTO8 = CAMPO8_BF = IIf(IsDBNull(R_BC.item(8)), "", R_BC.item(8))
        '            Else
        '                TEXTO8 = True
        '            End If
        '            If CAMPO9 = True Then
        '                TEXTO9 = CAMPO9_BF = IIf(IsDBNull(R_BC.item(9)), "", R_BC.item(9))
        '            Else
        '                TEXTO9 = True
        '            End If
        '            If CAMPO10 = True Then
        '                TEXTO10 = CAMPO10_BF = IIf(IsDBNull(R_BC.item(10)), "", R_BC.item(10))
        '            Else
        '                TEXTO10 = True
        '            End If
        '            '-----------------------------------------------------------------
        '            'Subtração e colocar dados na DT resultado

        '            If TEXTO1 And TEXTO2 And TEXTO3 And TEXTO4 And TEXTO5 And TEXTO6 And TEXTO7 And TEXTO8 _
        '                And TEXTO9 And TEXTO10 Then
        '                QUANTIDADE_BC = Decimal.Round(R_BC.Item(11), 4)
        '                VOC_BC = Decimal.Round(R_BC.Item(13), 4)
        '                DAC_BC = Decimal.Round(R_BC.Item(14), 4)
        '                'Valores unit.
        '                VOC_UNIT = VOC_BC / QUANTIDADE_BC
        '                DAC_UNIT = DAC_BC / QUANTIDADE_BC
        '                'Diminui BC
        '                If QUANTIDADE_BF >= QUANTIDADE_BC Then
        '                    'Var Conciliado
        '                    BF_BC_CONCIL = QUANTIDADE_BC
        '                    'Zera BC
        '                    Resultado_BC = 0
        '                Else
        '                    Resultado_BC = QUANTIDADE_BC - QUANTIDADE_BF
        '                    'Var Conciliado
        '                    BF_BC_CONCIL = QUANTIDADE_BF
        '                End If

        '                'Diminui BF
        '                Resultado_BF = QUANTIDADE_BF - QUANTIDADE_BC
        '                If Resultado_BF < 0 Then
        '                    Resultado_BF = 0
        '                End If

        '                R_BF.Item(11) = Resultado_BF
        '                QUANTIDADE_BF = Resultado_BF
        '                R_BC.Item(11) = Resultado_BC

        '                'Arrumar DT_BC - VOC e DAC
        '                R_BC.Item(13) = Resultado_BC * VOC_UNIT
        '                R_BC.Item(14) = Resultado_BC * DAC_UNIT
        '                'Preencher DT resultado
        '                Status = "CONCILIADO"
        '                n_CO += BF_BC_CONCIL
        '                DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
        '                                      R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
        '                                      R_BC.Item(10), R_BC.Item(12), VOC_UNIT * BF_BC_CONCIL, DAC_UNIT * BF_BC_CONCIL, BF_BC_CONCIL,
        '                                      Status, CHAVE_BF, CAMPO1_BF, CAMPO2_BF, CAMPO3_BF, CAMPO4_BF,
        '                                      CAMPO5_BF, CAMPO6_BF, CAMPO7_BF, CAMPO8_BF, CAMPO9_BF,
        '                                      CAMPO10_BF, BF_BC_CONCIL)

        '                '------------------------------------------------------------------
        '                'Limpar BC
        '                If R_BC.Item(11) <= 0 Then
        '                    R_BC.Item(11) = 0
        '                    Limp_BC.Add(n_BC - 1)
        '                End If
        '                'Limpar BF
        '                If R_BF.Item(11) <= 0 Then
        '                    R_BF.Item(11) = 0
        '                    QUANTIDADE_BF = 0
        '                End If
        '            End If
        'Prox_BC:
        '        Next
        '        DT_RESULTADO.EndLoadData()
        'Prox_BF:
        '        If Limp_BC.Count > 0 Then
        '            DT_BC.BeginLoadData()
        '            For k = 0 To Limp_BC.Count - 1
        '                DT_BC.Rows.RemoveAt(Limp_BC(k) - n_limpos_BC)
        '                n_limpos_BC += 1
        '            Next
        '            DT_BC.EndLoadData()
        '            n_limpos_BC = 0
        '            Limp_BC.Clear()
        '        End If
        '        Next
        'Err:
        '        'Limpar DT

        '        For Each R_BF In DT_BF.Rows
        '            Loop_n_BF += 1
        '            If R_BF.Item(11) <= 0 Then
        '                Limp_BF.Add(Loop_n_BF - 1)
        '            End If
        '        Next

        '        If Limp_BF.Count > 0 Then
        '            DT_BF.BeginLoadData()
        '            For k = 0 To Limp_BF.Count - 1
        '                DT_BF.Rows.RemoveAt(Limp_BF(k) - n_limpos_BF)
        '                n_limpos_BF += 1
        '            Next
        '            DT_BF.EndLoadData()
        '            n_limpos_BF = 0
        '            Limp_BF.Clear()
        '        End If

        '        Txt.Text = Txt.Text & IIf(Txt.Text = "", "", vbCrLf) & " | RODADA: " & n_Rodada & " | CONCILIADO: " &
        '            n_CO & " | CAMPOS: " & Campos
        '        'Devolver resultado para DGV
        '        DGV_RESULTADO.ItemsSource = DT_RESULTADO.DefaultView
        '        DGV_BC.ItemsSource = DT_BC.DefaultView
        '        DGV_BF.ItemsSource = DT_BF.DefaultView
        '        Menu.IsEnabled = True
        '        Menu.Header = "Arquivo"
    End Sub

    Public Sub DoEvents()
        Dim frame As New DispatcherFrame()
        Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, New DispatcherOperationCallback(AddressOf ExitFrame), frame)
        Windows.Threading.Dispatcher.PushFrame(frame)
    End Sub

    Public Function ExitFrame(ByVal f As Object) As Object
        CType(f, DispatcherFrame).Continue = False

        Return Nothing
    End Function

    Public Sub Zerar_Conciliacao(DgBF As DataGrid, DgBC As DataGrid, DgResultado As DataGrid)
        Try
            DT_BF.Clear()
            DT_BC.Clear()
            DT_RESULTADO.Clear()

            DgBF.ItemsSource = ""
            DgBC.ItemsSource = ""
            DgResultado.ItemsSource = ""

            DgBF.ItemsSource = DT_BF_Back.DefaultView
            DgBC.ItemsSource = DT_BC_Back.DefaultView

            DT_BF = DT_BF_Back
            DT_BC = DT_BC_Back

            n_Rodada = 0

        Catch ex As Exception

        End Try
    End Sub

    Public Function Analise_VOC() As Boolean
        Try
            Analise_VOC = True
            For x = 0 To DT_BF.Rows.Count - 1
                If DT_BF.Rows(x)(11) <= 0 Then
                    Analise_VOC = False
                    GoTo Fim
                End If
            Next
            For y = 0 To DT_BC.Rows.Count - 1
                If DT_BC.Rows(y)(11) <= 0 Or DT_BC.Rows(y)(13) <= 0 Then
                    Analise_VOC = False
                    GoTo Fim
                End If
            Next
Fim:
        Catch ex As Exception
            Analise_VOC = False
        End Try
    End Function

    'Private Delegate Sub UpdateProgressBarDelegate(ByVal dp As _
    '         System.Windows.DependencyProperty,
    '         ByVal value As Object)
    'Private Sub Process(ByRef Linhas As Single, ByRef Linhas_Totais As Single)
    '    Try

    '        'Create a new instance of our ProgressBar Delegate that points
    '        ' to the ProgressBar's SetValue method.
    '        value = (Linhas / (Linhas_Totais)) * 100
    '        Dim updatePbDelegate As New _
    '    UpdateProgressBarDelegate(AddressOf PB_W.PB.SetValue)
    '        PB_W.Dispatcher.Invoke(updatePbDelegate,
    '        System.Windows.Threading.DispatcherPriority.Background,
    '        New Object() {ProgressBar.ValueProperty, value})
    '    Catch
    '    End Try
    'End Sub
End Class
