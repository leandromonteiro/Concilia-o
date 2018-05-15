Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data
Imports Conciliação.MainWindow

Public Class BD
    Dim DS As New DataSet
    Dim DT_BF As New DataTable
    Dim DT_BC As New DataTable
    Public DT_RESULTADO As New DataTable
    Dim DT_BF_Back As New DataTable
    Dim DT_BC_Back As New DataTable
    Dim DV_Excel As New DataView

    Dim n_Rodada As Integer
    'Stores the value of the ProgressBar
    Public value As Double = 0
    Dim MW As New W_PB


    Public Sub Exportacao_SF_SC(Txt As TextBox)
        Try
            Dim Soma_BC As Single
            Dim Soma_BF As Single

            If DT_BC.Rows.Count > 0 Then
                For Each R_BC In DT_BC.Rows
                    Soma_BC += R_BC.item(11)
                Next
            Else
                Soma_BC = 0
            End If

            If DT_BF.Rows.Count > 0 Then
                For Each R_BF In DT_BF.Rows
                    Soma_BF += R_BF.item(11)
                Next
            Else
                Soma_BF = 0
            End If

            Txt.Text += vbCrLf & " | RODADA FINAL | SOBRA FÍSICA: " & Soma_BF & " | SOBRA CONTÁBIL " & Soma_BC
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
                For Each R_BF In DT_BF.Rows
                    DT_RESULTADO.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", 0, 0, 0,
                                          "SOBRA FÍSICA", R_BF.Item(0), R_BF.Item(1), R_BF.Item(2), R_BF.Item(3), R_BF.Item(4),
                                              R_BF.Item(5), R_BF.Item(6), R_BF.Item(7), R_BF.Item(8), R_BF.Item(9),
                                              R_BF.Item(10), R_BF.Item(11))
                Next
            End If
        Catch
        End Try

        Try
            If DT_BC.Rows.Count > 0 Then
                For Each R_BC In DT_BC.Rows
                    DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), R_BC.Item(12), R_BC.Item(13), R_BC.Item(14), R_BC.Item(11),
                                              "SOBRA CONTÁBIL", "", "", "", "", "", "", "", "", "", "", "", 0)
                Next
            End If
        Catch
        End Try
    End Sub

    Public Sub Exportar_Excel(Txt As TextBox, Casa_Decimal_Qtde As Integer, Casa_Decimal_Valor As Integer)
        Dim xlApp As Excel.Application
        Try
            Dim xlWorkBook As Excel.Workbook
            Dim StResultado As Excel.Worksheet
            Dim StRodadas As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim Formato_Qtde As String = ""
            Dim Formato_Valor As String = ""
            Dim i As Integer
            Dim j As Integer

            MW.Show()
            Select Case Casa_Decimal_Qtde
                Case 0
                    Formato_Qtde = "0"
                Case 1
                    Formato_Qtde = "0.0"
                Case 2
                    Formato_Qtde = "0.00"
                Case 3
                    Formato_Qtde = "0.000"
                Case 4
                    Formato_Qtde = "0.0000"
            End Select

            Select Case Casa_Decimal_Valor
                Case 0
                    Formato_Valor = "0"
                Case 1
                    Formato_Valor = "0.0"
                Case 2
                    Formato_Valor = "0.00"
                Case 3
                    Formato_Valor = "0.000"
                Case 4
                    Formato_Valor = "0.0000"
            End Select
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            StResultado = xlWorkBook.Sheets(1)

            'Colocando Títulos
            For k As Integer = 1 To DT_RESULTADO.Columns.Count
                StResultado.Cells(1, k) = DT_RESULTADO.Columns(k - 1).ColumnName
            Next
            Dim Contar_DT_Resultado As Integer = DT_RESULTADO.Rows.Count
            For i = 0 To Contar_DT_Resultado - 1
                Process(i, Contar_DT_Resultado - 1)
                For j = 0 To DT_RESULTADO.Columns.Count - 1
                    If j = 12 Or j = 13 Or j = 14 Or j = 27 Then
                        StResultado.Cells(i + 2, j + 1) = IIf(Not IsDBNull(DT_RESULTADO.Rows(i)(j)), CDec(DT_RESULTADO.Rows(i)(j)), "")
                    Else
                        StResultado.Cells(i + 2, j + 1) = DT_RESULTADO.Rows(i)(j)
                    End If
                    'Qtde
                    If j = 14 Or j = 27 Then
                        StResultado.Cells(i + 2, j + 1).numberformat = Formato_Qtde
                    End If
                    'Valor
                    If j = 12 Or j = 13 Then
                        StResultado.Cells(i + 2, j + 1).numberformat = Formato_Valor
                    End If
                Next
            Next

            StResultado.Range("a1:ab1").Font.Bold = True
            StResultado.Range("a1:ab1").Font.ColorIndex = 2
            StResultado.Range("p1").Font.ColorIndex = 1
            StResultado.Range("a1:o1").Interior.ColorIndex = 56
            StResultado.Range("p1").Interior.ColorIndex = 44
            StResultado.Range("q1:ab1").Interior.ColorIndex = 10

            StResultado.Columns.AutoFit()
            StResultado.Name = "Resultado"

            xlWorkBook.Sheets.Add(Before:=StResultado)
            StRodadas = xlWorkBook.Sheets(1)
            StRodadas.Cells(1, 1).value = Txt.Text
            StRodadas.Cells(1, 1).columnwidth = 100
            StRodadas.Cells(1, 1).VerticalAlignment = Excel.Constants.xlTop
            StRodadas.Name = "Rodadas"

            MW.Hide()
            xlApp.Visible = True
        Catch
            MW.Hide()
            xlApp.Quit()
            MsgBox("Erro ao exportar para Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Importar_Excel(ArquivoExcel As String, DGV_BF As DataGrid, DGV_BC As DataGrid)
        Try
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
            DA_BC.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE],FORMAT ([DATA],'dd/MM/yyyy') as DATA,[VOC],[DAC] FROM [Base Contábil$];", con)

            DA_BF.Fill(DS, "TB_BF")
            DA_BC.Fill(DS, "TB_BC")
            con.Close()
            'Dividir BF e BC em DataTables
            DT_BF = DS.Tables("TB_BF")
            DT_BC = DS.Tables("TB_BC")

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
        Catch
            MsgBox("Erro ao Carregar os Dados", MsgBoxStyle.Critical)
        End Try
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

    Public Sub Limpar_Limite(Limite_F As Single, Limite_C As Single, Txt As TextBox)
        On Error GoTo Err
        'Limpar DT
        Dim Limp_BF As New ArrayList
        Dim Limp_BC As New ArrayList
        Dim Loop_n_BF As Integer = 0
        Dim Loop_n_BC As Integer = 0
        Dim n_limpos_BF As Integer = 0
        Dim n_limpos_BC As Integer = 0

        For Each R_BF In DT_BF.Rows
            Loop_n_BF += 1
            If R_BF.Item(11) <= Limite_F Then
                Limp_BF.Add(Loop_n_BF - 1)
            End If
        Next

        For Each R_BC In DT_BC.Rows
            Loop_n_BC += 1
            If R_BC.Item(11) <= Limite_C Then
                Limp_BC.Add(Loop_n_BC - 1)
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
                DT_RESULTADO.Rows.Add(DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(0), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(1), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(2),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(3), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(4),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(5), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(6), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(7),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(8), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(9), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(10),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(12), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(13), DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(14),
                                      DT_BC.Rows(Limp_BC(k) - n_limpos_BC)(11), "SOBRA CONTÁBIL", "", "", "", "", "", "", "", "", "", "", "", 0)
                'Limpar Linha
                DT_BC.Rows.RemoveAt(Limp_BC(k) - n_limpos_BC)
                n_limpos_BC += 1
            Next
        End If
        Txt.Text = " | RODADA MÍN. QTDE | SOBRA FÍSICA: " & n_limpos_BF & " | SOBRA CONTÁBIL: " & n_limpos_BC
Err:
    End Sub

    Public Sub Conciliar(DGV_BF As DataGrid, DGV_BC As DataGrid, DGV_RESULTADO As DataGrid, CAMPO1 As Boolean,
                         CAMPO2 As Boolean, CAMPO3 As Boolean, CAMPO4 As Boolean, CAMPO5 As Boolean, CAMPO6 As Boolean,
                         CAMPO7 As Boolean, CAMPO8 As Boolean, CAMPO9 As Boolean, CAMPO10 As Boolean,
                         Txt As TextBox, Campos As String)

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

        Dim TEXTO1 As Boolean
        Dim TEXTO2 As Boolean
        Dim TEXTO3 As Boolean
        Dim TEXTO4 As Boolean
        Dim TEXTO5 As Boolean
        Dim TEXTO6 As Boolean
        Dim TEXTO7 As Boolean
        Dim TEXTO8 As Boolean
        Dim TEXTO9 As Boolean
        Dim TEXTO10 As Boolean

        Dim VOC_UNIT As Single
        Dim DAC_UNIT As Single
        Dim BF_BC_CONCIL As Single
        Dim Resultado_BF As Single
        Dim Resultado_BC As Single
        Dim Status As String

        Dim n_BF As Integer = 0
        Dim n_BC As Integer = 0

        Dim n_CO As Integer = 0

        Dim Linhas_Totais As Integer = DT_BF.Rows.Count
        n_Rodada += 1
        MW.Show()

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
            QUANTIDADE_BF = CDec(R_BF.Item(11))

            n_BF = n_BF + 1
            n_BC = 0
            Process(n_BF, Linhas_Totais)
            For Each R_BC In DT_BC.Rows
                n_BC = n_BC + 1
                On Error GoTo Err

                If QUANTIDADE_BF <= 0 Then
                    GoTo Prox_BF
                End If
                If R_BC.Item(11) <= 0 Then
                    GoTo Prox_BC
                End If

                'Dados selecionados
                If CAMPO1 = True Then
                    TEXTO1 = CAMPO1_BF = IIf(IsDBNull(R_BC.item(1)), "", R_BC.item(1))
                Else
                    TEXTO1 = True
                End If
                If CAMPO2 = True Then
                    TEXTO2 = CAMPO2_BF = IIf(IsDBNull(R_BC.item(2)), "", R_BC.item(2))
                Else
                    TEXTO2 = True
                End If
                If CAMPO3 = True Then
                    TEXTO3 = CAMPO3_BF = IIf(IsDBNull(R_BC.item(3)), "", R_BC.item(3))
                Else
                    TEXTO3 = True
                End If
                If CAMPO4 = True Then
                    TEXTO4 = CAMPO4_BF = IIf(IsDBNull(R_BC.item(4)), "", R_BC.item(4))
                Else
                    TEXTO4 = True
                End If
                If CAMPO5 = True Then
                    TEXTO5 = CAMPO5_BF = IIf(IsDBNull(R_BC.item(5)), "", R_BC.item(5))
                Else
                    TEXTO5 = True
                End If
                If CAMPO6 = True Then
                    TEXTO6 = CAMPO6_BF = IIf(IsDBNull(R_BC.item(6)), "", R_BC.item(6))
                Else
                    TEXTO6 = True
                End If
                If CAMPO7 = True Then
                    TEXTO7 = CAMPO7_BF = IIf(IsDBNull(R_BC.item(7)), "", R_BC.item(7))
                Else
                    TEXTO7 = True
                End If
                If CAMPO8 = True Then
                    TEXTO8 = CAMPO8_BF = IIf(IsDBNull(R_BC.item(8)), "", R_BC.item(8))
                Else
                    TEXTO8 = True
                End If
                If CAMPO9 = True Then
                    TEXTO9 = CAMPO9_BF = IIf(IsDBNull(R_BC.item(9)), "", R_BC.item(9))
                Else
                    TEXTO9 = True
                End If
                If CAMPO10 = True Then
                    TEXTO10 = CAMPO10_BF = IIf(IsDBNull(R_BC.item(10)), "", R_BC.item(10))
                Else
                    TEXTO10 = True
                End If
                '-----------------------------------------------------------------
                'Subtração e colocar dados na DT resultado

                If TEXTO1 And TEXTO2 And TEXTO3 And TEXTO4 And TEXTO5 And TEXTO6 And TEXTO7 And TEXTO8 _
                    And TEXTO9 And TEXTO10 Then

                    'Valores unit.
                    VOC_UNIT = CDec(R_BC.Item(13)) / CDec(R_BC.Item(11))
                    DAC_UNIT = CDec(R_BC.Item(14)) / CDec(R_BC.Item(11))
                    'Diminui BC
                    If QUANTIDADE_BF >= CDec(R_BC.Item(11)) Then
                        'Var Conciliado
                        BF_BC_CONCIL = CDec(R_BC.Item(11))
                        'Zera BC
                        Resultado_BC = 0
                    Else
                        Resultado_BC = CDec(R_BC.Item(11)) - QUANTIDADE_BF
                        'Var Conciliado
                        BF_BC_CONCIL = QUANTIDADE_BF
                    End If

                    'Diminui BF
                    Resultado_BF = QUANTIDADE_BF - CDec(R_BC.Item(11))
                    If QUANTIDADE_BF < 0 Then
                        Resultado_BF = 0
                    End If

                    R_BF.Item(11) = Resultado_BF
                    R_BC.Item(11) = Resultado_BC

                    'Arrumar DT_BC - VOC e DAC
                    R_BC.Item(13) = CDec(R_BC.Item(11)) * VOC_UNIT
                    R_BC.Item(14) = CDec(R_BC.Item(11)) * DAC_UNIT
                    'Preencher DT resultado
                    Status = "CONCILIADO"
                    n_CO += BF_BC_CONCIL
                    DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), R_BC.Item(12), VOC_UNIT * BF_BC_CONCIL, DAC_UNIT * BF_BC_CONCIL, BF_BC_CONCIL,
                                          Status, CHAVE_BF, CAMPO1_BF, CAMPO2_BF, CAMPO3_BF, CAMPO4_BF,
                                          CAMPO5_BF, CAMPO6_BF, CAMPO7_BF, CAMPO8_BF, CAMPO9_BF,
                                          CAMPO10_BF, BF_BC_CONCIL)

                    '------------------------------------------------------------------
                    'Limpar BC
                    If CDec(R_BC.Item(11)) <= 0 Then
                        R_BC.Item(11) = 0
                    End If
                End If
Prox_BC:
            Next
            'Limpar BF
            If CDec(R_BF.Item(11)) <= 0 Then
                R_BF.Item(11) = 0
            End If

Prox_BF:
        Next
Err:
        'Limpar DT
        Dim Limp_BF As New ArrayList
        Dim Limp_BC As New ArrayList
        Dim Loop_n_BF As Integer
        Dim Loop_n_BC As Integer
        Dim n_limpos_BF As Integer = 0
        Dim n_limpos_BC As Integer = 0

        For Each R_BF In DT_BF.Rows
            Loop_n_BF += 1
            If R_BF.Item(11) <= 0 Then
                Limp_BF.Add(Loop_n_BF - 1)
            End If
        Next

        For Each R_BC In DT_BC.Rows
            Loop_n_BC += 1
            If R_BC.Item(11) <= 0 Then
                Limp_BC.Add(Loop_n_BC - 1)
            End If
        Next

        If Limp_BF.Count > 0 Then
            For k = 0 To Limp_BF.Count - 1
                DT_BF.Rows.RemoveAt(Limp_BF(k) - n_limpos_BF)
                n_limpos_BF += 1
            Next
        End If

        If Limp_BC.Count > 0 Then
            For k = 0 To Limp_BC.Count - 1
                DT_BC.Rows.RemoveAt(Limp_BC(k) - n_limpos_BC)
                n_limpos_BC += 1
            Next
        End If

        Txt.Text = Txt.Text & IIf(Txt.Text = "", "", vbCrLf) & " | RODADA: " & n_Rodada & " | CONCILIADO: " &
            n_CO & " | CAMPOS: " & Campos
        'Devolver resultado para DGV
        DGV_RESULTADO.ItemsSource = DT_RESULTADO.DefaultView
        DGV_BC.ItemsSource = DT_BC.DefaultView
        DGV_BF.ItemsSource = DT_BF.DefaultView
        MW.Hide()
    End Sub

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

    Private Delegate Sub UpdateProgressBarDelegate(ByVal dp As _
             System.Windows.DependencyProperty,
             ByVal value As Object)
    Private Sub Process(ByRef Linhas As Single, ByRef Linhas_Totais As Single)
        Try

            'Create a new instance of our ProgressBar Delegate that points
            ' to the ProgressBar's SetValue method.
            value = (Linhas / (Linhas_Totais)) * 100
            Dim updatePbDelegate As New _
        UpdateProgressBarDelegate(AddressOf MW.PB.SetValue)
            MW.Dispatcher.Invoke(updatePbDelegate,
            System.Windows.Threading.DispatcherPriority.Background,
            New Object() {ProgressBar.ValueProperty, value})
        Catch
        End Try
    End Sub
End Class
