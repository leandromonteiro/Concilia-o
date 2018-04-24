Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data

Public Class BD
    Dim DS As New DataSet
    Public DT_BF As New DataTable
    Public DT_BC As New DataTable
    Public DT_RESULTADO As New DataTable
    Dim DV_Excel As New DataView

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

    Public Sub Exportar_Excel(DGV As DataGrid)
        Try
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer

            DV_Excel = DGV.ItemsSource

            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets(1)

            'Colocando Títulos
            For k As Integer = 1 To DGV.Columns.Count
                xlWorkSheet.Cells(1, k) = DGV.Columns(k - 1).Header
            Next
            For i = 0 To DGV.Items.Count - 1
                For j = 0 To DGV.Columns.Count - 1
                    If j = 3 Then
                        xlWorkSheet.Cells(i + 2, j + 1) = CDec(DV_Excel.Item(i)(j))
                    Else
                        xlWorkSheet.Cells(i + 2, j + 1) = DV_Excel.Item(i)(j)
                    End If
                Next
            Next

            xlWorkSheet.Range("a1:d1").Font.Bold = True
            xlWorkSheet.Range("a1:d1").Font.ColorIndex = 2
            xlWorkSheet.Range("a1:d1").Interior.ColorIndex = 30

            xlWorkSheet.Columns.AutoFit()
            xlWorkSheet.Name = "Resultado"
            xlApp.Visible = True
        Catch
            MsgBox("Erro ao exportar para Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Importar_Excel(ArquivoExcel As String, DGV_BF As DataGrid, DGV_BC As DataGrid)
        Try
            Dim con As New OleDbConnection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ArquivoExcel & "';Extended Properties= 'Excel 12.0';"
            Dim cmd As New OleDbCommand
            Dim DA_BF As New OleDbDataAdapter
            Dim DA_BC As New OleDbDataAdapter
            con.Open()
            DA_BF.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],abs([QUANTIDADE]) as QUANTIDADE,[PRIORIDADE] FROM [Base Física$];", con)
            DA_BC.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE],[DATA],abs([VOC]) as VOC,abs([DAC]) as DAC FROM [Base Contábil$];", con)

            DA_BF.Fill(DS, "TB_BF")
            DA_BC.Fill(DS, "TB_BC")
            con.Close()
            'Dividir BF e BC em DataTables
            'Limpar
            DT_BC.Clear()
            DT_BF.Clear()
            DGV_BF.ItemsSource = ""
            DGV_BC.ItemsSource = ""

            DT_BF = DS.Tables("TB_BF")
            DT_BC = DS.Tables("TB_BC")

            DGV_BF.ItemsSource = DS.Tables("TB_BF").DefaultView
            DGV_BC.ItemsSource = DS.Tables("TB_BC").DefaultView

            'Remover dados vazios
            'For s = 0 To DT_BC.Rows.Count - 1
            'If IsDBNull(DT_BC.Rows(s).Item(0)) Then
            'DT_BC.Rows(s).Delete()
            'End If
            'Next

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

    Public Sub Limpar_Limite(Limite_F As Single, Limite_C As Single, DGV_BF As DataGrid, DGV_BC As DataGrid)
        For Each R In DT_BF.Rows
            If R.Item(11) <= Limite_F Then
                R.Delete()
            End If
        Next

        For Each R In DT_BC.Rows
            If R.Item(11) <= Limite_C Then
                R.Delete()
            End If
        Next

        DGV_BF.ItemsSource = ""
        DGV_BC.ItemsSource = ""

        DGV_BF.ItemsSource = DT_BF.DefaultView
        DGV_BC.ItemsSource = DT_BC.DefaultView
    End Sub

    Public Sub Conciliar(DGV_BF As DataGrid, DGV_BC As DataGrid, DGV_RESULTADO As DataGrid, CAMPO1 As Boolean,
                         CAMPO2 As Boolean, CAMPO3 As Boolean, CAMPO4 As Boolean, CAMPO5 As Boolean, CAMPO6 As Boolean,
                         CAMPO7 As Boolean, CAMPO8 As Boolean, CAMPO9 As Boolean, CAMPO10 As Boolean, PB As ProgressBar)
        Dim TEXTO1 As Object
        Dim TEXTO2 As Object
        Dim TEXTO3 As Object
        Dim TEXTO4 As Object
        Dim TEXTO5 As Object
        Dim TEXTO6 As Object
        Dim TEXTO7 As Object
        Dim TEXTO8 As Object
        Dim TEXTO9 As Object
        Dim TEXTO10 As Object

        Dim VOC_UNIT As Single
        Dim DAC_UNIT As Single
        Dim BF_BC_CONCIL As Single
        Dim Resultado_BF As Single
        Dim Resultado_BC As Single
        Dim Deletar_BF As ArrayList
        Dim Status As String

        PB.Visibility = Visibility.Visible
        Dim n_BF As Integer
        Dim n_BC As Integer

        For Each R_BF In DT_BF.Rows
            n_BF = n_BF + 1
            n_BC = 0
            PB.Value = n_BF / DT_BF.Rows.Count * 100
            For Each R_BC In DT_BC.Rows
                n_BC = n_BC + 1
                On Error GoTo Err
                'Dados selecionados
                If CAMPO1 = True Then
                    TEXTO1 = R_BF.item(1) = R_BC.item(1)
                Else
                    TEXTO1 = R_BF.item(1) = R_BF.item(1)
                End If
                If CAMPO2 = True Then
                    TEXTO2 = R_BF.item(2) = R_BC.item(2)
                Else
                    TEXTO2 = R_BF.item(2) = R_BF.item(2)
                End If
                If CAMPO3 = True Then
                    TEXTO3 = R_BF.item(3) = R_BC.item(3)
                Else
                    TEXTO3 = R_BF.item(3) = R_BF.item(3)
                End If
                If CAMPO4 = True Then
                    TEXTO4 = R_BF.item(4) = R_BC.item(4)
                Else
                    TEXTO4 = R_BF.item(4) = R_BF.item(4)
                End If
                If CAMPO5 = True Then
                    TEXTO5 = R_BF.item(5) = R_BC.item(5)
                Else
                    TEXTO5 = R_BF.item(5) = R_BF.item(5)
                End If
                If CAMPO6 = True Then
                    TEXTO6 = R_BF.item(6) = R_BC.item(6)
                Else
                    TEXTO6 = R_BF.item(6) = R_BF.item(6)
                End If
                If CAMPO7 = True Then
                    TEXTO7 = R_BF.item(7) = R_BC.item(7)
                Else
                    TEXTO7 = R_BF.item(7) = R_BF.item(7)
                End If
                If CAMPO8 = True Then
                    TEXTO8 = R_BF.item(8) = R_BC.item(8)
                Else
                    TEXTO8 = R_BF.item(8) = R_BF.item(8)
                End If
                If CAMPO9 = True Then
                    TEXTO9 = R_BF.item(9) = R_BC.item(9)
                Else
                    TEXTO9 = R_BF.item(9) = R_BF.item(9)
                End If
                If CAMPO10 = True Then
                    TEXTO10 = R_BF.item(10) = R_BC.item(10)
                Else
                    TEXTO10 = R_BF.item(10) = R_BF.item(10)
                End If
                '-----------------------------------------------------------------
                'Subtração e colocar dados na DT resultado

                If TEXTO1 And TEXTO2 And TEXTO3 And TEXTO4 And TEXTO5 And TEXTO6 And TEXTO7 And TEXTO8 _
                    And TEXTO9 And TEXTO10 Then

                    If R_BC.Item(11) <= 0 Then
                        GoTo Prox
                    End If
                    'Valores unit.
                    VOC_UNIT = R_BC.Item(13) / R_BC.Item(11)
                    DAC_UNIT = R_BC.Item(14) / R_BC.Item(11)
                    'Diminui BC
                    If R_BF.Item(11) >= R_BC.Item(11) Then
                        'Var Conciliado
                        BF_BC_CONCIL = R_BC.Item(11)
                        'Zera BC
                        Resultado_BC = 0
                    Else
                        Resultado_BC = R_BC.Item(11) - R_BF.Item(11)
                        'Var Conciliado
                        BF_BC_CONCIL = R_BF.Item(11)
                    End If

                    'Diminui BF
                    Resultado_BF = R_BF.Item(11) - R_BC.Item(11)
                    If R_BF.Item(11) < 0 Then
                        Resultado_BF = 0
                    End If

                    R_BF.Item(11) = Resultado_BF
                    R_BC.Item(11) = Resultado_BC

                    'Preencher DT resultado
                    Status = "CONCILIADO"
                    DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), R_BC.Item(12), VOC_UNIT * BF_BC_CONCIL, DAC_UNIT * BF_BC_CONCIL, BF_BC_CONCIL,
                                          Status, R_BF.Item(0), R_BF.Item(1), R_BF.Item(2), R_BF.Item(3), R_BF.Item(4),
                                          R_BF.Item(5), R_BF.Item(6), R_BF.Item(7), R_BF.Item(8), R_BF.Item(9),
                                          R_BF.Item(10), BF_BC_CONCIL)
                    If R_BC.Item(11) > 0 Then
                            Status = "SOBRA CONTÁBIL"
                            DT_RESULTADO.Rows.Add(R_BC.Item(0), R_BC.Item(1), R_BC.Item(2), R_BC.Item(3), R_BC.Item(4),
                                          R_BC.Item(5), R_BC.Item(6), R_BC.Item(7), R_BC.Item(8), R_BC.Item(9),
                                          R_BC.Item(10), R_BC.Item(12), VOC_UNIT * R_BC.Item(11), DAC_UNIT * R_BC.Item(11), R_BC.Item(11),
                                          Status, "", "", "", "", "", "", "", "", "", "", "", "")
                        End If
                    End If

                    '------------------------------------------------------------------
                    'Limpar BC
                    If R_BC.Item(11) <= 0 Then
                    R_BC.Item(11) = 0
                End If
            Next
            'Limpar BF
            If R_BF.Item(11) <= 0 Then
                R_BF.Item(11) = 0
            Else
                Status = "SOBRA FÍSICA"
                'Preencher DT resultado
                DT_RESULTADO.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                          Status, R_BF.Item(0), R_BF.Item(1), R_BF.Item(2), R_BF.Item(3), R_BF.Item(4),
                                          R_BF.Item(5), R_BF.Item(6), R_BF.Item(7), R_BF.Item(8), R_BF.Item(9),
                                          R_BF.Item(10), R_BF.Item(11))
            End If
Prox:
        Next
Err:
        'Limpar DT
        For Each R_BF In DT_BF.Rows
            If R_BF.Item(11) = 0 Then
                R_BF.delete
            End If
        Next

        For Each R_BC In DT_BC.Rows
            If R_BC.Item(11) = 0 Then
                R_BC.delete
            End If
        Next

        PB.Visibility = Visibility.Hidden

        'Devolver resultado para DGV
        DGV_RESULTADO.ItemsSource = DT_RESULTADO.DefaultView
        DGV_BC.ItemsSource = DT_BC.DefaultView
        DGV_BF.ItemsSource = DT_BF.DefaultView
    End Sub

End Class
