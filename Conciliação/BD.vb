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
                         CAMPO2 As Boolean)
        Dim TEXTO1 As Object
        Dim TEXTO2 As Object

        For Each R_BF In DT_BF.Rows
            For Each R_BC In DT_BC.Rows
                'Fazer a subtração apenas dos campos selecionados
                If CAMPO1 = True Then
                    TEXTO1 = R_BF.item(1) = R_BC.item(1)
                Else
                    TEXTO1 = ""
                End If
                If CAMPO2 = True Then
                    TEXTO2 = R_BF.item(2) = R_BC.item(2)
                Else
                    TEXTO2 = ""
                End If
                If TEXTO1 And TEXTO2 Then
                    MsgBox("OK")
                End If
                'Limpar BF
                If R_BF.Item(11) <= 0 Then
                    R_BF.Delete()
                End If
            Next
        Next
    End Sub
End Class
