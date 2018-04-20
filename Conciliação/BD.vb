Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data

Public Class BD
    Dim DT_BF As New DataTable
    Dim DT_BC As New DataTable
    Dim DT_Rateio As New DataTable
    Dim alterado_positivo As Boolean
    Dim alterado_negativo As Boolean
    Dim DV_Excel As New DataView
    Dim N_BC As Integer
    Dim BF_Total As Double

    Dim BC_Seq_Atual As String
    Dim BC_Valor_Atual As Double

    Dim First_Loop As Integer = 0

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
            DA_BF.SelectCommand = New OleDbCommand("SELECT * FROM [Base Física$];", con)
            DA_BC.SelectCommand = New OleDbCommand("SELECT [CHAVE],[CAMPO1],[CAMPO2],[CAMPO3],[CAMPO4],[CAMPO5],[CAMPO6],[CAMPO7],[CAMPO8],[CAMPO9],[CAMPO10],[QUANTIDADE], left([DATA],2) & '/' & mid([DATA],4,2) & '/' & right([DATA],4) as DATA,[VOC],[DAC] FROM [Base Contábil$];", con)
            Dim DS As New DataSet
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

            'Ordenar tabelas
            'DT_BF = DT_BF.Select("", "CLASSIFICAÇÃO BASE FÍSICA,VALOR BASE FÍSICA asc").CopyToDataTable
            'DT_BC = DT_BC.Select("", "CLASSIFICAÇÃO BASE CONTÁBIL,VALOR BASE CONTÁBIL asc").CopyToDataTable

            MsgBox("Dados carregados com sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao Carregar os Dados", MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
