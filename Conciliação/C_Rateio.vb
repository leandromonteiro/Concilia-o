Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data
Imports System.Linq

Public Class C_Rateio
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
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Rateio_BF_BC"
            Sh_T.Range("a1").Value = "SEQUÊNCIAL BASE FÍSICA"
            Sh_T.Range("b1").Value = "CLASSIFICAÇÃO BASE FÍSICA"
            Sh_T.Range("c1").Value = "VALOR BASE FÍSICA"
            Sh_T.Range("d1").Value = "SEQUÊNCIAL BASE CONTÁBIL"
            Sh_T.Range("e1").Value = "CLASSIFICAÇÃO BASE CONTÁBIL"
            Sh_T.Range("f1").Value = "VALOR BASE CONTÁBIL"
            Sh_T.Columns.AutoFit()
            Sh_T.Range("a1:f1").Font.Bold = True
            Sh_T.Range("a1:f1").Font.ColorIndex = 2
            Sh_T.Range("a1:c1").Interior.ColorIndex = 12
            Sh_T.Range("d1:f1").Interior.ColorIndex = 14
            xlApp.Visible = True

        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Importar_Excel(ArquivoExcel As String, DGV As DataGrid)
        'Try
        Dim con As New OleDbConnection
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ArquivoExcel & "';Extended Properties= 'Excel 12.0';"
        Dim cmd As New OleDbCommand
        Dim DA As New OleDbDataAdapter
        con.Open()
        DA.SelectCommand = New OleDbCommand("SELECT [SEQUÊNCIAL BASE FÍSICA],[CLASSIFICAÇÃO BASE FÍSICA],format([VALOR BASE FÍSICA],'#,###.00') as [VALOR BASE FÍSICA]" &
            ",[SEQUÊNCIAL BASE CONTÁBIL],[CLASSIFICAÇÃO BASE CONTÁBIL],format([VALOR BASE CONTÁBIL],'#,###.00') as [VALOR BASE CONTÁBIL] FROM [Rateio_BF_BC$];", con)
        Dim DS As New DataSet
        DA.Fill(DS, "TB_Local")
        DGV.ItemsSource = DS.Tables("TB_Local").AsDataView
        con.Close()
        'Dividir BF e BC em DataTables
        DT_BC.Clear()
        DT_BF.Clear()

        DT_BF = DS.Tables("TB_Local").Clone
        DT_BC = DS.Tables("TB_Local").Clone

        For i = 0 To DS.Tables("TB_Local").Rows.Count - 1
            DT_BF.ImportRow(DS.Tables("TB_Local").Rows(i))
            DT_BC.ImportRow(DS.Tables("TB_Local").Rows(i))
        Next

        DT_BF.Columns.RemoveAt(5)
        DT_BF.Columns.RemoveAt(4)
        DT_BF.Columns.RemoveAt(3)

        DT_BC.Columns.RemoveAt(0)
        DT_BC.Columns.RemoveAt(0)
        DT_BC.Columns.RemoveAt(0)

        'Remover dados vazios
        For s = 0 To DT_BC.Rows.Count - 1
            If IsDBNull(DT_BC.Rows(s).Item(0)) Then
                DT_BC.Rows(s).Delete()
            End If
        Next
        For x = 0 To DT_BF.Rows.Count - 1
            If IsDBNull(DT_BF.Rows(x).Item(0)) Then
                DT_BF.Rows(x).Delete()
            End If
        Next
        'Ordenar tabelas
        DT_BF = DT_BF.Select("", "CLASSIFICAÇÃO BASE FÍSICA,VALOR BASE FÍSICA asc").CopyToDataTable
        DT_BC = DT_BC.Select("", "CLASSIFICAÇÃO BASE CONTÁBIL,VALOR BASE CONTÁBIL asc").CopyToDataTable

        MsgBox("Dados carregados com sucesso", MsgBoxStyle.Information)
        'Catch
        '    MsgBox("Erro ao Carregar os Dados", MsgBoxStyle.Critical)
        'End Try
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
                        xlWorkSheet.Cells(i + 2, j + 1) = IIf(Not IsDBNull(DV_Excel.Item(i)(j)), CDec(DV_Excel.Item(i)(j)), "")
                        xlWorkSheet.Cells(i + 2, j + 1).numberformat = "#,###.00"
                    Else
                        xlWorkSheet.Cells(i + 2, j + 1) = DV_Excel.Item(i)(j)
                    End If

                Next
            Next

            xlWorkSheet.Range("a1:d1").Font.Bold = True
            xlWorkSheet.Range("a1:d1").Font.ColorIndex = 2
            xlWorkSheet.Range("a1:d1").Interior.ColorIndex = 30

            xlWorkSheet.Columns.AutoFit()
            xlWorkSheet.Name = "Rateio"
            xlApp.Visible = True
        Catch
            MsgBox("Erro ao exportar para Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Calculo(DGV As DataGrid)
        If DT_Rateio.Columns.Count = 0 Then
            DT_Rateio.Columns.Add("SEQUÊNCIAL BASE FÍSICA")
            DT_Rateio.Columns.Add("SEQUÊNCIAL BASE CONTÁBIL")
            DT_Rateio.Columns.Add("CLASSIFICAÇÃO")
            DT_Rateio.Columns.Add("VALOR RATEADO")
        End If
        'Soma.Se das classif.
        Dim qry_Class = From dr In DT_BF.AsEnumerable
                        Group dr By grupo = dr("CLASSIFICAÇÃO BASE FÍSICA")
                      Into Group
                        Select Grupos = grupo

        Dim qry_Valor = From dr In DT_BF
                        Group dr By grupo = dr("CLASSIFICAÇÃO BASE FÍSICA")
                      Into Group
                        Select Total = Group.Sum(Function(dr) Decimal.Parse(dr("VALOR BASE FÍSICA")))

        Dim Linhas_Total As Integer = qry_Class.Count

        Dim Classifica As Array = qry_Class.ToArray
        Dim Valor As Array = qry_Valor.ToArray

        'Ratear por classif. multip. valor BC com a % BF

        For F = 0 To DT_BF.Rows.Count - 1
            For C = 0 To DT_BC.Rows.Count - 1
                For L = 0 To Linhas_Total - 1
                    If CStr(DT_BF.Rows(F)(1)) = CStr(DT_BC.Rows(C)(1)) And CStr(DT_BC.Rows(C)(1)) = CStr(Classifica(L)) Then
                        DT_Rateio.Rows.Add(DT_BF.Rows(F)(0), DT_BC.Rows(C)(0), Classifica(L), Format((DT_BF.Rows(F)(2) / Valor(L)) * DT_BC.Rows(C)(2), "#,###.00"))
                    End If
                Next
            Next
        Next
        DGV.ItemsSource = ""
        DGV.ItemsSource = DT_Rateio.AsDataView
    End Sub

End Class
