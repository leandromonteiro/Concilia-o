Imports Conciliação.BD
Imports System.Data
Imports Microsoft.Win32
Imports System.ComponentModel

Class MainWindow

    Public FileName As String
    Dim BD As New BD
    Dim Limite_Primeira As Boolean

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        'Botao Limpar
        BD.Modelo_Excel()
    End Sub

    Private Sub BtnLimparCb_Click(sender As Object, e As RoutedEventArgs) Handles BtnLimparCb.Click
        Cb1.IsChecked = False
        Cb2.IsChecked = False
        Cb3.IsChecked = False
        Cb4.IsChecked = False
        Cb5.IsChecked = False
        Cb6.IsChecked = False
        Cb7.IsChecked = False
        Cb8.IsChecked = False
        Cb9.IsChecked = False
        Cb10.IsChecked = False
    End Sub

    Private Sub MenuItem_Click_1(sender As Object, e As RoutedEventArgs)
        MiArquivo.IsEnabled = False
        MiArquivo.Header = "Aguarde a Importação do Arquivo"
        Limite_Primeira = False

        Dim OFD As New OpenFileDialog
        OFD.DefaultExt = ".xlsx"
        OFD.Filter = "Documentos Excel (.xlsx)|*.xlsx"
        Dim result As Nullable(Of Boolean) = OFD.ShowDialog()

        If result = True Then
            FileName = OFD.FileName
        Else
            MiArquivo.IsEnabled = True
            MiArquivo.Header = "Arquivo"
            Exit Sub
        End If

        BD.Importar_Excel(FileName, DgBF, DgBC)
        DgResultado.ItemsSource = ""
        TxtRodadas.Text = ""
        MiArquivo.IsEnabled = True
        MiArquivo.Header = "Arquivo"
    End Sub

    Private Sub BtnSelecionar_Click(sender As Object, e As RoutedEventArgs) Handles BtnSelecionar.Click
        Cb1.IsChecked = True
        Cb2.IsChecked = True
        Cb3.IsChecked = True
        Cb4.IsChecked = True
        Cb5.IsChecked = True
        Cb6.IsChecked = True
        Cb7.IsChecked = True
        Cb8.IsChecked = True
        Cb9.IsChecked = True
        Cb10.IsChecked = True
    End Sub

    Private Sub BtnConciliar_Click(sender As Object, e As RoutedEventArgs) Handles BtnConciliar.Click
        Dim Array_Campos As New ArrayList
        Dim Campos As String = ""

        Me.Hide()
        'Validação
        If Validar() = False Then
            Exit Sub
        End If
        BtnConciliar.IsEnabled = False
        BD.Classificar_BD(CmbPrioridade.Text, CmbOrdem.Text)
        'Limpar limite somente na primeira vez
        If Limite_Primeira = False Then
            BD.Limpar_Limite(TxtMinFis.Text, TxtMinCont.Text, TxtRodadas)
        End If
        Limite_Primeira = True

        'Campos Conciliação
        If Cb1.IsChecked = True Then Array_Campos.Add("1")
        If Cb2.IsChecked = True Then Array_Campos.Add("2")
        If Cb3.IsChecked = True Then Array_Campos.Add("3")
        If Cb4.IsChecked = True Then Array_Campos.Add("4")
        If Cb5.IsChecked = True Then Array_Campos.Add("5")
        If Cb6.IsChecked = True Then Array_Campos.Add("6")
        If Cb7.IsChecked = True Then Array_Campos.Add("7")
        If Cb8.IsChecked = True Then Array_Campos.Add("8")
        If Cb9.IsChecked = True Then Array_Campos.Add("9")
        If Cb10.IsChecked = True Then Array_Campos.Add("10")
        For K = 0 To Array_Campos.Count - 1
            If K = 0 Then
                Campos = Array_Campos(K)
            Else
                Campos = Campos & " , " & Array_Campos(K)
            End If
        Next


        'Conciliar
        BD.Conciliar(DgBF, DgBC, DgResultado, Cb1.IsChecked, Cb2.IsChecked, Cb3.IsChecked, Cb4.IsChecked,
                     Cb5.IsChecked, Cb6.IsChecked, Cb7.IsChecked, Cb8.IsChecked, Cb9.IsChecked, Cb10.IsChecked,
                     TxtRodadas, Campos)
        BtnConciliar.IsEnabled = True
        Me.Show()
    End Sub
    Private Function Validar() As Boolean
        On Error Resume Next

        If Cb1.IsChecked = False And Cb2.IsChecked = False And Cb3.IsChecked = False And Cb4.IsChecked = False _
        And Cb5.IsChecked = False And Cb6.IsChecked = False And Cb7.IsChecked = False And Cb8.IsChecked = False _
        And Cb9.IsChecked = False And Cb10.IsChecked = False Then
            MsgBox("Selecione o(s) campo(s).", vbInformation)
            Validar = False
            Exit Function
        End If

        If CmbPrioridade.Text = "" Then
            MsgBox("Selecione a prioridade.", vbInformation)
            Validar = False
            Exit Function
        End If

        If CmbOrdem.Text = "" Then
            MsgBox("Selecione a ordem.", vbInformation)
            Validar = False
            Exit Function
        End If

        If TxtMinCont.Text = "" Or IsNumeric(TxtMinCont.Text) = False Then
            MsgBox("Mínimo valor para a base contábil incorreta.", vbInformation)
            Validar = False
            Exit Function
        End If

        If TxtMinFis.Text = "" Or IsNumeric(TxtMinFis.Text) = False Then
            MsgBox("Mínimo valor para a base física incorreta.", vbInformation)
            Validar = False
            Exit Function
        End If

        If DgBF.Items.Count = 0 Or DgBC.Items.Count = 0 Then
            MsgBox("Insira a base física e/ou base contábil.", vbInformation)
            Validar = False
            Exit Function
        End If
        Validar = True
    End Function

    Private Sub MenuItem_Click_2(sender As Object, e As RoutedEventArgs)
        If TxtRodadas.Text = "" Then
            Exit Sub
        End If
        Me.Hide()
        MiArquivo.IsEnabled = False
        BD.Exportacao_SF_SC(TxtRodadas)
        BD.Juntar_DT()
        BD.Exportar_Excel(TxtRodadas, CInt(Slide_Qtd.Value), CInt(Slider_Valor.Value))
        MiArquivo.IsEnabled = True
        Me.Show()
    End Sub

    Private Sub MenuItem_Click_3(sender As Object, e As RoutedEventArgs)
        BD.Zerar_Conciliacao(DgBF, DgBC, DgResultado)
        TxtRodadas.Text = ""
    End Sub

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        CmbPrioridade.Items.Add("Valor")
        CmbPrioridade.Items.Add("Data")
        CmbOrdem.Items.Add("Crescente")
        CmbOrdem.Items.Add("Decrescente")
        BD.Criar_DT_Resultado()
    End Sub

    Private Sub Slide_Qtd_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles Slide_Qtd.ValueChanged
        Slide_Qtd.ToolTip = Slide_Qtd.Value
    End Sub

    Private Sub Slider_Valor_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles Slider_Valor.ValueChanged
        Slider_Valor.ToolTip = Slider_Valor.Value
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Application.Current.Shutdown()
    End Sub
End Class
