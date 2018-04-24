Imports Conciliação.BD
Imports System.Data
Imports Microsoft.Win32
Class MainWindow
    Public FileName As String
    Dim BD As New BD
    Dim Limite_Primeira As Boolean
    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
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

    Private Sub PanelPrincipal_Initialized(sender As Object, e As EventArgs) Handles PanelPrincipal.Initialized
        CmbPrioridade.Items.Add("Valor")
        CmbPrioridade.Items.Add("Data")
        CmbOrdem.Items.Add("Crescente")
        CmbOrdem.Items.Add("Decrescente")
        BD.Criar_DT_Resultado()
    End Sub

    Private Sub MenuItem_Click_1(sender As Object, e As RoutedEventArgs)
        Dim OFD As New OpenFileDialog
        OFD.DefaultExt = ".xlsx"
        OFD.Filter = "Documentos Excel (.xlsx)|*.xlsx"
        Dim result As Nullable(Of Boolean) = OFD.ShowDialog()

        If result = True Then
            FileName = OFD.FileName
        Else
            Exit Sub
        End If
        Limite_Primeira = False
        BD.Importar_Excel(FileName, DgBF, DgBC)
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
        'Validação
        Validar()
        BD.Classificar_BD(CmbPrioridade.Text, CmbOrdem.Text)
        'Limpar limite somente na primeira vez
        If Limite_Primeira = True Then
            BD.Limpar_Limite(TxtMinFis.Text, TxtMinCont.Text, DgBF, DgBC)
        End If
        Limite_Primeira = True

        'Conciliar
        BD.Conciliar(DgBF, DgBC, DgResultado, Cb1.IsChecked, Cb2.IsChecked, Cb3.IsChecked, Cb4.IsChecked,
                     Cb5.IsChecked, Cb6.IsChecked, Cb7.IsChecked, Cb8.IsChecked, Cb9.IsChecked, Cb10.IsChecked, PbConciliar)
    End Sub
    Private Sub Validar()
        On Error Resume Next

        If Cb1.IsChecked = False And Cb2.IsChecked = False And Cb3.IsChecked = False And Cb4.IsChecked = False _
        And Cb5.IsChecked = False And Cb6.IsChecked = False And Cb7.IsChecked = False And Cb8.IsChecked = False _
        And Cb9.IsChecked = False And Cb10.IsChecked = False Then
            MsgBox("Selecione o(s) campo(s).", vbInformation)
            Exit Sub
        End If

        If CmbPrioridade.Text = "" Then
            MsgBox("Selecione a prioridade.", vbInformation)
            Exit Sub
        End If

        If CmbOrdem.Text = "" Then
            MsgBox("Selecione a ordem.", vbInformation)
            Exit Sub
        End If

        If TxtMinCont.Text = "" Or IsNumeric(TxtMinCont.Text) = False Then
            MsgBox("Mínimo valor para a base contábil incorreta.", vbInformation)
            Exit Sub
        End If

        If TxtMinFis.Text = "" Or IsNumeric(TxtMinFis.Text) = False Then
            MsgBox("Mínimo valor para a base física incorreta.", vbInformation)
            Exit Sub
        End If

        If DgBF.Items.Count = 0 Or DgBC.Items.Count = 0 Then
            MsgBox("Insira a base física e/ou base contábil.", vbInformation)
            Exit Sub
        End If
    End Sub
End Class
