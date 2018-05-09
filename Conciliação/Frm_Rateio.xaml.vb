Imports System.Windows.Media.Effects
Imports System.Windows.Media.Animation
Imports Microsoft.Win32

Public Class Frm_Rateio
    Dim Magia_F As New C_Rateio
    Dim ArquivoExcel As String
    Dim OFD As New OpenFileDialog
    'Dim Frm_Magia_Principal As New UserControl

    Private Sub Frm_Rateio_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        BtnCalcular.IsEnabled = False
        BtnExcel.IsEnabled = False
    End Sub

    Private Sub BtnModelo_Click(sender As Object, e As RoutedEventArgs) Handles BtnModelo.Click
        BtnModelo.IsEnabled = False
        'Gerar modelo Excel
        Magia_F.Modelo_Excel()
        BtnModelo.IsEnabled = True
    End Sub
    Private Sub BtnCarga_Click(sender As Object, e As RoutedEventArgs) Handles BtnCarga.Click
        Try
            OFD.Title = "Selecione a carga"
            OFD.Filter = "Excel (*.xlsx)|*.xlsx"
            OFD.FileName = ""
            OFD.ShowDialog()
            ArquivoExcel = OFD.FileName
        Catch
        End Try
        'Sem o arquivo
        If ArquivoExcel = "" Then
            Exit Sub
        End If
        Magia_F.Importar_Excel(ArquivoExcel, DGV_Magia)
        BtnCalcular.IsEnabled = True
    End Sub
    Private Sub BtnCalcular_Click(sender As Object, e As RoutedEventArgs) Handles BtnCalcular.Click
        'Ordenar
        Tab_Rateio.IsSelected = True
        Magia_F.Calculo(DGV_Rateio)
        BtnExcel.IsEnabled = True
        BtnCalcular.IsEnabled = False
    End Sub
    Private Sub BtnExcel_Click(sender As Object, e As RoutedEventArgs) Handles BtnExcel.Click
        'Exporta para Excel
        Me.Cursor = Cursors.Wait
        BtnExcel.IsEnabled = False
        BtnExcel.Content = "Aguarde"
        Magia_F.Exportar_Excel(DGV_Rateio)
        BtnExcel.IsEnabled = True
        BtnExcel.Content = "Excel"
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub Efeito_Borrado()
        Dim blurEffect As New BlurEffect()
        blurEffect.Radius = 2
        BtnModelo.Effect = blurEffect
    End Sub

    Private Sub Efeito_Sombra(Sombra As Integer, Btn As Button)
        Dim DSE As New DropShadowEffect
        DSE.BlurRadius = Sombra
        DSE.ShadowDepth = Sombra
        Btn.Effect = DSE
    End Sub
    Private Sub BtnCalcular_MouseLeave(sender As Object, e As MouseEventArgs) Handles BtnCalcular.MouseLeave
        Efeito_Sombra(0, BtnCalcular)
    End Sub
    Private Sub BtnCalcular_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnCalcular.MouseMove
        Efeito_Sombra(5, BtnCalcular)
    End Sub
    Private Sub BtnCarga_MouseLeave(sender As Object, e As MouseEventArgs) Handles BtnCarga.MouseLeave
        Efeito_Sombra(0, BtnCarga)
    End Sub
    Private Sub BtnCarga_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnCarga.MouseMove
        Efeito_Sombra(5, BtnCarga)
    End Sub
    Private Sub BtnModelo_MouseLeave(sender As Object, e As MouseEventArgs) Handles BtnModelo.MouseLeave
        Efeito_Sombra(0, BtnModelo)
    End Sub
    Private Sub BtnModelo_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnModelo.MouseMove
        Efeito_Sombra(5, BtnModelo)
    End Sub
    Private Sub BtnExcel_MouseLeave(sender As Object, e As MouseEventArgs) Handles BtnExcel.MouseLeave
        Efeito_Sombra(0, BtnExcel)
    End Sub
    Private Sub BtnExcel_MouseMove(sender As Object, e As MouseEventArgs) Handles BtnExcel.MouseMove
        Efeito_Sombra(5, BtnExcel)
    End Sub

End Class
