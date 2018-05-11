Public Class Frm_Escolher
    Dim Ss As New SsConciliacao
    Dim Conciliacao As New MainWindow
    Dim Rateio As New Frm_Rateio
    Private Sub Btn_Conciliar_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Conciliar.Click
        Conciliacao.Show()
        Me.Hide()
    End Sub

    Private Sub Frm_Conciliar_Rateio_Initialized(sender As Object, e As EventArgs) Handles Frm_Conciliar_Rateio.Initialized
        Ss.Show()
    End Sub

    Private Sub Btn_Rateio_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Rateio.Click
        Rateio.Show()
        Me.Hide()
    End Sub

    Private Sub Frm_Escolher_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Application.Current.Shutdown()
    End Sub
End Class
