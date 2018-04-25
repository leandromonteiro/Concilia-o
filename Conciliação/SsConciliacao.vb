Public NotInheritable Class SsConciliacao
    Private Sub SsConciliacao_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title
        Else
            'Se o título da aplicação estiver faltando, utiliza o nome da aplicação sem a extensão
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)
        'Informação de Copyright
        Copyright.Text = My.Application.Info.Copyright
    End Sub

    Private Sub MainLayoutPanel_Paint(sender As Object, e As Forms.PaintEventArgs) Handles MainLayoutPanel.Paint

    End Sub
End Class
