Public Class MainWindow

    Private Sub MainWindow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Text = "Hello World!"
    End Sub

    Private Sub mainList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItem = "Typography Editor" Then
            Button1.Text = "Open Typography Editor"
        End If
        If ListBox1.SelectedItem = "Hyperlink Generator" Then
            Button1.Text = "Open Hyperlink Generator"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Open Hyperlink Generator" Then
            HyperlinkGenerator.Show()
        End If
    End Sub
End Class
