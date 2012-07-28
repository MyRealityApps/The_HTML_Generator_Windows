Public Class HyperlinkGenerator

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
        ElseIf CheckBox3.Checked = False Then
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
        End If
    End Sub

    Private Sub HyperlinkGenerator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox3.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Enabled = True
        ElseIf CheckBox1.Checked = False Then
            TextBox3.Enabled = False
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ErrorMessage As String = "You must include a _"
        Dim TitleError As String = "Error: Missing Information"
        If TextBox1.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the website's address."), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the link's clickable text."), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class