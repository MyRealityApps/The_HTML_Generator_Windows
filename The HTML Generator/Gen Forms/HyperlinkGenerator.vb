Public Class HyperlinkGenerator

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
        ElseIf CheckBox3.Checked = False Then
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox7.Enabled = False
        End If
    End Sub

    Private Sub HyperlinkGenerator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Enabled = True
        ElseIf CheckBox1.Checked = False Then
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox4.Enabled = True
        ElseIf CheckBox2.Checked = False Then
            TextBox4.Enabled = False
        End If
    End Sub
End Class