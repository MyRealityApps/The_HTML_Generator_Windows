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
        'The following "Dim's" declare their representative strings. Declaring objects make it easier to reuse them in the same action.
        Dim ErrorMessage As String = "You must include a _"
        Dim TitleError As String = "Error: Missing Information"
        Dim TextLink As String = "<a href=""address"">clickable_text</a>"
        Dim ImageLink As String = "<a href=""address""><img src=""image_file"" alt=""ALT"" height=""heightPX"" width=""widthPX"" /></a>"
        Dim LinkWithClass As String = "class=""the_class"""

        'The Following Code checks the length of TextBox1 and TextBox2. If the two textboxes are NOT empty, then generate a text hyperlink.
        If TextBox1.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the website's address."), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the link's clickable text"), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            'This generates the text hyperlink. Notice how easy it is to reuse the declared strings and just replace bits and pieces of them.
            GeneratedCode.RichTextBox1.Text = TextLink.Replace("address", TextBox1.Text).Replace("clickable_text", TextBox2.Text)
            GeneratedCode.Show()
        End If

    End Sub

    Private Sub HyperlinkGen_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim CloseResult As DialogResult = MessageBox.Show("Are you sure you want to close? Any inserted data will be lost!", "Hyperlink Generator: Closing", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If CloseResult = Windows.Forms.DialogResult.Yes Then
            e.Cancel = False
        ElseIf CloseResult = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub
End Class