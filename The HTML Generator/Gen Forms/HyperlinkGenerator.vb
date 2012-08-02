Public Class HyperlinkGenerator
#Region "Custom Functions"
    'A region for custom functions.

    Private Sub HyperlinkGen_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'Declares a DialogResult then uses the dialogresult to show a MessageBox. The MessageBox asks the user if they want to close the app.
        Dim CloseResult As DialogResult = MessageBox.Show("Are you sure you want to close? Any inserted data will be lost!", "Hyperlink Generator: Closing", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If CloseResult = Windows.Forms.DialogResult.Yes Then
            e.Cancel = False
        ElseIf CloseResult = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

#End Region

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        'The following code sees if CheckBox1 is enabled. If it is enabled, then enable TextBox3.
        If CheckBox1.Checked = True Then
            TextBox3.Enabled = True
        ElseIf CheckBox1.Checked = False Then
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Declare strings so we can easily use them later.
        Dim ErrorMessage As String = "You must include a _"
        Dim TitleError As String = "Error: Missing Information"
        Dim TextLink As String = "<a href=""address"">clickable_text</a>"
        Dim ImageLink As String = "<a href=""address""><img src=""image_file"" alt=""ALT"" height=""heightPX"" width=""widthPX"" /></a>"
        Dim LinkWithClass As String = "class=""the_class"""

        'The Following Code checks the length of TextBox1 and TextBox2. If the two textboxes are NOT empty, then generate a text hyperlink.
        If TextBox1.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the address of the website."), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox2.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("a _", "the clickable text of the link."), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            'This generates the text hyperlink.
            GeneratedCode.RichTextBox1.Text = TextLink.Replace("address", TextBox1.Text).Replace("clickable_text", TextBox2.Text)
            If TextBox3.Enabled = True Then
                'This generates the text hyperlink with a class value.
                GeneratedCode.RichTextBox1.Text = TextLink.Replace("address", TextBox1.Text).Replace("clickable_text", TextBox2.Text).Replace("<a", "<a " & LinkWithClass).Replace("the_class", TextBox3.Text)
            End If
            GeneratedCode.Show()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Declare strings so we can easily use them later.
        Dim ErrorMessage As String = "You must include _"
        Dim TitleError As String = "Error: Missing Information"
        Dim ImageLink As String = "<a href=""address""><img src=""image_file"" alt=""alt_text"" height=""heightPX"" width=""widthPX"" /></a>"
        Dim LinkWithClass As String = "class=""the_class"""

        'Check the length of the textboxes
        If TextBox4.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("_", "the address of the website"), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox5.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("_", "the name of the image. This is vital!"), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox6.Text.Length = 0 Or TextBox7.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("_", "either the height or the width of the image"), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf TextBox7.Text.Length = 0 Then
            MessageBox.Show(ErrorMessage.Replace("_", "the ""alt"" tag. Recommended as the ""alt"" tag displays text if the image cannot be found"), TitleError, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            GeneratedCode.RichTextBox1.Text = ImageLink.Replace("address", TextBox4.Text).Replace("image_file", TextBox5.Text).Replace("alt_text", TextBox8.Text).Replace("heightPX", TextBox6.Text).Replace("widthPX", TextBox7.Text)
            If TextBox9.Enabled = True Then
                GeneratedCode.RichTextBox1.Text = ImageLink.Replace("<a", "<a " & LinkWithClass.Replace("the_class", TextBox9.Text)).Replace("address", TextBox4.Text).Replace("image_file", TextBox5.Text).Replace("alt_text", TextBox8.Text).Replace("heightPX", TextBox6.Text).Replace("widthPX", TextBox7.Text)
            End If
            GeneratedCode.Show()
        End If

    End Sub

    Private Sub HyperlinkGenerator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox3.Enabled = False
        TextBox9.Enabled = False
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox9.Enabled = True
        Else
            TextBox9.Enabled = False
        End If
    End Sub
End Class