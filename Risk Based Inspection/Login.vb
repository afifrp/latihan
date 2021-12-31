Public Class Login
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Button3.Visible = False
        Button2.Visible = False
        TextBox6.Visible = False
        TextBox5.Visible = False
        TextBox4.Visible = False
        TextBox3.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        LinkLabel2.Visible = False
        LinkLabel1.Visible = True
        Button1.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        Label3.Visible = True

        PictureBox2.Visible = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Button3.Visible = True
        Button2.Visible = True
        TextBox6.Visible = True
        TextBox5.Visible = True
        TextBox4.Visible = True
        TextBox3.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        LinkLabel2.Visible = True
        LinkLabel1.Visible = False
        Button1.Visible = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        Label3.Visible = False

        PictureBox2.Visible = False
    End Sub

End Class