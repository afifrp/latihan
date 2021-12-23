Public Class Equipment
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Insert_Equipment.Button7.Visible = True
        Insert_Equipment.Button8.Visible = False
        Insert_Equipment.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_Equipment.Button7.Visible = False
        Insert_Equipment.Button8.Visible = True
        Insert_Equipment.ShowDialog()
    End Sub
End Class