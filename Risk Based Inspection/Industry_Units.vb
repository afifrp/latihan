Public Class Industry_Units
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Insert_industry.Button1.Visible = True
        Insert_industry.Button2.Visible = False
        Insert_industry.ShowDialog()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_industry.Button1.Visible = False
        Insert_industry.Button2.Visible = True
        Insert_industry.ShowDialog()

    End Sub
End Class