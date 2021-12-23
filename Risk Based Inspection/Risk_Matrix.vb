Public Class Risk_Matrix


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Button1.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Risk_Matrix_Load(sender As Object, e As EventArgs) Handles Me.Load
        RadioButton1.Checked = True Or RadioButton2.Checked = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton1.Checked = True Then
            Me.Close()
        ElseIf RadioButton2.Checked = True Then
            Me.Close()
        End If
    End Sub
End Class