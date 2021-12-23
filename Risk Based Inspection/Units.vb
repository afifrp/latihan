Public Class Units
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub SI()
        RadioButton1.Checked = True
        RadioButton6.Checked = True
        RadioButton9.Checked = True
        RadioButton12.Checked = True
        RadioButton10.Checked = True
        RadioButton14.Checked = True
        RadioButton16.Checked = True
        RadioButton18.Checked = True
        RadioButton21.Checked = True
        RadioButton30.Checked = True

    End Sub

    Private Sub US()
        RadioButton2.Checked = True
        RadioButton4.Checked = True
        RadioButton8.Checked = True
        RadioButton11.Checked = True
        RadioButton7.Checked = True
        RadioButton13.Checked = True
        RadioButton15.Checked = True
        RadioButton17.Checked = True
        RadioButton23.Checked = True
        RadioButton28.Checked = True
    End Sub
    Private Sub Units_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ComboBox1.Text = "SI" Then
            Call SI()
        ElseIf ComboBox1.Text = "US Customary" Then
            Call US()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "SI" Then
            Call SI()
        ElseIf ComboBox1.Text = "US Customary" Then
            Call US()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "SI" Then
            Me.Close()
        ElseIf ComboBox1.Text = "US Customary" Then
            Me.Close()
        End If
    End Sub
End Class