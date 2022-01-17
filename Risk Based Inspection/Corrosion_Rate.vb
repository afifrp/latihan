Public Class Corrosion_Rate

    Private Sub checkedfalse()
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
        Panel15.Visible = False
        Panel16.Visible = False
        Panel17.Visible = False
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
        CheckBox8.Enabled = True
        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            Panel5.Visible = True
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = True
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = True
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = True
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = True
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = True
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = True
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = True
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = True
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = True
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = True
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = True
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = True
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = True
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = True
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = True
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = True
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = True
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = True
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = True
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = True
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = True
            CheckBox2.Enabled = False
            CheckBox3.Enabled = True
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = True
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = True
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub panggillocalorgeneral()
        If Home.RadioButton1.Checked = True Then
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = False
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = False
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = True
            CheckBox12.Enabled = True
            CheckBox13.Enabled = True
        ElseIf Home.RadioButton2.Checked = True Then
            CheckBox1.Enabled = False
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = False
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = True
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
            MsgBox("Select thining type local/general !")
        End If
    End Sub

    Private Sub Corrosion_Rate_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call panggillocalorgeneral()
    End Sub
End Class