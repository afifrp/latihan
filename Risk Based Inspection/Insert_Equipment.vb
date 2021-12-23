Imports MySql.Data.MySqlClient

Public Class Insert_Equipment

    Function execCommand(ByVal cmd As MySqlCommand) As Boolean

        If constring.State = ConnectionState.Closed Then
            constring.Open()
        End If

        Try
            If cmd.ExecuteNonQuery() = 1 Then
                Return True

            Else
                Return False
            End If
        Catch ex As Exception

            MessageBox.Show("ERROR")
            Return False

        End Try

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Function

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PictureBox1.Visible = True
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""

        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("Pressure Vessel")
        ComboBox6.Items.Add("Compressors")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PictureBox1.Visible = False
        PictureBox2.Visible = True
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""

        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("Heat Exchanger")
        ComboBox6.Items.Add("AirFin Heat Exchanger Header Boxes")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = True
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""

        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("Pipes & Tubes")
        ComboBox6.Items.Add("Pumps")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = True
        PictureBox5.Visible = False
        PictureBox6.Visible = False

        ComboBox2.Visible = False
        ComboBox3.Visible = False
        ComboBox4.Visible = False
        ComboBox5.Visible = True
        TextBox2.Visible = False
        TextBox2.Text = ""

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""

        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("Atmospheric Storage Tank - Shell Courses")
        ComboBox6.Items.Add("Atmospheric Storage Tank - Bottom Plates")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = True
        PictureBox6.Visible = False

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""

        ComboBox6.Items.Clear()
        ComboBox6.Items.Add("Pressure Relief Device")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = True

        ComboBox6.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        TextBox2.Text = ""
        ComboBox6.Items.Clear()

        ComboBox6.Items.Add("Heat Exchanger Tube Bundles")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "Pressure Vessel" Then
            ComboBox2.Visible = True
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "Compressors" Then
            ComboBox2.Visible = True
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "Heat Exchanger" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "AirFin Heat Exchanger Header Boxes" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = True
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "Pipes & Tubes" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = True
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "Pumps" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = True
            ComboBox5.Visible = False
            TextBox2.Visible = False
            TextBox2.Text = ""
        ElseIf ComboBox6.Text = "Pressure Relief Device" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = True
            TextBox2.Text = "Pressure Relief Device"
        ElseIf ComboBox6.Text = "Heat Exchanger Tube Bundles" Then
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            ComboBox5.Visible = False
            TextBox2.Visible = True
            TextBox2.Text = "Heat Exchanger Tube Bundles"
        End If
    End Sub
End Class