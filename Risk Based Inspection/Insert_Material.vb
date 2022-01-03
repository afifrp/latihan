Imports MySql.Data.MySqlClient

Public Class Insert_Material

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

    'Koneksi Database ------------------------------------------------------------------------------------

    Private Sub simpandatamaterial()
        Call koneksi()
        Dim simpan_command As New MySqlCommand("INSERT INTO `tbl_material`(`units`, `kode_material`, `nama_material`, `yield_strength`, `tensile_strength`, `design_pressure`, `max_operating_pressure`, `design_temperature`, `max_design_temperature`, `cost_factor`) VALUES (@unit,@kode,@nama,@ys,@ts,@designpressure,@maxoperpressure,@designtemp,@maxdesigntemp,@cost)", constring)
        simpan_command.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        simpan_command.Parameters.Add("@kode", MySqlDbType.VarChar).Value = TextBox1.Text
        simpan_command.Parameters.Add("@nama", MySqlDbType.VarChar).Value = TextBox2.Text
        simpan_command.Parameters.Add("@ys", MySqlDbType.Double).Value = TextBox3.Text
        simpan_command.Parameters.Add("@ts", MySqlDbType.Double).Value = TextBox4.Text
        simpan_command.Parameters.Add("@designpressure", MySqlDbType.Double).Value = TextBox5.Text
        simpan_command.Parameters.Add("@maxoperpressure", MySqlDbType.Double).Value = TextBox6.Text
        simpan_command.Parameters.Add("@designtemp", MySqlDbType.Double).Value = TextBox7.Text
        simpan_command.Parameters.Add("@maxdesigntemp", MySqlDbType.Double).Value = TextBox8.Text
        simpan_command.Parameters.Add("@cost", MySqlDbType.Double).Value = TextBox9.Text

        If execCommand(simpan_command) Then
            MessageBox.Show("Data Saved")
        Else
            MessageBox.Show("Data Not Saved")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Public Sub updatedatamaterial()
        Call koneksi()

        Dim update_command As New MySqlCommand("UPDATE `tbl_material` SET `nama_material`=@nama,`yield_strength`=@ys,`tensile_strength`=@ts,`design_pressure`=@designpressure,`max_operating_pressure`=@maxoperatingpressure,`design_temperature`=@designtemp,`max_design_temperature`=@maxdesigntemp,`cost_factor`=@cost WHERE `kode_material`=@kode AND `units`=@unit", constring)
        update_command.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        update_command.Parameters.Add("@kode", MySqlDbType.VarChar).Value = TextBox1.Text
        update_command.Parameters.Add("@nama", MySqlDbType.VarChar).Value = TextBox2.Text
        update_command.Parameters.Add("@ys", MySqlDbType.Double).Value = TextBox3.Text
        update_command.Parameters.Add("@ts", MySqlDbType.Double).Value = TextBox4.Text
        update_command.Parameters.Add("@designpressure", MySqlDbType.Double).Value = TextBox5.Text
        update_command.Parameters.Add("@maxoperatingpressure", MySqlDbType.Double).Value = TextBox6.Text
        update_command.Parameters.Add("@designtemp", MySqlDbType.Double).Value = TextBox7.Text
        update_command.Parameters.Add("@maxdesigntemp", MySqlDbType.Double).Value = TextBox8.Text
        update_command.Parameters.Add("@cost", MySqlDbType.Double).Value = TextBox9.Text

        If execCommand(update_command) Then
            MessageBox.Show("Data Updated")
        Else
            MessageBox.Show("Data Not Updated")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    'Load ------------------------------------------------------------------------------------------------


    'Button Coding ---------------------------------------------------------------------------------------

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
        Materials.loaddatamaterial()
        Materials.aturDGV()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call simpandatamaterial()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""

        Materials.loaddatamaterial()
        Materials.aturDGV()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call updatedatamaterial()

        Materials.loaddatamaterial()
        Materials.aturDGV()

        Me.Close()
    End Sub
End Class