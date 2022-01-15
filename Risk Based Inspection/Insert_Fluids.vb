Imports MySql.Data.MySqlClient

Public Class Insert_Fluids

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
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub simpandatfluida()
        Call koneksi()
        Dim simpan_command As New MySqlCommand("INSERT INTO `tbl_fluida`(`units`, `kode_fluida`, `nama_fluida`, `density`, `nbp`, `mw`, `ait`, `liquid_flow_velocity`, `viscosity`, `igc_a`, `igc_b`, `igc_c`, `igc_d`, `igc_e`, `fluid_type`, `ambient_state`, `liquid_dynamic_viscosity`) VALUES (@unit,@kode,@nama,@density,@nbp,@mw,@ait,@liquidflowvelocity,@viscosity,@a,@b,@c,@d,@e,@fluidtype,@ambientstate,@liquiddynamicviscosity)", constring)
        simpan_command.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        simpan_command.Parameters.Add("@kode", MySqlDbType.VarChar).Value = TextBox1.Text
        simpan_command.Parameters.Add("@nama", MySqlDbType.VarChar).Value = TextBox2.Text
        simpan_command.Parameters.Add("@density", MySqlDbType.Double).Value = TextBox3.Text
        simpan_command.Parameters.Add("@nbp", MySqlDbType.Double).Value = TextBox4.Text
        simpan_command.Parameters.Add("@mw", MySqlDbType.Double).Value = TextBox5.Text
        simpan_command.Parameters.Add("@ait", MySqlDbType.Double).Value = TextBox6.Text
        simpan_command.Parameters.Add("@liquidflowvelocity", MySqlDbType.Double).Value = TextBox7.Text
        simpan_command.Parameters.Add("@viscosity", MySqlDbType.Double).Value = TextBox8.Text
        simpan_command.Parameters.Add("@a", MySqlDbType.Double).Value = TextBox9.Text
        simpan_command.Parameters.Add("@b", MySqlDbType.Double).Value = TextBox10.Text
        simpan_command.Parameters.Add("@c", MySqlDbType.Double).Value = TextBox11.Text
        simpan_command.Parameters.Add("@d", MySqlDbType.Double).Value = TextBox12.Text
        simpan_command.Parameters.Add("@e", MySqlDbType.Double).Value = TextBox13.Text
        simpan_command.Parameters.Add("@fluidtype", MySqlDbType.VarChar).Value = ComboBox1.Text
        simpan_command.Parameters.Add("@ambientstate", MySqlDbType.VarChar).Value = ComboBox2.Text
        simpan_command.Parameters.Add("@liquiddynamicviscosity", MySqlDbType.Double).Value = TextBox14.Text

        If execCommand(simpan_command) Then
            MessageBox.Show("Data Saved")
        Else
            MessageBox.Show("Data Not Saved")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call simpandatfluida()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""

        Fluids.loaddatafluida()
        Fluids.aturDGV()

    End Sub

    Public Sub updatedatafluida()
        Call koneksi()

        Dim update_command As New MySqlCommand("UPDATE `tbl_fluida` SET `nama_fluida`=@nama,`density`=@density,`nbp`=@nbp,`mw`=@mw,`ait`=@ait,`liquid_flow_velocity`=@liquidflowvelocity,`viscosity`=@viscosity,`igc_a`=@a,`igc_b`=@b,`igc_c`=@c,`igc_d`=@d,`igc_e`=@e,`fluid_type`=@fluidtype,`ambient_state`=@ambient,`liquid_dynamic_viscosity`=@liquiddynamicviscosity WHERE `kode_fluida`=@kode AND `units`=@unit", constring)
        update_command.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        update_command.Parameters.Add("@kode", MySqlDbType.VarChar).Value = TextBox1.Text
        update_command.Parameters.Add("@nama", MySqlDbType.VarChar).Value = TextBox2.Text
        update_command.Parameters.Add("@density", MySqlDbType.Double).Value = TextBox3.Text
        update_command.Parameters.Add("@nbp", MySqlDbType.Double).Value = TextBox4.Text
        update_command.Parameters.Add("@mw", MySqlDbType.Double).Value = TextBox5.Text
        update_command.Parameters.Add("@ait", MySqlDbType.Double).Value = TextBox6.Text
        update_command.Parameters.Add("@liquidflowvelocity", MySqlDbType.Double).Value = TextBox7.Text
        update_command.Parameters.Add("@viscosity", MySqlDbType.Double).Value = TextBox8.Text
        update_command.Parameters.Add("@a", MySqlDbType.Double).Value = TextBox9.Text
        update_command.Parameters.Add("@b", MySqlDbType.Double).Value = TextBox10.Text
        update_command.Parameters.Add("@c", MySqlDbType.Double).Value = TextBox11.Text
        update_command.Parameters.Add("@d", MySqlDbType.Double).Value = TextBox12.Text
        update_command.Parameters.Add("@e", MySqlDbType.Double).Value = TextBox13.Text
        update_command.Parameters.Add("@fluidtype", MySqlDbType.VarChar).Value = ComboBox1.Text
        update_command.Parameters.Add("@ambient", MySqlDbType.VarChar).Value = ComboBox2.Text
        update_command.Parameters.Add("@liquiddynamicviscosity", MySqlDbType.Double).Value = TextBox14.Text

        If execCommand(update_command) Then
            MessageBox.Show("Data Updated")
        Else
            MessageBox.Show("Data Not Updated")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call updatedatafluida()
        Fluids.loaddatafluida()
        Fluids.aturDGV()

        Me.Close()
    End Sub

    Private Sub Insert_Fluids_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Units.ComboBox1.Text = "SI" Then
            Label15.Text = "(kg/m³)"
            Label16.Text = "(°C)"
            Label17.Text = "(kg/kg-mol)"
            Label18.Text = "(K)"
            Label19.Text = "(m/s)"
            Label20.Text = "(Pa.s)"
            Label21.Text = "(N-s/m²)"
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            Label15.Text = "(lb/ft³)"
            Label16.Text = "(°F)"
            Label17.Text = "(lb/lb-mol)"
            Label18.Text = "(°R)"
            Label19.Text = "(ft/s)"
            Label20.Text = "(Pa.s)"
            Label21.Text = "(N-s/m²)"
        Else
            Label15.Text = ""
            Label16.Text = ""
            Label17.Text = ""
            Label18.Text = ""
            Label19.Text = ""
            Label20.Text = ""
            Label21.Text = ""
        End If
    End Sub
End Class