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

    'Koneksi Database ------------------------------------------------------------------------------------------

    Public Sub tampildataindustrycombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `industry_name` FROM `tbl_industry`", constring)

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox1.Items.Clear()
        Do While rd.Read
            ComboBox1.Items.Add(rd.Item(0))
        Loop
    End Sub

    Private Sub simpandataequipment()
        Call koneksi()
        Dim simpan_command As New MySqlCommand("INSERT INTO `tbl_equipment`(`industry_name`, `equipment_code`, `equipment_description`, `equipment_type`, `component_type`) VALUES (@name,@code,@desc,@type,@comp)", constring)
        simpan_command.Parameters.Add("@name", MySqlDbType.VarChar).Value = ComboBox1.Text
        simpan_command.Parameters.Add("@code", MySqlDbType.VarChar).Value = TextBox1.Text
        simpan_command.Parameters.Add("@desc", MySqlDbType.VarChar).Value = TextBox2.Text
        simpan_command.Parameters.Add("@type", MySqlDbType.VarChar).Value = ComboBox6.Text
        simpan_command.Parameters.Add("@comp", MySqlDbType.VarChar).Value = ComboBox2.Text

        If execCommand(simpan_command) Then
            MessageBox.Show("Data Saved")
        Else
            MessageBox.Show("Data Not Saved")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Public Sub updatedataequipment()
        Call koneksi()

        Dim update_command As New MySqlCommand("UPDATE `tbl_equipment` SET `industry_name`=@name,`equipment_description`=@desc,`equipment_type`=@type,`component_type`=@comp WHERE `equipment_code`=@code", constring)
        update_command.Parameters.Add("@name", MySqlDbType.VarChar).Value = ComboBox1.Text
        update_command.Parameters.Add("@code", MySqlDbType.VarChar).Value = TextBox1.Text
        update_command.Parameters.Add("@desc", MySqlDbType.VarChar).Value = TextBox2.Text
        update_command.Parameters.Add("@type", MySqlDbType.VarChar).Value = ComboBox6.Text
        update_command.Parameters.Add("@comp", MySqlDbType.VarChar).Value = ComboBox2.Text

        If execCommand(update_command) Then
            MessageBox.Show("Data Updated")
        Else
            MessageBox.Show("Data Not Updated")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    'Load ------------------------------------------------------------------------------------------------------
    Private Sub Insert_Equipment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call tampildataindustrycombobox()
    End Sub

    'Button Coding ---------------------------------------------------------------------------------------------

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ComboBox2.Items.Clear()

        If ComboBox6.Text = "Pressure Vessel" Then
            ComboBox2.Items.Add("KODRUM")
            ComboBox2.Items.Add("DRUM")
            ComboBox2.Items.Add("REACTOR")
            ComboBox2.Items.Add("COLTOP")
            ComboBox2.Items.Add("COLMID")
            ComboBox2.Items.Add("COLBTM")
        ElseIf ComboBox6.Text = "Compressors" Then
            ComboBox2.Items.Add("COMPC")
            ComboBox2.Items.Add("COMPR")
        ElseIf ComboBox6.Text = "Heat Exchanger" Then
            ComboBox2.Items.Add("HEXSS")
            ComboBox2.Items.Add("HEXTS")
            ComboBox2.Items.Add("HEXTUBE")
        ElseIf ComboBox6.Text = "AirFin Heat Exchanger Header Boxes" Then
            ComboBox2.Items.Add("FINFAN")
            ComboBox2.Items.Add("FILTER")
        ElseIf ComboBox6.Text = "Pipes & Tubes" Then
            ComboBox2.Items.Add("PIPE-1")
            ComboBox2.Items.Add("PIPE-3")
            ComboBox2.Items.Add("PIPE-4")
            ComboBox2.Items.Add("PIPE-6")
            ComboBox2.Items.Add("PIPE-8")
            ComboBox2.Items.Add("PIPE-10")
            ComboBox2.Items.Add("PIPE-13")
            ComboBox2.Items.Add("PIPE-16")
            ComboBox2.Items.Add("PIPEGT16")
        ElseIf ComboBox6.Text = "Pumps" Then
            ComboBox2.Items.Add("PUMP3S")
            ComboBox2.Items.Add("PUMP1S")
            ComboBox2.Items.Add("PUMPR")
        ElseIf ComboBox6.Text = "Atmospheric Storage Tank - Shell Courses" Then
            ComboBox2.Items.Add("COURSES-1-10")
        ElseIf ComboBox6.Text = "Atmospheric Storage Tank - Bottom Plates" Then
            ComboBox2.Items.Add("TANKBOTTOM")
        ElseIf ComboBox6.Text = "Pressure Relief Devices" Then
            ComboBox2.Items.Add("Pressure Relief Devices")
        ElseIf ComboBox6.Text = "Heat Exchanger Tube Bundles" Then
            ComboBox2.Items.Add("Heat Exchanger Tube Bundles")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox6.Text = ""
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call simpandataequipment()

        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox6.Text = ""

        Call Equipment.loaddataequipment()
        Call Equipment.aturDGV()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call updatedataequipment()

        Call Equipment.loaddataequipment()
        Call Equipment.aturDGV()

        Me.Close()
    End Sub
End Class