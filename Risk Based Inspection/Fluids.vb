Imports MySql.Data.MySqlClient

Public Class Fluids

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


    Public Sub loaddatafluida()
        Call koneksi()
        Dim dt As New DataTable()

        Dim CMD As New MySqlCommand("SELECT `kode_fluida`, `nama_fluida`, `density`, `nbp`, `mw`, `ait`, `liquid_flow_velocity`, `viscosity`, `igc_a`, `igc_b`, `igc_c`, `igc_d`, `igc_e`, `fluid_type`, `ambient_state`, `liquid_dynamic_viscosity` FROM `tbl_fluida` WHERE `units`=@unit", constring)
        CMD.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text

        Dim adapter As New MySqlDataAdapter(CMD)

        adapter.Fill(dt)

        DataGridView1.DataSource = dt

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Sub aturDGV()
        Try
            DataGridView1.Columns(0).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(0).Width = 130
            DataGridView1.Columns(1).Width = 170
            DataGridView1.Columns(2).Width = 100
            DataGridView1.Columns(3).Width = 110
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(5).Width = 150
            DataGridView1.Columns(6).Width = 120
            DataGridView1.Columns(7).Width = 100
            DataGridView1.Columns(8).Width = 110
            DataGridView1.Columns(9).Width = 110
            DataGridView1.Columns(10).Width = 110
            DataGridView1.Columns(11).Width = 110
            DataGridView1.Columns(12).Width = 110
            DataGridView1.Columns(13).Width = 100
            DataGridView1.Columns(14).Width = 100
            DataGridView1.Columns(15).Width = 110

            DataGridView1.Columns(0).HeaderText = "Fluids Code"
            DataGridView1.Columns(1).HeaderText = "Fluids Name"
            DataGridView1.Columns(2).HeaderText = "Density (kg/m3)"
            DataGridView1.Columns(3).HeaderText = "Normal Boiling Point (⁰C)"
            DataGridView1.Columns(4).HeaderText = "Molecular Weight"
            DataGridView1.Columns(5).HeaderText = "Auto Ignition Temperature (⁰C)"
            DataGridView1.Columns(6).HeaderText = "Liquid Flow Velocity (m/s)"
            DataGridView1.Columns(7).HeaderText = "Viscosity of Liquid"
            DataGridView1.Columns(8).HeaderText = "Ideal Gas Constant [a]"
            DataGridView1.Columns(9).HeaderText = "Ideal Gas Constant [b]"
            DataGridView1.Columns(10).HeaderText = "Ideal Gas Constant [c]"
            DataGridView1.Columns(11).HeaderText = "Ideal Gas Constant [d]"
            DataGridView1.Columns(12).HeaderText = "Ideal Gas Constant [e]"
            DataGridView1.Columns(13).HeaderText = "Fluid Type"
            DataGridView1.Columns(14).HeaderText = "Ambient State"
            DataGridView1.Columns(15).HeaderText = "Liquid Dynamic Viscosity (N-s/m2)"

            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(11).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(13).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(14).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(10).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(11).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(12).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(13).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(14).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(15).SortMode = DataGridViewColumnSortMode.NotSortable

            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue

        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Insert_Fluids.TextBox1.Text = ""
        Insert_Fluids.TextBox2.Text = ""
        Insert_Fluids.TextBox3.Text = ""
        Insert_Fluids.TextBox4.Text = ""
        Insert_Fluids.TextBox5.Text = ""
        Insert_Fluids.TextBox6.Text = ""
        Insert_Fluids.TextBox7.Text = ""
        Insert_Fluids.TextBox8.Text = ""
        Insert_Fluids.TextBox9.Text = ""
        Insert_Fluids.TextBox10.Text = ""
        Insert_Fluids.TextBox11.Text = ""
        Insert_Fluids.TextBox12.Text = ""
        Insert_Fluids.TextBox13.Text = ""
        Insert_Fluids.TextBox14.Text = ""
        Insert_Fluids.ComboBox1.Text = ""
        Insert_Fluids.ComboBox2.Text = ""

        Insert_Fluids.Button1.Visible = True
        Insert_Fluids.Button2.Visible = False

        Insert_Fluids.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_Fluids.Button1.Visible = False
        Insert_Fluids.Button2.Visible = True

        Insert_Fluids.TextBox1.ReadOnly = True

        Insert_Fluids.ShowDialog()

    End Sub

    Private Sub Fluids_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call loaddatafluida()
        Call aturDGV()
        TextBox1.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call loaddatafluida()
        Call aturDGV()
        TextBox1.Text = ""
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        DataGridView1.Rows(e.RowIndex).HeaderCell.Value = CStr(e.RowIndex + 1)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim index As String
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)

        TextBox1.Text = selectedRow.Cells(0).Value.ToString()
        Insert_Fluids.TextBox1.Text = selectedRow.Cells(0).Value.ToString()
        Insert_Fluids.TextBox2.Text = selectedRow.Cells(1).Value.ToString()
        Insert_Fluids.TextBox3.Text = selectedRow.Cells(2).Value.ToString()
        Insert_Fluids.TextBox4.Text = selectedRow.Cells(3).Value.ToString()
        Insert_Fluids.TextBox5.Text = selectedRow.Cells(4).Value.ToString()
        Insert_Fluids.TextBox6.Text = selectedRow.Cells(5).Value.ToString()
        Insert_Fluids.TextBox7.Text = selectedRow.Cells(6).Value.ToString()
        Insert_Fluids.TextBox8.Text = selectedRow.Cells(7).Value.ToString()
        Insert_Fluids.TextBox9.Text = selectedRow.Cells(8).Value.ToString()
        Insert_Fluids.TextBox10.Text = selectedRow.Cells(9).Value.ToString()
        Insert_Fluids.TextBox11.Text = selectedRow.Cells(10).Value.ToString()
        Insert_Fluids.TextBox12.Text = selectedRow.Cells(11).Value.ToString()
        Insert_Fluids.TextBox13.Text = selectedRow.Cells(12).Value.ToString()
        Insert_Fluids.ComboBox1.Text = selectedRow.Cells(13).Value.ToString()
        Insert_Fluids.ComboBox2.Text = selectedRow.Cells(14).Value.ToString()
        Insert_Fluids.TextBox14.Text = selectedRow.Cells(15).Value.ToString()
    End Sub

    Public Sub deletedatafluida()
        Call koneksi()
        Dim delete_command As New MySqlCommand("DELETE FROM `tbl_fluida` WHERE `kode_fluida`=@kode AND `units`=@unit", constring)
        delete_command.Parameters.Add("@kode", MySqlDbType.VarChar).Value = TextBox1.Text
        delete_command.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text

        If execCommand(delete_command) Then
            MessageBox.Show("Data Deleted")
        Else
            MessageBox.Show("Data Not Deleted")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call deletedatafluida()
        Call loaddatafluida()
        Call aturDGV()
        TextBox1.Text = ""
    End Sub
End Class