Imports MySql.Data.MySqlClient
Public Class Materials

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

    Public Sub loaddatamaterial()
        Call koneksi()
        Dim dt As New DataTable()

        Dim CMD As New MySqlCommand("SELECT `kode_material`, `nama_material`, `yield_strength`, `tensile_strength`, `design_pressure`, `max_operating_pressure`, `design_temperature`, `max_design_temperature`, `cost_factor` FROM `tbl_material` WHERE `units`=@unit", constring)
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
            DataGridView1.Columns(4).Width = 110
            DataGridView1.Columns(5).Width = 150
            DataGridView1.Columns(6).Width = 120
            DataGridView1.Columns(7).Width = 160
            DataGridView1.Columns(8).Width = 100

            DataGridView1.Columns(0).HeaderText = "Material Code"
            DataGridView1.Columns(1).HeaderText = "Material Name"
            DataGridView1.Columns(2).HeaderText = "Yield Strength (MPa)"
            DataGridView1.Columns(3).HeaderText = "Tensile Strength (MPa)"
            DataGridView1.Columns(4).HeaderText = "Design Pressure (MPa)"
            DataGridView1.Columns(5).HeaderText = "Max. Operating Pressure (kpa)"
            DataGridView1.Columns(6).HeaderText = "Design Temperature (⁰C)"
            DataGridView1.Columns(7).HeaderText = "Max. Operating Temperature (⁰C)"
            DataGridView1.Columns(8).HeaderText = "Cost Factor ($)"

            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable

            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue

        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Insert_Material.TextBox1.Text = ""
        Insert_Material.TextBox2.Text = ""
        Insert_Material.TextBox3.Text = ""
        Insert_Material.TextBox4.Text = ""
        Insert_Material.TextBox5.Text = ""
        Insert_Material.TextBox6.Text = ""
        Insert_Material.TextBox7.Text = ""
        Insert_Material.TextBox8.Text = ""
        Insert_Material.TextBox9.Text = ""

        Insert_Material.Button1.Visible = True
        Insert_Material.Button2.Visible = False

        Insert_Material.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_Material.Button1.Visible = False
        Insert_Material.Button2.Visible = True

        Insert_Material.TextBox1.ReadOnly = True

        Insert_Material.ShowDialog()
    End Sub

    Private Sub Materials_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call loaddatamaterial()
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
        Insert_Material.TextBox1.Text = selectedRow.Cells(0).Value.ToString()
        Insert_Material.TextBox2.Text = selectedRow.Cells(1).Value.ToString()
        Insert_Material.TextBox3.Text = selectedRow.Cells(2).Value.ToString()
        Insert_Material.TextBox4.Text = selectedRow.Cells(3).Value.ToString()
        Insert_Material.TextBox5.Text = selectedRow.Cells(4).Value.ToString()
        Insert_Material.TextBox6.Text = selectedRow.Cells(5).Value.ToString()
        Insert_Material.TextBox7.Text = selectedRow.Cells(6).Value.ToString()
        Insert_Material.TextBox8.Text = selectedRow.Cells(7).Value.ToString()
        Insert_Material.TextBox9.Text = selectedRow.Cells(8).Value.ToString()
    End Sub


    Public Sub deletedatamaterial()
        Call koneksi()
        Dim delete_command As New MySqlCommand("DELETE FROM `tbl_material` WHERE `kode_material`=@kode AND `units`=@unit", constring)
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
        Call deletedatamaterial()
        Call loaddatamaterial()
        Call aturDGV()
        TextBox1.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call loaddatamaterial()
        Call aturDGV()
        TextBox1.Text = ""
    End Sub
End Class