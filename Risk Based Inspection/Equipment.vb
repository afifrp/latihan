Imports MySql.Data.MySqlClient
Public Class Equipment

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

    'Koneksi database --------------------------------------------------------------------------------------
    Public Sub loaddataequipment()
        Call koneksi()
        Dim dt As New DataTable()

        Dim CMD As New MySqlCommand("SELECT * FROM `tbl_equipment`", constring)

        Dim adapter As New MySqlDataAdapter(CMD)

        adapter.Fill(dt)

        DataGridView1.DataSource = dt

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Sub aturDGV()
        Try

            DataGridView1.Columns(0).HeaderText = "Industry Name"
            DataGridView1.Columns(1).HeaderText = "Equipment Code"
            DataGridView1.Columns(2).HeaderText = "Equipment Description"
            DataGridView1.Columns(3).HeaderText = "Equipment Type"
            DataGridView1.Columns(4).HeaderText = "Component Type"

            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable

            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue

        Catch ex As Exception
        End Try
    End Sub

    'Button coding ----------------------------------------------------------------------------------------
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Insert_Equipment.Button7.Visible = True
        Insert_Equipment.Button8.Visible = False
        Insert_Equipment.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_Equipment.Button7.Visible = False
        Insert_Equipment.Button8.Visible = True
        Insert_Equipment.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call deletedataequipment()
        Call loaddataequipment()
        Call aturDGV()
    End Sub

    Private Sub Equipment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call loaddataequipment()
        Call aturDGV()
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        DataGridView1.Rows(e.RowIndex).HeaderCell.Value = CStr(e.RowIndex + 1)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim index As String
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)

        TextBox1.Text = selectedRow.Cells(1).Value.ToString()
        Insert_Equipment.ComboBox1.Text = selectedRow.Cells(0).Value.ToString()
        Insert_Equipment.TextBox1.Text = selectedRow.Cells(1).Value.ToString()
        Insert_Equipment.TextBox2.Text = selectedRow.Cells(2).Value.ToString()
        Insert_Equipment.ComboBox6.Text = selectedRow.Cells(3).Value.ToString()
        Insert_Equipment.ComboBox2.Text = selectedRow.Cells(4).Value.ToString()
    End Sub

    Public Sub deletedataequipment()
        Call koneksi()
        Dim delete_command As New MySqlCommand("DELETE FROM `tbl_equipment` WHERE `equipment_code`=@code", constring)
        delete_command.Parameters.Add("@code", MySqlDbType.VarChar).Value = TextBox1.Text


        If execCommand(delete_command) Then
            MessageBox.Show("Data Deleted")
        Else
            MessageBox.Show("Data Not Deleted")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub


End Class