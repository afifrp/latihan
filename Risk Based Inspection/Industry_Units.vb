Imports MySql.Data.MySqlClient
Public Class Industry_Units
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
        Insert_industry.Button1.Visible = True
        Insert_industry.Button2.Visible = False
        Insert_industry.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Insert_industry.Button1.Visible = False
        Insert_industry.Button2.Visible = True
        Insert_industry.ShowDialog()
    End Sub

    Public Sub loaddataindustry()
        Call koneksi()
        Dim dt As New DataTable()

        Dim CMD As New MySqlCommand("SELECT `industry_name` FROM `tbl_industry`", constring)

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

            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue

        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim index As String
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)

        TextBox1.Text = selectedRow.Cells(0).Value.ToString()
    End Sub

    Public Sub deletedataindustry()
        Call koneksi()
        Dim delete_command As New MySqlCommand("DELETE FROM `tbl_industry` WHERE `industry_name`=@name", constring)
        delete_command.Parameters.Add("@name", MySqlDbType.VarChar).Value = TextBox1.Text

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
        Call deletedataindustry()
        Call loaddataindustry()
        Call aturDGV()
    End Sub

    Private Sub Industry_Units_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call loaddataindustry()
        Call aturDGV()
    End Sub
End Class