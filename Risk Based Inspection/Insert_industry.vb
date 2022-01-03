Imports MySql.Data.MySqlClient

Public Class Insert_industry

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

    'Koneksi Database --------------------------------------------------------------------------------------------
    Private Sub simpandatindustry()
        Call koneksi()
        Dim simpan_command As New MySqlCommand("INSERT INTO `tbl_industry`(`industry_code`, `industry_name`) VALUES (@code,@name)", constring)
        simpan_command.Parameters.Add("@code", MySqlDbType.VarChar).Value = TextBox1.Text
        simpan_command.Parameters.Add("@name", MySqlDbType.VarChar).Value = TextBox2.Text

        If execCommand(simpan_command) Then
            MessageBox.Show("Data Saved")
        Else
            MessageBox.Show("Data Not Saved")
        End If

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    Public Sub tampilcodeindustry()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT * FROM `tbl_industry` WHERE `industry_code`IN (SELECT MAX(`industry_code`) FROM `tbl_industry`)", constring)
        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        Dim urutan As String
        Dim hitung As Long

        rd = CMD.ExecuteReader

        rd.Read()

        If Not rd.HasRows Then
            urutan = "P" + "001"
        Else
            hitung = Microsoft.VisualBasic.Right(rd.GetString(0), 3) + 1
            urutan = "P" + Microsoft.VisualBasic.Right("000" & hitung, 3)
        End If

        TextBox1.Text = urutan

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Sub

    'Load --------------------------------------------------------------------------------------------------------
    Private Sub Insert_industry_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call tampilcodeindustry()
    End Sub

    'Button Coding -----------------------------------------------------------------------------------------------

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
        Call Industry_Units.loaddataindustry()
        Call Industry_Units.aturDGV()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox2.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call simpandatindustry()

        TextBox1.Text = ""
        TextBox2.Text = ""

        Call tampilcodeindustry()
        Call Industry_Units.loaddataindustry()
        Call Industry_Units.aturDGV()
    End Sub
End Class