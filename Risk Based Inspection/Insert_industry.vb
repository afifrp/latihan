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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox2.Text = ""
    End Sub
End Class