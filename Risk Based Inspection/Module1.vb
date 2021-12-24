Imports MySql.Data.MySqlClient
Module Module1
    Public constring As MySqlConnection
    Public cm As MySqlCommand
    Public da As MySql.Data.MySqlClient.MySqlDataAdapter
    Public dt As New DataTable
    Public comBuilderDB As New MySql.Data.MySqlClient.MySqlCommandBuilder

    Sub koneksi()
        Try
            constring = New MySqlConnection("server=localhost;user id=root;database=db_rbi")
            If constring.State = ConnectionState.Closed Then
                constring.Open()
            End If

        Catch ex As Exception
            MsgBox("Connection Failed")
        End Try

    End Sub


End Module
