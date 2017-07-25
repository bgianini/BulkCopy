Imports System.Data.SqlClient

Public Class Form1



    Dim connectionstring = "Server=Servidor;Database=Teste;User Id=USUARIO;Password=SENHA;"


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        BulkCopy()
    End Sub



    ' Aqui é de tabela pra tabela
    Private Sub BulkCopy()
        Dim origem As SqlConnection = New SqlConnection(connectionstring)
        Dim destino As SqlConnection = New SqlConnection(connectionstring)

        Dim cmd As SqlCommand = New SqlCommand("DELETE FROM tbUsuarioCopia", destino)

        Try

            origem.Open()
            destino.Open()

            cmd.ExecuteNonQuery()

            cmd = New SqlCommand("SELECT * FROM tbUsuario", origem)

            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim bulkData As SqlBulkCopy = New SqlBulkCopy(destino)

            bulkData.DestinationTableName = "tbUsuarioCopia"
            bulkData.WriteToServer(reader)

            bulkData.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            destino.Close()
            origem.Close()
            MessageBox.Show("feito")
        End Try
    End Sub



    'Aqui é do DataTable para o SQL SERVER
    Private Sub BulkCopyDataTable(ByVal dt As DataTable)
        If dt.Rows.Count > 0 Then
            Dim consString As String = connectionstring
            Using con As New SqlConnection(consString)
                Using sqlBulkCopy As New SqlBulkCopy(con)
                    'Set the database table name
                    sqlBulkCopy.DestinationTableName = "dbo.tbUsuarioCopia"

                    '[OPTIONAL]: Map the DataTable columns with that of the database table
                    sqlBulkCopy.ColumnMappings.Add("Nome", "Nome")
                    sqlBulkCopy.ColumnMappings.Add("Idade", "Idade")
                    sqlBulkCopy.ColumnMappings.Add("Email", "Email")
                    sqlBulkCopy.BulkCopyTimeout = 14400
                    con.Open()
                    sqlBulkCopy.WriteToServer(dt)
                    con.Close()
                End Using
            End Using
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dt As New DataTable
        dt.Columns.Add("Nome")
        dt.Columns.Add("Idade")
        dt.Columns.Add("Email")


        Dim maximo As Integer = 1000
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = maximo

        For index = 1 To maximo
            dt.Rows.Add("Nome", index.ToString, "email@email.com")
            ProgressBar1.Value = index
        Next


        BulkCopyDataTable(dt)
        MessageBox.Show("Fim do Bulk...")


    End Sub
End Class
