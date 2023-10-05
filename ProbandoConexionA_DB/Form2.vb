Imports System.Data.SqlClient

Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim leg As Integer
        leg = TextBox1.Text
        Dim resultado As Boolean
        Dim transa As SqlTransaction = Nothing
        Dim mensaje As String = Nothing
        Dim resul1 As DataTable = Nothing

        'Dim libsql As New ClsLibreriaSql
        libSql.AbrirConexion(resultado, mensaje)
        If resultado Then
            libSql.Consulta("dbo.AAAA_BORRAR " & leg, resul1, transa, resultado, mensaje)
            If resultado Then
                If resul1.Rows.Count > 0 Then

                    DataGridView1.Rows(0).Cells("Apellido").Value = resul1.Rows(0).Item("Apellido")
                    DataGridView1.Rows(0).Cells("Nombre").Value = resul1.Rows(0).Item("Nombre")
                    DataGridView1.Rows(0).Cells("EDAD").Value = resul1.Rows(0).Item("EDAD")
                    DataGridView1.Rows(0).Cells("ANTIGUEDAD").Value = resul1.Rows(0).Item("ANTIGUEDAD")
                    DataGridView1.Rows(0).Cells("Estado").Value = resul1.Rows(0).Item("Estado")



                End If
                'DataGridView1.DataSource = resul1


            End If
        Else
            MessageBox.Show("NO se conecto")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class