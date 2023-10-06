Imports System.Data.SqlClient

Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim leg As String
        leg = TextBox1.Text
        Dim resultado As Boolean
        Dim transa As SqlTransaction = Nothing
        Dim mensaje As String = Nothing
        Dim resul1 As DataTable = Nothing
        If leg = "" Then
            MessageBox.Show("No ha ingresado un nuemro de legajo")
            Exit Sub
        End If
        'Dim libsql As New ClsLibreriaSql
        'Dim leg1 As Integer = 22

        libSql.AbrirConexion(resultado, mensaje)

        If resultado Then
            SQL = "dbo.AAAA_BORRAR " + leg
            libSql.Consulta(SQL, resul1, transa, resultado, mensaje)
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

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim Ape As String
        Dim Nom As String
        Dim Edad As Integer
        Dim Sex As String
        Dim resultado As Boolean
        Dim transa As SqlTransaction = Nothing
        Dim mensaje As String = Nothing
        Dim resul2 As DataTable = Nothing
        Dim devuelve As Integer

        Ape = TextBox2.Text
        Nom = TextBox3.Text
        Edad = TextBox4.Text
        Sex = TextBox5.Text

        'MessageBox.Show(Ape & ", " & ", " & Nom & ", " & Edad & ", " & " ," & Sex)
        libSql.AbrirConexion(resultado, mensaje)

        If resultado Then


            SQL = "INSERT INTO [Sistema].[dbo].[AAAA_Emplados_Sisitema]
           ([Apellido_sis]
           ,[Nombre_sis]
           ,[Edad_sis]
           ,[Sexo_sis])
            VALUES
           ('" & Ape & "', '" & Nom & "', '" & Edad & "', '" & Sex & "')"
            libSql.Ejecutar(SQL, True,devuelve,transa,resultado,mensaje)

            'libSql.Consulta(SQL, resul2, transa, resultado, mensaje)
        Else
            MessageBox.Show("NO se conecto")

        End If

        libSql.AbrirConexion(resultado, mensaje)

        SQL = "SELECT * FROM AAAA_Emplados_Sisitema WHERE id = '" & devuelve & "';"

        libSql.Consulta(SQL, resul2, transa, resultado, mensaje)

        'DataGridView2.Rows(0).Cells("id").Value = resul2.Rows(0).Item("id")
        DataGridView2.Rows(0).Cells("Apellido_sis").Value = resul2.Rows(0).Item("Apellido_sis")
        DataGridView2.Rows(0).Cells("Nombre_sis").Value = resul2.Rows(0).Item("Nombre_sis")
        DataGridView2.Rows(0).Cells("Edad_sis").Value = resul2.Rows(0).Item("Edad_sis")
        DataGridView2.Rows(0).Cells("Sexo_sis").Value = resul2.Rows(0).Item("Sexo_sis")


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub


End Class