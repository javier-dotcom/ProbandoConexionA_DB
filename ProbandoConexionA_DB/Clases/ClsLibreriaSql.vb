Imports System.Data.SqlClient
Imports System.Net

Public Class ClsLibreriaSql
    Public CadCon As String = Conexion.CadCon
    Private NombrePc As String
    Private IpLocal As String
    Private Server As String
    Private Base As String
    Private UserBase As String
    Private PassBase As String
    'Public ReadOnly conn As New SqlConnection
    Public ReadOnly connLog As New SqlConnection

    Private Sub DefinirServidor()
        If ENTORNO = "Produccion" Then
            'Me.Server = "server"
            'Me.Base = "Sistema"
            'Me.UserBase = "sa"
            'Me.PassBase = "asql123bd"
        Else
            Me.Server = "serversqltst"
            Me.Base = "Sistema"
            Me.UserBase = "sa"
            Me.PassBase = "sql.test"
            'Me.Server = "localhost"
            'Me.Base = "Sistema"
            'Me.UserBase = "sa"
            'Me.PassBase = "sasasa"

        End If
        Me.NombrePc = Environment.MachineName
        Me.IpLocal = Dns.GetHostEntry(My.Computer.Name).AddressList.FirstOrDefault(Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork).ToString()

        Try
            conn.Close()
        Catch ex As Exception

        End Try

        'Me.CadCon = "server=" & Me.IpLocal & ";database=" & Me.Base & ";uid=" & Me.UserBase & ";pwd=" & Me.PassBase
        Me.CadCon = "server=" & Me.Server & ";database=" & Me.Base & ";uid=" & Me.UserBase & ";pwd=" & Me.PassBase
        conn.ConnectionString = Me.CadCon
        'ConexionReporte = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=sql.test;Initial Catalog=" + base + ";Data Source=" + Ip
        connLog.ConnectionString = Me.CadCon
    End Sub

    Public Sub GetDatosConexion(ByRef Optional Server As String = "", ByRef Optional Base As String = "", ByRef Optional UserBase As String = "",
                                ByRef Optional PassBase As String = "", ByRef Optional CadCon As String = "", ByRef Optional IpLocal As String = "",
                                ByRef Optional NombrePc As String = "")
        Server = Me.Server
        Base = Me.Base
        UserBase = Me.UserBase
        PassBase = Me.PassBase
        IpLocal = Me.IpLocal
        NombrePc = Me.NombrePc
        CadCon = Me.CadCon

    End Sub

    Public Sub AbrirConexion(ByRef Resultado As Boolean, ByRef Mensaje As String)
        Try
            DefinirServidor()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
                Resultado = True
                Mensaje = "OK"
            End If
        Catch ex As Exception
            Resultado = False
            Mensaje = "Hubo un error : " & ex.Message
        End Try

    End Sub

    Public Sub CerrarConexion()
        conn.Close()

    End Sub

    Public Sub LiberarObjeto()
        Me.Finalize()
    End Sub

    Public Sub Consulta(strsql As String, ByRef ts As DataTable, Transa As SqlTransaction,
                        ByRef resultado As Boolean, ByRef mensaje As String)

        If IsNothing(Transa) Then
            Try
                Dim da As New SqlDataAdapter(strsql, CadCon)
                Dim ds As New DataSet
                da.Fill(ds, "tabla")
                ts = ds.Tables("tabla")

                resultado = True
                mensaje = "OK"
            Catch ex As Exception
                resultado = False
                mensaje = "Hubo un error : " & ex.Message
            End Try
        Else
            Dim ds As New DataSet()
            Try
                Dim comando As SqlCommand = conn.CreateCommand()
                Dim band As Boolean = False
                comando.Connection = conn
                comando.Transaction = Transa
                comando.CommandText = strsql
                comando.ExecuteNonQuery()
                If comando.CommandText.Contains(" ") Then
                    comando.CommandType = CommandType.Text
                Else
                    comando.CommandType = CommandType.StoredProcedure
                End If
                Dim sdadapter As New SqlDataAdapter(comando)
                sdadapter.SelectCommand.CommandTimeout = 0
                sdadapter.Fill(ds, "tabla")
                ts = ds.Tables("tabla")
            Catch ex As Exception
                resultado = False
                mensaje = "Hubo un error : " & ex.Message
            End Try
        End If


    End Sub
    'Public Function ConsultaTr(ByVal STRSQL As String, ByVal TRANSACCION As SqlTransaction) As DataTable

    '    Dim ds As New DataSet()

    '    Try

    '        Dim cmd As SqlCommand = conn.CreateCommand()
    '        Dim band As Boolean = False
    '        cmd.Connection = conn
    '        cmd.Transaction = TRANSACCION
    '        cmd.CommandText = strsql
    '        cmd.ExecuteNonQuery()
    '        'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
    '        If cmd.CommandText.Contains(" ") Then
    '            cmd.CommandType = CommandType.Text
    '        Else
    '            cmd.CommandType = CommandType.StoredProcedure
    '        End If

    '        Dim adapter As New SqlDataAdapter(cmd)
    '        adapter.SelectCommand.CommandTimeout = 0
    '        adapter.Fill(ds, "tabla")


    '    Catch ex As Exception
    '        ' The connection failed. Display an error message.
    '        Throw New Exception("Database Error: " & ex.Message)
    '    End Try

    '    Return ds.Tables("tabla")

    'End Function


    Public Sub Ejecutar(strsql As String, DevuelveId As Boolean, ByRef ValorId As Int32,
                                      Transa As SqlTransaction, ByRef Resultado As Boolean, ByRef Mensaje As String)
        If Resultado Then
            Try
                Dim comando = conn.CreateCommand()
                If DevuelveId Then
                    strsql += "; SELECT SCOPE_IDENTITY()"
                    comando.CommandText = strsql
                    If Not IsNothing(Transa) Then comando.Transaction = Transa
                    ValorId = Convert.ToInt32(comando.ExecuteScalar())
                    Resultado = True
                    Mensaje = "OK"
                Else
                    comando.CommandText = strsql
                    If Not IsNothing(Transa) Then comando.Transaction = Transa
                    comando.ExecuteNonQuery()
                    Resultado = True
                    Mensaje = "OK"
                    ValorId = 0
                End If
            Catch ex As Exception
                Resultado = False
                Mensaje = "Hubo un error " + ex.Message
                'Throw New System.Exception("Se produjo un error")
            End Try
        End If

    End Sub

    Public Sub IniciarTransaccion(ByRef Transa As SqlTransaction, ByRef Resultado As Boolean, ByRef Mensaje As String)
        'AbrirConexion(Resultado, Mensaje)
        'If Resultado Then
        Try
            Transa = conn.BeginTransaction
            Resultado = True
            Mensaje = "OK"
        Catch ex As Exception
            Resultado = False
            Mensaje = "Hubo un error : " & ex.Message
        End Try
        'End If
    End Sub

    Public Sub ConfirmaTransaccion(ByRef Transa As SqlTransaction, ByRef Resultado As Boolean, ByRef Mensaje As String)
        Try
            Transa.Commit()
            Resultado = True
            Mensaje = "OK"
        Catch ex As Exception
            Resultado = False
            Mensaje = "Hubo un error : " & ex.Message
        End Try
    End Sub

    Public Sub DeshaceTransacion(ByRef Transa As SqlTransaction, ByRef ResultadoRB As Boolean, ByRef MensajeRB As String)
        Try
            Transa.Rollback()
            ResultadoRB = True
            MensajeRB = "OK"
        Catch ex As Exception
            ResultadoRB = False
            MensajeRB = "Hubo un error : " & ex.Message
        End Try
    End Sub


End Class
