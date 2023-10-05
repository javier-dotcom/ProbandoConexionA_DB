

Imports System.Data.SqlClient
Module Conexion
    Public strDsn As String
    Public strUid As String
    Public strPsw As String
    Public CentroCosto As Integer
    Public CentroCostoCasaCtral As Integer
    Public CentroCostoAEC As Integer
    Public NombreCentro As String
    Public conn As New SqlConnection
    Public cmd As New SqlCommand
    Public prm As New SqlParameter
    Public conSecundaria As New SqlConnection

    'Parámetros de búsqueda de personas

    Public GApellido As String
    Public GDomicilio As String
    Public GNroCalleD As String
    Public GNroCalleH As String
    Public GFechaIngD As String
    Public GFechaIngH As String

    'Public Rta As Integer

    Public rsdatos As DataTable
    Public rsDatos1 As DataTable
    Public rsDatos2 As DataTable
    Public rsDatos3 As DataTable
    Public rsDatos4 As DataTable
    Public rsDatos5 As DataTable 'Agregado el 10/02/2020
    Public rol As Integer
    Public usuario As String
    Public legajoUsuario As Integer
    Public RegistrosAfectados As Double


    Public Ip As String
    Public base As String
    Public ConexionReporte As String
    Public UserDb As String
    Public PassDb As String
    Public IpSec As String
    Public baseSec As String
    Public UserDbSec As String
    Public PassDbSec As String



    Public CadCon As String
    Dim CadConSecundaria As String
    Public colorCon As Color
    Public backCon As Color
    Private pServidor As String
    Private pEntorno As String
    Private pCentroCosto As Integer
    Private pidDep As Integer

    Public SQL As String
    Public RsTmp As DataTable
    Public JERARQUIAUSUARIO As String
    Public DetalleEvento, TipoEvento As String

    Public Property Deposito() As Integer
        Get
            Return pidDep
        End Get
        Set(ByVal value As Integer)
            pidDep = value
        End Set
    End Property

    Public Property CCosto() As Integer
        Get
            Return pCentroCosto
        End Get
        Set(ByVal value As Integer)
            pCentroCosto = value
        End Set
    End Property
    Public Property Entorno() As String
        Get
            Return pEntorno
        End Get
        Set(ByVal value As String)
            pEntorno = value
        End Set
    End Property
    Public prueba1 As Integer

    'Validacion Anmat
    'Type Anmat_Valores
    '    Rta_Validado As Boolean
    '    Rta_NroTrans As String
    '    Rta_ErrorDesc  As String
    '    Rta_ErrorCod  As String
    '    val_CodBarra As String
    '    val_Serie As String
    '    val_FVto As String
    '    val_Lote As String
    'End Type





    Public Sub DefinirCC()

        CentroCosto = 540
        'CentroCostoCasaCtral = 760
        'CentroCostoAEC = 0
    End Sub

    Sub Inicio()
        '
        '    strDsn = "conexionDesa"
        '    strUid = "desa"
        '    strPsw = "desa"
    End Sub

    Public Sub RealizarConexion()



        ''--------------ALIAS SERVIDOR - BASE SISTEMAS ------------
        'Ip = "Server"
        'base = "Sistema"
        'CadCon = "server=" & Ip & ";database=" & base & ";uid=sa;pwd=asql123bd"
        'conn.ConnectionString = CadCon
        'colorCon = Color.Black
        'backCon = Color.Green
        'UserDb = "sa"
        'PassDb = "asql123bd"


        'conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=asql123bd;Initial Catalog=" + Trim(base) + ";Data Source=" + Ip
        'ConexionReporte = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=asql123bd;Initial Catalog=" + Trim(base) + ";Data Source=" + Ip
        '----------------------------------
        '
        '

        ''--------------serversqltst ------------
        Ip = "serversqltst"
        base = "Sistema"
        ' base = "ProveeAEC"
        UserDb = "sa"
        PassDb = "sql.test"
        colorCon = Color.White
        backCon = Color.Red
        CadCon = "server=" & Ip & ";database=" & base & ";uid=sa;pwd=sql.test"
        conn.ConnectionString = CadCon
        '----------------------------------

        '--------------ALIAS SERVIDOR - NUEVOS---------------------
        'Ip = "serversqltst"
        'base = "Sistema"
        'conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=sql.test;Initial Catalog=" + base + ";Data Source=" + Ip
        'ConexionReporte = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=sql.test;Initial Catalog=" + base + ";Data Source=" + Ip

        ''----------------------------------

        '-------------- ALIAS SERVIDOR 3 - BASE SISTEMAS ------------
        '    Ip = "Server3"
        '    base = "Sistema"
        '    conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=asql123bd;Initial Catalog=" + Trim(base) + ";Data Source=" + Ip
        '    ConexionReporte = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; pwd=asql123bd;Initial Catalog=" + Trim(base) + ";Data Source=" + Ip
        '----------------------------------



        Try
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error de Conexion")
        End Try

        'conn.Curso = adUseClient


    End Sub

    Public Sub CerrarConexion()

        ' Close the database.
        conn.Close()

    End Sub

    Public Sub CerrarConexion_Secundaria()

        ' Close the database.
        conSecundaria.Close()

    End Sub



    Public Function Get_parametros_iniciales() As Boolean
        Dim respuesta As Boolean = True

        Try
            pCentroCosto = -1
            pEntorno = "D"
            pServidor = "SERVERSQLTST"
            pidDep = -1
        Catch ex As Exception
            Throw New Exception("Error al Obtener parámetros desde INI " & ex.Message)
            respuesta = False
        End Try
        Return respuesta
    End Function



    Public Function Consulta(ByVal STRSQL As String) As DataTable
        Dim da As New SqlDataAdapter(STRSQL, CadCon)
        Dim ds As New DataSet
        Try
            da.Fill(ds, "tabla")
        Catch ex As Exception
            MsgBox(ex.Message)
            da = Nothing
        End Try
        ' Defino variables

        ' Cargo la tabla virtual

        ' Retorno el dataset cargado
        Return ds.Tables("tabla")
    End Function


    Public Function ExecSQL(ByVal sSql As String) As Boolean
        Dim command As SqlCommand = conn.CreateCommand()
        Dim band As Boolean = False
        command.Connection = conn
        command.CommandText = sSql
        Try
            command.ExecuteNonQuery()
            band = True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        ExecSQL = band
    End Function

    Public Function ExecSQL_CON_TRANSAC(ByVal sSql As String, ByVal TRANSACCION As SqlTransaction) As Boolean
        Dim command As SqlCommand = conn.CreateCommand()
        Dim band As Boolean = False
        command.Connection = conn
        command.Transaction = TRANSACCION
        command.CommandText = sSql
        Try
            command.ExecuteNonQuery()
            band = True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        ExecSQL_CON_TRANSAC = band
    End Function

    Public Function Consulta_CON_TRANSAC(ByVal STRSQL As String, ByVal TRANSACCION As SqlTransaction) As DataTable
        'VERRRRR
        Dim ds As New DataSet()

        Try

            Dim cmd As SqlCommand = conn.CreateCommand()
            Dim band As Boolean = False
            cmd.Connection = conn
            cmd.Transaction = TRANSACCION
            cmd.CommandText = STRSQL
            cmd.ExecuteNonQuery()
            'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
            If cmd.CommandText.Contains(" ") Then
                cmd.CommandType = CommandType.Text
            Else
                cmd.CommandType = CommandType.StoredProcedure
            End If

            Dim adapter As New SqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 0
            adapter.Fill(ds, "tabla")


        Catch ex As Exception
            ' The connection failed. Display an error message.
            Throw New Exception("Database Error: " & ex.Message)
        End Try

        Return ds.Tables("tabla")

    End Function
    Public Function ExecConsulta_SP() As DataTable
        Dim ds As New DataSet()

        Try
            Dim connection As New SqlConnection(conn.ConnectionString)
            cmd.Connection = connection
            'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
            If cmd.CommandText.Contains(" ") Then
                cmd.CommandType = CommandType.Text
            Else
                cmd.CommandType = CommandType.StoredProcedure
            End If

            Dim adapter As New SqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 0
            adapter.Fill(ds, "tabla")
            connection.Close()

        Catch ex As Exception
            ' The connection failed. Display an error message.
            Throw New Exception("Database Error: " & ex.Message)
        End Try

        Return ds.Tables("tabla")
    End Function

    Public Function Consulta_Secundaria(ByVal STRSQL As String) As DataTable
        Dim da As New SqlDataAdapter(STRSQL, CadConSecundaria)
        Dim ds As New DataSet
        Try
            da.Fill(ds, "tabla")
        Catch ex As Exception
            MsgBox(ex.Message)
            da = Nothing
        End Try
        ' Defino variables

        ' Cargo la tabla virtual

        ' Retorno el dataset cargado
        Return ds.Tables("tabla")
    End Function


    Public Function ExecSQL_Secundaria(ByVal sSql As String) As Boolean
        Dim command As SqlCommand = conSecundaria.CreateCommand()
        Dim band As Boolean = False
        command.Connection = conSecundaria
        command.CommandText = sSql
        Try
            command.ExecuteNonQuery()
            band = True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return band
    End Function

    Public Function ExecSQL_CON_TRANSAC_secundaria(ByVal sSql As String, ByVal TRANSACCION As SqlTransaction) As Boolean
        Dim command As SqlCommand = conSecundaria.CreateCommand()
        Dim band As Boolean = False
        command.Connection = conSecundaria
        command.Transaction = TRANSACCION
        command.CommandText = sSql
        Try
            command.ExecuteNonQuery()
            band = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return band
    End Function

    Public Function Consulta_CON_TRANSAC_Secundaria(ByVal STRSQL As String, ByVal TRANSACCION As SqlTransaction) As DataTable
        'VERRRRR
        Dim ds As New DataSet()

        Try

            Dim cmd As SqlCommand = conSecundaria.CreateCommand()
            Dim band As Boolean = False
            cmd.Connection = conSecundaria
            cmd.Transaction = TRANSACCION
            cmd.CommandText = STRSQL
            cmd.ExecuteNonQuery()
            'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
            If cmd.CommandText.Contains(" ") Then
                cmd.CommandType = CommandType.Text
            Else
                cmd.CommandType = CommandType.StoredProcedure
            End If

            Dim adapter As New SqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 0
            adapter.Fill(ds, "tabla")


        Catch ex As Exception
            ' The connection failed. Display an error message.
            Throw New Exception("Database Error: " & ex.Message)
        End Try

        Return ds.Tables("tabla")

    End Function
    Public Function ExecConsulta_SP_Secundaria() As DataTable
        Dim ds As New DataSet()

        Try
            Dim connection As New SqlConnection(conSecundaria.ConnectionString)
            cmd.Connection = connection
            'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
            If cmd.CommandText.Contains(" ") Then
                cmd.CommandType = CommandType.Text
            Else
                cmd.CommandType = CommandType.StoredProcedure
            End If

            Dim adapter As New SqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 0
            adapter.Fill(ds, "tabla")
            connection.Close()

        Catch ex As Exception
            ' The connection failed. Display an error message.
            Throw New Exception("Database Error: " & ex.Message)
        End Try

        Return ds.Tables("tabla")
    End Function




End Module
