Imports System.Net
Imports System.IO
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.Office.Interop
Imports System.Drawing.Printing
Imports QRCoder

Module Utiles
    Public libSql As New ClsLibreriaSql

    Public PId As Long
    Public strDsn As String
    Public strUid As String
    Public strPsw As String
    Public LlevaForm As Integer
    Public rutaDestinoFotos As String = AppPath() + "\FotosArticulos"

    'Parámetros de búsqueda de personas

    Public GApellido As String
    Public GDomicilio As String
    Public GNroCalleD As String
    Public GNroCalleH As String
    Public GFechaIngD As String
    Public GFechaIngH As String


    Public vpc As String
    Public vservidor As String
    Public vip As String
    Public vbd As String
    Public RegistrosAfectados As Double
    Public ColorResalteTexto As Color = Color.Goldenrod
    Public ColorFondoFoco As Color = Color.WhiteSmoke
    Public TeclaBuscar As Keys = Keys.F1
    Public tablaFiltro As DataTable
    Public cambioDatos As Boolean
    Public DirectorioFotos As String = AppPath() + "\FOTOS"

    Public utilidad As Double
    Public idArticuloNuevo As Integer
    Public artPrecio As Double
    Public utilizarListaBlanca As Boolean
    Public identregaM As Integer


    'Obtener IP
    Public Function getIp() As String

        Dim valorIp As String

        valorIp = Dns.GetHostEntry(My.Computer.Name).AddressList.FirstOrDefault(Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork).ToString()

        Return valorIp

    End Function
    Public Function AppPath(
            Optional ByVal backSlash As Boolean = False
            ) As String
        ' System.Reflection.Assembly.GetExecutingAssembly...
        Dim s As String =
            IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetCallingAssembly.Location)
        ' si hay que añadirle el backslash
        If backSlash Then
            s &= "\"
        End If
        Return s
    End Function

    Public Function EstaAbierto(ByVal Myform As Form)
        Dim objForm As Form
        Dim blnAbierto As Boolean = False
        blnAbierto = False
        For Each objForm In My.Application.OpenForms
            If (Trim(objForm.Name) = Trim(Myform.Name)) Then
                blnAbierto = True
            End If
        Next
        Return blnAbierto
    End Function
    Public Function GET_MUEVE_STOCK(ByVal idArt As Long) As Boolean
        rsdatos = Consulta("SELECT     MueveStock
                                FROM         ArticulosCentroCosto
                                WHERE     (IdArticulo = " & idArt & ") AND (IdCentroCosto = " & CentroCosto & ")")
        If rsdatos.Rows.Count > 0 Then
            Return IIf(IsDBNull(rsdatos.Rows(0).Item("MueveStock")), False, rsdatos.Rows(0).Item("MueveStock"))
        Else
            Return False
        End If
    End Function

    Public Function Rellenar(ByVal Cadena As String, ByVal Caracter As String, ByVal Cantidad As Integer, ByVal Izquierda As Boolean) As String
        Dim Cuantos As Integer
        Dim LongCadena As Integer
        LongCadena = Len(Cadena)
        Cuantos = Cantidad - LongCadena
        If Cuantos > 0 Then
            If Izquierda = True Then
                Return Cadena.PadLeft(Cantidad, Caracter)
            Else
                Return Cadena.PadRight(Cantidad, Caracter)
            End If
        ElseIf Cuantos = 0 Then
            Return Cadena
        Else
            If Izquierda = True Then
                Return Left(Cadena, Cantidad)
            Else
                Return Right(Cadena, Cantidad)
            End If
        End If
    End Function

    'Public Function GeneraExcell(titulo As String, ByRef Grid As DataGridView, ByVal subTitulo As String) As Boolean
    '    'Nunca se imprimen las filas de imagenes
    '    Dim resultado As Boolean
    '    Try
    '        If Grid.Rows.Count > 0 Then

    '            Dim xcelApp As New Excel.Application
    '            Dim i As Integer
    '            Dim j As Integer
    '            Dim col As Integer


    '            xcelApp.Application.Workbooks.Add(Type.Missing)
    '            xcelApp.Cells(1, 1) = titulo
    '            xcelApp.Cells(2, 1) = subTitulo
    '            col = 1
    '            For i = 0 To Grid.Columns.Count - 1
    '                If Grid.Columns(i).Visible = True Then
    '                    'xcelApp.Cells(4, i + 1) = Grid.Columns(i).HeaderText
    '                    xcelApp.Cells(4, col) = Grid.Columns(i).HeaderText
    '                    col += 1
    '                End If
    '            Next

    '            i = 0
    '            While i < Grid.Rows.Count
    '                j = 0
    '                col = 1
    '                While j < Grid.Columns.Count
    '                    If IsNothing(Grid.Rows(i).Cells(j).Value) Or IsDBNull(IsNothing(Grid.Rows(i).Cells(j).Value)) Then
    '                    Else
    '                        If Grid.Rows(i).Cells(j).Visible = True Then
    '                            'xcelApp.Cells(i + 4, j + 1) = Grid.Rows(i).Cells(j).Value.ToString()
    '                            xcelApp.Cells(i + 5, col) = Grid.Rows(i).Cells(j).Value.ToString()
    '                            col += 1
    '                        End If
    '                    End If
    '                    j += 1
    '                End While
    '                i += 1
    '            End While
    '            xcelApp.Columns.AutoFit()
    '            xcelApp.Visible = True
    '            resultado = True
    '        End If
    '    Catch ex As Exception
    '        resultado = False
    '    End Try
    '    Return resultado
    'End Function

    Public Function GeneraCsv(ByVal Grilla As DataGridView, ByRef Resultado As Boolean, ByRef Mensaje As String) As Boolean
        'Dim resultado As Boolean
        Try
            Dim strExport As String
            Dim ruta As String
            strExport = ""
            For Each c As DataGridViewColumn In Grilla.Columns
                If c.Visible = True Then
                    strExport &= """" & c.HeaderText & ""","
                End If
            Next
            strExport = strExport.Substring(0, strExport.Length - 1)
            strExport &= Environment.NewLine
            For Each r As DataGridViewRow In Grilla.Rows
                For Each c As DataGridViewColumn In Grilla.Columns
                    If c.Visible = True Then
                        If Grilla.Item(c.Index, r.Index).Value.ToString <> "" Then
                            strExport &= """" & Grilla.Item(c.Index, r.Index).Value.ToString & ""","
                        Else
                            strExport &= """" & "" & ""","
                        End If
                    End If
                Next
                strExport = strExport.Substring(0, strExport.Length - 1)
                strExport &= Environment.NewLine
            Next
            Dim SaveFileDialog As SaveFileDialog = New SaveFileDialog
            SaveFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            SaveFileDialog.Filter = "Archivos CSV (*.CSV)| *.CSV"
            SaveFileDialog.FilterIndex = 2
            If SaveFileDialog.ShowDialog = DialogResult.OK Then
                ruta = SaveFileDialog.FileName
                'MsgBox("Exportado Correctamente", MsgBoxStyle.Information)
                Dim tw As IO.TextWriter = New StreamWriter(ruta)
                tw.Write(strExport)
                tw.Close()
                Resultado = True
            Else
                Resultado = False
                Mensaje = "El Usuario canceló la exportación"
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
            Mensaje = "Se produjo un error " & ex.Message
            Resultado = False
        End Try
        Return Resultado
    End Function

    Public Function GeneraTxt(Grilla As DataGridView, titulo As String, subtitulo As String, ByRef resultado As Boolean, ByRef Mensaje As String) As Boolean
        'Dim resultado As Boolean
        Try
            Dim strExport As String
            Dim ruta As String
            strExport = ""
            strExport &= titulo & Environment.NewLine
            strExport &= subtitulo & Environment.NewLine
            For Each c As DataGridViewColumn In Grilla.Columns
                If c.Visible = True Then
                    strExport &= c.HeaderText & ","
                End If
            Next
            strExport = strExport.Substring(0, strExport.Length - 1)
            strExport &= Environment.NewLine
            For Each r As DataGridViewRow In Grilla.Rows
                For Each c As DataGridViewColumn In Grilla.Columns
                    If c.Visible = True Then
                        strExport &= Grilla.Item(c.Index, r.Index).Value.ToString & ","
                    End If
                Next
                strExport = strExport.Substring(0, strExport.Length - 1)
                strExport &= Environment.NewLine
            Next
            Dim SaveFileDialog As SaveFileDialog = New SaveFileDialog
            SaveFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            SaveFileDialog.Filter = "Archivos TXT (*.TXT)| *.TXT"
            SaveFileDialog.FilterIndex = 2
            If SaveFileDialog.ShowDialog = DialogResult.OK Then
                ruta = SaveFileDialog.FileName
                'MsgBox("Exportado Correctamente", MsgBoxStyle.Information)
                Dim tw As IO.TextWriter = New StreamWriter(ruta)
                tw.Write(strExport)
                tw.Close()
                resultado = True
                Mensaje = "Los artículos se exportaron con éxito"
            Else
                resultado = False
                Mensaje = "El Usuario canceló la exportacion"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            resultado = False
        End Try
        Return resultado
    End Function

    'Public Function ExportToPDF(rpt As ReportDocument, NombreArchivo As String) As String
    '    Dim vFileName As String = String.Empty
    '    Dim diskOpts As New DiskFileDestinationOptions()

    '    Try
    '        rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
    '        rpt.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat

    '        'Este es la ruta donde se guardara tu archivo.
    '        vFileName = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & NombreArchivo
    '        If File.Exists(vFileName) Then
    '            File.Delete(vFileName)
    '        End If
    '        diskOpts.DiskFileName = vFileName
    '        rpt.ExportOptions.DestinationOptions = diskOpts
    '        rpt.Export()
    '    Catch ex As Exception
    '        vFileName = String.Empty
    '        Throw ex
    '    End Try
    '    Return vFileName
    'End Function
    Public Function Validacuit(ByVal CUIT As String) As Boolean
        Dim suma As Integer
        Dim valido As Boolean

        CUIT = CUIT.Replace("-", "")
        If IsNumeric(CUIT) Then
            If CUIT.Length <> 11 Then
                Return False
            Else
                suma = 0
                suma += CInt(CUIT.Substring(0, 1)) * 5
                suma += CInt(CUIT.Substring(1, 1)) * 4
                suma += CInt(CUIT.Substring(2, 1)) * 3
                suma += CInt(CUIT.Substring(3, 1)) * 2
                suma += CInt(CUIT.Substring(4, 1)) * 7
                suma += CInt(CUIT.Substring(5, 1)) * 6
                suma += CInt(CUIT.Substring(6, 1)) * 5
                suma += CInt(CUIT.Substring(7, 1)) * 4
                suma += CInt(CUIT.Substring(8, 1)) * 3
                suma += CInt(CUIT.Substring(9, 1)) * 2
                suma += CInt(CUIT.Substring(10, 1)) * 1
            End If

            If Math.Round(suma / 11, 0) = (suma / 11) Then
                valido = True
            Else
                valido = False
            End If
        Else
            valido = False
        End If
        Return (valido)
    End Function
    Public Sub AgregaArchivoAuditoriaDonde(Texto As String, Archivo As String)

        '------------------- Agrega valor al log --------------
        Dim ruta As String
        Dim FechaHoy As String
        Dim strTExto As String
        ruta = ""
        If (Trim(Archivo) <> "") Then
            ruta = AppPath() + "\Auditoria\" + Archivo
        Else
            ruta = AppPath() + "\Auditoria\ErrorSentencia.TXT"
        End If
        FechaHoy = FormatDateTime(Date.Now(), DateFormat.ShortDate)
        strTExto = FechaHoy & ";" & Texto
        If File.Exists(ruta) Then
            Dim escritor As StreamWriter
            escritor = File.AppendText(ruta)
            escritor.Write(strTExto)
            escritor.Flush()
            escritor.Close()
        Else
            Dim tw As IO.TextWriter = New StreamWriter(ruta)
            tw.Write(strTExto)
            tw.Close()
        End If
    End Sub

    Public Function Redondear(ByVal vdValor As Double, Optional viDecimales As Integer = 2) As Double
        Try
            Return CDbl(Int(vdValor * 10 ^ viDecimales + 0.501) / 10 ^ viDecimales)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return vdValor
        End Try
    End Function

    Public Function Calcular_Edad(Fecha_Nacimiento As Object) As Integer
        Dim Años As Object
        ' comprueba si el valor no es nulo
        If IsDBNull(Fecha_Nacimiento) = True Then
            Return 0
        End If

        Años = DateDiff("yyyy", Fecha_Nacimiento, Date.Now)

        If Date.Now < DateAndTime.DateSerial(Year(Date.Now), Month(Fecha_Nacimiento), DateAndTime.Day(Fecha_Nacimiento)) Then
            Años -= 1
        End If
        Return CInt(Años)
    End Function
    Public Function Get_Nombre_proveedor(ByVal idProv As Integer) As String
        Dim nProv As String
        Try
            rsDatos5 = Consulta("SELECT (CONVERT(varchar(10), Proveedores.IdProveedor) + ' - ' + PersonaJuridica.RazonSocial) as dscPRov
                                    FROM Proveedores INNER JOIN PersonaJuridica ON Proveedores.IdPersonaJuridica = PersonaJuridica.IdPersJuridica
                                WHERE (Proveedores.IdProveedor = " & idProv & ")")
            If rsDatos5.Rows.Count > 0 Then nProv = rsDatos5.Rows(0).Item("dscPRov") Else nProv = "Proveedor no encontrado"
        Catch ex As Exception
            MsgBox(ex.Message)
            nProv = "ERROR"
        End Try
        Return nProv
    End Function
    Public Function Get_mail_proveedor(ByVal idProv As Integer) As String
        Dim nProv As String
        Try
            rsDatos5 = Consulta("SELECT isnull(PersonaJuridica.email1,isnull(PersonaJuridica.email2,'')) as email FROM Proveedores 
                                INNER JOIN PersonaJuridica ON Proveedores.IdPersonaJuridica = PersonaJuridica.IdPersJuridica
                                WHERE (Proveedores.IdProveedor = " & idProv & ")")
            If rsDatos5.Rows.Count > 0 Then nProv = rsDatos5.Rows(0).Item("email") Else nProv = ""
        Catch ex As Exception
            MsgBox(ex.Message)
            nProv = ""
        End Try
        Return nProv
    End Function
    Public Function Get_NombreCentro(ByVal idCentro As Integer) As String
        Dim nCentro As String
        Try
            rsDatos5 = Consulta("SELECT DscCentroCosto FROM CentrosCostos where idCentroCosto= " & idCentro)
            nCentro = rsDatos5.Rows(0).Item("DscCentroCosto")
        Catch ex As Exception
            MsgBox(ex.Message)
            nCentro = "ERROR"
        End Try
        Return nCentro
    End Function
    Public Sub Get_Centro_ListaBlanca()
        Try
            rsDatos5 = Consulta("SELECT isnull((SELECT UsaListaBlanca FROM CentroCostoConfiguracion where idCtroCtoHijo= " & CentroCosto & "),0) as UsaListaBlanca")
            utilizarListaBlanca = IIf(IsDBNull(rsDatos5.Rows(0).Item("UsaListaBlanca")), False, rsDatos5.Rows(0).Item("UsaListaBlanca"))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error al consultar lista Blanca", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function SoloNum_0_9(val As Integer) As Boolean
        'Valida numeros y - (guion del medio)
        Return (val >= 48 And val <= 57)
    End Function

    Public Function SoloNum_0_9_punto(val As Integer) As Boolean
        'Valida numeros y - (guion del medio)
        Return (val >= 48 And val <= 57) Or (val = 46)
    End Function

    Public Function SoloNum_0_9_guion(val As Integer) As Boolean
        'Valida numeros y - (guion del medio)
        Return (val >= 48 And val <= 57) Or (val = 45)
    End Function
    Public Function AZaz_SPC(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 / Spacio
        Return (val >= 97 And val <= 122) Or (val >= 65 And val <= 90) Or (val = 32)
    End Function
    Public Function AZaz_SINSPC(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 
        Return (val >= 97 And val <= 122) Or (val >= 65 And val <= 90)
    End Function
    Public Function AZaz09_SPC(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 / Spacio
        Return (val >= 97 And val <= 122) Or (val >= 65 And val <= 90) Or (val >= 48 And val <= 57) Or (val = 32)
    End Function
    Public Function AZaz09_SINSPC(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 / Spacio
        Return (val >= 97 And val <= 122) Or (val >= 65 And val <= 90) Or (val >= 48 And val <= 57) Or (val = 8)
    End Function

    Public Function ASCII_Extendido_32_126(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 / Spacio
        Return (val >= 32 And val <= 126)
    End Function

    Public Function ASCII_Extendido_acentos_32_126(val As Integer) As Boolean
        'Valida A..Z / a..z / 0..9 / Spacio
        Return (val >= 32 And val <= 165) Or (val = 241 Or val = 209 Or val = 225 Or val = 233 Or val = 237 Or val = 243 Or val = 250)
    End Function

    Public Function Valida_DataCol_Long(parCol As DataRow, ColName As String) As Long
        'Valida un campo del DataRow y devuelve 0 - Long
        If IsDBNull(parCol(ColName)) Then
            Return 0
        Else
            Return CLng(parCol(ColName))
        End If
    End Function

    Public Function Valida_DataCol_Integer(parCol As DataRow, ColName As String) As Integer
        'Valida un campo del DataRow y devuelve 0 Integer
        If IsDBNull(parCol(ColName)) Then
            Return 0
        Else
            Return CInt(parCol(ColName))
        End If
    End Function

    Public Function Valida_DataCol_Date(parCol As DataRow, ColName As String) As DateTime
        'Valida un campo del DataRow y devuelve un Date
        If IsDBNull(parCol(ColName)) Then
            Return ""
        Else
            Return parCol(ColName)
        End If
    End Function

    Public Function Valida_DataCol_Str(parCol As DataRow, ColName As String) As String
        'Valida un campo del DataRow y devuelve String 
        If IsDBNull(parCol(ColName)) Then
            Return ""
        Else
            Return CStr(parCol(ColName))
        End If
    End Function

    Public Function Valida_DataCol_Double(parCol As DataRow, ColName As String) As Double
        'Valida un campo del DataRow y devuelve String 
        If IsDBNull(parCol(ColName)) Then
            Return 0.00
        Else
            Return CDbl(parCol(ColName))
        End If
    End Function
    Public Function Valida_EMAIL(Par_Email As String) As Boolean
        Try
            Dim mail = New System.Net.Mail.MailAddress(Par_Email)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    'Public Function GuardaLog(ByVal descripcion As String, ByVal operacion As String) As Boolean
    '    Dim resultado As Boolean
    '    Dim oLog As New ClsLogs
    '    Try
    '        resultado = oLog.ADD(descripcion, operacion)
    '        If resultado = False Then Throw New Exception("Error al guardar LOG")
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '        resultado = False
    '    End Try
    '    Return resultado
    'End Function

    'Public Function GuardaLog_CT(ByVal descripcion As String, ByVal operacion As String, ByVal transac As SqlTransaction) As Boolean
    '    Dim resultado As Boolean
    '    Dim oLog As New ClsLogs
    '    Try
    '        resultado = oLog.ADD_CT(descripcion, operacion, transac)
    '        If resultado = False Then Throw New Exception("Error al guardar LOG")
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '        resultado = False
    '    End Try
    '    Return resultado
    'End Function

    Public Function Obtener_nro_col(ByVal letra As String) As Integer
        Dim respuesta As Integer
        Try
            If Trim(letra) <> "" Then
                Dim letraViene As String = Trim(Strings.Left(letra, 1)).ToUpper
                respuesta = Asc(letraViene)
                respuesta -= 64
            Else
                respuesta = 0
            End If
        Catch ex As Exception
            respuesta = 1
        End Try
        Return respuesta
    End Function

    Public Sub CentrarForm(ByVal contenedor As Form, ByVal formHijo As Form)
        formHijo.Left = (contenedor.Width - formHijo.Width) / 2
        formHijo.Top = (contenedor.Height - formHijo.Height) / 2
    End Sub

    Public Sub Cambiar_color_fila_datagrid(ByVal dg As DataGridView, ByVal fila As Integer, ByVal color As Color)
        Try
            With dg
                .Rows(fila).DefaultCellStyle.BackColor = color
            End With
        Catch ex As Exception

        End Try
    End Sub
    'Public Function Enviarmail_Cta_pred(ByVal destinatario As String, ByVal cuerpo As String, ByVal asunto As String,
    '                         Optional ByVal pArchivo As String = "") As Boolean
    '    Dim oConf As New ClsConfigMail
    '    Dim resultado As Boolean
    '    Try
    '        If oConf.Get_Predet Then
    '            Dim correo As New System.Net.Mail.MailMessage()
    '            correo.From = New System.Net.Mail.MailAddress(oConf.Cuenta_Envio) 'remitente
    '            correo.Subject = asunto  'asunto
    '            correo.To.Add(destinatario)
    '            If pArchivo <> "" Then
    '                Dim archivo As New System.Net.Mail.Attachment(pArchivo) ''  ruta del achivo adjunto
    '                correo.Attachments.Add(archivo) ''adjuntar archivos
    '            End If
    '            correo.Body = cuerpo
    '            Dim Servidor As New System.Net.Mail.SmtpClient
    '            Servidor.Host = oConf.Dir_SMTP
    '            Servidor.Port = oConf.Puerto_Smtp
    '            Servidor.EnableSsl = oConf.Habilita_SSL
    '            Servidor.Credentials = New System.Net.NetworkCredential(oConf.Usuario, oConf.Password_CTA)
    '            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    '            Servidor.Send(correo)
    '            resultado = True
    '        Else
    '            MsgBox("No hay una cuenta de mail predeterminada para el envio de correos ", MsgBoxStyle.Information, "Enviar Correos")
    '            resultado = False
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '        resultado = False
    '    End Try
    '    Return resultado
    'End Function
    'Public Function Enviarmail_Cta_Selec(ByVal id_cta As Integer, ByVal destinatario As String, ByVal cuerpo As String, ByVal asunto As String,
    '                         Optional ByVal pArchivo As String = "") As Boolean
    '    Dim oConf As New ClsConfigMail
    '    Dim resultado As Boolean
    '    Try
    '        oConf.ID_Config = id_cta
    '        If oConf.Get_by_id Then
    '            Dim correo As New System.Net.Mail.MailMessage()
    '            correo.From = New System.Net.Mail.MailAddress(oConf.Cuenta_Envio) 'remitente
    '            correo.Subject = asunto  'asunto
    '            correo.To.Add(destinatario)
    '            If pArchivo <> "" Then
    '                Dim archivo As New System.Net.Mail.Attachment(pArchivo) ''  ruta del achivo adjunto
    '                correo.Attachments.Add(archivo) ''adjuntar archivos
    '            End If
    '            correo.Body = cuerpo
    '            Dim Servidor As New System.Net.Mail.SmtpClient
    '            Servidor.Host = oConf.Dir_SMTP
    '            Servidor.Port = oConf.Puerto_Smtp
    '            Servidor.EnableSsl = oConf.Habilita_SSL
    '            Servidor.Credentials = New System.Net.NetworkCredential(oConf.Usuario, oConf.Password_CTA)
    '            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    '            Servidor.Send(correo)
    '            resultado = True
    '        Else
    '            MsgBox("No hay una cuenta de mail predeterminada para el envio de correos ", MsgBoxStyle.Information, "Enviar Correos")
    '            resultado = False
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '        resultado = False
    '    End Try
    '    Return resultado
    'End Function
    Public Function Get_precio_ult_cpra(ByVal idArticulo As Integer, ByVal idProveedor As Integer) As Double
        Dim sql As String
        Dim respuesta As Double
        Try
            sql = "select ISNULL((SELECT TOP (1) FactCompraDetalle.PrecioUnitario FROM FactCompraDetalle INNER JOIN "
            sql += " FactCompraCabecera ON FactCompraDetalle.idFac = FactCompraCabecera.idFac "
            sql += " WHERE (FactCompraDetalle.IdProveedor = " & idProveedor & ") AND (FactCompraDetalle.Idarticulo = " & idArticulo & ") "
            sql += " AND (FactCompraCabecera.idCentroCosto = " & CentroCosto & ") ORDER BY FactCompraDetalle.idFac DESC),0) as precio"
            rsDatos4 = Consulta(sql)
            If rsDatos4.Rows.Count > 0 Then
                respuesta = rsDatos4.Rows(0).Item("precio")
            End If
        Catch ex As Exception
            Throw New Exception("Error al consultar precio artículo" & ex.Message)
            respuesta = 0
        End Try
        Return respuesta
    End Function
    Public Function Get_precio_ult_OC(ByVal idArticulo As Integer, ByVal idProveedor As Integer) As Double
        Dim respuesta As Double
        Dim sql As String
        Try
            sql = "select ISNULL((SELECT  top(1)  nuevoCosto FROM NotasPedidoCCDet "
            sql += " WHERE (idArticulo = " & idArticulo & ") AND (idCentroCosto = " & CentroCosto & ") AND (idProveedor = " & idProveedor & ")"
            sql += " ORDER BY nroNotaPedido DESC),0) as precio"
            rsDatos4 = Consulta(sql)
            If rsDatos4.Rows.Count > 0 Then
                respuesta = rsDatos4.Rows(0).Item("precio")
            End If
        Catch ex As Exception
            Throw New Exception("Error al consultar precio artículo" & ex.Message)
            respuesta = 0
        End Try
        Return respuesta
    End Function
    Public Function Get_last_NotaPedido(ByVal idArticulo As Integer, ByVal idProveedor As Integer) As DataTable
        Dim sql As String
        Try
            sql = "SELECT top (1)  NotasPedidoCCCab.nroNotaPedido, NotasPedidoCCCab.fechaHoraDia, NotasPedidoCCDet.nuevoCosto, "
            sql += " NotasPedidoCCCab.idEstadoDoc FROM NotasPedidoCCCab INNER JOIN "
            sql += " NotasPedidoCCDet ON NotasPedidoCCCab.idCentroCosto = NotasPedidoCCDet.idCentroCosto "
            sql += " And NotasPedidoCCCab.nroNotaPedido = NotasPedidoCCDet.nroNotaPedido And "
            sql += " NotasPedidoCCCab.idProveedor = NotasPedidoCCDet.idProveedor "
            sql += " WHERE     (NotasPedidoCCCab.idCentroCosto = " & CentroCosto & ") AND "
            sql += " (NotasPedidoCCCab.idProveedor = " & idProveedor & ") AND (NotasPedidoCCDet.idArticulo = " & idArticulo & ") "
            sql += " AND (NotasPedidoCCCab.idEstadoDoc = 1) order by fechaHoraDia desc"
            rsDatos4 = Consulta(sql)
        Catch ex As Exception
            Throw New Exception("Error al consultar precio artículo" & ex.Message)
        End Try
        Return rsDatos4
    End Function



    'LLena combos
    Public Sub CargarCombo(ByVal cb As ComboBox, strSql As String, display As String, value As String)
        Try

            rsdatos = Consulta(strSql)

            If rsdatos.Rows.Count > 0 Then

                With cb
                    .DataSource = rsdatos
                    .DisplayMember = display
                    .ValueMember = value
                    .Refresh()
                End With
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error al cargar combo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Sub AlternaColorGrilla(ByVal Grilla As DataGridView)
        Try
            With Grilla
                .RowsDefaultCellStyle.BackColor = Color.LightGray
                .AlternatingRowsDefaultCellStyle.BackColor = Color.White
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GET_ARTICULOS_X_CODIGOBARRA(ByVal CODBARRA As String) As DataTable
        GET_ARTICULOS_X_CODIGOBARRA = Consulta("SELECT     Articulos.IdArticulo, Articulos.Descripcion, Articulos.CodBarra,
                                                ArticulosCentroCosto.IdCentroCosto, ArticulosCentroCosto.PrecioUnitarioVta, 
                                                ArticulosCentroCosto.PrecioUnitarioCosto, ArticulosCentroCosto.cantidad, 
                                                ArticulosCentroCosto.PorcUtilidad, ArticulosCentroCosto.BajaLogica, 
                                                 ArticulosCentroCosto.MueveStock, ArticulosCentroCosto.idUnidadMedidaProduccion
                                                FROM         Articulos INNER JOIN
                                                ArticulosCentroCosto ON Articulos.IdArticulo = ArticulosCentroCosto.IdArticulo
                                                WHERE     (Articulos.CodBarra = '" & CODBARRA & "') AND (ArticulosCentroCosto.BajaLogica = 0) AND 
                                                (ArticulosCentroCosto.IdCentroCosto = " & CentroCosto & ")")
    End Function
    Public Function GET_UNIDAD_MEDIDA(ByVal IDUNIDAD As Integer) As String
        rsDatos1 = Consulta("SELECT     idUnidadMedida, dscUnidadMedida
                            FROM         UnidadesMedida where idUnidadMedida = " & IDUNIDAD)
        If rsDatos1.Rows.Count > 0 Then
            Return rsDatos1.Rows(0).Item("dscUnidadMedida")
        Else
            Return "No Especificada"
        End If
    End Function
    'Public Function GET_STOCK_ART(ByVal IdArt As Integer) As Double
    '    Dim respuesta As Double
    '    Try
    '        rsDatos1 = Consulta("SELECT ISNULL((SELECT Cantidad FROM ArticulosCentroCosto WHERE (IdArticulo = " & IdArt & ") AND (IdCentroCosto = " & CentroCosto & ")),0) as cantidad")
    '        respuesta = rsDatos1.Rows(0).Item("cantidad")
    '    Catch ex As Exception
    '        respuesta = 0
    '    End Try
    '    Return respuesta
    'End Function
    Public Function GET_STOCK_ART(ByVal IdArt As Integer, ByVal idDep As Integer) As Double
        Dim respuesta As Double
        Try
            rsDatos1 = Consulta("SELECT ISNULL((SELECT cantidad FROM Articulos_Stock WHERE (idArticulo = " & IdArt & ") AND (idDep = " & idDep & ")),0) as cantidad")
            respuesta = rsDatos1.Rows(0).Item("cantidad")
        Catch ex As Exception
            respuesta = 0
        End Try
        Return respuesta
    End Function


    Public Function Get_proveedor_tiene_art(idProv As Integer) As Boolean
        Dim respuesta As Boolean
        Try
            rsDatos2 = Consulta("SELECT ISNULL((SELECT count(IdArticulo) as cant FROM ArticulosCentroCostoProveedor where IdCentroCosto = " & CentroCosto & " AND  IdProveedor = " & idProv & "),0) as cantidad")
            respuesta = IIf(rsDatos2.Rows(0).Item("cantidad") > 0, True, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error al consultar artículos proveedor", MessageBoxButtons.OK, MessageBoxIcon.Error)
            respuesta = False
        End Try
        Return respuesta
    End Function

    Public Sub Marca_si_hay_texo(chkB As CheckBox, textB As TextBox)
        chkB.Checked = IIf(Len(textB.Text) > 0, True, False)
    End Sub

    Public Function Get_ccosto_deposito(idDep As Integer) As Integer
        Dim respuesta As Integer
        Dim sql As String
        Try
            sql = "select ISNULL ((SELECT idCentroCosto FROM Depositos_Mercaderia WHERE (estado = 1) AND (idDep = " & idDep & ")),0)  as ccosto"
            rsDatos4 = Consulta(sql)
            If rsDatos4.Rows.Count > 0 Then
                respuesta = rsDatos4.Rows(0).Item("ccosto")
            End If
        Catch ex As Exception
            respuesta = 0
        End Try
        Return respuesta
    End Function
    Public Function Cantidad_registros(tabla As String) As Integer
        Dim respuesta As Integer
        Try
            rsDatos5 = Consulta("Select * from " & tabla)
            respuesta = rsDatos5.Rows.Count
        Catch ex As Exception
            Throw New Exception("Error al obtener cuenta predeterminada " & ex.Message)
            respuesta = 1
        End Try
        Return respuesta
    End Function


    Public Function DecodeBase64ToString(valor As String) As String
        Dim myBase64ret As Byte() = Convert.FromBase64String(valor)
        Dim myStr As String = System.Text.Encoding.UTF8.GetString(myBase64ret)
        Return myStr
    End Function
    Public Function EncodeStrToBase64(valor As String) As String
        Dim myByte As Byte() = System.Text.Encoding.UTF8.GetBytes(valor)
        Dim myBase64 As String = Convert.ToBase64String(myByte)
        Return myBase64
    End Function

    Public Function ArmaStringQR(ByVal VER As String, ByVal fecha As String, ByVal CUIT As String, ByVal ptoVta As String,
                            ByVal TipoComp As String, ByVal Nro As String, ByVal IMPORTE As String, ByVal Moneda As String,
                            ByVal Cotiz As String, ByVal TipoDoc As String, ByVal NroDoc As String, ByVal TipoCodAut As String,
                            ByVal CAE As String) As String
        Dim strver As String
        Dim strFECHA As String
        Dim strCuit As String
        Dim strPtoVta As String
        Dim strTipoComp As String
        Dim strNro As String
        Dim strImporte As String
        Dim strMoneda As String
        Dim strCotiz As String
        Dim strTipoDoc As String
        Dim strNroDoc As String
        Dim strTipoCodAut As String
        Dim strCae As String
        Dim oCotiza As String
        Dim oImporte As String
        Dim oCuit As String
        Dim oNroDoc As String
        Dim respuesta As String
        Try
            oCotiza = Replace(Cotiz, ",", "")
            oCotiza = Replace(oCotiza, ".", "")
            oImporte = Replace(IMPORTE, ",", "")
            oImporte = Replace(oImporte, ".", "")
            oCuit = Replace(CUIT, "-", "")
            oNroDoc = Replace(NroDoc, "-", "")
            strver = """ver:""" & VER
            strFECHA = """fecha:""" & "" & Format(CDate(fecha), "yyyy-MM-dd") & ""
            strCuit = """cuit:""" & oCuit
            strPtoVta = """ptoVta:""" & ptoVta
            strTipoComp = """tipoCmp:""" & TipoComp
            strNro = """nroCmp:""" & Nro
            strImporte = """importe:""" & oImporte
            strMoneda = """moneda:""" & """" & Moneda & """"
            strCotiz = """ctz:""" & oCotiza
            strTipoDoc = """tipoDocRec:""" & TipoDoc
            strNroDoc = """nroDocRec:""" & oNroDoc
            strTipoCodAut = """tipoCodAut:""" & """" & TipoCodAut & """"
            strCae = """codAut:""" & CAE
            respuesta = "{" & strver & "," & strFECHA & "," & strCuit & "," & strPtoVta & "," & strTipoComp & "," & strNro & "," & strImporte & "," & strMoneda & "," & strCotiz & "," & strTipoDoc & "," & strNroDoc & "," & strTipoCodAut & "," & strCae & "}"
            respuesta = "https://www.afip.gob.ar/fe/qr/?p=" & EncodeStrToBase64(respuesta)
        Catch ex As Exception
            MsgBox(ex.Message)
            respuesta = "ERROR"
        End Try
        Return respuesta
    End Function


End Module
