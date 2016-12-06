Attribute VB_Name = "FCN_Huella"
Public Type rawImage
 img As Variant
 width As Long
 height As Long
 res As Long
End Type

Public Type TTemplate
 tpt() As Byte
 Size As Long
End Type

Public raw As rawImage
Public template(3) As TTemplate
Public Function dame_cadena_conexion(server As String, db As String, us As String, ps As String) As String
dame_cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver};" _
            & "SERVER=" & server & ";" _
            & "DATABASE=" & db & ";" _
            & "UID=" & us & ";" _
            & "PWD=" & ps & ";" _
            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
            'En SERVER el nombre o la IP Pública del servidor de datos
            'Si usamos la misma máquina sería localhost o 127.0.0.1
            
            'en DATABASE el nombre de la BASE DE DATOS
            
            'en UID el nombre del usuario
            
            'en PWD el password
            
            'hay que dejar las opciones porque suele funcionar mejor
            
End Function

'''''''Aqui para Buscar la HUELLA
    Public Function Identificar(Formulario As Form, Numero As Integer, Nombre As Label, Area As Label) As Integer
     'On Error Resume Next
     Dim ret As Integer
     Dim tpt() As Byte
     Dim Cuantos As Integer

     Identificar = 0

     ret = Formulario.grFinger.IdentifyPrepare(template(Numero).tpt, GR_DEFAULT_CONTEXT)
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    conn.CursorLocation = adUseClient
    conn.ConnectionString = dame_cadena_conexion(serv1, db1, us1, psw1)
    conn.Open
    
    rs.Open "Select * from per_tipo where pertp_huella1 is not null", conn, adOpenStatic, adLockReadOnly ' Set Resultado = BD.OpenRecordset("SELECT * FROM usuarios")
     rs.MoveLast 'Resultado.MoveLast
     Cuantos = rs.RecordCount 'Resultado.RecordCount
     rs.MoveFirst ' Resultado.MoveFirst
     For i = 1 To Cuantos
       ' Revisar si es la Huella1
        tpt = rs!pertp_huella1 'Resultado.Fields("pertp_huella1")
        ret = Formulario.grFinger.Identify(tpt, 0, GR_DEFAULT_CONTEXT)
        If ret = GR_MATCH Then
             Nombre = rs.Fields("PERTP_CODIGO_MEMBRESIA")
             Area = rs!PERTP_PER_ID 'Resultado.Fields("IdUsuario")
             Identificar = 1
             Exit Function
        End If
         ' Revisar si es la Huella2
            tpt = rs!pertp_huella2 'Resultado.Fields("pertp_huella2")
            ret = Formulario.grFinger.Identify(tpt, 0, GR_DEFAULT_CONTEXT)
            If ret = GR_MATCH Then
               Nombre = rs.Fields("PERTP_CODIGO_MEMBRESIA")
               Area = rs!PERTP_PER_ID 'Resultado.Fields("IdUsuario")
               Identificar = 1
               Exit Function
            Else
              rs.MoveNext 'Resultado.MoveNext
            End If
     Next i
     rs.Close
     Set rs = Nothing
     Nombre = ""
     Area = ""
    End Function

Public Function Inicializar(Formulario As Form) As Integer
 Err = Formulario.grFinger.Initialize
 If Err < 0 Then
  Inicializar = Err
  Exit Function
 End If
 Inicializar = Formulario.grFinger.CapInitialize
End Function

Public Sub CapturaHuella(ByVal biometricDisplay As Boolean, ByVal context As Integer, Formulario As Form, LaImagen As Image, Numero As Integer)
 Dim handle As IPictureDisp
 Dim ret As Integer

 If biometricDisplay Then
   Formulario.grFinger.biometricDisplay template(Numero).tpt, raw.img, raw.width, raw.height, raw.res, Formulario.hDC, handle, context
 Else
   Formulario.grFinger.CapRawImageToHandle raw.img, raw.width, raw.height, Formulario.hDC, handle
 End If

 If Not (handle Is Nothing) Then
   LaImagen.Picture = handle
 End If
End Sub

Public Function EncuentraPuntos(Formulario As Form, ControlMensajes As Label, LaImagen As Image, Numero As Integer) As Boolean
 Dim ret As Integer
 template(Numero).Size = GR_MAX_SIZE_TEMPLATE
 ReDim Preserve template(Numero).tpt(template(Numero).Size)
 ret = Formulario.grFinger.Extract(raw.img, raw.width, raw.height, raw.res, template(Numero).tpt, template(Numero).Size, GR_DEFAULT_CONTEXT)
 If ret < 0 Then template(Numero).Size = 0
 ReDim Preserve template(Numero).tpt(template(Numero).Size)
   
 If ret = GR_BAD_QUALITY Then
   ControlMensajes = "Huella detectada pero con baja calidad. Intentalo nuevamente"
   LaImagen.Picture = LoadPicture()
 ElseIf ret = GR_MEDIUM_QUALITY Then
   ControlMensajes = "Huella detectada con calidad mediana"
 ElseIf ret = GR_HIGH_QUALITY Then
   ControlMensajes = "Huella detectada con buena calidad"
 End If
 
 If ret >= 1 Then
   CapturaHuella True, GR_NO_CONTEXT, Formulario, LaImagen, Numero
   EncuentraPuntos = True
 Else
   EncuentraPuntos = False
 End If
End Function





