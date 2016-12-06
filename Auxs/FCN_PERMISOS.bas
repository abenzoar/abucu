Attribute VB_Name = "FCN_PERMISOS"
Dim SQL1 As String
Dim RES1 As Recordset
Public permAcceso As String
Public permAdd As String
Public permEdit As String
Public Sub checarPermisos(pantalla As String, tipoUsuario As Long)
    permisos = "NO"

    SQL1 = "SELECT * FROM VIEW_PERMISOS WHERE PANTALLA = '" & pantalla & "' AND CLAVE_TIPOUSUARIO = ' " & tipoUsuario & "'"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        permAcceso = RES1.Fields("ACCESO")
        permAdd = RES1.Fields("CREACION")
        permEdit = RES1.Fields("MODIFICAR")
    Else
        MsgBox "No se han asigando permisos para el tipo de cuenta actual. Verifique. ", vbInformation
    End If
    


End Sub
