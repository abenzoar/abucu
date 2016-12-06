Attribute VB_Name = "FCN_DbConn"
Public rs As ADODB.Recordset
Public rsSuc As ADODB.Recordset
Public con As Connection
Public conTest As Connection
Public conSuc As Connection
Public serv1 As String
Public us1 As String
Public psw1 As String
Public db1 As String
Public prt1 As String
Public Sub buscarConexiones(tipo As String)
    'On Error Resume Next
    
    Dim DirFile As String
    Dim a As String
    Dim actual As String
    Dim num1, num2 As Long
    Dim carga As Boolean
    Dim dirDir As String
    
    Dim numDb As Long
    Dim numDbN As Long
    
    actual = ""
    txtFileDir

    dirDir = Dir(direccionSistema & "\Com\Dat.dat", vbArchive)
    If dirDir = "" Then
        MsgBox "No se encuentra un archivo de carga inicial. No podrá iniciar el sistema. Por favor verifique.", vbInformation, "Información SIGEEST"
        End
    Else
        If tipo = "actual" Then
            FRM_Acceso.cmbSucur.Clear
        End If
        Open direccionSistema & "\Com\Dat.dat" For Input As #1
            numDb = 1
            numDbN = 1
            num1 = 0
            num2 = 0
            carga = False
            Do While Not EOF(1)
                Line Input #1, a

                ''''Para la conexión actual
                ''''Para la conexión actual
                If tipo = "actual" Then
                    If num1 = 1 Then
                        actual = a
                        num1 = 0
                    End If
                    If a = "<" & tipo & ">" Then
                        num1 = 1
                    End If
                Else
                    actual = tipo
                    num1 = 0
                End If
                If actual <> "" Then
                    If num2 = 5 Then
                        prt1 = a
                        num2 = num2 + 1
                        num2 = 0
                        carga = True
                        Call ConexionDB(serv1, db1, us1, psw1, prt1)
                        'Exit Sub
                    End If
                    If num2 = 4 Then
                        psw1 = a
                        num2 = num2 + 1
'                        num2 = 0
'                        carga = True
'                        Call ConexionDB(serv1, db1, us1, psw1, prt1)
                        'Exit Sub
                    End If
                    If num2 = 3 Then
                        us1 = a
                        num2 = num2 + 1
                    End If
                    If num2 = 2 Then
                        serv1 = a
                        num2 = num2 + 1
                    End If
                    If num2 = 1 Then
                        db1 = a
                        dbActual = a
                        num2 = num2 + 1
                    End If
                    If a = "<" & actual & ">" Then
                        num2 = num2 + 1
                    End If
                End If
                ''''Para la conexión actual
                ''''Para la conexión actual
                ''''Para cargar en el combo las conexiones
                ''''Para cargar en el combo las conexiones
                
                If tipo = "actual" Then
                    If numDbN = 2 Then
                        FRM_Acceso.cmbSucur.AddItem a
                        FRM_Acceso.cmbSucur.ItemData(FRM_Acceso.cmbSucur.ListCount - 1) = numDb - 1
                        numDbN = 1
                    End If
                    If a = "<" & numDb & ">" Then
                        numDbN = 2
                        numDb = numDb + 1
                    End If
                End If
                ''''Para cargar en el combo las conexiones
                ''''Para cargar en el combo las conexiones
            Loop
            

'            FRM_Acceso.cmbSucur.ListIndex = 0
            If carga = False Then
                MsgBox "No se puede realizar la conexión con los parámetros proporcionados. Verifique.", vbInformation
                End
            End If
            
        Close #1
    End If
        
    If FRM_Acceso.cmbSucur.ListCount > 0 Then
        'FRM_Acceso.cmbSucur.ListIndex = 0
        FRM_Acceso.cmbSucur.Text = dbActual
        loadDb = False
    End If
End Sub
Public Sub ConexionDB(server As String, db As String, us As String, ps As String, prt As String)
    Dim a As String
    On Error Resume Next
                       
    Set conTest = New ADODB.Connection
    conTest.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & server & ";uid=" & us & "; Port=" & prt & "; pwd=9807288;database=test;OPTION=16427"
    conTest.CursorLocation = adUseClient
    conTest.Open
    
    If Err.Number = (-2147467259) Then
        a = MsgBox("No se ha realizado la conexion con el servidor. Posibles causas: " & vbCrLf & vbCrLf & _
        "- Conexión de red dañada o incorrecta " & vbCrLf & _
        "- Nombre incorrecto de la base de datos. " & vbCrLf & _
        "- Nombre del Servidor o dirección IP incorrecto" & vbCrLf & vbCrLf & _
        "Descripción: " & Err.Description & vbCrLf & vbCrLf & _
        "Verifica datos siguientes Servidor: " & server & " Us: " & us & " DataBase: " & db & " Puerto: " & prt & vbCrLf & vbCrLf & _
        "¿Salir?", vbYesNo + vbQuestion, "Información SIGEEST")
    
        Open App.Path & "\Com\Dat\LogErr.txt" For Append As #1
        Print #1, Date & "  " & Time & " Ventana: " & "Frm_Acceso Modulo: Conectar" & " Motivo: " & Err.Description
        Close #1
            
        If a = vbYes Then
            End
        Else
            End
            'Exit Sub
        End If
    End If
    
    Dim sql1 As String
    Dim RES1 As Recordset
    sql1 = "SELECT MD5('9807288') PASS"
    Set RES1 = conTest.Execute(sql1)

    If Not RES1.EOF Then
        If ps <> RES1.Fields("pass") Then
            MsgBox "Las credenciales no coinciden. Verifique con el administrador.", vbCritical
            End
        Else
            ps = "9807288"
        End If
    End If
    
    Set con = New ADODB.Connection
    con.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & server & ";uid=" & us & "; Port=" & prt & "; pwd=" & ps & ";database=" & db & ";OPTION=16427"
    con.CursorLocation = adUseClient
    con.Open
    
    'MsgBox Err.Number
    If Err.Number = (-2147467259) Then
        a = MsgBox("No se ha realizado la conexion con el servidor. Posibles causas: " & vbCrLf & vbCrLf & _
        "- Conexión de red dañada o incorrecta " & vbCrLf & _
        "- Nombre incorrecto de la base de datos. " & vbCrLf & _
        "- Nombre del Servidor o dirección IP incorrecto" & vbCrLf & vbCrLf & _
        "Descripción: " & Err.Description & vbCrLf & vbCrLf & _
        "Verifica datos siguientes " & vbCrLf & vbCrLf & _
        "Servidor: " & server & " Us: " & us & vbCrLf & _
        " DataBase: " & db & " Puerto: " & prt & vbCrLf & vbCrLf & _
        "¿Salir?", vbYesNo + vbQuestion, "Información SIGEEST")
    
        Open App.Path & "\Com\Dat\LogErr.txt" For Append As #1
        Print #1, Date & "  " & Time & " Ventana: " & "Frm_Acceso Modulo: Conectar" & " Motivo: " & Err.Description
        Close #1
            
        If a = vbYes Then
            End
        Else
            End
            'Exit Sub
        End If
    End If

'    FRM_Acceso.cmbSucur.AddItem db
'    FRM_Acceso.cmbSucur.ListIndex = 0
End Sub

Public Sub ConexionDB_Suc(server As String, db As String, us As String, ps As String, prt As String)
    Dim a As String
    On Error Resume Next
    
    Set conSuc = New ADODB.Connection
    conSuc.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & server & ";uid=" & us & "; Port=" & prt & "; pwd=" & ps & ";database=" & db & ";OPTION=16427"
    conSuc.CursorLocation = adUseClient
    conSuc.Open
    
    'MsgBox Err.Number
    If Err.Number = (-2147467259) Then
        a = MsgBox("No se ha realizado la conexion con el servidor. Posibles causas: " & vbCrLf & vbCrLf & _
        "- Conexión de red dañada o incorrecta " & vbCrLf & _
        "- Nombre incorrecto de la base de datos. " & vbCrLf & _
        "- Nombre del Servidor o dirección IP incorrecto" & vbCrLf & vbCrLf & _
        "Descripción: " & Err.Description & vbCrLf & vbCrLf & _
        "Verifica el Servidor: " & server & " Us: " & us & " DB " & db & vbCrLf & vbCrLf & _
        "¿Salir?", vbYesNo + vbQuestion, "Información SIGEEST")
    
        Open App.Path & "\Com\Dat\LogErr.txt" For Append As #1
        Print #1, Date & "  " & Time & " Ventana: " & "Frm_Acceso Modulo: Conectar" & " Motivo: " & Err.Description
        Close #1
            
        If a = vbYes Then
            End
        Else
            End
            'Exit Sub
        End If
    End If

'    FRM_Acceso.cmbSucur.AddItem db
'    FRM_Acceso.cmbSucur.ListIndex = 0
End Sub

Sub Registro()
'Set rs = New ADODB.Recordset
'    With rs
'        .ActiveConnection = con
'        .CursorLocation = adUseClient
'        .CursorType = adOpenDynamic
'        .LockType = adLockOptimistic
'        '.Open "select * from Usuarios"
'    End With
End Sub


Public Sub txtFileDir()

    Dim DirFile As String
    Dim a As String
    Dim dirDir As String

    dirDir = Dir(App.Path & "\DatDirInstall.txt", vbArchive)
    If dirDir = "" Then
        MsgBox "No se encuentra un archivo de carga inicial. No podrá iniciar el sistema. Por favor verifique.", vbInformation, "Información SIGEEST"
        End
    Else
        Open App.Path & "\DatDirInstall.txt" For Input As #1
        Line Input #1, a
        Close #1
        direccionSistema = a
    End If

End Sub

