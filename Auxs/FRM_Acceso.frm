VERSION 5.00
Begin VB.Form FRM_Acceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9060
   Icon            =   "FRM_Acceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8520
      Top             =   2640
   End
   Begin VB.ComboBox cmbSucur 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Selecciona la sucursal a la que deseas ingresar"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   7440
      Picture         =   "FRM_Acceso.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   6000
      Picture         =   "FRM_Acceso.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "informes: contacto@abucu.com.mx     www.abucu.com.mx      Hecho en México"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   7680
      Width           =   7335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2011. Abucu. Derechos reservados."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   8040
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FRM_Acceso.frx":1A5E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   960
      TabIndex        =   10
      Top             =   6600
      Width           =   7335
   End
   Begin VB.Label lDato 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.127"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lFecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   8595
      Left            =   0
      Picture         =   "FRM_Acceso.frx":1B87
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   9120
   End
End
Attribute VB_Name = "FRM_Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Option Explicit
    Dim sql1 As String
    Dim RES1 As Recordset
    Dim RES2 As Recordset
    Dim Sucursal As Boolean
    Dim Salir As Boolean
    'Dim TT1 As New clss_ToolTipText

Private Sub cmbSucur_Click()
Dim ques As String
    If loadDb = False Then
        If cmbSucur.Text <> dbActual Then
            ques = MsgBox("Cambiará de conexión de fuente de información. ¿Continuar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                buscarConexiones (cmbSucur.ItemData(cmbSucur.ListIndex))
            Else
                loadDb = True
                cmbSucur.Text = dbActual
                loadDb = False
            End If
            
        End If
    End If
End Sub

Private Sub cmdLogin_Click(Index As Integer)
    Dim a As String
    
    If Index = 0 Then
        validarUsuario
    Else
        a = MsgBox("Saldra de la aplicación. ¿Continuar?", vbQuestion + vbYesNo)
        If a = vbYes Then
            con.Close
            End
        End If
    End If
    
End Sub
Private Sub validarUsuario()
    
    Sucursal = False
    checkSucursal
    'Sucursal = True
    If Sucursal = True Then
    
        sql1 = "SELECT PERTP_USUARIO, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_TIPO_ID, CTPT_TIPO, " & _
        "PER_ID, PER_FOTO, SUC_RAZON_SOCIAL, SUC_NOMBRE, SUC_HORAENTRADA, SUC_HORASALIDA, SUC_MESAS, SUC_FOTO, SUC_DIA_CIERRE, SUC_EstadosOper, SUC_TICKETCOBRO " & _
        "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3, SUCURSAL T4 " & _
        "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' " & _
        "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'U' " & _
        "AND T2.PERTP_USUARIO = '" & txtUsuario(0).Text & "' AND T2.PERTP_PASSWORD = MD5('" & txtUsuario(1).Text & "') "
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
            If Sucursal = True Then
                If RES1.Fields("suc_mesas") = "S" Then
                    mesas = True
                Else
                    mesas = False
                End If
                If RES1.Fields("SUC_TICKETCOBRO") = "CR" Then
                    tipoTicket = "CORTO"
                Else
                    tipoTicket = "LARGO"
                End If
                FRM_Menu.menuBarra2.Panels(3).Text = "Sucursal: " & RES1.Fields("SUC_NOMBRE") & "" 'RES1.Fields("PERTP_TIPO_ID")
                FRM_Menu.menuBarra2.Panels(4).Text = RES1.Fields("PERTP_USUARIO")
                FRM_Menu.menuBarra2.Panels(5).Text = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
                FRM_Menu.menuBarra2.Panels(6).Text = RES1.Fields("CTPT_TIPO")
                FRM_Menu.menuBarra2.Panels(7).Text = RES1.Fields("PER_ID")
                FRM_Menu.menuBarra2.Panels(8).Text = RES1.Fields("PERTP_TIPO_ID")
                FRM_Menu.menuBarra2.Panels(9).Text = RES1.Fields("SUC_RAZON_sOCIAL") & "" 'RES1.Fields("PERTP_TIPO_ID")
                FRM_Menu.menuBarra2.Panels(10).Text = "Horario de: " & Format(RES1.Fields("SUC_HORAENTRADA"), "Short Time") & " a " & Format(RES1.Fields("SUC_HORASALIDA"), "Short Time")
                FRM_Menu.menuBarra2.Panels(11).Text = Format(RES1.Fields("SUC_HORAENTRADA"), "Long Time")
                FRM_Menu.menuBarra2.Panels(12).Text = Format(RES1.Fields("SUC_HORASALIDA"), "Long Time")
                FRM_Menu.menuBarra2.Panels(13).Text = RES1.Fields("SUC_DIA_CIERRE")
                FRM_Menu.menuBarra2.Panels(14).Text = RES1.Fields("SUC_EstadosOper") & ""
                FRM_Menu.lInfo(1).Caption = RES1.Fields("PERTP_USUARIO")
                If IsNull(RES1.Fields("PER_fOTO")) = False Then
                    Dim Imagen1 As Stream
                    Set Imagen1 = New Stream
                    Imagen1.Type = adTypeBinary
                    checarCarpetaTemp
                    Imagen1.Open
                    Imagen1.Write RES1.Fields("PER_FOTO")
                    Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
                    Imagen1.Close
                    FRM_Menu.imgInfo(0).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
                Else
                    FRM_Menu.imgInfo(0).Picture = LoadPicture("")
                End If
                
                If IsNull(RES1.Fields("SUC_fOTO")) = False Then
                    Set Imagen1 = New Stream
                    Imagen1.Type = adTypeBinary
                    checarCarpetaTemp
                    Imagen1.Open
                    Imagen1.Write RES1.Fields("SUC_FOTO")
                    Imagen1.SaveToFile direccionSistema & "\Temp\TempSucur.dat", adSaveCreateOverWrite
                    Imagen1.Close
                    FRM_Menu.imgInfo(1).Picture = LoadPicture(direccionSistema & "\Temp\TempSucur.dat")
                Else
                    FRM_Menu.imgInfo(1).Picture = LoadPicture("")
                End If
                
                Salir = True
                
                FRM_Menu.Show
                
                Unload Me
            Else
                FRM_DatosSuc.lbStatus.Caption = "Agregando sucursal"
                FRM_DatosSuc.Show
            End If
        Else
            txtUsuario(1).Text = ""
            txtUsuario(0).SetFocus
            TT1.Style = TTBalloon
            TT1.Icon = TTIconError
            TT1.Title = "Error en la información"
            TT1.TipText = "Por favor verifique su usuario y contraseña." & vbCrLf & "La información proporcionada no es correcta."
            TT1.PopupOnDemand = True
            TT1.CreateToolTip txtUsuario(0).hWnd
            TT1.Show (txtUsuario(0).Left / Screen.TwipsPerPixelX - 1) + 400, (txtUsuario(0).Top / Screen.TwipsPerPixelY - 1) + 200
        End If
    Else
        MsgBox "No se puede iniciar sesión por falta de datos. Verifique.", vbInformation
    End If
End Sub

'Private Sub Command1_Click()
'    Dim Path As String
'
'    Path = InputBox(" Ruta del archivo para obtener la versión", _
'                    " Averiguar Versión ")
'
'    If Path = vbNullString Then Exit Sub
'
'      Muestra la versión del fichero
'    MsgBox Obtener_Version(Path), vbInformation, " Versión del archivo "
'End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TT1.Destroy
End Sub

Private Sub Form_Paint()
    TT1.Destroy
End Sub

Private Sub Form_Resize()
    TT1.Destroy
End Sub
Private Sub cmdLogin_GotFocus(Index As Integer)
    cmdLogin(Index).BackColor = vbWhite
End Sub

Private Sub cmdLogin_LostFocus(Index As Integer)
    cmdLogin(Index).BackColor = &HE0E0E0
End Sub

Private Sub cmdLogin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdLogin(Index).BackColor = vbWhite
    
End Sub

Private Sub cmdLogin_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    cmdLogin(Index).BackColor = vbWhite
End Sub

Function Obtener_Version(Path_File As String) As Variant
      
    On Local Error GoTo ErrSub
      
    Dim Fso As Object
      
    ' Crea un Nuevo objeto FSO
    Set Fso = CreateObject("Scripting.FileSystemObject")
      
    'Ejecuta el método  GetFileVersion
    Obtener_Version = Fso.GetFileVersion(Path_File)
       
    Set Fso = Nothing
  
Exit Function
  
'Error
ErrSub:
  
MsgBox Err.Description, vbCritical
  
End Function

Private Sub Form_Load()
    
    If App.PrevInstance Then
        MsgBox "La aplicación Auxs System ya se encuentra en ejecución. " & vbCrLf & vbCrLf & "Verifique", vbExclamation
        Exit Sub
        End ' Solo ejecutar la App una vez
    End If
    
      
    lDato(1).Caption = "Versión: " & Obtener_Version(App.Path & "\AuxsSis.exe")
    
    Salir = False
    usuarioInicial = False
    ConectarDB
    'checkMac
    checkPeriodo
    checkUsuarios
    'checkSucursal
    checkFecha
    cargaToolTips
        
    
End Sub
Private Sub checkMac()

On Error Resume Next
Dim Devices As Object
Dim Device As Object
Dim Temp As Variant
Dim Info As String
Dim mac As String
Dim sql1 As String
Dim valor As Boolean
valor = False
Set Devices = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapter")
For Each Device In Devices
 For Each Temp In Device.Properties_
' If Temp.Name = "MACAddress" Then GetMACAddress = CStr(Temp)
' If Temp.Name = "MACAddress" Then Text1.Text = CStr(Temp)
    If Temp.Name = "MACAddress" Then
        mac = CStr(Temp)
       If mac <> "" Then
            mac = CStr(Temp)
'            MsgBox mac
           valor = True
            GoTo validar
       End If
    End If
 Next
Next Device

validar:

If valor = True Then
        sql1 = "SELECT AES_DECRYPT(suc_mc, '9807288') MAC FROM SUCURSAL"
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
            If IsNull(RES1.Fields("MAC")) = True Then
                sql1 = "UPDATE SUCURSAL SET SUC_MC = AES_ENCRYPT('" & mac & "', '9807288')"
                con.Execute (sql1)
                
                sql1 = "UPDATE SUCURSAL SET SUC_PERIODO = AES_ENCRYPT( (DATE_add(NOW(), INTERVAL 7 DAY)) , '9807288')"
                con.Execute (sql1)
                
                MsgBox "Se ha detectado una actualización indevida de datos. " & vbCrLf & vbCrLf & "Se ha activado el periodo de prueba." & _
                vbCrLf & vbCrLf & "Comuníquese con su proveedor para cualquier aclaración.", vbExclamation
                
            Else
                If RES1.Fields("MAC") <> mac Then
                    sql1 = "UPDATE SUCURSAL SET SUC_MC = AES_ENCRYPT('" & mac & "', '9807288')"
                    con.Execute (sql1)
                    
                    sql1 = "UPDATE SUCURSAL SET SUC_PERIODO = AES_ENCRYPT( (DATE_add(NOW(), INTERVAL 7 DAY)) , '9807288')"
                    con.Execute (sql1)
                    
                    MsgBox "Se ha detectado que un cambio de equipo. " & vbCrLf & vbCrLf & "Se ha activado el periodo de prueba." & _
                    vbCrLf & vbCrLf & "Comuníquese con su proveedor para cualquier aclaración.", vbExclamation
                End If
            End If
        End If

End If

End Sub

Private Sub checkPeriodo()
    On Error Resume Next
    sql1 = "SELECT AES_DECRYPT(suc_periodo, '9807288') FECHA FROM SUCURSAL"
    Set RES1 = con.Execute(sql1)

    If Not RES1.EOF Then
        If Format(Date, "yyyy-MM-dd") >= RES1.Fields("FECHA") Then
            MsgBox "La fecha límite para utilizar el sistema ha llegado. " & vbCrLf & vbCrLf & _
            "Por favor renueve su licencia para continuar usando el sistema", vbExclamation
        End If
        If Format(Date, "yyyy-MM-dd") > RES1.Fields("FECHA") Then
            End
        End If
    End If
    

End Sub

Private Sub cargaToolTips()
        
    TT3.Title = "Nombre de Usuario"
    TT3.TipText = "Escribe el nombre de usuario de tu cuenta para accesar"
    TT3.Style = TTBalloon
    TT3.Icon = TTIconError
    TT3.ForeColor = vbWhite
    TT3.BackColor = &HCE7110
    TT3.PopupOnDemand = False
    TT3.VisibleTime = 6000
    TT3.CreateToolTip txtUsuario(0).hWnd
    
    TT4.Title = "Contraseña de Usuario"
    TT4.TipText = "Escribe la contraseña de tu cuenta para accesar"
    TT4.Style = TTBalloon
    TT4.Icon = TTIconError
    TT4.ForeColor = vbWhite
    TT4.BackColor = &HCE7110
    TT4.PopupOnDemand = False
    TT4.VisibleTime = 6000
    TT4.CreateToolTip txtUsuario(1).hWnd

End Sub

Private Sub checkFecha()
    lFecha(0).Caption = Format(Date, "DDDD-MMMM-DD-yyyy") & " " & Format(Time, "HH:MM")
End Sub
Private Sub checkUsuarios()
    sql1 = "SELECT COUNT(*) NUM FROM PER_TIPO WHERE PERTP_PER_TIPO = 'U' AND PERTP_STATUS = 'A'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        If RES1.Fields("NUM") = 0 Then
            a = MsgBox("Bienvenido a Auxs Sis. " & vbCrLf & vbCrLf & _
            "Se ha detectado que es la primera vez que va a iniciar sesión en el Sistema. " & _
            "Deberá registrarse como usuario para poder iniciar sesión. " & _
            "¿Continuar?", vbQuestion + vbYesNo)
            If a = vbYes Then
                usuarioInicial = True
                FRM_Usuarios.Show vbModal
            Else
                con.Close
                End
            End If
        End If
    End If
End Sub
Private Sub checkSucursal()
    sql1 = "SELECT COUNT(*) NUM FROM SUCURSAL WHERE SUC_LOCAL = 'S'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        If RES1.Fields("NUM") = 0 Then
            MsgBox "Bienvenido a Auxs Sis " & txtUsuario(0).Text & ". " & vbCrLf & vbCrLf & _
            "No se han asignado los datos de identificación el negocio. " & vbCrLf & _
            "Deberá completar la información que se solicita.", vbInformation
            Sucursal = False
            FRM_DatosSuc.Show vbModal
        Else
            Sucursal = True
        End If
    End If
End Sub
Private Sub ConectarDB()
'    Call ConexionDB("localhost", "auXs_Db", "root", "9807288")
    loadDb = True
    Call buscarConexiones("actual")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
Dim ques As String

If Salir = False Then

    ques = MsgBox("¿Salir?", vbYesNo + vbQuestion)
    
    If ques = vbYes Then
        Cancel = 0
        TT1.Destroy
        TT2.Destroy
        TT3.Destroy
        TT4.Destroy
    Else
        Cancel = 1
    
    End If
End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdLogin(0).BackColor = -2147483633
    cmdLogin(1).BackColor = -2147483633
End Sub

Private Sub Timer1_Timer()
    checkFecha
End Sub

Private Sub txtUsuario_GotFocus(Index As Integer)
    txtUsuario(Index).SelStart = 0
    txtUsuario(Index).SelLength = Len(txtUsuario(Index).Text)
End Sub

Private Sub txtUsuario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If CapsLockOn Then
        TT2.Style = TTBalloon
        TT2.Icon = TTIconWarning
        TT2.Title = "Mayúsculas activadas"
        TT2.TipText = "Verifica si las mayúsuculas están activas..."
        TT2.CreateToolTip txtUsuario(Index).hWnd
        TT2.Show 0, txtUsuario(Index).height / Screen.TwipsPerPixelX - 1
    Else
        TT2.Destroy
        TT1.Destroy
    End If
End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        txtUsuario(Index).PasswordChar = "*"
        txtUsuario(Index).FontBold = True
        txtUsuario(Index).FontSize = 16
    End If
    If KeyAscii = 13 Then
        cmdLogin_Click (0)
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub txtUsuario_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TT1.Destroy
End Sub
