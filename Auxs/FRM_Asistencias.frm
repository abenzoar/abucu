VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Asistencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistencias"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19005
   Icon            =   "FRM_Asistencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   19005
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   16200
      TabIndex        =   18
      Top             =   1095
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   14640
      Top             =   360
   End
   Begin VB.CheckBox chkHuella 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   13680
      TabIndex        =   7
      Top             =   1110
      Width           =   135
   End
   Begin VB.CommandButton cmdHuella 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      Picture         =   "FRM_Asistencias.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12480
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdButon 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      Picture         =   "FRM_Asistencias.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer tHorFecha 
      Interval        =   1000
      Left            =   13920
      Top             =   240
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   9135
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   16113
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   $"FRM_Asistencias.frx":1A5E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox grFinger 
      Height          =   240
      Left            =   10080
      ScaleHeight     =   180
      ScaleWidth      =   1500
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activar impresión de turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   4
      Left            =   16440
      TabIndex        =   20
      Top             =   1125
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activar lector de huellas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   3
      Left            =   13920
      TabIndex        =   19
      Top             =   1125
      Width           =   2175
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   5
      Left            =   17640
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   4
      Left            =   15720
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   3
      Left            =   13800
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   2
      Left            =   17640
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   1
      Left            =   15720
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información de la cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   2
      Left            =   13800
      TabIndex        =   17
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   2
      X1              =   13800
      X2              =   18360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   13800
      X2              =   18360
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fotos de usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   13800
      TabIndex        =   16
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   13800
      X2              =   18360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información de huella digital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   13800
      TabIndex        =   15
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   9
      X1              =   120
      X2              =   3600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lInfo2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   13560
      TabIndex        =   13
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Label lInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   13560
      TabIndex        =   12
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label lHuella 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   17520
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label NombreVerificar 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   255
      Left            =   15600
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label AreaVerificar 
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      Height          =   255
      Left            =   15600
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imagenHuella 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   13800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1080
      Index           =   0
      Left            =   18600
      Picture         =   "FRM_Asistencias.frx":1AEF
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   960
      Index           =   1
      Left            =   18600
      Picture         =   "FRM_Asistencias.frx":23B9
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   840
   End
   Begin VB.Image fotoUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   0
      Left            =   13800
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lFecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Miércoles, 31 de Septiembre del 2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   480
      Left            =   9240
      TabIndex        =   3
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label lHora 
      BackStyle       =   0  'Transparent
      Caption         =   "22:00 P.M."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Mensajes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   15600
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   10095
      Index           =   1
      Left            =   0
      Picture         =   "FRM_Asistencias.frx":2C83
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18975
   End
End
Attribute VB_Name = "FRM_Asistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim res1 As Recordset
Dim RES2 As Recordset
Dim TT1 As New clss_ToolTipText

Private Sub chkHuella_Click()
On Error Resume Next
    If chkHuella.value = Checked Then
        Dim Error As Integer
        Error = Inicializar(Me)
    Else
'        grFinger.CapStopCapture (idSensor)
'        grFinger.CapFinalize
    End If
    
    If Err.Number <> 0 Then
        MsgBox "No se puede iniciar el lecto de huellas. Verifique.", vbInformation
    End If
End Sub

Private Sub chkPrint_Click()
    If chkPrint.value = Checked Then
        txtUsuario(0).SetFocus
    End If
End Sub

Private Sub cmdButon_Click()
    modBusqueda = "Asistencia"
    tipoBusqueda = "C"
    BUSQ_Usuarios.Show vbModal
End Sub

Public Sub cmdCheck_Click()
    checkUsuario
End Sub

Private Sub Form_Load()
    lHora.Caption = Format(Time, "Medium Time")
    lFecha.Caption = Format(Date, "Long Date")
    Image1(0).Visible = False
    Image1(1).Visible = False
    'chkHuella.Value = Checked
    'FRM_Menu.Visible = False
    cargaLista
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ques As String
    ques = MsgBox("¿Salir?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        If chkHuella.value = Checked Then
'                grFinger_SensorUnplug (idSensor)
'                grFinger.CapStopCapture (idSensor)
'                grFinger.CapFinalize
'                grFinger.Finalize
        End If
        Cancel = 0
        'End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Lista_Click()
    'muestraInfo (lista.TextMatrix(lista.Row, 1))
End Sub
Private Sub muestraInfo(perId As String)

    fotoUser(0).Picture = LoadPicture("")
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
        
    sql1 = "SELECT PER_ID, PER_PATERNO, PER_MATERNO, PER_NOMBRE, PERTP_USUARIO, PER_FEC_NAC, if(PERTP_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, " & _
    "(YEAR(CURDATE()) - YEAR(PER_FEC_NAC)) EDAD, PER_FOTO, PER_EMAIL, PER_TEL1, PER_TEL2, CTPT_TIPO " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_TIPO_ID = T3.CTPT_ID  AND T2.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
    "AND T2.PERTP_CODIGO_MEMBRESIA = '" & perId & "'"
    Set res1 = con.Execute(sql1)
    If Not res1.EOF Then
        If IsNull(res1.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write res1.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            fotoUser(0).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            fotoUser(0).Picture = LoadPicture("")
        End If
'        lInfo(0).Caption = "Usuario: " & RES1.Fields("PERTP_USUARIO")
'        lInfo(1).Caption = "Clave: " & RES1.Fields("PER_ID")
'        lInfo(2).Caption = "Nombre: " & RES1.Fields("PER_NOMBRE")
'        lInfo(3).Caption = "Apellido: " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
'        lInfo(4).Caption = "Fecha de nac: " & RES1.Fields("PER_FEC_NAC")
'        lInfo(5).Caption = "Edad: " & RES1.Fields("EDAD")
'        lInfo(6).Caption = "Estatus: " & RES1.Fields("STATUS")
'        lInfo(7).Caption = "Email: " & RES1.Fields("PER_EMAIL")
'        lInfo(8).Caption = "Teléfonos: " & RES1.Fields("PER_TEL1") & " " & RES1.Fields("PER_TEL2")
'        lInfo(9).Caption = "Cargo: " & RES1.Fields("CTPT_TIPO")
        
    Else
        fotoUser(0).Picture = LoadPicture("")
'        lInfo(0).Caption = "Usuario: "
'        lInfo(1).Caption = "Clave: "
'        lInfo(2).Caption = "Nombre: "
'        lInfo(3).Caption = "Apellido: "
'        lInfo(4).Caption = "Fecha de nac: "
'        lInfo(5).Caption = "Edad: "
'        lInfo(6).Caption = "Estatus: "
'        lInfo(7).Caption = "Email: "
'        lInfo(8).Caption = "Teléfonos: "
'        lInfo(9).Caption = "Cargo: "
    
    End If


End Sub

Private Sub tHorFecha_Timer()
    If Me.WindowState = vbMaximized Then
        lHora.Caption = Format(Time, "Medium Time")
        lFecha.Caption = Format(Date, "Long Date")
        lista.height = Me.height - 1500
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Image2(1).width = Me.width
    Image2(1).height = Me.height

End Sub

Private Sub txtUsuario_GotFocus(Index As Integer)
               
        
        TT1.Title = "Código del cliente"
        TT1.TipText = "Escribe o digita el código del cliente. " & vbCrLf & vbCrLf & "Si tiene un escaner coloce el código de barras en el lector"
        TT1.Style = TTBalloon
        TT1.Icon = TTIconError
        TT1.ForeColor = vbWhite
        TT1.BackColor = &HCE7110
        TT1.PopupOnDemand = False
        TT1.CreateToolTip txtUsuario(Index).hwnd

End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        'TopForm_Asistencia.Show
        checkUsuario
    End If
End Sub
Private Sub checkUsuario()
    lInfo2.Caption = ""
    
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    
    sql1 = "SELECT T2.PER_ID, T1.PERTP_CODIGO_MEMBRESIA, T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FOTO, " & _
    "T1.PERTP_TIPO_ID, T1.PERTP_PER_ID, T1.PERTP_PER_TIPO " & _
    "FROM PER_TIPO T1, PERSONA T2, CAT_TIPO T3 " & _
    "WHERE T1.PERTP_PER_ID = T2.PER_ID AND T1.PERTP_TIPO_ID = T3.CTPT_ID AND T1.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
    "AND T1.PERTP_CODIGO_MEMBRESIA = '" & txtUsuario(0).Text & "' AND T1.PERTP_STATUS = 'A'"
    Set res1 = con.Execute(sql1)
        
    If Not res1.EOF Then
        lInfo2.Caption = res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO") & vbCrLf & _
        "Clave persona: " & res1.Fields("PER_ID")
        If res1.Fields("PERTP_PER_TIPO") = "C" Then
                    
            sql1 = "SELECT T2.PER_ID, T1.PERTP_CODIGO_MEMBRESIA, T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FOTO, " & _
            "T1.PERTP_TIPO_ID, T1.PERTP_PER_ID, T1.PERTP_PER_TIPO, T4.mbr_ctmbId, T4.mbr_VentaFolio, T4.MBR_FIN, DATEDIFF(MBR_fIN, CURDATE()) DIAS  " & _
            "FROM PER_TIPO T1, PERSONA T2, CAT_TIPO T3, MEMBRESIAS T4 " & _
            "WHERE T1.PERTP_PER_ID = T2.PER_ID AND T1.PERTP_TIPO_ID = T3.CTPT_ID AND T1.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
            "AND MBR_PERTP_PER_ID = T1.PERTP_PER_ID AND MBR_PERTP_TIPO_ID = T1.PERTP_TIPO_ID " & _
            "AND CURDATE() BETWEEN T4.MBR_INICIO AND T4.MBR_FIN AND MBR_STATUS = 'A'  " & _
            "AND T1.PERTP_CODIGO_MEMBRESIA = '" & txtUsuario(0).Text & "'"
            Set res1 = con.Execute(sql1)
            If Not res1.EOF Then
                sql1 = "INSERT INTO ASISTENCIAS (ast_PertpTipoId, ast_PertpPerId, ast_PertpPerTipo, ast_FechaHora, ast_membresia, ast_ventfolio) " & _
                "VALUES " & _
                "('" & res1.Fields("PERTP_TIPO_ID") & "', '" & res1.Fields("PERTP_PER_ID") & "', '" & res1.Fields("PERTP_PER_TIPO") & "', " & _
                "NOW(), '" & res1.Fields("MBR_CTMBID") & "',  '" & res1.Fields("MBR_VENTAFOLIO") & "')"
                con.Execute (sql1)
                                        
                cargaLista
                
                If chkPrint.value = Checked Then
                    imprimirTicket
                Else
                    ''''
                End If
                
                Unload TopForm_Asistencia
                If IsNull(res1.Fields("PER_fOTO")) = False Then
                    checarCarpetaTemp
                    Imagen1.Open
                    Imagen1.Write res1.Fields("PER_FOTO")
                    Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
                    Imagen1.Close
                    TopForm_Asistencia.iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
                Else
                    TopForm_Asistencia.iFoto.Picture = LoadPicture("")
                End If
                        
                TopForm_Asistencia.lDatos.Caption = res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO") & _
                vbCrLf & res1.Fields("PERTP_CODIGO_MEMBRESIA") & vbCrLf & "Tu membresia vence: " & res1.Fields("MBR_FIN") & vbCrLf & _
                "Días disponibles: " & res1.Fields("DIAS")
                TopForm_Asistencia.Show
                txtUsuario(0).Text = ""
                txtUsuario(0).SetFocus
                Image1(0).Visible = True
                Image1(1).Visible = False
                lInfo.Caption = "Bienvenido " & vbCrLf & res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO") & _
                vbCrLf & "Tu membresia vence: " & res1.Fields("MBR_FIN") & vbCrLf & _
                "Días disponibles: " & res1.Fields("DIAS")
                lInfo2.Caption = ""
            
            Else
                'MsgBox "El usuario no cuenta con membresía o no está vigente. " & vbCrLf & vbCrLf & "Verifique.", vbExclamation
                lInfo.Caption = "El usuario no cuenta con membresía o no está vigente."
                Image1(1).Visible = True
                Image1(0).Visible = False
            End If
        Else
            If res1.Fields("PERTP_PER_TIPO") = "U" Then
                sql1 = "SELECT T2.PER_ID, T1.PERTP_CODIGO_MEMBRESIA, T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FOTO, " & _
                "T1.PERTP_TIPO_ID, T1.PERTP_PER_ID, T1.PERTP_PER_TIPO " & _
                "FROM PER_TIPO T1, PERSONA T2, CAT_TIPO T3 " & _
                "WHERE T1.PERTP_PER_ID = T2.PER_ID AND T1.PERTP_TIPO_ID = T3.CTPT_ID AND T1.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
                "AND T1.PERTP_CODIGO_MEMBRESIA = '" & txtUsuario(0).Text & "' AND PERTP_PER_TIPO = 'U' AND PERTP_STATUS = 'A' "
                Set res1 = con.Execute(sql1)
                If Not res1.EOF Then
                    sql1 = "INSERT INTO ASISTENCIAS (ast_PertpTipoId, ast_PertpPerId, ast_PertpPerTipo, ast_FechaHora ) " & _
                    "VALUES " & _
                    "('" & res1.Fields("PERTP_TIPO_ID") & "', '" & res1.Fields("PERTP_PER_ID") & "', '" & res1.Fields("PERTP_PER_TIPO") & "', " & _
                    "NOW())"
                    con.Execute (sql1)
                    
                    cargaLista
                    
                    Unload TopForm_Asistencia
                    If IsNull(res1.Fields("PER_fOTO")) = False Then
                        checarCarpetaTemp
                        Imagen1.Open
                        Imagen1.Write res1.Fields("PER_FOTO")
                        Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
                        Imagen1.Close
                        TopForm_Asistencia.iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
                    Else
                        TopForm_Asistencia.iFoto.Picture = LoadPicture("")
                    End If
                            
                    TopForm_Asistencia.lDatos.Caption = res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO")
                    TopForm_Asistencia.Show
                    txtUsuario(0).Text = ""
                    txtUsuario(0).SetFocus
                    Image1(0).Visible = True
                    Image1(1).Visible = False
                    lInfo.Caption = "Bienvenido " & vbCrLf & res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO")
                    lInfo2.Caption = ""
                Else
                    'MsgBox "El usuario no cuenta con membresía o no está vigente. " & vbCrLf & vbCrLf & "Verifique.", vbExclamation
                    lInfo.Caption = "Usuario no encontrado."
                    Image1(1).Visible = True
                    Image1(0).Visible = False
                End If
            End If
        End If
    End If
    
End Sub
Private Sub imprimirTicket()
    
    Printer.KillDoc
    Printer.Font = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "Bienvenido"
    Printer.Print "--------------------"
    Printer.Print "Turno: "
    Printer.FontSize = 36
    Printer.Print lista.Rows - 1
    Printer.FontSize = 10
    Printer.Print "--------------------"
    Printer.Print Format(Date, "Short Date")
    Printer.Print Format(Time, "Short Time")
    Printer.EndDoc

End Sub

Private Sub cargaLista()
    On Error Resume Next
    Dim numFoto As Long

    For b1 = 0 To 5
        fotoUser(b1).Picture = LoadPicture("")
    Next b1

    sql1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA, if(PERTP_PER_TIPO= 'C', 'Cliente', 'USUARIO') TIPO, ast_FechaHora, " & _
    "T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FOTO, " & _
    "IF(((SELECT COUNT(*) FROM ASISTENCIAS TA WHERE TA.ast_PertpPerId = T1.ast_PertpPerId AND " & _
    "DATE_FORMAT(TA.AST_FECHAHORA, '%d/%m/%y') = DATE_FORMAT(NOW(), '%d/%m/%y') AND " & _
    "TA.AST_FECHAHORA <= T1.AST_FECHAHORA) % 2)='1', 'ENTRADA', 'SALIDA') TIPO_AS " & _
    "FROM ASISTENCIAS T1, PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
    "WHERE T1.ast_PertpTipoId = T4.PERTP_TIPO_ID AND T1.ast_PertpPerId = T4.PERTP_PER_ID AND T1.ast_PertpPerTipo = T4.PERTP_PER_TIPO " & _
    "AND T1.ast_PertpPerId = T2.PER_ID AND T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
    "AND DATE_FORMAT(T1.AST_FECHAHORA, '%d/%m/%y') = DATE_FORMAT(NOW(), '%d/%m/%y') ORDER BY AST_FECHAHORA DESC"
    Set RES2 = con.Execute(sql1)
    
    lista.Rows = 1
    numFoto = 0
    
    lista.Redraw = False
    
    Do While Not RES2.EOF
        
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = lista.Rows - 1
        lista.TextMatrix(lista.Rows - 1, 1) = RES2.Fields("PERTP_CODIGO_MEMBRESIA")
        lista.TextMatrix(lista.Rows - 1, 2) = RES2.Fields("PER_NOMBRE") & " " & RES2.Fields("PER_PATERNO") & " " & RES2.Fields("PER_MATERNO")
        lista.TextMatrix(lista.Rows - 1, 3) = RES2.Fields("TIPO")
        lista.TextMatrix(lista.Rows - 1, 4) = Format(RES2.Fields("ast_FechaHora"), "Medium time")
        lista.TextMatrix(lista.Rows - 1, 5) = RES2.Fields("TIPO_AS")
        
                
        If numFoto < 6 Then
            If IsNull(RES2.Fields("PER_fOTO")) = False Then
                Dim Imagen1 As Stream
                Set Imagen1 = New Stream
                Imagen1.Type = adTypeBinary
                checarCarpetaTemp
                Imagen1.Open
                Imagen1.Write RES2.Fields("PER_FOTO")
                Imagen1.SaveToFile direccionSistema & "\Temp\TempClie.dat", adSaveCreateOverWrite
                Imagen1.Close
                fotoUser(numFoto).Picture = LoadPicture(direccionSistema & "\Temp\TempClie.dat")
            Else
                fotoUser(numFoto).Picture = LoadPicture("")
            End If
        End If
        
        numFoto = numFoto + 1
        
        RES2.MoveNext
    Loop
    
    If lista.Rows > 1 Then
        lista.Row = lista.Rows - 1
        lista.RowSel = lista.Rows - 1
        For b1 = 0 To lista.Cols - 1
            lista.Col = b1
            lista.CellForeColor = vbWhite
            lista.CellBackColor = &H8000000D
        Next b1
    End If
    
    lista.Redraw = True
    
'    If lista.Rows > 25 Then
'        lista.TopRow = lista.Rows - 1
'    End If

End Sub

''''''DE AQUI PARA LA HUELLA
''''''DE AQUI PARA LA HUELLA
''''''DE AQUI PARA LA HUELLA
Private Sub grFinger_FingerDown(ByVal idSensor As String)
 ' Aqui detecta cuando pones el dedo (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Detector = "Huella detectada"
End Sub

Private Sub grFinger_FingerUp(ByVal idSensor As String)
 ' Aqui detecta cuando pones el dedo (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Detector = "Huella removida"
End Sub

Private Sub grFinger_ImageAcquired(ByVal idSensor As String, ByVal width As Long, ByVal height As Long, rawImage As Variant, ByVal res As Long)
 ' Capturar Imagen (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Mensajes = "Capturando imagen..."
 
 With raw
   .img = rawImage
   .height = height
   .width = width
   .res = res
 End With
 
 
CapturaHuella False, GR_DEFAULT_CONTEXT, Me, Me.imagenHuella(3), 3
If EncuentraPuntos(Me, Mensajes, imagenHuella(3), 3) = True Then
    'El numero 3 es por el Template que es el numero 3
    CambiaFoco Identificar(Me, 3, Me.NombreVerificar, Me.AreaVerificar)
End If

End Sub
Private Sub grFinger_SensorPlug(ByVal idSensor As String)
 ' Inicializar la Captura del dispositivo
' grFinger.CapStartCapture (idSensor)
End Sub
Private Sub grFinger_SensorUnplug(ByVal idSensor As String)
 ' Finalizar la Captura del dispositivo
' grFinger.CapStopCapture (idSensor)
End Sub

Private Sub CambiaFoco(Color As Integer)
Dim Cadena As String
 If Color = 1 Then
    txtUsuario(0).Text = NombreVerificar.Caption
    lHuella.Caption = ""
    lInfo.Caption = ""
    checkUsuario
 Else
    lHuella.Caption = "Huella no encontrada. Verifique."
    lInfo.Caption = ""
    Image1(1).Visible = False
    Image1(0).Visible = False
    'MsgBox "Huella no detectada. Verfique.", vbInformation
 End If
End Sub

Private Sub txtUsuario_LostFocus(Index As Integer)
    TT1.Destroy
End Sub
