VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ADD_Cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar cliente "
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10035
   Icon            =   "ADD_Cliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   6720
      Width           =   6495
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   4440
      MaxLength       =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   2295
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5760
      Width           =   3975
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   7560
      MaxLength       =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   11
      Left            =   240
      MaxLength       =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   4575
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6840
      Picture         =   "ADD_Cliente.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guardar cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   360
      Picture         =   "ADD_Cliente.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   2775
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2760
      Width           =   3375
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3840
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63176705
      CurrentDate     =   40783
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   5040
      MaxLength       =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   12
      Left            =   240
      Top             =   6720
      Width           =   6525
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Recomendado por "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   28
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Código membresía *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   26
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   11
      Left            =   4440
      Top             =   5760
      Width           =   2325
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Membresía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   24
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   10
      Left            =   240
      Top             =   5760
      Width           =   4005
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   9
      Left            =   7560
      Top             =   4800
      Width           =   2325
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   8
      Left            =   5040
      Top             =   4800
      Width           =   2325
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   7
      Left            =   240
      Top             =   4800
      Width           =   4605
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   6
      Left            =   4680
      Top             =   3840
      Width           =   2565
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   5
      Left            =   240
      Top             =   3840
      Width           =   4005
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   4
      Left            =   5640
      Top             =   2760
      Width           =   3405
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   3
      Left            =   2880
      Top             =   2760
      Width           =   2445
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   2
      Left            =   240
      Top             =   2760
      Width           =   2325
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   1
      Left            =   240
      Top             =   1920
      Width           =   3525
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   5565
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del contacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   2
      Left            =   240
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Razón social - Nombre empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   20
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   19
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de nacimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   17
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido materno *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido paterno *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Género *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   13
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de cliente *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   9855
      Index           =   0
      Left            =   -600
      Picture         =   "ADD_Cliente.frx":171E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "ADD_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim res1 As Recordset
Dim checkError As Boolean
Dim mayus As Boolean
    
Private Sub cargaTipoCliente()
       Dim tipo As String
    If tipoPersona = "CLIENTE" Or tipoPersona = "CLIENTE_DEVO" Or tipoPersona = "CLIENTE_EDIT" Then
        tipo = "C"
    Else
        If tipoPersona = "PROVEEDOR" Or tipoPersona = "PROVEEDOR_G" Then
            tipo = "V"
        End If
    End If
    
    sql1 = ("SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = '" & tipo & "' ORDER BY CTPT_TIPO")
    Set res1 = con.Execute(sql1)
    cmbUser(5).Clear
    Do While Not res1.EOF
        cmbUser(5).AddItem res1.Fields("CTPT_TIPO")
        cmbUser(5).ItemData(cmbUser(5).ListCount - 1) = res1.Fields("CTPT_ID")
        res1.MoveNext
    Loop
    
    
    If cmbUser(5).ListCount > 0 Then
        cmbUser(5).ListIndex = 0
    End If
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            checarCampos
            If checkError = False Then
                'crearCliente
                If tipoPersona = "CLIENTE" Or tipoPersona = "CLIENTE_DEVO" Then
                    crearCliente
                Else
                    If tipoPersona = "CLIENTE_EDIT" Then
                        editarClientE
                    Else
                        If tipoPersona = "PROVEEDOR" Or tipoPersona = "PROVEEDOR_G" Then
                            crearCliente
                        End If
                    End If
                End If
'                If tipoPersona = "PROVEEDOR" Then
'                    crearProveedor
'                Else
'                    If tipoPersona = "CLIENTE" Then
'                        crearCliente
'                    End If
'                End If
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
            End If
        Case 1:
            Dim ques As String
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Unload Me
            End If
    End Select
End Sub
Private Sub editarClientE()

  If txtUsuario(3).Text = "" Then
        tel1 = "null"
    Else
        tel1 = txtUsuario(3).Text
    End If
        
    If txtUsuario(4).Text = "" Then
        tel2 = "null"
    Else
        tel2 = txtUsuario(4).Text
    End If

    sql1 = "UPDATE PERSONA SET PER_NOMBRE = '" & txtUsuario(0).Text & "',  " & _
    "PER_PATERNO = '" & txtUsuario(1).Text & "', " & _
    "PER_MATERNO = '" & txtUsuario(2).Text & "', " & _
    "PER_GENERO = '" & Left(cmbUser(2).Text, 1) & "'," & _
    "PER_EMAIL = '" & txtUsuario(11).Text & "', " & _
    "PER_TEL1 = " & tel1 & ", " & _
    "PER_TEL2 = " & tel2 & " " & _
    "WHERE PER_ID = '" & BUSQ_Usuarios.lista.TextMatrix(BUSQ_Usuarios.lista.Row, 4) & "' "
    con.Execute (sql1)
    
    sql1 = "UPDATE PER_TIPO SET PERTP_TIPO_ID = " & cmbUser(5).ItemData(cmbUser(5).ListIndex) & ", " & _
    "PERTP_MEMBRESIA  = '" & Left(cmbUser(0).Text, 1) & "', PERTP_CODIGO_mEMBRESIA = '" & txtUsuario(6).Text & "' " & _
    "WHERE PERTP_PER_ID = '" & BUSQ_Usuarios.lista.TextMatrix(BUSQ_Usuarios.lista.Row, 4) & "'"
    con.Execute (sql1)
    
    MsgBox "Edición realizada. Verfique.", vbInformation
    
    BUSQ_Usuarios.buscarUsuario
    Unload Me
    
End Sub

Private Sub Form_Load()
    If tipoPersona = "PROVEEDOR" Then
        txtUsuario(5).Visible = True
        lUsuario(5).Visible = True
        lProd(12).Caption = "Datos del contacto"
        lUsuario(26).Caption = "Tipo de proveedor"
        Me.Caption = "Agregando proveedor - Formato rápido"
        Borde(0).Visible = True
        lUsuario(31).Visible = False
        dtFecha(0).Visible = False
        Borde(6).Visible = False
        'txtUsuario(5).SetFocus
        cmbUser(0).Visible = False
        txtUsuario(6).Visible = False
        Borde(11).Visible = False
        Borde(10).Visible = False
        cmBoton(0).Caption = "Guardar proveedor"
            lUsuario(7).Visible = False
            lUsuario(6).Visible = False
        cmbUser(1).Visible = False
        lUsuario(8).Visible = False
        Borde(12).Visible = False
    Else
        If tipoPersona = "CLIENTE" Or tipoPersona = "CLIENTE_DEVO" Or tipoPersona = "CLIENTE_EDIT" Then
            txtUsuario(5).Visible = False
            lUsuario(5).Visible = False
            lProd(12).Caption = "Datos del cliente"
            lUsuario(26).Caption = "Tipo de cliente"
            Me.Caption = "Agregando cliente - Formato rápido"
            Borde(0).Visible = False
            lUsuario(31).Visible = False
            dtFecha(0).Visible = False
            Borde(6).Visible = False
            cmbUser(0).Visible = True
            cmBoton(0).Caption = "Guardar cliente"
            txtUsuario(6).Visible = True
            Borde(11).Visible = True
            Borde(10).Visible = True
            lUsuario(7).Visible = True
            lUsuario(6).Visible = True
            cmbUser(1).Visible = True
            lUsuario(8).Visible = True
            Borde(12).Visible = True
            
            'txtUsuario(0).SetFocus
        End If
    End If
    
    
    cmbUser(2).Clear
    cmbUser(2).AddItem "FEMENINO"
    cmbUser(2).AddItem "MASCULINO"
    cmbUser(2).ListIndex = 0
    
    cmbUser(0).Clear
    cmbUser(0).AddItem "SI"
    cmbUser(0).AddItem "NO"
        
    
    cargaTipoCliente
    cargarCliente
    checkMayus
    
    If tipoPersona = "CLIENTE_EDIT" Then
        cargaEdit
    End If
End Sub
Private Sub cargarCliente()
    cmbUser(1).Clear
    
    sql1 = "SELECT * fROM VIEW_PERSONA WHERE TIPO = 'CLIENTE' AND STATUS = 'ACTIVO' and tipo = 'CLIENTE' ORDER BY NOMBRE ASC "
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(1).AddItem res1.Fields("NOMBRE") & " " & res1.Fields("PATERNO") & " " & res1.Fields("MATERNO") & " ID: " & res1.Fields("ID")
        cmbUser(1).ItemData(cmbUser(1).ListCount - 1) = res1.Fields("ID")
        res1.MoveNext
    Loop


End Sub
Private Sub cargaEdit()
'    SQL1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA, if(PERTP_PER_TIPO= 'C', 'Cliente', 'USUARIO') TIPO, " & _
'    "T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FEC_NAC, T2.PER_EMAIL, T2.PER_TEL1, T2.PER_TEL2, T3.CTPT_TIPO TIPO, IF(T2.PER_GENERO='M', 'MASCULINO', 'FEMENINO') GENERO, T2.PER_ID, T4.PERTP_TIPO_ID, T2.PER_DESCRIPCION, IF(T4.PERTP_MEMBRESIA = 'S', 'SI', 'NO') MEMBRESIA " & _
'    "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
'    "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
'    " " & _
'    "AND T4.PERTP_PER_ID = '" & BUSQ_Usuarios.lista.TextMatrix(BUSQ_Usuarios.lista.Row, 4) & "' "
    sql1 = "SELECT * FROM VIEW_PERSONA WHERE ID = '" & BUSQ_Usuarios.lista.TextMatrix(BUSQ_Usuarios.lista.Row, 4) & "' "
    
    Set res1 = con.Execute(sql1)
    
    If Not res1.EOF Then
        txtUsuario(0).Text = res1.Fields("NOMBRE")
        txtUsuario(1).Text = res1.Fields("PATERNO")
        txtUsuario(2).Text = res1.Fields("MATERNO")
        cmbUser(2).Text = res1.Fields("GENERO")
        cmbUser(5).Text = res1.Fields("ROL")
        dtFecha(0) = res1.Fields("NACIMIENTO")
        txtUsuario(11).Text = res1.Fields("EMAIL") & ""
        txtUsuario(3).Text = res1.Fields("TEL1") & ""
        txtUsuario(4).Text = res1.Fields("TEL2") & ""
        txtUsuario(6).Text = res1.Fields("MEMBRESIA") & ""
        cmbUser(0).Text = res1.Fields("CON_MEMBRESIA")
        If IsNull(res1.Fields("RECOMENDADO_POR")) Then
            'cmbUser(1).AddItem ""
            'cmbUser(1).Text = ""
            
        Else
            cmbUser(1).Text = res1.Fields("RECOMENDADO_POR")
        End If
        cmbUser(1).Enabled = False
    Else
        MsgBox "No se ha encontrado información. Verifique.", vbInformation
    End If
End Sub

Private Sub checkMayus()
    sql1 = "SELECT SUC_MAYUSCULAS FROM SUCURSAL"
    Set res1 = con.Execute(sql1)
    If Not res1.EOF Then
        If res1.Fields("SUC_MAYUSCULAS") = "1" Then
            mayus = True
        Else
            mayus = False
        End If
    End If
    
End Sub
Private Sub checarCampos()
    checkError = False
    
    
    If txtUsuario(0).Text = "" Then
        checkError = True
        lUsuario(0).ForeColor = vbRed
        Exit Sub
    End If
    
    If tipoPersona = "CLIENTE_EDIT" Then
        If txtUsuario(6).Text = "" Then
            checkError = True
            'lUsuario(0).ForeColor = vbRed
            Exit Sub
        End If
    End If
    
    If tipoPersona = "PROVEEDOR" Then
        If txtUsuario(5).Text = "" Then
            checkError = True
            lUsuario(5).ForeColor = vbRed
            Exit Sub
        End If
    End If
    
    For b1 = 1 To 2
        If txtUsuario(b1).Text = "" Then
            txtUsuario(b1).Text = "-"
        End If
    Next b1
    
    
    
    If checkError = False Then
        If dtFecha(0) > Date Then
            checkError = True
            lUsuario(31).ForeColor = vbRed
        Else
            If cmbUser(2).Text = "" Then
                checkError = True
                lUsuario(14).ForeColor = vbRed
            Else
                If cmbUser(5).Text = "" Then
                    checkError = True
                    lUsuario(26).ForeColor = vbRed
                End If
            End If
        End If
    End If
End Sub

Private Sub crearCliente()

    Dim status As String
    Dim idEstado As String
    Dim idMunicipio As String
    Dim idEstadoNac As String
    Dim genero As String
    Dim cp As String
    Dim tel1 As String
    Dim tel2 As String
    Dim telAccdte As String
    Dim membresia As String
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    Dim membresiaCodigo As String
    Dim recomiendaId As Integer
    Dim recomiendaTipoid As Integer
    Dim recomiendaTipo As String
        
    genero = Left(cmbUser(2).Text, 1)
    If txtUsuario(3).Text = "" Then
        tel1 = "null"
    Else
        tel1 = txtUsuario(3).Text
    End If
    If txtUsuario(4).Text = "" Then
        tel2 = "null"
    Else
        tel2 = txtUsuario(4).Text
    End If
                    
    If tipoPersona = "PROVEEDOR" Or tipoPersona = "PROVEEDOR_G" Then
            
            sql1 = "INSERT INTO PERSONA (PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_FEC_NAC, " & _
            "PER_EMAIL, PER_FECHA_SISTEMA, PER_GENERO, PER_TEL1, PER_TEL2, PER_ALIAS) VALUES " & _
            "('" & txtUsuario(0).Text & "', '" & txtUsuario(1).Text & "', '" & txtUsuario(2).Text & "', '" & Format(dtFecha(0), "yyyy-MM-dd") & "', " & _
            "'" & txtUsuario(11).Text & "', now(), '" & genero & "', " & tel1 & ", " & tel2 & ", '" & txtUsuario(5).Text & "' )"
            con.Execute (sql1)
            
            sql1 = "select last_insert_id() perId"
            Set res1 = con.Execute(sql1)
            If Not res1.EOF Then
                perId = res1.Fields("perId")
            End If
            
            membresiaCodigo = perId
            
            sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA, PERTP_CODIGO_MEMBRESIA, " & _
            "PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_PERALTA_FECHA) " & _
            "VALUES " & _
            "(" & cmbUser(5).ItemData(cmbUser(5).ListIndex) & ", " & perId & ", now(), 'V', 'A', now(), " & _
            "'" & membresiaCodigo & "', " & _
            "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW())"
            con.Execute (sql1)
    
    Else
        If tipoPersona = "CLIENTE" Then
                
            recomiendaId = 0
            recomiendaTipoid = 0
            recomiendaTipo = ""
                                    
            If cmbUser(1).Text <> "" Then
                sql1 = "SELECT * FROM VIEW_PERSONA WHERE ID = '" & cmbUser(1).ItemData(cmbUser(1).ListIndex) & "' AND TIPO = 'CLIENTE'"
                Set res1 = con.Execute(sql1)
                If Not res1.EOF Then
                    recomiendaId = res1.Fields("perid")
                    recomiendaTipoid = res1.Fields("tipoId")
                    recomiendaTipo = res1.Fields("tipo_tipo")
                End If
            End If
            
            sql1 = "INSERT INTO PERSONA (PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_FEC_NAC, " & _
            "PER_EMAIL, PER_FECHA_SISTEMA, PER_GENERO, PER_TEL1, PER_TEL2) VALUES " & _
            "('" & txtUsuario(0).Text & "', '" & txtUsuario(1).Text & "', '" & txtUsuario(2).Text & "', '" & Format(dtFecha(0), "yyyy-MM-dd") & "', " & _
            "'" & txtUsuario(11).Text & "', now(), '" & genero & "', " & tel1 & ", " & tel2 & ")"
            con.Execute (sql1)
            
            sql1 = "select last_insert_id() perId"
            Set res1 = con.Execute(sql1)
            If Not res1.EOF Then
                perId = res1.Fields("perId")
            End If
            
            If txtUsuario(6).Text = "" Then
                membresiaCodigo = perId
            Else
                membresiaCodigo = txtUsuario(6).Text
            End If
            
            
            If recomiendaTipo <> "" Then
                sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA,  " & _
                "PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_PERALTA_FECHA, PERTP_MEMBRESIA, PERTP_CODIGO_MEMBRESIA, PERTP_RECOMENDADO_ID, PERTP_RECOMENDADO_TIPO_ID, PERTP_RECOMENDADO_TIPO, PERTP_RECOMENDADO_FECHA) " & _
                "VALUES " & _
                "(" & cmbUser(5).ItemData(cmbUser(5).ListIndex) & ", " & perId & ", now(), 'C', 'A', now(), " & _
                "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW(), '" & Left(cmbUser(0).Text, 1) & "', '" & membresiaCodigo & "', '" & recomiendaId & "', '" & recomiendaTipoid & "', '" & recomiendaTipo & "', now() )"
                con.Execute (sql1)
            Else
                sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA,  " & _
                "PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_PERALTA_FECHA, PERTP_MEMBRESIA, PERTP_CODIGO_MEMBRESIA) " & _
                "VALUES " & _
                "(" & cmbUser(5).ItemData(cmbUser(5).ListIndex) & ", " & perId & ", now(), 'C', 'A', now(), " & _
                "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW(), '" & Left(cmbUser(0).Text, 1) & "', '" & membresiaCodigo & "' )"
                con.Execute (sql1)
            End If
            BUSQ_Usuarios.buscarUsuario
        End If
    End If
    
    
    
    MsgBox "Información guardada.", vbInformation
    
    If tipoPersona = "PROVEEDOR" Then
        FRM_Productos.cmdProveed_Click
    Else
        If tipoPersona = "PROVEEDOR_G" Then
            FRM_Gastos.cargaProveedor
        End If
    End If

    Unload Me
    
    
    
End Sub


Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Or Index = 4 Then
        Call Numeros(KeyAscii)
    End If

     If mayus = True Then
        Call Mayusculas(KeyAscii)
     End If

End Sub
