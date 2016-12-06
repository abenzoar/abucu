VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_DatosSuc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la sucursal"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   Icon            =   "FRM_DatosSuc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Datos de la sucursal"
      TabPicture(0)   =   "FRM_DatosSuc.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lProd(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lProd(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lUsuario(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lUsuario(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lUsuario(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lUsuario(8)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lUsuario(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lUsuario(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lUsuario(120)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lUsuario(130)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lUsuario(21)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lProd(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lProd(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lProd(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lProd(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lProd(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lProd(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lProd(8)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbStatus"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lUsuario(25)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "iFoto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lProd(11)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lProd(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lProd(19)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lProd(20)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lProd(21)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "dtTime1(1)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cMd1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtProd(5)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtProd(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtProd(11)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtProd(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtProd(9)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtProd(14)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtProd(13)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtProd(12)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmbUser(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbUser(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmbUser(5)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtProd(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtProd(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtProd(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtProd(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtProd(6)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtProd(7)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtProd(8)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmBoton(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cmBoton(0)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "pFoto"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TimerFoto"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cmBoton(5)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cmBoton(4)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cmBoton(6)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtProd(15)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtProd(16)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "dtTime1(0)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmbUser(3)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "Conexión de la sucursal"
      TabPicture(1)   =   "FRM_DatosSuc.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmBoton(3)"
      Tab(1).Control(1)=   "cmBoton(2)"
      Tab(1).Control(2)=   "txtUsuario(4)"
      Tab(1).Control(3)=   "txtUsuario(3)"
      Tab(1).Control(4)=   "txtUsuario(2)"
      Tab(1).Control(5)=   "txtUsuario(1)"
      Tab(1).Control(6)=   "txtUsuario(0)"
      Tab(1).Control(7)=   "txtProd(92)"
      Tab(1).Control(8)=   "cmbUser(2)"
      Tab(1).Control(9)=   "MSFlexGrid1"
      Tab(1).Control(10)=   "lProd(10)"
      Tab(1).Control(11)=   "lUsuario(11)"
      Tab(1).Control(12)=   "lUsuario(4)"
      Tab(1).Control(13)=   "lUsuario(3)"
      Tab(1).Control(14)=   "lUsuario(2)"
      Tab(1).Control(15)=   "lUsuario(1)"
      Tab(1).Control(16)=   "lProd(9)"
      Tab(1).Control(17)=   "lUsuario(0)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Ticket"
      TabPicture(2)   =   "FRM_DatosSuc.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmBoton(8)"
      Tab(2).Control(1)=   "cmBoton(7)"
      Tab(2).Control(2)=   "chkInfo(5)"
      Tab(2).Control(3)=   "chkInfo(4)"
      Tab(2).Control(4)=   "chkInfo(3)"
      Tab(2).Control(5)=   "chkInfo(2)"
      Tab(2).Control(6)=   "chkInfo(1)"
      Tab(2).Control(7)=   "chkInfo(0)"
      Tab(2).Control(8)=   "Label1"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Configuración correo"
      TabPicture(3)   =   "FRM_DatosSuc.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmBoton(10)"
      Tab(3).Control(1)=   "cmBoton(9)"
      Tab(3).Control(2)=   "cmbMail(1)"
      Tab(3).Control(3)=   "cmbMail(0)"
      Tab(3).Control(4)=   "txtMail(5)"
      Tab(3).Control(5)=   "txtMail(4)"
      Tab(3).Control(6)=   "txtMail(3)"
      Tab(3).Control(7)=   "txtMail(2)"
      Tab(3).Control(8)=   "txtMail(1)"
      Tab(3).Control(9)=   "txtMail(0)"
      Tab(3).Control(10)=   "lUsuario(13)"
      Tab(3).Control(11)=   "lUsuario(12)"
      Tab(3).Control(12)=   "lProd(18)"
      Tab(3).Control(13)=   "lProd(17)"
      Tab(3).Control(14)=   "lProd(16)"
      Tab(3).Control(15)=   "lProd(15)"
      Tab(3).Control(16)=   "lProd(14)"
      Tab(3).Control(17)=   "lProd(13)"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Configuración general"
      TabPicture(4)   =   "FRM_DatosSuc.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkConfig(2)"
      Tab(4).Control(1)=   "chkConfig(1)"
      Tab(4).Control(2)=   "chkConfig(0)"
      Tab(4).ControlCount=   3
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
         Index           =   3
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   7680
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtTime1 
         Height          =   375
         Index           =   0
         Left            =   8280
         TabIndex        =   98
         Top             =   6840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   103415810
         CurrentDate     =   41250
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Notificar en el menu con ventana emergente cuando se realiza una asistencia en un equipo remoto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -74400
         TabIndex        =   97
         Top             =   3240
         Width           =   4815
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Solicitar contraseña de administrador al realizar descuentos en operaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74400
         TabIndex        =   96
         Top             =   2400
         Width           =   3735
      End
      Begin VB.CheckBox chkConfig 
         Caption         =   "Tipo de letra en ""Mayúsculas"" al guardar información"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -74400
         TabIndex        =   95
         Top             =   1560
         Width           =   3735
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
         Index           =   10
         Left            =   -67440
         Picture         =   "FRM_DatosSuc.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
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
         Index           =   9
         Left            =   -69240
         Picture         =   "FRM_DatosSuc.frx":1220
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   7200
         Width           =   1695
      End
      Begin VB.ComboBox cmbMail 
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   7560
         Width           =   2535
      End
      Begin VB.ComboBox cmbMail 
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   6600
         Width           =   2535
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   87
         Top             =   5760
         Width           =   2655
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   85
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   83
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   81
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   79
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtMail 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   77
         Top             =   1560
         Width           =   3975
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
         Index           =   8
         Left            =   -72480
         Picture         =   "FRM_DatosSuc.frx":1AEA
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   7440
         Width           =   1695
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
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
         Index           =   7
         Left            =   -74280
         Picture         =   "FRM_DatosSuc.frx":23B4
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   7440
         Width           =   1695
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Incluir código de barra del ticket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -74280
         TabIndex        =   73
         Top             =   3720
         Width           =   4215
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Incluir información adicional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74280
         TabIndex        =   72
         Top             =   3240
         Width           =   4455
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Incluir teléfonos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74280
         TabIndex        =   71
         Top             =   2760
         Width           =   3255
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Incluir domicilio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74280
         TabIndex        =   70
         Top             =   2280
         Width           =   3255
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Incluir logotipo del negocio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74280
         TabIndex        =   69
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CheckBox chkInfo 
         Caption         =   "Imprimir ticket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74280
         TabIndex        =   68
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtProd 
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
         Index           =   16
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   15
         Top             =   6720
         Width           =   2175
      End
      Begin VB.TextBox txtProd 
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
         Index           =   15
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   16
         Top             =   7440
         Width           =   2175
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buscar imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   2760
         Picture         =   "FRM_DatosSuc.frx":2C7E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tomar foto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   2760
         Picture         =   "FRM_DatosSuc.frx":3548
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Iniciar cámara"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   2760
         Picture         =   "FRM_DatosSuc.frx":3E12
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer TimerFoto 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   3240
         Top             =   3240
      End
      Begin VB.PictureBox pFoto 
         BackColor       =   &H00E0E0E0&
         Height          =   2175
         Left            =   240
         ScaleHeight     =   2115
         ScaleWidth      =   2355
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2415
         Begin VB.Label lCamara 
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciando cámara"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   64
            Top             =   3600
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   -74760
         Picture         =   "FRM_DatosSuc.frx":46DC
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   7560
         Width           =   1695
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
         Height          =   735
         Index           =   2
         Left            =   -72960
         Picture         =   "FRM_DatosSuc.frx":4FA6
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   7560
         Width           =   1695
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
         Left            =   -68520
         MaxLength       =   120
         TabIndex        =   53
         Top             =   6360
         Width           =   1455
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
         Left            =   -66240
         MaxLength       =   15
         TabIndex        =   52
         Top             =   5640
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
         Left            =   -71520
         MaxLength       =   75
         TabIndex        =   51
         Top             =   5640
         Width           =   2895
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
         Left            =   -71520
         MaxLength       =   75
         TabIndex        =   50
         Top             =   6720
         Width           =   1575
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
         Left            =   -68520
         MaxLength       =   6
         TabIndex        =   49
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox txtProd 
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
         Index           =   92
         Left            =   -74760
         MaxLength       =   65
         TabIndex        =   46
         Top             =   5640
         Width           =   3135
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
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   6360
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   44
         Top             =   1080
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FormatString    =   $"FRM_DatosSuc.frx":5870
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
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
         Left            =   240
         Picture         =   "FRM_DatosSuc.frx":590D
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   7800
         Width           =   1695
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
         Left            =   2040
         Picture         =   "FRM_DatosSuc.frx":61D7
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   8
         Left            =   8280
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   4560
         Width           =   4815
      End
      Begin VB.TextBox txtProd 
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
         Index           =   7
         Left            =   8280
         MaxLength       =   250
         TabIndex        =   21
         Top             =   3840
         Width           =   4815
      End
      Begin VB.TextBox txtProd 
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
         Left            =   8280
         MaxLength       =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtProd 
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
         Left            =   8280
         MaxLength       =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtProd 
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
         Left            =   8280
         MaxLength       =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtProd 
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
         Left            =   8280
         MaxLength       =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtProd 
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
         TabIndex        =   0
         Top             =   960
         Width           =   4095
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
         TabIndex        =   1
         Top             =   1680
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
         Index           =   1
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
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
         Index           =   0
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtProd 
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
         Index           =   12
         Left            =   4560
         MaxLength       =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   3495
      End
      Begin VB.TextBox txtProd 
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
         Index           =   13
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   13
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
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
         Index           =   14
         Left            =   4560
         MaxLength       =   15
         TabIndex        =   14
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
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
         Index           =   9
         Left            =   4560
         MaxLength       =   75
         TabIndex        =   9
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtProd 
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
         Index           =   10
         Left            =   4560
         MaxLength       =   75
         TabIndex        =   10
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtProd 
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
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   11
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtProd 
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
         Left            =   240
         MaxLength       =   13
         TabIndex        =   3
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtProd 
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
         MaxLength       =   75
         TabIndex        =   2
         Top             =   2400
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog cMd1 
         Left            =   240
         Top             =   6480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtTime1 
         Height          =   375
         Index           =   1
         Left            =   8280
         TabIndex        =   99
         Top             =   7680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96600066
         CurrentDate     =   41250
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Día de cierre "
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
         Index           =   21
         Left            =   10080
         TabIndex        =   103
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Horario Cierre"
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
         Index           =   20
         Left            =   8280
         TabIndex        =   101
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Horario Inicio"
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
         Index           =   19
         Left            =   8280
         TabIndex        =   100
         Top             =   6600
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "SSL"
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
         Index           =   13
         Left            =   -74640
         TabIndex        =   92
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Autentificación"
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
         Index           =   12
         Left            =   -74640
         TabIndex        =   90
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto SMTP"
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
         Index           =   18
         Left            =   -74640
         TabIndex        =   88
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio POP"
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
         Index           =   17
         Left            =   -74640
         TabIndex        =   86
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio SMTP"
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
         Index           =   16
         Left            =   -74640
         TabIndex        =   84
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         Index           =   15
         Left            =   -74640
         TabIndex        =   82
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   -74640
         TabIndex        =   80
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta de correo"
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
         Index           =   13
         Left            =   -74640
         TabIndex        =   78
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Información relacionada con la impresión del ticket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   74
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label lProd 
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
         Index           =   12
         Left            =   4560
         TabIndex        =   67
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label lProd 
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
         Index           =   11
         Left            =   4560
         TabIndex        =   66
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Image iFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Foto"
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
         Index           =   25
         Left            =   240
         TabIndex        =   65
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lbStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus:"
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
         Left            =   3960
         TabIndex        =   62
         Top             =   8280
         Width           =   4695
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Conexiones a sucursales"
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
         Index           =   10
         Left            =   -74760
         TabIndex        =   59
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   -68520
         TabIndex        =   58
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Db Nombre"
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
         Left            =   -66240
         TabIndex        =   57
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor"
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
         Left            =   -71520
         TabIndex        =   56
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto"
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
         Left            =   -71520
         TabIndex        =   55
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   -68520
         TabIndex        =   54
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Sucursal *"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   48
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo *"
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
         Left            =   -74760
         TabIndex        =   47
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Información adicional"
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
         Left            =   8280
         TabIndex        =   41
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Eslogan *"
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
         Left            =   8280
         TabIndex        =   40
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Página web"
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
         Left            =   8280
         TabIndex        =   39
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Twitter"
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
         Left            =   8280
         TabIndex        =   38
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Facebook"
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
         Left            =   8280
         TabIndex        =   37
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lProd 
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
         Index           =   1
         Left            =   8280
         TabIndex        =   36
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la Sucursal *"
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
         TabIndex        =   35
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo *"
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
         Index           =   21
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio *"
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
         Index           =   130
         Left            =   4560
         TabIndex        =   33
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado *"
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
         Index           =   120
         Left            =   4560
         TabIndex        =   32
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Calle *"
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
         Index           =   10
         Left            =   4560
         TabIndex        =   31
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Número exterior *"
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
         Index           =   9
         Left            =   4560
         TabIndex        =   30
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Número interior *"
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
         Left            =   4560
         TabIndex        =   29
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad *"
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
         Left            =   4560
         TabIndex        =   28
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia *"
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
         Left            =   4560
         TabIndex        =   27
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Código postal *"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "RFC"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Razón social"
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
         TabIndex        =   24
         Top             =   2160
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FRM_DatosSuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim sucId As Long
Dim errorDat As Boolean

Private Sub buscarImagen()
    cMd1.DialogTitle = "Buscando imagen..."
    cMd1.Filter = "Archivos de Imagenes|*.jpg*||*.bmp*||*.gif*||*.wmf*||*.emf*|"
    cMd1.FileName = ""
    cMd1.ShowOpen
    If cMd1.FileName <> "" Then
        guardarImagen
    End If
End Sub
Private Sub guardarImagen()
    With cMd1
        iFoto.Visible = True
        iFoto.Picture = LoadPicture(.FileName)
        
    End With
End Sub

Private Sub cargaEstados()
    
    SQL1 = ("SELECT CT_EST_ID, CT_EST_NOMBRE FROM CAT_ESTADO ORDER BY CT_eST_NOMBRE")
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        cmbUser(0).AddItem RES1.Fields("CT_EST_NOMBRE")
        cmbUser(0).ItemData(cmbUser(0).ListCount - 1) = RES1.Fields("CT_EST_ID")
        RES1.MoveNext
    Loop
    
End Sub
Sub STOPCAM()
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
TimerFoto.Enabled = False
pFoto.Visible = False
'cTomarFoto.Caption = "Tomar foto"
End Sub
Sub STARTCAM()
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hwnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0
TimerFoto.Enabled = True
End Sub

Private Sub chkInfo_Click(Index As Integer)

    Select Case Index
        Case 0:
            checkInfoTicket
    End Select

End Sub
Private Sub checkInfoTicket()
    If chkInfo(0).value = 0 Then
        For b1 = 1 To 5
            chkInfo(b1).Enabled = False
        Next b1
    Else
        For b1 = 1 To 5
            chkInfo(b1).Enabled = True
        Next b1
    End If
End Sub
Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 1
            Unload Me
        Case 10
            Unload Me
        Case 6
        buscarImagen
        Case 0
            checkDatos
            If errorDat = False Then
                crearSucursal
            Else
                MsgBox "Falta información. Por favor verifique.", vbInformation
            End If
        Case 7
            checkDatos
            If errorDat = False Then
                crearSucursal
            Else
                MsgBox "Falta información. Por favor verifique.", vbInformation
            End If
        Case 9
            checkDatosMail
            If errorDat = False Then
                crearMailInfo
            Else
                MsgBox "Falta información. Por favor verifique.", vbInformation
            End If
        Case 4
            If cmBoton(4).Caption = "Tomar foto" Then
                iFoto.Visible = False
                pFoto.Visible = True
                cmBoton_Click (5)
                cmBoton(4).Caption = "Capturar"
            Else
                iFoto.Visible = True
                iFoto.Picture = pFoto.Picture
                cmBoton(4).Caption = "Tomar foto"
                STOPCAM
                checarCarpetaTemp
                SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempSuc.dat")
            End If
        Case 5
            STARTCAM
    End Select

End Sub
Private Sub crearMailInfo()
    
    If lbStatus.Caption = "Editando datos" Then
        Dim auten As String
        Dim ssl As String
        If cmbMail(0).Text = "Si" Then
            auten = "True"
        Else
            auten = "False"
        End If
        If cmbMail(1).Text = "Si" Then
            ssl = "True"
        Else
            ssl = "False"
        End If
    
        SQL1 = "UPDATE SUCURSAL SET SUC_MAIL_SMTP = '" & txtMail(3).Text & "', " & _
        "SUC_MAIL_POP = '" & txtMail(4).Text & "', SUC_MAIL_CORREO = '" & txtMail(0).Text & "', SUC_MAIL_USUARIO = '" & txtMail(1).Text & "', " & _
        "SUC_MAIL_PASS = '" & txtMail(2).Text & "', SUC_MAIL_PUERTO='" & txtMail(5).Text & "', " & _
        "SUC_MAIL_AUTEN = '" & auten & "', SUC_MAIL_SSL = '" & ssl & "' " & _
        "WHERE SUC_ID = " & sucId & ""
        con.Execute (SQL1)
        MsgBox "Información guardada.", vbInformation
        lbStatus.Caption = "Editando datos"
    Else
        MsgBox "No se puede agregar la información de confirguración de mail.", vbInformation
    End If

End Sub

Private Sub checkDatosMail()
    errorDat = False
    For b1 = 0 To 5
        If txtMail(b1).Text = "" Then
            errorDat = True
            Exit Sub
        End If
    Next b1
    For b1 = 0 To 1
        If cmbMail(b1).Text = "" Then
            errorDat = True
            Exit Sub
        End If
    Next b1
End Sub
Private Sub checkDatos()
    errorDat = False
    If txtProd(0).Text = "" Then
        errorDat = True
        Exit Sub
    End If
    For b1 = 9 To 14
        If txtProd(b1).Text = "" Then
            errorDat = True
        Exit Sub
        End If
    Next b1
    
End Sub
Private Sub crearSucursal()
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    Dim tipoSuc As String
    Dim idEstado As String
    Dim idMunicipio As String
    Dim cp As String
    Dim tel1 As String
    Dim tel2 As String
    
    If cmbUser(5).Text <> "" Then
        tipoSuc = Left(cmbUser(5).Text, 1)
    End If
    
    If cmbUser(0).Text = "" Then
        idEstado = "null"
    Else
        idEstado = cmbUser(0).ItemData(cmbUser(0).ListIndex)
    End If
    
    If cmbUser(1).Text = "" Then
        idMunicipio = "null"
    Else
        idMunicipio = cmbUser(1).ItemData(cmbUser(1).ListIndex)
    End If
    If txtProd(11).Text = "" Then
        cp = "null"
    Else
        cp = txtProd(11).Text
    End If
        
    If txtProd(16).Text = "" Then
        tel1 = "null"
    Else
        tel1 = txtProd(16).Text
    End If
        
    If txtProd(15).Text = "" Then
        tel2 = "null"
    Else
        tel2 = txtProd(15).Text
    End If
        
    If lbStatus.Caption = "Agregando datos" Then
        'agregar
        SQL1 = "INSERT INTO SUCURSAL (SUC_NOMBRE, SUC_TIPO, SUC_RAZON_SOCIAL, SUC_RFC, " & _
        "SUC_DIR_EST_ID, SUC_DIR_MUN_ID, SUC_dIR_CIUDAD, SUC_DIR_COLONIA, SUC_DIR_CP, SUC_DIR_CALLE, " & _
        "SUC_DIR_NUM_EXT, SUC_DIR_NUM_INT, SUC_EMAIL, SUC_FACEBOOK, SUC_TWITTER, SUC_PAGINA_WEB, " & _
        "SUC_SLOGAN, SUC_INFORMACION, SUC_TEL1, SUC_TEL2, SUC_ESTATUSTICKET, " & _
        "SUC_TICKETLOGO, SUC_TICKETDOMICILIO, SUC_TICKETINFOADICIONAL, SUC_TICKETCODIGOBARRA, SUC_TICKETFON, SUC_LOCAL, suc_horaentrada, suc_horasalida, suc_dia_cierre) VALUES " & _
        "('" & txtProd(0).Text & "', '" & tipoSuc & "', '" & txtProd(5).Text & "', " & _
        "'" & txtProd(4).Text & "', '" & idEstado & "', '" & idMunicipio & "', " & _
        "'" & txtProd(9).Text & "', '" & txtProd(10).Text & "', '" & cp & "', " & _
        "'" & txtProd(12).Text & "', '" & txtProd(13).Text & "', '" & txtProd(14).Text & "', " & _
        "'" & txtProd(1).Text & "', '" & txtProd(2).Text & "', '" & txtProd(3).Text & "', " & _
        "'" & txtProd(6).Text & "', '" & txtProd(7).Text & "', '" & txtProd(8).Text & "', " & _
        "" & tel1 & ", " & tel2 & ", " & _
        "'" & chkInfo(0).value & "', '" & chkInfo(1).value & "', '" & chkInfo(2).value & "', " & _
        "'" & chkInfo(4).value & "', '" & chkInfo(5).value & "', '" & chkInfo(3).value & "', 'S', '" & Format(dtTime1(0).value, "Short Time") & "', '" & Format(dtTime1(1).value, "Short Time") & "', '" & Left(cmbUser(3).Text, 1) & "')"
        MsgBox SQL1
        con.Execute (SQL1)
        
        SQL1 = "select last_insert_id() sucId"
        Set RES1 = con.Execute(SQL1)
        If Not RES1.EOF Then
            sucId = RES1.Fields("SUCID")
        End If
        
    Else
        SQL1 = "UPDATE SUCURSAL SET SUC_NOMBRE = '" & txtProd(0).Text & "', " & _
        "SUC_TIPO = '" & tipoSuc & "', SUC_RAZON_SOCIAL = '" & txtProd(5).Text & "', SUC_RFC = '" & txtProd(4).Text & "', " & _
        "SUC_DIR_EST_ID = '" & idEstado & "', SUC_DIR_MUN_ID='" & idMunicipio & "', " & _
        "SUC_dIR_CIUDAD = '" & txtProd(9).Text & "', SUC_DIR_COLONIA = '" & txtProd(10).Text & "', " & _
        "SUC_DIR_CP = '" & cp & "', SUC_DIR_CALLE ='" & txtProd(12).Text & "', " & _
        "SUC_DIR_NUM_EXT='" & txtProd(13).Text & "', SUC_DIR_NUM_INT='" & txtProd(14).Text & "', " & _
        "SUC_EMAIL='" & txtProd(1).Text & "', SUC_FACEBOOK='" & txtProd(2).Text & "', " & _
        "SUC_TWITTER='" & txtProd(3).Text & "', SUC_PAGINA_WEB='" & txtProd(6).Text & "', " & _
        "SUC_SLOGAN='" & txtProd(7).Text & "', SUC_INFORMACION='" & txtProd(8).Text & "', " & _
        "SUC_TEL1=" & tel1 & ", SUC_TEL2=" & tel2 & ", " & _
        "SUC_ESTATUSTICKET='" & chkInfo(0).value & "', " & _
        "SUC_TICKETLOGO='" & chkInfo(1).value & "', " & _
        "SUC_TICKETDOMICILIO='" & chkInfo(2).value & "', " & _
        "SUC_TICKETINFOADICIONAL='" & chkInfo(4).value & "', " & _
        "SUC_TICKETCODIGOBARRA='" & chkInfo(5).value & "', " & _
        "SUC_TICKETFON='" & chkInfo(3).value & "', SUC_HORAENTRADA='" & Format(dtTime1(0).value, "Short Time") & "', suc_horasalida='" & Format(dtTime1(1).value, "Short Time") & "', suc_dia_cierre='" & Left(cmbUser(3).Text, 1) & "' " & _
        "WHERE SUC_ID = " & sucId & ""
        ''MsgBox SQL1
        con.Execute (SQL1)
    End If
    
    SQ1 = ""
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempSuc.dat")
            res.Open "SELECT * FROM SUCURSAL WHERE SUC_ID = '" & sucId & "'", con, adOpenStatic, adLockOptimistic
            If res.EOF Then
            Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempSuc.dat")
                res.Fields("SUC_FOTO") = Imagen1.Read
                res.Update
            End If
    End If
    MsgBox "Información guardada.", vbInformation
    lbStatus.Caption = "Editando datos"
End Sub
Private Sub Salir()
End Sub
Private Sub cmbUser_Click(Index As Integer)
Select Case Index
    Case 0:
    
    cmbUser(1).Clear
    
    SQL1 = ("SELECT CTMUN_ID, CTMUN_NOMBRE FROM CAT_MUNICIPIO WHERE CTMUN_EST_ID = " & cmbUser(0).ItemData(cmbUser(0).ListIndex) & "")
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        cmbUser(1).AddItem RES1.Fields("CTMUN_NOMBRE")
        cmbUser(1).ItemData(cmbUser(1).ListCount - 1) = RES1.Fields("CTMUN_ID")
        RES1.MoveNext
    Loop
        

End Select


End Sub

Private Sub Form_Load()

    SSTab1.Tab = 0
    cargaEstados
    cmbUser(5).Clear
    cmbUser(5).AddItem "SUCURSAL"
    cmbUser(5).AddItem "MATRIZ"
    cmbUser(2).Clear
    cmbUser(2).AddItem "SUCURSAL"
    cmbUser(2).AddItem "MATRIZ"
    iFoto.Picture = LoadPicture("")
    pFoto.Visible = False
    cmbMail(0).AddItem "Si"
    cmbMail(0).AddItem "No"
    cmbMail(1).AddItem "Si"
    cmbMail(1).AddItem "No"
    cmbUser(3).AddItem "Mismo día"
    cmbUser(3).AddItem "Día siguiente"
    checkSucursal
    errorDat = False
    checkInfoTicket
    
    
End Sub
Private Sub checkSucursal()
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
'    pFoto.Visible = False
    iFoto.Visible = True

    SQL1 = "SELECT SUC_ID, SUC_NOMBRE, SUC_TIPO, SUC_RAZON_SOCIAL, SUC_RFC, " & _
        "SUC_DIR_EST_ID, SUC_DIR_MUN_ID, SUC_dIR_CIUDAD, SUC_DIR_COLONIA, SUC_DIR_CP, SUC_DIR_CALLE, " & _
        "SUC_DIR_NUM_EXT, SUC_DIR_NUM_INT, SUC_EMAIL, SUC_FACEBOOK, SUC_TWITTER, SUC_PAGINA_WEB, " & _
        "SUC_SLOGAN, SUC_INFORMACION, SUC_TEL1, SUC_TEL2, SUC_FOTO, SUC_ESTATUSTICKET, " & _
        "SUC_TICKETLOGO, SUC_TICKETDOMICILIO, SUC_TICKETINFOADICIONAL, SUC_TICKETCODIGOBARRA, SUC_TICKETFON,  " & _
        "SUC_MAIL_SMTP, SUC_MAIL_POP, SUC_MAIL_CORREO, SUC_MAIL_USUARIO, SUC_MAIL_PASS, SUC_MAIL_PUERTO, SUC_MAIL_AUTEN, SUC_MAIL_SSL, SUC_MAIL_NOMBRE, " & _
        "SUC_HORAENTRADA, SUC_HORASALIDA, IF(SUC_DIA_CIERRE='D', 'Día siguiente', 'Mismo día') DIA_CIERRE FROM SUCURSAL"
        Set RES2 = con.Execute(SQL1)
    
    If Not RES2.EOF Then
        lbStatus.Caption = "Editando datos"
        
        chkInfo(0).value = RES2.Fields("SUC_ESTATUSTICKET")
        chkInfo(1).value = RES2.Fields("SUC_TICKETLOGO")
        chkInfo(2).value = RES2.Fields("SUC_TICKETDOMICILIO")
        chkInfo(3).value = RES2.Fields("SUC_TICKETFON")
        chkInfo(4).value = RES2.Fields("SUC_TICKETINFOADICIONAL")
        chkInfo(5).value = RES2.Fields("SUC_TICKETCODIGOBARRA")
        
        If IsNull(RES2.Fields("SUC_DIR_EST_ID")) Then
        Else
            cmbUser(0).ListIndex = (RES2.Fields("SUC_DIR_EST_ID") - 1)
        End If
        If IsNull(RES2.Fields("SUC_DIR_MUN_ID")) Then
        Else
            For b1 = 0 To cmbUser(1).ListCount - 1
                If cmbUser(1).ItemData(b1) = RES2.Fields("SUC_DIR_MUN_ID") Then
                    cmbUser(1).ListIndex = b1
                    Exit For
                End If
            Next b1
        End If
        
        txtProd(0).Text = RES2.Fields("SUC_NOMBRE")
        'txtProd(0).Text = RES1.Fields("SUC_TIPO")
        If RES2.Fields("SUC_TIPO") = "M" Then
            cmbUser(5).ListIndex = 1
        Else
            cmbUser(5).ListIndex = 0
        End If
        txtProd(5).Text = RES2.Fields("SUC_RAZON_SOCIAL")
        txtProd(4).Text = RES2.Fields("SUC_RFC")
        txtProd(9).Text = RES2.Fields("SUC_DIR_CIUDAD")
        txtProd(10).Text = RES2.Fields("SUC_DIR_COLONIA")
        txtProd(11).Text = RES2.Fields("SUC_DIR_CP")
        txtProd(12).Text = RES2.Fields("SUC_DIR_CALLE")
        txtProd(13).Text = RES2.Fields("SUC_dIR_NUM_EXT")
        txtProd(14).Text = RES2.Fields("SUC_dIR_NUM_INT")
        txtProd(1).Text = RES2.Fields("SUC_EMAIL")
        txtProd(2).Text = RES2.Fields("SUC_FACEBOOK")
        txtProd(3).Text = RES2.Fields("SUC_TWITTER")
        txtProd(6).Text = RES2.Fields("SUC_PAGINA_WEB")
        txtProd(7).Text = RES2.Fields("SUC_SLOGAN")
        txtProd(8).Text = RES2.Fields("SUC_INFORMACION")
        txtProd(16).Text = RES2.Fields("SUC_TEL1") & ""
        txtProd(15).Text = RES2.Fields("SUC_TEL2") & ""
        
        dtTime1(0).value = RES2.Fields("suc_HORAENTRADA")
        dtTime1(1).value = RES2.Fields("suc_HORASALIDA")
        cmbUser(3).Text = RES2.Fields("DIA_CIERRE")
        
        
        
        txtMail(0).Text = RES2.Fields("SUC_MAIL_CORREO") & ""
        txtMail(1).Text = RES2.Fields("SUC_MAIL_USUARIO") & ""
        txtMail(2).Text = RES2.Fields("SUC_MAIL_PASS") & ""
        txtMail(3).Text = RES2.Fields("SUC_MAIL_SMTP") & ""
        txtMail(4).Text = RES2.Fields("SUC_MAIL_POP") & ""
        txtMail(5).Text = RES2.Fields("SUC_MAIL_PUERTO") & ""
        
        If RES2.Fields("SUC_MAIL_AUTEN") = "True" Then
            cmbMail(0).Text = "Si"
        Else
            cmbMail(0).Text = "No"
        End If
        If RES2.Fields("SUC_MAIL_SSL") = "True" Then
            cmbMail(1).Text = "Si"
        Else
            cmbMail(1).Text = "No"
        End If
        
        
        sucId = RES2.Fields("SUC_ID")
        
        If IsNull(RES2.Fields("SUC_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES2.Fields("SUC_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempSuc.dat", adSaveCreateOverWrite
            Imagen1.Close
            iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempSuc.dat")
        Else
            iFoto.Picture = LoadPicture("")
        End If
        SSTab1.TabEnabled(3) = True
        
    Else
        lbStatus.Caption = "Agregando datos"
        SSTab1.TabEnabled(3) = False
    End If
    checkInfoTicket
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
        a = MsgBox("Saldra sin guardar la información. ¿Cancelar?", vbYesNo + vbQuestion)
        If a = vbYes Then
            Cancel = 0
        Else
            Cancel = 1
        End If
        
End Sub

Private Sub txtMail_Change(Index As Integer)
    If Index = 2 Then
        txtMail(Index).PasswordChar = "*"
        txtMail(Index).FontBold = True
        txtMail(Index).FontSize = 16
    End If

End Sub

Private Sub txtMail_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        txtMail(Index).PasswordChar = "*"
        txtMail(Index).FontBold = True
        txtMail(Index).FontSize = 16
    End If

End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 Or Index = 15 Then
        Call Numeros(KeyAscii)
    End If
End Sub
