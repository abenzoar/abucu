VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   9375
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   17295
   Icon            =   "FRM_Clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   17295
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   697
      TabCaption(0)   =   "  Lista de clientes"
      TabPicture(0)   =   "FRM_Clientes.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2(1)"
      Tab(0).Control(1)=   "Shape1(7)"
      Tab(0).Control(2)=   "lInfo(10)"
      Tab(0).Control(3)=   "lBus(3)"
      Tab(0).Control(4)=   "lBus(2)"
      Tab(0).Control(5)=   "lBus(1)"
      Tab(0).Control(6)=   "lBus(0)"
      Tab(0).Control(7)=   "lBus(4)"
      Tab(0).Control(8)=   "lProd(16)"
      Tab(0).Control(9)=   "Shape1(6)"
      Tab(0).Control(10)=   "Borde(15)"
      Tab(0).Control(11)=   "Borde(0)"
      Tab(0).Control(12)=   "Borde(1)"
      Tab(0).Control(13)=   "Borde(2)"
      Tab(0).Control(14)=   "fotoUser"
      Tab(0).Control(15)=   "lUsuario(20)"
      Tab(0).Control(16)=   "lBus(5)"
      Tab(0).Control(17)=   "Borde(3)"
      Tab(0).Control(18)=   "Borde(29)"
      Tab(0).Control(19)=   "ListaUsers"
      Tab(0).Control(20)=   "textBus(3)"
      Tab(0).Control(21)=   "textBus(2)"
      Tab(0).Control(22)=   "textBus(1)"
      Tab(0).Control(23)=   "textBus(0)"
      Tab(0).Control(24)=   "textBus(4)"
      Tab(0).Control(25)=   "timeCarga"
      Tab(0).Control(26)=   "cmbUser(8)"
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "  Datos generales"
      TabPicture(1)   =   "FRM_Clientes.frx":0E64
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "iFoto"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lUsuario(25)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lUsuario(13)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lUsuario(12)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lUsuario(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lUsuario(130)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lUsuario(120)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lUsuario(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lUsuario(9)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lUsuario(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lUsuario(7)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lUsuario(6)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lUsuario(5)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lUsuario(44)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lUsuario(3)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lUsuario(31)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lUsuario(2)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lUsuario(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lUsuario(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "lUsuario(14)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lUsuario(15)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lUsuario(16)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lUsuario(4)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lUsuario(17)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lUsuario(18)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "lUsuario(26)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "lUsuario(19)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lUsuario(21)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Borde(4)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Borde(5)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Borde(6)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Borde(7)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Borde(8)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Borde(9)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Borde(10)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Borde(11)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Borde(12)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Borde(13)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Borde(14)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Borde(16)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Borde(17)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Borde(18)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Borde(19)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Borde(20)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Borde(21)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Borde(22)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Borde(23)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Borde(24)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Borde(25)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Borde(26)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Borde(27)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Borde(28)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Shape1(1)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lbStatus"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "dtFecha(1)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "dtFecha(0)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "cMd1"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "cmdTipoUsuario"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "cmBoton(2)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "cmBoton(1)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "cmBoton(0)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "txtUsuario(13)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "txtUsuario(12)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "txtUsuario(11)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "cmbUser(1)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "cmbUser(0)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "txtUsuario(10)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "txtUsuario(9)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "txtUsuario(8)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "txtUsuario(7)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "txtUsuario(6)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "txtUsuario(5)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "txtUsuario(4)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "txtUsuario(3)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "txtUsuario(2)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "txtUsuario(1)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "txtUsuario(0)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "cmbUser(2)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "txtUsuario(14)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "cmbUser(3)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "cmbUser(4)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "txtUsuario(15)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "cmBoton(3)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "cmBoton(4)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "cmBoton(5)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "TimerFoto"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "pFoto"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "cmbUser(5)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "cmbUser(6)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "Check1"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "cmbUser(7)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "cmBoton(6)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).ControlCount=   94
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
         Index           =   8
         Left            =   -65640
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Timer timeCarga 
         Interval        =   25
         Left            =   -62280
         Top             =   720
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Huella digital"
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
         Index           =   6
         Left            =   5640
         Picture         =   "FRM_Clientes.frx":13FE
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   7920
         Width           =   1695
      End
      Begin VB.TextBox textBus 
         Height          =   285
         Index           =   4
         Left            =   -61440
         TabIndex        =   75
         Text            =   "50"
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
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
         Index           =   7
         Left            =   14160
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   7200
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Relacionar venta de Servicio/Producto para obtener membresía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14160
         TabIndex        =   72
         Top             =   6120
         Visible         =   0   'False
         Width           =   3615
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
         Index           =   6
         Left            =   14160
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   5640
         Visible         =   0   'False
         Width           =   2895
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5400
         Width           =   2895
      End
      Begin VB.PictureBox pFoto 
         BackColor       =   &H00E0E0E0&
         Height          =   2775
         Left            =   11880
         ScaleHeight     =   2715
         ScaleWidth      =   2355
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
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
            TabIndex        =   68
            Top             =   3600
            Width           =   2415
         End
      End
      Begin VB.Timer TimerFoto 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   14880
         Top             =   120
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
         Left            =   14520
         Picture         =   "FRM_Clientes.frx":1CC8
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3000
         Visible         =   0   'False
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
         Left            =   14520
         Picture         =   "FRM_Clientes.frx":2592
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar y agregar cliente"
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
         Index           =   3
         Left            =   3840
         Picture         =   "FRM_Clientes.frx":2E5C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7920
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
         Height          =   3015
         Index           =   15
         Left            =   11760
         MaxLength       =   2450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   5160
         Width           =   4575
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
         Index           =   4
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4680
         Width           =   2895
      End
      Begin VB.ComboBox cmbUser 
         Enabled         =   0   'False
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3120
         Width           =   3375
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
         Index           =   14
         Left            =   8040
         MaxLength       =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   2895
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5280
         Width           =   3375
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
         Left            =   360
         MaxLength       =   50
         TabIndex        =   0
         Top             =   960
         Width           =   3495
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
         Left            =   360
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1680
         Width           =   3495
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
         Left            =   360
         MaxLength       =   50
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
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
         Left            =   360
         MaxLength       =   13
         TabIndex        =   4
         Top             =   3840
         Width           =   2535
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
         Left            =   360
         MaxLength       =   18
         TabIndex        =   5
         Top             =   4560
         Width           =   2535
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
         Left            =   4200
         MaxLength       =   6
         TabIndex        =   11
         Top             =   3840
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
         Index           =   6
         Left            =   4200
         MaxLength       =   75
         TabIndex        =   10
         Top             =   3120
         Width           =   3495
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
         Index           =   7
         Left            =   4200
         MaxLength       =   75
         TabIndex        =   9
         Top             =   2400
         Width           =   3495
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
         Index           =   8
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   14
         Top             =   6000
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
         Index           =   9
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   13
         Top             =   5280
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
         Index           =   10
         Left            =   4200
         MaxLength       =   120
         TabIndex        =   12
         Top             =   4560
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
         Index           =   0
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3375
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
         Left            =   8040
         MaxLength       =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   3495
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
         Index           =   12
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
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
         Index           =   13
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Width           =   1935
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
         Picture         =   "FRM_Clientes.frx":3726
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7920
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
         Picture         =   "FRM_Clientes.frx":3FF0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7920
         Width           =   1695
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
         Index           =   2
         Left            =   14520
         Picture         =   "FRM_Clientes.frx":48BA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -74640
         TabIndex        =   34
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -72600
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -70440
         TabIndex        =   32
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -68400
         TabIndex        =   31
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton cmdTipoUsuario 
         Caption         =   "Command1"
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   7560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog cMd1 
         Left            =   2160
         Top             =   6960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   3120
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
         Format          =   112328705
         CurrentDate     =   40783
      End
      Begin MSFlexGridLib.MSFlexGrid ListaUsers 
         Height          =   6855
         Left            =   -74640
         TabIndex        =   35
         Top             =   1680
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12091
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"FRM_Clientes.frx":5184
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Index           =   1
         Left            =   14160
         TabIndex        =   20
         Top             =   4920
         Visible         =   0   'False
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
         Format          =   112590849
         CurrentDate     =   40783
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   2835
         Index           =   29
         Left            =   -60480
         Top             =   1680
         Width           =   2205
      End
      Begin VB.Label lbStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   9000
         Width           =   4695
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   120
         Top             =   9000
         Width           =   15615
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   2835
         Index           =   28
         Left            =   11880
         Top             =   840
         Width           =   2445
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   3075
         Index           =   27
         Left            =   11760
         Top             =   5160
         Width           =   4605
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   26
         Left            =   4200
         Top             =   6000
         Width           =   1725
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   25
         Left            =   4200
         Top             =   5280
         Width           =   1725
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   24
         Left            =   4200
         Top             =   3840
         Width           =   1605
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   23
         Left            =   8040
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   22
         Left            =   8040
         Top             =   960
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   21
         Left            =   360
         Top             =   4560
         Width           =   2565
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   20
         Left            =   360
         Top             =   3840
         Width           =   2565
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   19
         Left            =   360
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   18
         Left            =   8040
         Top             =   5400
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   17
         Left            =   8040
         Top             =   4680
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   16
         Left            =   8040
         Top             =   3840
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   14
         Left            =   8040
         Top             =   2400
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   13
         Left            =   8040
         Top             =   3120
         Width           =   3405
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   12
         Left            =   360
         Top             =   5280
         Width           =   3405
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   11
         Left            =   4200
         Top             =   4560
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   10
         Left            =   4200
         Top             =   960
         Width           =   3405
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   9
         Left            =   4200
         Top             =   1680
         Width           =   3405
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   8
         Left            =   4200
         Top             =   3120
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   7
         Left            =   4200
         Top             =   2400
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   6
         Left            =   360
         Top             =   2400
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   5
         Left            =   360
         Top             =   1680
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   360
         Top             =   960
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   -65640
         Top             =   1080
         Width           =   3045
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -65640
         TabIndex        =   80
         Top             =   840
         Width           =   1815
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
         Index           =   20
         Left            =   -60480
         TabIndex        =   79
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Image fotoUser 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   -60480
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   -68400
         Top             =   1080
         Width           =   2565
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   1
         Left            =   -70440
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   -72600
         Top             =   1080
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   15
         Left            =   -74640
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   -74640
         Top             =   720
         Width           =   11655
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Campos de búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   -74640
         TabIndex        =   78
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Núm reg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -61440
         TabIndex        =   76
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto/Servicio"
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
         Left            =   14160
         TabIndex        =   74
         Top             =   6960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de membresía"
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
         Left            =   14160
         TabIndex        =   71
         Top             =   5400
         Visible         =   0   'False
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
         Left            =   8040
         TabIndex        =   69
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha incio membresia"
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
         Left            =   14160
         TabIndex        =   65
         Top             =   4680
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Left            =   11760
         TabIndex        =   64
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus *"
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
         Left            =   8040
         TabIndex        =   63
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Membresía *"
         Enabled         =   0   'False
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
         Left            =   8040
         TabIndex        =   62
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Código membresia"
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
         Left            =   8040
         TabIndex        =   61
         Top             =   3600
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
         Left            =   360
         TabIndex        =   60
         Top             =   5040
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
         Left            =   360
         TabIndex        =   59
         Top             =   720
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
         Left            =   360
         TabIndex        =   58
         Top             =   1440
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
         Left            =   360
         TabIndex        =   57
         Top             =   2160
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
         Left            =   360
         TabIndex        =   56
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lUsuario 
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
         Index           =   3
         Left            =   360
         TabIndex        =   55
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "CURP"
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
         Index           =   44
         Left            =   360
         TabIndex        =   54
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Código postal"
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
         Left            =   4200
         TabIndex        =   53
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia"
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
         Left            =   4200
         TabIndex        =   52
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
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
         Left            =   4200
         TabIndex        =   51
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Número interior"
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
         Left            =   4200
         TabIndex        =   50
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Número exterior"
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
         Left            =   4200
         TabIndex        =   49
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Calle"
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
         Left            =   4200
         TabIndex        =   48
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   4200
         TabIndex        =   47
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio"
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
         Left            =   4200
         TabIndex        =   46
         Top             =   1440
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
         Left            =   8040
         TabIndex        =   45
         Top             =   2160
         Width           =   2415
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
         Index           =   12
         Left            =   8040
         TabIndex        =   44
         Top             =   1440
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
         Index           =   13
         Left            =   8040
         TabIndex        =   43
         Top             =   720
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   2775
         Index           =   0
         Left            =   11880
         Top             =   840
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
         Left            =   11880
         TabIndex        =   42
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   -74640
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido paterno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   -72600
         TabIndex        =   40
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido materno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   -70440
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   -68400
         TabIndex        =   38
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image iFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   11880
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Productos en lista:"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   8880
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   -74640
         Top             =   8880
         Width           =   15615
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   1
         Left            =   -75000
         Picture         =   "FRM_Clientes.frx":52DE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   0
         Left            =   0
         Picture         =   "FRM_Clientes.frx":1231E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
   End
   Begin VB.Menu mn_Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu mn_historialClie 
         Caption         =   "Historial de cliente"
      End
      Begin VB.Menu mn_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mn_Delete 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu mn_Cat 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_TipoClie 
         Caption         =   "Tipo de cliente"
      End
   End
   Begin VB.Menu mn_Opc 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Imprimir 
         Caption         =   "Exportar"
      End
   End
End
Attribute VB_Name = "FRM_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sql1 As String
    Dim res1 As Recordset
    Dim RES2 As Recordset
    Dim RES3 As Recordset
    Dim checkError As Boolean
    Dim perId As Long
    Dim save As Boolean
    Dim mayus As Boolean

    
Sub STOPCAM()
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
TimerFoto.Enabled = False
pFoto.Visible = False
'cTomarFoto.Caption = "Tomar foto"
End Sub
Sub STARTCAM()
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hWnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0
TimerFoto.Enabled = True
End Sub
    
Private Sub CargaGeneral()
    SSTab1.Tab = 0
    'Image1(0).Visible = False
    'Image1(1).Visible = False
    iFoto.Picture = LoadPicture("")
    pFoto.Visible = False
    SSTab1.TabEnabled(1) = False
    cmbUser(4).Clear
    cmbUser(4).AddItem "ACTIVO"
    cmbUser(4).AddItem "INACTIVO"
    cmbUser(2).Clear
    cmbUser(2).AddItem "MASCULINO"
    cmbUser(2).AddItem "FEMENINO"
    cmbUser(3).Clear
    cmbUser(3).AddItem "SI"
    cmbUser(3).AddItem "NO"
    cmbUser(3).Text = "NO"
    txtUsuario(14).Enabled = True
    dtFecha(1).Enabled = False
   
End Sub
Private Sub cargaTipoCliente()
    
    sql1 = ("SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'C' ORDER BY CTPT_TIPO")
    Set res1 = con.Execute(sql1)
    
    cmbUser(5).Clear
    cmbUser(8).Clear
    cmbUser(8).AddItem "TODOS"
    Do While Not res1.EOF
        cmbUser(5).AddItem res1.Fields("CTPT_TIPO")
        cmbUser(5).ItemData(cmbUser(5).ListCount - 1) = res1.Fields("CTPT_ID")
        cmbUser(8).AddItem res1.Fields("CTPT_TIPO")
        cmbUser(8).ItemData(cmbUser(5).ListCount - 1) = res1.Fields("CTPT_ID")
        res1.MoveNext
    Loop
    
End Sub

Private Sub Check1_Click()
    If Check1.value = Checked Then
        cmbUser(7).Enabled = True
    Else
        cmbUser(7).Enabled = False
    End If
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            checarCampos
            If checkError = False Then
                crearCliente
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
            End If
        Case 2:
            buscarImagen
        Case 1:
            Dim ques As String
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        Case 3:
            checarCampos
            If checkError = False Then
                crearCliente
                lbStatus.Caption = "Estatus: Agregando cliente"
                SSTab1.TabEnabled(1) = True
                SSTab1.Tab = 1
                SSTab1.TabEnabled(0) = False
                txtUsuario(0).SetFocus
                save = False
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
            End If
    
        Case 4:
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
                SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempClie.dat")
            End If
        Case 5:
            STARTCAM
        Case 6: checkHuellas

    End Select

End Sub
Private Sub checkHuellas()
    Dim ques As String
    If lbStatus.Caption = "Estatus: Agregando cliente" Then
        ques = MsgBox("Se guardará la información del cliente y se procederá a agregar las huellas" & vbCrLf & _
        vbCrLf & "¿Continuar?", vbQuestion + vbYesNo)
        If ques = vbYes Then
            checarCampos
            If checkError = False Then
                crearCliente
                addHuellas
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
                Exit Sub
            End If
        End If
    Else
        addHuellas
    End If
End Sub
Private Sub addHuellas()
    tipoHuellas = "Clientes"
    ADD_HuellaDig.Show vbModal
End Sub
Private Sub cargaLista()
    ListaUsers.Rows = 1
    Dim texto1 As String
    
    texto1 = ""
    If cmbUser(8).Text <> "TODOS" Then
        texto1 = texto1 & "AND upper(ROL) LIKE upper('%" & cmbUser(8).Text & "%') "
    End If
        
    
    sql1 = "SELECT * FROM VIEW_PERSONA WHERE TIPO = 'CLIENTE' " & _
    "AND ID LIKE '" & textBus(0).Text & "%' " & _
    "AND upper(PATERNO) LIKE upper('%" & textBus(1).Text & "%') " & _
    "AND upper(MATERNO) LIKE upper('%" & textBus(2).Text & "%') " & _
    "AND upper(NOMBRE) LIKE upper('%" & textBus(3).Text & "%') " & _
    texto1
    


    Set res1 = con.Execute(sql1)
    ListaUsers.Redraw = False
    Do While Not res1.EOF
        ListaUsers.AddItem ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 0) = res1.Fields("ID")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 1) = res1.Fields("MEMBRESIA") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 2) = res1.Fields("PATERNO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 3) = res1.Fields("MATERNO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 4) = res1.Fields("NOMBRE")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 5) = res1.Fields("TIPO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 6) = res1.Fields("FOTO_SN")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = res1.Fields("NACIMIENTO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 8) = res1.Fields("EDAD")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 9) = res1.Fields("STATUS")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 10) = res1.Fields("TEL1") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 11) = res1.Fields("TEL2") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 12) = res1.Fields("EMAIL") & ""
        
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 13) = res1.Fields("COLONIA") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = res1.Fields("CP") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 15) = res1.Fields("cIUDAD") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 16) = res1.Fields("CALLE") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 17) = res1.Fields("NUM_EXT") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 18) = res1.Fields("NUM_INT") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 19) = res1.Fields("ALTA_SISTEMA") & ""
        
        
        If res1.Fields("STATUS") = "INACTIVO" Then
            ListaUsers.Row = ListaUsers.Rows - 1
            For b1 = 0 To ListaUsers.Cols - 1
                ListaUsers.Col = b1
                ListaUsers.CellForeColor = vbRed
            Next b1
        End If
        
        res1.MoveNext
    Loop
    lInfo(10).Caption = "Clientes en lista: " & ListaUsers.Rows - 1
    ListaUsers.Redraw = True
End Sub


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
    
    status = Left(cmbUser(4).Text, 1)
    genero = Left(cmbUser(2).Text, 1)
    membresia = Left(cmbUser(3).Text, 1)
    
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
                        
    If txtUsuario(5).Text = "" Then
        cp = "null"
    Else
        cp = txtUsuario(5).Text
    End If
        
    If txtUsuario(13).Text = "" Then
        tel1 = "null"
    Else
        tel1 = txtUsuario(13).Text
    End If
        
    If txtUsuario(12).Text = "" Then
        tel2 = "null"
    Else
        tel2 = txtUsuario(12).Text
    End If
                                
                
    If lbStatus.Caption = "Estatus: Agregando cliente" Then
        sql1 = "INSERT INTO PERSONA (PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_FEC_NAC, PER_RFC, PER_CURP, PER_DIR_EST_ID, PER_DIR_MUN_ID, " & _
        "PER_DIR_CIUDAD, PER_DIR_COLONIA, PER_DIR_CP, PER_DIR_CALLE, PER_DIR_NUM_EXT, PER_DIR_NUM_INT, PER_TEL1, PER_TEL2, PER_EMAIL, " & _
        "PER_FECHA_SISTEMA, PER_GENERO) VALUES " & _
        "('" & txtUsuario(0).Text & "', '" & txtUsuario(1).Text & "', '" & txtUsuario(2).Text & "', '" & Format(dtFecha(0), "yyyy-MM-dd") & "', " & _
        "'" & txtUsuario(3).Text & "', '" & txtUsuario(4).Text & "', " & idEstado & ", " & idMunicipio & ", " & _
        "'" & txtUsuario(7).Text & "', '" & txtUsuario(6).Text & "', " & cp & ", '" & txtUsuario(10).Text & "', '" & txtUsuario(9).Text & "', '" & txtUsuario(8).Text & "',  " & _
        "" & tel1 & ", " & tel2 & ", '" & txtUsuario(11).Text & "', " & _
        "now(), '" & genero & "' )"
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() perId"
        Set res1 = con.Execute(sql1)
        If Not res1.EOF Then
            perId = res1.Fields("perId")
            idUserHuella = res1.Fields("perId")
        End If
        
        If txtUsuario(14).Text = "" Then
            membresiaCodigo = perId
        Else
            membresiaCodigo = txtUsuario(14).Text
        End If
        
        sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA, PERTP_FECHA_MEMBRESIA, PERTP_MEMBRESIA, PERTP_CODIGO_MEMBRESIA, " & _
        "PERTP_COMENTARIOS, PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_PERALTA_FECHA) " & _
        "VALUES " & _
        "(" & cmbUser(5).ItemData(cmbUser(5).ListIndex) & ", " & perId & ", now(), 'C', '" & status & "', '" & Format(dtFecha(1), "yyyy-MM-dd") & "', '" & Format(dtFecha(1), "yyyy-MM-dd") & "', '" & membresia & "', " & _
        "'" & membresiaCodigo & "', '" & txtUsuario(15).Text & "', " & _
        "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW())"
        'MsgBox SQL1
        con.Execute (sql1)
    Else
        sql1 = "UPDATE PERSONA SET PER_NOMBRE = '" & txtUsuario(0).Text & "', " & _
        "PER_PATERNO = '" & txtUsuario(1).Text & "', " & _
        "PER_MATERNO = '" & txtUsuario(2).Text & "', " & _
        "PER_FEC_NAC = '" & Format(dtFecha(0), "yyyy-MM-dd") & "', " & _
        "PER_RFC = '" & txtUsuario(3).Text & "', " & _
        "PER_CURP = '" & txtUsuario(4).Text & "', " & _
        "PER_DIR_EST_ID = " & idEstado & ", " & _
        "PER_DIR_MUN_ID = " & idMunicipio & ", " & _
        "PER_DIR_CIUDAD = '" & txtUsuario(7).Text & "', " & _
        "PER_DIR_COLONIA = '" & txtUsuario(6).Text & "', " & _
        "PER_DIR_CP = " & cp & ", " & _
        "PER_DIR_CALLE = '" & txtUsuario(10).Text & "', " & _
        "PER_DIR_NUM_EXT = '" & txtUsuario(9).Text & "', " & _
        "PER_DIR_NUM_INT = '" & txtUsuario(8).Text & "', " & _
        "PER_TEL1 = " & tel1 & ", " & _
        "PER_TEL2 = " & tel2 & ", " & _
        "PER_EMAIL = '" & txtUsuario(11).Text & "' " & _
        "WHERE PER_ID = " & perId & ""
        con.Execute (sql1)
        
        idUserHuella = perId
        
        If txtUsuario(14).Text = "" Then
            membresiaCodigo = perId
        Else
            membresiaCodigo = txtUsuario(14).Text
        End If
        
        sql1 = "UPDATE PER_TIPO SET  " & _
        "PERTP_STATUS = '" & status & "', PERTP_FECHA_MEMBRESIA = '" & Format(dtFecha(1), "yyyy-MM-dd") & "', " & _
        "PERTP_MEMBRESIA = '" & membresia & "', PERTP_CODIGO_MEMBRESIA = '" & membresiaCodigo & "', " & _
        "PERTP_COMENTARIOS = '" & txtUsuario(15).Text & "', PERTP_TIPO_ID = " & cmbUser(5).ItemData(cmbUser(5).ListIndex) & "  " & _
        "WHERE PERTP_PER_ID = " & perId & " AND PERTP_PER_TIPO = 'C'"
        'MsgBox SQL1
        con.Execute (sql1)
        
    End If
    'Para la fotoi
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempClie.dat")
            res.Open "SELECT * FROM Persona WHERE per_id = '" & perId & "'", con, adOpenStatic, adLockOptimistic
            If res.EOF Then
            Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempClie.dat")
                res.Fields("Per_Foto") = Imagen1.Read
                res.Update
            End If
    End If
    
    MsgBox "Información guardada.", vbInformation
    save = True
    cancelar
    
End Sub


Private Sub checarCampos()
    checkError = False
    
    For b1 = 0 To 2
        If txtUsuario(b1).Text = "" Then
            checkError = True
            lUsuario(b1).ForeColor = vbRed
            Exit For
        End If
    Next b1
    
    If checkError = False Then
        If dtFecha(0) > Date Then
            checkError = True
            lUsuario(31).ForeColor = vbRed
        Else
            If cmbUser(3).Text = "" Then
                checkError = True
                lUsuario(16).ForeColor = vbRed
            Else
                If cmbUser(2).Text = "" Then
                    checkError = True
                    lUsuario(14).ForeColor = vbRed
                Else
                    If cmbUser(4).Text = "" Then
                        checkError = True
                        lUsuario(4).ForeColor = vbRed
                    Else
                        If cmbUser(5).Text = "" Then
                            checkError = True
                            lUsuario(26).ForeColor = vbRed
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub limpiarCampos()
    For b1 = 0 To 15
        txtUsuario(b1).Text = ""
    Next b1
    
    For b1 = 0 To 4
        cmbUser(b1).Clear
    Next b1

End Sub

Private Sub cancelar()
    limpiarCampos
    CargaGeneral
    cargaEstados
    cargaTipoCliente
    cargaLista
    STOPCAM
End Sub
Private Sub cmbUser_Click(Index As Integer)
Select Case Index
    Case 0:
    
    cmbUser(1).Clear
    
    sql1 = ("SELECT CTMUN_ID, CTMUN_NOMBRE FROM CAT_MUNICIPIO WHERE CTMUN_EST_ID = " & cmbUser(0).ItemData(cmbUser(0).ListIndex) & "")
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(1).AddItem res1.Fields("CTMUN_NOMBRE")
        cmbUser(1).ItemData(cmbUser(1).ListCount - 1) = res1.Fields("CTMUN_ID")
        res1.MoveNext
    Loop
        
    Case 2:
        If lUsuario(14).ForeColor = vbRed Then
            lUsuario(14).ForeColor = vbBlack
        End If
    
    Case 4:
        If lUsuario(4).ForeColor = vbRed Then
            lUsuario(4).ForeColor = vbBlack
        End If
    Case 5:
        If lUsuario(26).ForeColor = vbRed Then
            lUsuario(26).ForeColor = vbBlack
        End If

'    Case 3:
'        If Left(cmbUser(3).Text, 1) = "S" Then
'            txtUsuario(14).Enabled = True
'            dtFecha(1).Enabled = True
'            cmbUser(6).Enabled = True
'            Check1.Enabled = True
'            cmbUser(7).Enabled = True
'        Else
'            txtUsuario(14).Enabled = False
'            dtFecha(1).Enabled = False
'            cmbUser(6).Enabled = False
'            Check1.Enabled = False
'            cmbUser(7).Enabled = False
'        End If

        If lUsuario(16).ForeColor = vbRed Then
            lUsuario(16).ForeColor = vbBlack
        End If
        
    Case 6:
        cargaProdSer_Periodo (cmbUser(6).ItemData(cmbUser(6).ListIndex))
    Case 8:
        cargaLista

End Select


End Sub
Private Sub cargaProdSer_Periodo(periodoId As Long)
    sql1 = "SELECT PROD_NOMBRE, PROD_ID FROM PRODUCTO WHERE PROD_CTIDPERIODO = '" & periodoId & "'"
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(7).AddItem res1.Fields("PROD_NOMBRE")
        cmbUser(7).ItemData(cmbUser(7).ListCount - 1) = res1.Fields("PROD_ID")
        res1.MoveNext
    Loop
    
    
End Sub
Private Sub cargaEstados()
    
    sql1 = ("SELECT CT_EST_ID, CT_EST_NOMBRE FROM CAT_ESTADO ORDER BY CT_eST_NOMBRE")
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(0).AddItem res1.Fields("CT_EST_NOMBRE")
        cmbUser(0).ItemData(cmbUser(0).ListCount - 1) = res1.Fields("CT_EST_ID")
        res1.MoveNext
    Loop
    
End Sub


Public Sub cmdTipoUsuario_Click()
    cargaTipoCliente
End Sub

Private Sub Form_Load()
    CargaGeneral
    cargaEstados
    cargaTipoCliente
    cargaPeriodo
    cargaLista
    checkMayus
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
Private Sub cargaPeriodo()
    sql1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS FROM CAT_PERIODO"
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(6).AddItem res1.Fields("CTPR_PERIODO")
        cmbUser(6).ItemData(cmbUser(6).ListCount - 1) = res1.Fields("CTID_PERIODO")
        res1.MoveNext
    Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If SSTab1.Tab = 1 And save = False Then
        a = MsgBox("No se guardarán los cambios. ¿Salir?", vbYesNo + vbQuestion)
        If a = vbYes Then
            If pFoto.Visible = True Then
                STOPCAM
            End If
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub ListaUsers_Click()

    muestraInfo (ListaUsers.TextMatrix(ListaUsers.Row, 0))


End Sub
Private Sub muestraInfo(perId As Long)
'On Error Resume Next
    fotoUser.Picture = LoadPicture("")
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    sql1 = "SELECT PER_FOTO FROM PERSONA WHERE PER_ID = '" & perId & "'"
    Set res1 = con.Execute(sql1)
    If Not res1.EOF Then
        If IsNull(res1.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write res1.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            fotoUser.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            fotoUser.Picture = LoadPicture("")
        End If
        
    Else
        fotoUser.Picture = LoadPicture("")
    
    End If


End Sub

Private Sub ListaUsers_DblClick()
    If ListaUsers.MouseRow = 0 Then
        Call ordenarLista(ListaUsers)
    End If
    
End Sub

Private Sub ListaUsers_GotFocus()
    ConScroll ListaUsers
End Sub

Private Sub ListaUsers_LostFocus()
    SinScroll ListaUsers
End Sub

Private Sub ListaUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ListaUsers.Rows > 1 Then
        If Button = vbRightButton Then
            mn_Add.Enabled = True
            mn_Edit.Enabled = True
            mn_Delete.Enabled = True
            PopupMenu mn_CLientes, vbPopupMenuLeftAlign
        End If
    Else
            mn_Add.Enabled = True
            mn_Edit.Enabled = False
            mn_Delete.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_CLientes, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub ListaUsers_SelChange()
    ListaUsers_Click
End Sub

Private Sub mn_Add_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea agregar un cliente?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Agregando cliente"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            txtUsuario(0).SetFocus
            save = False
        End If


End Sub

Private Sub mn_Edit_Click()
    Dim ques As String
    
    ques = MsgBox("Desea editar al cliente: " & ListaUsers.TextMatrix(ListaUsers.Row, 0) & vbCrLf & _
            ListaUsers.TextMatrix(ListaUsers.Row, 1) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 2) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 3), vbYesNo + vbQuestion)
        If ques = vbYes Then
            perId = ListaUsers.TextMatrix(ListaUsers.Row, 0)
            idUserHuella = ListaUsers.TextMatrix(ListaUsers.Row, 0)

            lbStatus.Caption = "Estatus: Editando cliente"
            cargaEdit
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            save = True
        End If

End Sub
Private Sub cargaEdit()
'On Error Resume Next
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
'    pFoto.Visible = False
    iFoto.Visible = True
    sql1 = "SELECT PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_RFC, PER_CURP, PER_FEC_NAC, " & _
    "PER_DIR_EST_ID, PER_DIR_MUN_ID, PER_DIR_CIUDAD, PER_DIR_COLONIA, PER_DIR_CP, PER_FOTO, " & _
    "PER_DIR_CALLE, PER_DIR_NUM_EXT, PER_DIR_NUM_INT, PER_TEL1, PER_TEL2, PER_EMAIL, PER_PER_CASO_ACCIDTE, " & _
    "PER_TEL_CASO_ACCDTE, PER_DIR_ESTADO_NAC, PERTP_STATUS, PERTP_ALTA, PERTP_TIPO_ID, PER_GENERO, PERTP_MEMBRESIA, " & _
    "PERTP_CODIGO_MEMBRESIA, PERTP_COMENTARIOS, PERTP_TIPOPERIODO, PERTP_PERIODO_pRODID, PERTP_PERIODO_PRODTIPO, PERTP_PERIODO_ID " & _
    "fROM PERSONA, PER_TIPO WHERE PER_ID = " & perId & " AND PER_ID = PERTP_PER_ID AND PERTP_PER_TIPO = 'C'"
    Set RES2 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES2.EOF Then
        txtUsuario(0).Text = RES2.Fields("PER_NOMBRE")
        txtUsuario(1).Text = RES2.Fields("PER_PATERNO")
        txtUsuario(2).Text = RES2.Fields("PER_MATERNO")
        txtUsuario(3).Text = RES2.Fields("PER_RFC") & ""
        txtUsuario(4).Text = RES2.Fields("PER_CURP") & ""
        dtFecha(0) = RES2.Fields("PER_FEC_NAC")
        If IsNull(RES2.Fields("PER_DIR_EST_ID")) Then
        Else
            cmbUser(0).ListIndex = (RES2.Fields("PER_DIR_EST_ID") - 1)
        End If
        If IsNull(RES2.Fields("PER_DIR_MUN_ID")) Then
        Else
            For b1 = 0 To cmbUser(1).ListCount - 1
                If cmbUser(1).ItemData(b1) = RES2.Fields("PER_DIR_MUN_ID") Then
                    cmbUser(1).ListIndex = b1
                    Exit For
                End If
            Next b1
        End If
        txtUsuario(7).Text = RES2.Fields("PER_DIR_CIUDAD") & ""
        txtUsuario(6).Text = RES2.Fields("PER_DIR_COLONIA") & ""
        txtUsuario(5).Text = "" & RES2.Fields("PER_DIR_CP") & ""
        txtUsuario(10).Text = RES2.Fields("PER_DIR_CALLE") & ""
        txtUsuario(9).Text = RES2.Fields("PER_DIR_NUM_EXT") & ""
        txtUsuario(8).Text = RES2.Fields("PER_DIR_NUM_INT") & ""
        txtUsuario(13).Text = "" & RES2.Fields("PER_TEL1") & ""
        txtUsuario(12).Text = "" & RES2.Fields("PER_TEL2") & ""
        txtUsuario(11).Text = RES2.Fields("PER_EMAIL") & ""
        txtUsuario(14).Text = "" & RES2.Fields("PERTP_CODIGO_MEMBRESIA") & ""
        txtUsuario(15).Text = "" & RES2.Fields("PERTP_COMENTARIOS") & ""
        If RES2.Fields("PERTP_STATUS") = "A" Then
            cmbUser(4).Text = "ACTIVO"
        Else
            cmbUser(4).Text = "INACTIVO"
        End If
        If RES2.Fields("PER_GENERO") = "M" Then
            cmbUser(2).Text = "MASCULINO"
        Else
            cmbUser(2).Text = "FEMENINO"
        End If
        If RES2.Fields("PERTP_MEMBRESIA") = "S" Then
            cmbUser(3).Text = "SI"
        Else
            cmbUser(3).Text = "NO"
        End If
        For b1 = 0 To cmbUser(5).ListCount - 1
            If cmbUser(5).ItemData(b1) = RES2.Fields("PERTP_TIPO_ID") Then
                cmbUser(5).ListIndex = b1
                Exit For
            End If
        Next b1
        
        If IsNull(RES2.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES2.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempClie.dat", adSaveCreateOverWrite
            Imagen1.Close
            iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempClie.dat")
        Else
            iFoto.Picture = LoadPicture("")
        End If
        
    End If
    
End Sub

Private Sub mn_historialClie_Click()
    FRM_HistoClie.Show vbModal
End Sub

Private Sub mn_Imprimir_Click()
    Dim ques As String
    ques = MsgBox("¿Exportar la lista a excel?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(ListaUsers)
    End If
    
End Sub

Private Sub mn_TipoClie_Click()
    tipoCatTipo = "C"
    CAT_Tipo.Show vbModal

End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaLista
    End If
    Numeros (textBus(4).Text)
End Sub

Private Sub timeCarga_Timer()
    timeCarga.Enabled = False
    Image2(1).width = Me.width
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    Image2(0).width = Me.width
    Image2(0).height = Me.height
    Image2(1).width = Me.width
    Image2(1).height = Me.height
    ListaUsers.width = Me.width - 3200
    lUsuario(20).Left = Me.width - 2700
    fotoUser.Left = Me.width - 2700
    Borde(29).Left = Me.width - 2700
End Sub

Private Sub TimerFoto_Timer()
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    pFoto.Picture = Clipboard.GetData
    Clipboard.Clear
    pFoto.AutoRedraw = True
    If pFoto.Picture = 0 Then
        cmBoton(4).Caption = "Tomar foto"
        MsgBox "No se ha detectado un dispositivo de captura de imágenes. Verifique.", vbInformation
        STOPCAM
    Else
        pFoto.PaintPicture pFoto.Picture, 0, 0, pFoto.ScaleWidth, pFoto.ScaleHeight
    End If

End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
     If mayus = True Then
        Call Mayusculas(KeyAscii)
     End If
End Sub
