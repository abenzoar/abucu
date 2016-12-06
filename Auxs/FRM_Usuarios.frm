VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   20400
   Icon            =   "FRM_Usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   20400
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   20415
      _ExtentX        =   36010
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   697
      TabCaption(0)   =   "   Lista de usuarios"
      TabPicture(0)   =   "FRM_Usuarios.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lBus(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lBus(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lBus(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lBus(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fotoUser"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lUsuario(20)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lInfo(10)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Shape1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lProd(16)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Borde(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Borde(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Borde(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Borde(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Borde(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lBus(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ListaUsers"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textBus(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textBus(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textBus(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textBus(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "timeCarga"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "   Datos generales"
      TabPicture(1)   =   "FRM_Usuarios.frx":0E64
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmBoton(6)"
      Tab(1).Control(1)=   "Check1"
      Tab(1).Control(2)=   "TimerFoto"
      Tab(1).Control(3)=   "pFoto"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmBoton(2)"
      Tab(1).Control(5)=   "cmBoton(4)"
      Tab(1).Control(6)=   "cmBoton(5)"
      Tab(1).Control(7)=   "cmbUser(5)"
      Tab(1).Control(8)=   "cmdTipoUsuario"
      Tab(1).Control(9)=   "cMd1"
      Tab(1).Control(10)=   "cmbUser(4)"
      Tab(1).Control(11)=   "cmBoton(1)"
      Tab(1).Control(12)=   "cmBoton(0)"
      Tab(1).Control(13)=   "cmbUser(3)"
      Tab(1).Control(14)=   "txtUsuario(18)"
      Tab(1).Control(15)=   "cmbUser(2)"
      Tab(1).Control(16)=   "txtUsuario(17)"
      Tab(1).Control(17)=   "txtUsuario(16)"
      Tab(1).Control(18)=   "txtUsuario(15)"
      Tab(1).Control(19)=   "txtUsuario(14)"
      Tab(1).Control(20)=   "txtUsuario(13)"
      Tab(1).Control(21)=   "txtUsuario(12)"
      Tab(1).Control(22)=   "txtUsuario(11)"
      Tab(1).Control(23)=   "cmbUser(1)"
      Tab(1).Control(24)=   "cmbUser(0)"
      Tab(1).Control(25)=   "txtUsuario(10)"
      Tab(1).Control(26)=   "txtUsuario(9)"
      Tab(1).Control(27)=   "txtUsuario(8)"
      Tab(1).Control(28)=   "txtUsuario(7)"
      Tab(1).Control(29)=   "txtUsuario(6)"
      Tab(1).Control(30)=   "txtUsuario(5)"
      Tab(1).Control(31)=   "txtUsuario(4)"
      Tab(1).Control(32)=   "txtUsuario(3)"
      Tab(1).Control(33)=   "dtFecha(0)"
      Tab(1).Control(34)=   "txtUsuario(2)"
      Tab(1).Control(35)=   "txtUsuario(1)"
      Tab(1).Control(36)=   "txtUsuario(0)"
      Tab(1).Control(37)=   "dtFecha(1)"
      Tab(1).Control(38)=   "lUsuario(21)"
      Tab(1).Control(39)=   "lbStatus"
      Tab(1).Control(40)=   "iFoto"
      Tab(1).Control(41)=   "Image1(1)"
      Tab(1).Control(42)=   "Image1(0)"
      Tab(1).Control(43)=   "lUsuario(26)"
      Tab(1).Control(44)=   "lUsuario(25)"
      Tab(1).Control(45)=   "Shape1(0)"
      Tab(1).Control(46)=   "lUsuario(24)"
      Tab(1).Control(47)=   "lUsuario(18)"
      Tab(1).Control(48)=   "lUsuario(22)"
      Tab(1).Control(49)=   "lUsuario(17)"
      Tab(1).Control(50)=   "lUsuario(16)"
      Tab(1).Control(51)=   "lUsuario(19)"
      Tab(1).Control(52)=   "lUsuario(15)"
      Tab(1).Control(53)=   "lUsuario(14)"
      Tab(1).Control(54)=   "lUsuario(13)"
      Tab(1).Control(55)=   "lUsuario(12)"
      Tab(1).Control(56)=   "lUsuario(11)"
      Tab(1).Control(57)=   "lUsuario(130)"
      Tab(1).Control(58)=   "lUsuario(120)"
      Tab(1).Control(59)=   "lUsuario(10)"
      Tab(1).Control(60)=   "lUsuario(9)"
      Tab(1).Control(61)=   "lUsuario(8)"
      Tab(1).Control(62)=   "lUsuario(7)"
      Tab(1).Control(63)=   "lUsuario(6)"
      Tab(1).Control(64)=   "lUsuario(5)"
      Tab(1).Control(65)=   "lUsuario(4)"
      Tab(1).Control(66)=   "lUsuario(3)"
      Tab(1).Control(67)=   "lUsuario(31)"
      Tab(1).Control(68)=   "lUsuario(2)"
      Tab(1).Control(69)=   "lUsuario(1)"
      Tab(1).Control(70)=   "lUsuario(0)"
      Tab(1).Control(71)=   "Image2(0)"
      Tab(1).ControlCount=   72
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   11400
         TabIndex        =   79
         Top             =   1320
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Timer timeCarga 
         Interval        =   25
         Left            =   12720
         Top             =   600
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
         Left            =   -71040
         Picture         =   "FRM_Usuarios.frx":13FE
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   7920
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usuario aparece en agenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63120
         TabIndex        =   27
         Top             =   5040
         Width           =   3375
      End
      Begin VB.Timer TimerFoto 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   -60000
         Top             =   2760
      End
      Begin VB.PictureBox pFoto 
         BackColor       =   &H00E0E0E0&
         Height          =   2775
         Left            =   -63120
         ScaleHeight     =   2715
         ScaleWidth      =   2355
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   6000
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
            TabIndex        =   76
            Top             =   3600
            Width           =   2415
         End
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
         Left            =   -60600
         Picture         =   "FRM_Usuarios.frx":1CC8
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6000
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
         Left            =   -60600
         Picture         =   "FRM_Usuarios.frx":2592
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7200
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
         Left            =   -60600
         Picture         =   "FRM_Usuarios.frx":2E5C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   8400
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5280
         Width           =   3375
      End
      Begin VB.CommandButton cmdTipoUsuario 
         Caption         =   "Command1"
         Height          =   255
         Left            =   -65880
         TabIndex        =   73
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   6240
         TabIndex        =   68
         Top             =   1200
         Width           =   2295
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
         Left            =   4200
         TabIndex        =   66
         Top             =   1200
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
         Left            =   2040
         TabIndex        =   64
         Top             =   1200
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
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog cMd1 
         Left            =   -74640
         Top             =   6480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   -63120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4560
         Width           =   2895
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
         Left            =   -72960
         Picture         =   "FRM_Usuarios.frx":3726
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7920
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
         Index           =   0
         Left            =   -74760
         Picture         =   "FRM_Usuarios.frx":3FF0
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7920
         Width           =   1695
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
         Index           =   3
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   6720
         Width           =   3375
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   -63120
         MaxLength       =   8
         TabIndex        =   24
         Top             =   3120
         Width           =   1935
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
         Left            =   -63120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3840
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
         Index           =   17
         Left            =   -63120
         MaxLength       =   12
         TabIndex        =   22
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   -63120
         MaxLength       =   8
         TabIndex        =   23
         Top             =   2400
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
         Index           =   15
         Left            =   -66960
         MaxLength       =   10
         TabIndex        =   20
         Top             =   3840
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
         Index           =   14
         Left            =   -66960
         MaxLength       =   100
         TabIndex        =   19
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
         Index           =   13
         Left            =   -66960
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
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
         Index           =   12
         Left            =   -66960
         MaxLength       =   10
         TabIndex        =   17
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
         Index           =   11
         Left            =   -66960
         MaxLength       =   120
         TabIndex        =   18
         Top             =   2400
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
         Index           =   1
         Left            =   -70800
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
         Left            =   -70800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
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
         Index           =   10
         Left            =   -70800
         MaxLength       =   120
         TabIndex        =   12
         Top             =   4560
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
         Index           =   9
         Left            =   -70800
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
         Index           =   8
         Left            =   -70800
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
         Index           =   7
         Left            =   -70800
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
         Index           =   6
         Left            =   -70800
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
         Index           =   5
         Left            =   -70800
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
         Index           =   4
         Left            =   -74640
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
         Index           =   3
         Left            =   -74640
         MaxLength       =   13
         TabIndex        =   4
         Top             =   3840
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Index           =   0
         Left            =   -74640
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
         Format          =   120324097
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
         Index           =   2
         Left            =   -74640
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
         Index           =   1
         Left            =   -74640
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
         Index           =   0
         Left            =   -74640
         MaxLength       =   50
         TabIndex        =   0
         Top             =   960
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid ListaUsers 
         Height          =   6855
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   12091
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Usuarios.frx":48BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   -63120
         TabIndex        =   21
         Top             =   960
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
         Format          =   120324097
         CurrentDate     =   40783
      End
      Begin VB.Label lBus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ver solo activos"
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
         Left            =   11520
         TabIndex        =   80
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         FillColor       =   &H000040C0&
         Height          =   2595
         Index           =   3
         Left            =   13800
         Top             =   1800
         Width           =   2085
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         FillColor       =   &H000040C0&
         Height          =   435
         Index           =   2
         Left            =   6240
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         FillColor       =   &H000040C0&
         Height          =   435
         Index           =   1
         Left            =   4200
         Top             =   1200
         Width           =   1845
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         FillColor       =   &H000040C0&
         Height          =   435
         Index           =   0
         Left            =   2040
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000040C0&
         BorderWidth     =   4
         FillColor       =   &H000040C0&
         Height          =   435
         Index           =   15
         Left            =   120
         Top             =   1200
         Width           =   1725
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
         Left            =   120
         TabIndex        =   78
         Top             =   480
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   120
         Top             =   720
         Width           =   11655
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
         Index           =   21
         Left            =   -74640
         TabIndex        =   74
         Top             =   5040
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
         Left            =   120
         TabIndex        =   72
         Top             =   8640
         Width           =   5775
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
         Left            =   -68040
         TabIndex        =   71
         Top             =   8520
         Width           =   4695
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
         Left            =   13800
         TabIndex        =   70
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Image fotoUser 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   13800
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Image iFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   -63120
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   2415
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
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   69
         Top             =   960
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
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   67
         Top             =   960
         Width           =   1815
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
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   65
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave usuario"
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
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   -60960
         Picture         =   "FRM_Usuarios.frx":49D1
         Top             =   3000
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   -60960
         Picture         =   "FRM_Usuarios.frx":529B
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de usuario *"
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
         Left            =   -63120
         TabIndex        =   61
         Top             =   4320
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
         Left            =   -63120
         TabIndex        =   60
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   2775
         Index           =   0
         Left            =   -63120
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado de nacimiento"
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
         Index           =   24
         Left            =   -70800
         TabIndex        =   59
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirme password *"
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
         Left            =   -63120
         TabIndex        =   58
         Top             =   2880
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
         Index           =   22
         Left            =   -63120
         TabIndex        =   57
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario*"
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
         Left            =   -63120
         TabIndex        =   56
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Password *"
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
         Left            =   -63120
         TabIndex        =   55
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso al negocio *"
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
         Left            =   -63120
         TabIndex        =   54
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono persona aviso"
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
         Left            =   -66960
         TabIndex        =   53
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Persona aviso en caso accidente"
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
         Left            =   -66960
         TabIndex        =   52
         Top             =   2880
         Width           =   3495
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
         Left            =   -66960
         TabIndex        =   51
         Top             =   720
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
         Left            =   -66960
         TabIndex        =   50
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
         Left            =   -66960
         TabIndex        =   49
         Top             =   2160
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
         Left            =   -70800
         TabIndex        =   48
         Top             =   1440
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
         Left            =   -70800
         TabIndex        =   47
         Top             =   720
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
         Left            =   -70800
         TabIndex        =   46
         Top             =   4320
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
         Left            =   -70800
         TabIndex        =   45
         Top             =   5040
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
         Left            =   -70800
         TabIndex        =   44
         Top             =   5760
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
         Left            =   -70800
         TabIndex        =   43
         Top             =   2160
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
         Left            =   -70800
         TabIndex        =   42
         Top             =   2880
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
         Left            =   -70800
         TabIndex        =   41
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
         Index           =   4
         Left            =   -74640
         TabIndex        =   40
         Top             =   4320
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
         Left            =   -74640
         TabIndex        =   39
         Top             =   3600
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
         Left            =   -74640
         TabIndex        =   38
         Top             =   2880
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
         Left            =   -74640
         TabIndex        =   37
         Top             =   2160
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   1440
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
         Left            =   -74640
         TabIndex        =   35
         Top             =   720
         Width           =   2415
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   0
         Left            =   -75000
         Picture         =   "FRM_Usuarios.frx":5B65
         Stretch         =   -1  'True
         Top             =   480
         Width           =   20415
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   120
         Top             =   8640
         Width           =   15615
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   1
         Left            =   0
         Picture         =   "FRM_Usuarios.frx":12BA5
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
   End
   Begin VB.Menu mn_Usuario 
      Caption         =   "Usuarios"
      Begin VB.Menu mn_Adduser 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_EditUser 
         Caption         =   "Editar"
      End
      Begin VB.Menu mn_delUser 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu mn_Catalogo 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_TipoUsuario 
         Caption         =   "Tipo de usuario"
      End
   End
End
Attribute VB_Name = "FRM_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sql1 As String
    Dim RES1 As Recordset
    Dim RES2 As Recordset
    Dim RES3 As Recordset
    Dim checkError As Boolean
    Dim perId As Long
    Dim save As Boolean
    
Private Sub CargaGeneral()
    If usuarioInicial = False Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
    Else
        lbStatus.Caption = "Estatus: Agregando usuario"
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        Borde(3).Visible = False
        lUsuario(20).Visible = False
        fotoUser.Visible = False
    
    End If
        Image1(0).Visible = False
        Image1(1).Visible = False
        iFoto.Picture = LoadPicture("")
        pFoto.Visible = False
        cmbUser(2).Clear
        cmbUser(2).AddItem "ACTIVO"
        cmbUser(2).AddItem "INACTIVO"
        cmbUser(5).Clear
        cmbUser(5).AddItem "MASCULINO"
        cmbUser(5).AddItem "FEMENINO"
        save = False
        dtFecha(0) = Date - 7300
        
End Sub
Private Sub crearUsuario()

    Dim status As String
    Dim idEstado As String
    Dim idMunicipio As String
    Dim idEstadoNac As String
    Dim cp As String
    Dim tel1 As String
    Dim tel2 As String
    Dim telAccdte As String
    Dim genero As String
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    
    status = Left(cmbUser(2).Text, 1)
    genero = Left(cmbUser(5).Text, 1)

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
        
    If cmbUser(3).Text = "" Then
        idEstadoNac = "null"
    Else
        idEstadoNac = cmbUser(3).ItemData(cmbUser(3).ListIndex)
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
        
    If txtUsuario(15).Text = "" Then
        telAccdte = "null"
    Else
        telAccdte = txtUsuario(15).Text
    End If
        
    If lbStatus.Caption = "Estatus: Agregando usuario" Then
        sql1 = "INSERT INTO PERSONA (PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_FEC_NAC, PER_RFC, PER_CURP, PER_DIR_EST_ID, PER_DIR_MUN_ID, " & _
        "PER_DIR_CIUDAD, PER_DIR_COLONIA, PER_DIR_CP, PER_DIR_CALLE, PER_DIR_NUM_EXT, PER_DIR_NUM_INT, PER_TEL1, PER_TEL2, PER_EMAIL, " & _
        "PER_FECHA_SISTEMA, PER_DIR_ESTADO_NAC, PER_PER_CASO_ACCIDTE, PER_TEL_CASO_ACCDTE, PER_GENERO) VALUES " & _
        "('" & txtUsuario(0).Text & "', '" & txtUsuario(1).Text & "', '" & txtUsuario(2).Text & "', '" & Format(dtFecha(0), "yyyy-MM-dd") & "', " & _
        "'" & txtUsuario(3).Text & "', '" & txtUsuario(4).Text & "', " & idEstado & ", " & idMunicipio & ", " & _
        "'" & txtUsuario(7).Text & "', '" & txtUsuario(6).Text & "', " & cp & ", '" & txtUsuario(10).Text & "', '" & txtUsuario(9).Text & "', '" & txtUsuario(8).Text & "',  " & _
        "" & tel1 & ", " & tel2 & ", '" & txtUsuario(11).Text & "', " & _
        "now(), " & idEstadoNac & ", '" & txtUsuario(14).Text & "', " & telAccdte & ", '" & genero & "')"
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() perId"
        Set RES1 = con.Execute(sql1)
        If Not RES1.EOF Then
            perId = RES1.Fields("perId")
            idUserHuella = RES1.Fields("perId")
        End If
        
        If usuarioInicial = True Then
                sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA, PERTP_USUARIO, PERTP_PASSWORD, " & _
                "PERTP_PERALTA_FECHA, PERTP_CODIGO_MEMBRESIA, PERTP_AGENDA) " & _
                "VALUES " & _
                "(" & cmbUser(4).ItemData(cmbUser(4).ListIndex) & ", " & perId & ", now(), 'U', '" & status & "', '" & Format(dtFecha(1), "yyyy-MM-dd") & "', " & _
                "'" & txtUsuario(17).Text & "', MD5('" & txtUsuario(18).Text & "'), NOW(), '" & perId & "', '" & Check1.value & "' )"
            con.Execute (sql1)
        Else
                sql1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA, PERTP_USUARIO, PERTP_PASSWORD, " & _
                "PERTP_PERALTA_FECHA, PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_CODIGO_MEMBRESIA, PERTP_AGENDA) " & _
                "VALUES " & _
                "(" & cmbUser(4).ItemData(cmbUser(4).ListIndex) & ", " & perId & ", now(), 'U', '" & status & "', '" & Format(dtFecha(1), "yyyy-MM-dd") & "', " & _
                "'" & txtUsuario(17).Text & "', MD5('" & txtUsuario(18).Text & "'), NOW(), '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', '" & perId & "', '" & Check1.value & "' )"

            con.Execute (sql1)
        End If
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
        "PER_EMAIL = '" & txtUsuario(11).Text & "', " & _
        "PER_DIR_ESTADO_NAC = " & idEstadoNac & ", " & _
        "PER_PER_CASO_ACCIDTE = '" & txtUsuario(14).Text & "', " & _
        "PER_TEL_CASO_ACCDTE = " & telAccdte & ", PER_GENERO = '" & genero & "' " & _
        "WHERE PER_ID = " & perId & ""
        con.Execute (sql1)
        
        
        idUserHuella = perId
        
        If txtUsuario(16).Text <> "NOCAMBIO" Then
            sql1 = "UPDATE PER_TIPO SET PERTP_TIPO_ID = '" & cmbUser(4).ItemData(cmbUser(4).ListIndex) & "', " & _
            "PERTP_STATUS = '" & status & "', PERTP_ALTA = '" & Format(dtFecha(1), "yyyy-MM-dd") & "', " & _
            "PERTP_USUARIO = '" & txtUsuario(17).Text & "', PERTP_PASSWORD = MD5('" & txtUsuario(18).Text & "'), " & _
            "PERTP_AGENDA = '" & Check1.value & "' " & _
            "WHERE PERTP_PER_ID = " & perId & " AND PERTP_PER_TIPO = 'U'"
        Else
            sql1 = "UPDATE PER_TIPO SET PERTP_TIPO_ID = '" & cmbUser(4).ItemData(cmbUser(4).ListIndex) & "', " & _
            "PERTP_STATUS = '" & status & "', PERTP_ALTA = '" & Format(dtFecha(1), "yyyy-MM-dd") & "', " & _
            "PERTP_USUARIO = '" & txtUsuario(17).Text & "', " & _
            "PERTP_AGENDA = '" & Check1.value & "' " & _
            "WHERE PERTP_PER_ID = " & perId & " AND PERTP_PER_TIPO = 'U'"
        End If
        con.Execute (sql1)
        
    End If
    'Para la fotoi
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempUser.dat")
            res.Open "SELECT * FROM Persona WHERE per_id = '" & perId & "'", con, adOpenStatic, adLockOptimistic
            If res.EOF Then
            Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempUser.dat")
                res.Fields("Per_Foto") = Imagen1.Read
                res.Update
            End If
    End If
    
    MsgBox "Información guardada.", vbInformation
    save = True
    If usuarioInicial = False Then
        cancelar
    Else
        MsgBox "Ahora puede inciar sesión como usuario en el sistema.", vbInformation
        usuarioInicial = False
        Unload Me
    End If
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
        If dtFecha(1) > Date Then
            checkError = True
            lUsuario(19).ForeColor = vbRed
        Else
            If dtFecha(0) > Date Then
                checkError = True
                lUsuario(31).ForeColor = vbRed
            Else
                If txtUsuario(17).Text = "" Then
                    checkError = True
                    lUsuario(17).ForeColor = vbRed
                Else
                    If Image1(1).Visible = True Then
                        checkError = True
                        lUsuario(18).ForeColor = vbRed
                    Else
                        If Image1(0).Visible = False And Image1(1).Visible = False Then
                            checkError = True
                            lUsuario(18).ForeColor = vbRed
                        Else
                            If cmbUser(2).Text = "" Then
                                checkError = True
                                lUsuario(22).ForeColor = vbRed
                            Else
                                If cmbUser(4).Text = "" Then
                                    checkError = True
                                    lUsuario(26).ForeColor = vbRed
                                Else
                                    If cmbUser(5).Text = "" Then
                                        checkError = True
                                        lUsuario(21).ForeColor = vbRed
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub Check2_Click()
    cargaLista
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            checarCampos
            If checkError = False Then
                crearUsuario
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
        Case 5
            STARTCAM
        Case 6: checkHuellas

    End Select

End Sub
Private Sub checkHuellas()
    Dim ques As String
    If lbStatus.Caption = "Estatus: Agregando usuario" Then
        ques = MsgBox("Se guardará la información del usuario y se procederá a agregar las huellas" & vbCrLf & _
        vbCrLf & "¿Continuar?", vbQuestion + vbYesNo)
        If ques = vbYes Then
            checarCampos
            If checkError = False Then
                crearUsuario
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
    tipoHuellas = "Usuarios"
    ADD_HuellaDig.Show vbModal
End Sub
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

Private Sub cancelar()
    limpiarCampos
    CargaGeneral
    cargaEstados
    cargaTipoUsuario
    cargaLista

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

Private Sub cmbUser_Click(Index As Integer)
Select Case Index
    Case 0:
    
    cmbUser(1).Clear
    
    sql1 = ("SELECT CTMUN_ID, CTMUN_NOMBRE FROM CAT_MUNICIPIO WHERE CTMUN_EST_ID = " & cmbUser(0).ItemData(cmbUser(0).ListIndex) & "")
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        cmbUser(1).AddItem RES1.Fields("CTMUN_NOMBRE")
        cmbUser(1).ItemData(cmbUser(1).ListCount - 1) = RES1.Fields("CTMUN_ID")
        RES1.MoveNext
    Loop
        
    Case 2:
        If lUsuario(22).ForeColor = vbRed Then
            lUsuario(22).ForeColor = vbBlack
        End If
    
    Case 4:
        If lUsuario(26).ForeColor = vbRed Then
            lUsuario(26).ForeColor = vbBlack
        End If


End Select


End Sub


Private Sub Command1_Click()
MsgBox cmbUser(1).ListCount
End Sub

Public Sub cmdTipoUsuario_Click()
    cargaTipoUsuario
End Sub

Private Sub dtFecha_Click(Index As Integer)
    If Index = 1 Then
        If lUsuario(19).ForeColor = vbRed Then
            lUsuario(19).ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub Form_Load()
'    txtFileDir
'    ConectarDB
    CargaGeneral
    cargaEstados
    cargaTipoUsuario
    cargaLista
    
End Sub

'Private Sub ConectarDB()
'    Call ConexionDB("localhost", "db_gym", "root", "9807288")
'End Sub

Private Sub cargaLista()
    Dim texto1 As String
    texto1 = " "
    If Check2.value = Checked Then
        texto1 = texto1 & " AND PERTP_STATUS = 'A' "
    End If
    
    
    ListaUsers.Rows = 1
    ListaUsers.Redraw = False

    sql1 = "SELECT * FROM   VIEW_PERSONA WHERE TIPO = 'USUARIO' " & _
    "AND ID LIKE '" & textBus(0).Text & "%' " & _
    "AND upper(PATERNO) LIKE upper('%" & textBus(1).Text & "%') " & _
    "AND upper(MATERNO) LIKE upper('%" & textBus(2).Text & "%') " & _
    "AND upper(NOMBRE) LIKE upper('%" & textBus(3).Text & "%') " & texto1


    Set RES1 = con.Execute(sql1)
        
    Do While Not RES1.EOF
        ListaUsers.AddItem ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 0) = RES1.Fields("ID")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 1) = RES1.Fields("PATERNO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 2) = RES1.Fields("MATERNO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 3) = RES1.Fields("NOMBRE")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 4) = RES1.Fields("STATUS")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 5) = RES1.Fields("ROL")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 6) = RES1.Fields("FOTO_SN")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = RES1.Fields("TEL1") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 8) = RES1.Fields("TEL2") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 9) = RES1.Fields("NACIMIENTO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 10) = RES1.Fields("EDAD")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 11) = RES1.Fields("AVISO_PERSONA") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 12) = RES1.Fields("TEL_AVISO") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 13) = RES1.Fields("FECHA_INGRESO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = RES1.Fields("ALTA_SISTEMA")
        
        If RES1.Fields("STATUS") = "INACTIVO" Then
            ListaUsers.Row = ListaUsers.Rows - 1
            For b1 = 0 To ListaUsers.Cols - 1
                ListaUsers.Col = b1
                ListaUsers.CellForeColor = vbRed
            Next b1
        End If
        
        RES1.MoveNext
    Loop
    lInfo(10).Caption = "Usuarios en lista: " & ListaUsers.Rows - 1
    ListaUsers.Redraw = True
    
End Sub
Private Sub limpiarCampos()
    For b1 = 0 To 18
        txtUsuario(b1).Text = ""
    Next b1
    
    For b1 = 0 To 4
        cmbUser(b1).Clear
    Next b1

End Sub

Private Sub cargaEstados()
    
    sql1 = ("SELECT CT_EST_ID, CT_EST_NOMBRE FROM CAT_ESTADO ORDER BY CT_eST_NOMBRE")
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        cmbUser(0).AddItem RES1.Fields("CT_EST_NOMBRE")
        cmbUser(0).ItemData(cmbUser(0).ListCount - 1) = RES1.Fields("CT_EST_ID")
        cmbUser(3).AddItem RES1.Fields("CT_EST_NOMBRE")
        cmbUser(3).ItemData(cmbUser(3).ListCount - 1) = RES1.Fields("CT_EST_ID")
        RES1.MoveNext
    Loop
    
End Sub

Private Sub cargaTipoUsuario()
    
    sql1 = ("SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'U' ORDER BY CTPT_TIPO")
    Set RES1 = con.Execute(sql1)
    
    cmbUser(4).Clear
    Do While Not RES1.EOF
        cmbUser(4).AddItem RES1.Fields("CTPT_TIPO")
        cmbUser(4).ItemData(cmbUser(4).ListCount - 1) = RES1.Fields("CTPT_ID")
        RES1.MoveNext
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SSTab1.Tab = 1 And save = False Then
    a = MsgBox("Se perderan los cambios. ¿Salir?", vbYesNo + vbQuestion)
    If a = vbYes Then
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

    fotoUser.Picture = LoadPicture("")
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    sql1 = "SELECT PER_FOTO FROM PERSONA T2 " & _
    "WHERE T2.PER_ID = '" & perId & "'"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        If IsNull(RES1.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PER_FOTO")
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

Private Sub ListaUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ListaUsers.Rows > 1 Then
        If Button = vbRightButton Then
            mn_Adduser.Enabled = True
            mn_EditUser.Enabled = True
            mn_delUser.Enabled = True
            PopupMenu mn_Usuario, vbPopupMenuLeftAlign
        End If
    Else
            mn_Adduser.Enabled = True
            mn_EditUser.Enabled = False
            mn_delUser.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_Usuario, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Adduser_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea agregar un usuario?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Agregando usuario"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            txtUsuario(0).SetFocus
            save = False
        End If

End Sub

Private Sub mn_EditUser_Click()
    Dim ques As String
    
    ques = MsgBox("Desea editar al usuario: " & ListaUsers.TextMatrix(ListaUsers.Row, 0) & vbCrLf & _
            ListaUsers.TextMatrix(ListaUsers.Row, 1) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 2) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 3), vbYesNo + vbQuestion)
        If ques = vbYes Then
            perId = ListaUsers.TextMatrix(ListaUsers.Row, 0)
            lbStatus.Caption = "Estatus: Editando usuario"
            cargaEdit
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            save = False
        End If
    
End Sub
Private Sub cargaEdit()
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
'    pFoto.Visible = False
    iFoto.Visible = True
    sql1 = "SELECT PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_RFC, PER_CURP, PER_FEC_NAC, " & _
    "PER_DIR_EST_ID, PER_DIR_MUN_ID, PER_DIR_CIUDAD, PER_DIR_COLONIA, PER_DIR_CP, PER_FOTO, " & _
    "PER_DIR_CALLE, PER_DIR_NUM_EXT, PER_DIR_NUM_INT, PER_TEL1, PER_TEL2, PER_EMAIL, PER_PER_CASO_ACCIDTE, " & _
    "PER_TEL_CASO_ACCDTE, PER_DIR_ESTADO_NAC, PERTP_STATUS, PERTP_ALTA, PERTP_USUARIO, PERTP_TIPO_ID, PERTP_PASSWORD, PER_GENERO, PERTP_AGENDA " & _
    "FROM PERSONA, PER_TIPO WHERE PER_ID = " & perId & " AND PER_ID = PERTP_PER_ID AND PERTP_PER_TIPO = 'U'"
    Set RES2 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES2.EOF Then
        idUserHuella = perId
        
        txtUsuario(0).Text = RES2.Fields("PER_NOMBRE")
        txtUsuario(1).Text = RES2.Fields("PER_PATERNO")
        txtUsuario(2).Text = RES2.Fields("PER_MATERNO")
        txtUsuario(3).Text = RES2.Fields("PER_RFC")
        txtUsuario(4).Text = RES2.Fields("PER_CURP")
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
        txtUsuario(7).Text = RES2.Fields("PER_DIR_CIUDAD")
        txtUsuario(6).Text = RES2.Fields("PER_DIR_COLONIA")
        txtUsuario(5).Text = "" & RES2.Fields("PER_DIR_CP")
        txtUsuario(10).Text = RES2.Fields("PER_DIR_CALLE")
        txtUsuario(9).Text = RES2.Fields("PER_DIR_NUM_EXT")
        txtUsuario(8).Text = RES2.Fields("PER_DIR_NUM_INT")
        txtUsuario(13).Text = "" & RES2.Fields("PER_TEL1")
        txtUsuario(12).Text = "" & RES2.Fields("PER_TEL2")
        txtUsuario(11).Text = RES2.Fields("PER_EMAIL")
        txtUsuario(14).Text = RES2.Fields("PER_PER_CASO_ACCIDTE")
        txtUsuario(15).Text = "" & RES2.Fields("PER_TEL_CASO_ACCDTE")
        If IsNull(RES2.Fields("PERTP_AGENDA")) Then
            Check1.value = Unchecked
        Else
            Check1.value = RES2.Fields("PERTP_AGENDA")
        End If
        If IsNull(RES2.Fields("PER_DIR_ESTADO_NAC")) Then
        Else
            cmbUser(3).ListIndex = (RES2.Fields("PER_DIR_ESTADO_NAC") - 1)
        End If
        If RES2.Fields("PERTP_STATUS") = "A" Then
            cmbUser(2).Text = "ACTIVO"
        Else
            cmbUser(2).Text = "INACTIVO"
        End If
        If RES2.Fields("PER_GENERO") = "M" Then
            cmbUser(5).Text = "MASCULINO"
        Else
            cmbUser(5).Text = "FEMENINO"
        End If
        txtUsuario(17).Text = RES2.Fields("PERTP_USUARIO")
        txtUsuario(16).Text = "NOCAMBIO"
        txtUsuario(18).Text = "NOCAMBIO"
        dtFecha(1) = RES2.Fields("PERTP_ALTA")
            For b1 = 0 To cmbUser(4).ListCount - 1
                If cmbUser(4).ItemData(b1) = RES2.Fields("PERTP_TIPO_ID") Then
                    cmbUser(4).ListIndex = b1
                    Exit For
                End If
            Next b1
        If IsNull(RES2.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES2.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            iFoto.Picture = LoadPicture("")
        End If
        
    End If
    
End Sub

Private Sub mn_TipoUsuario_Click()
    tipoCatTipo = "U"
    CAT_Tipo.Show vbModal
    
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaLista
    End If
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
    Borde(3).Left = Me.width - 2700
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

Private Sub txtUsuario_Change(Index As Integer)
    Select Case Index
        Case 16:
            txtUsuario(Index).PasswordChar = "*"
            txtUsuario(Index).FontBold = True
            txtUsuario(Index).FontSize = 16
            txtUsuario(18).Text = ""
        Case 18:
            txtUsuario(Index).PasswordChar = "*"
            txtUsuario(Index).FontBold = True
            txtUsuario(Index).FontSize = 16
            If txtUsuario(18).Text = txtUsuario(16).Text Then
                Image1(0).Visible = True
                Image1(1).Visible = False
                
            Else
                Image1(1).Visible = True
                Image1(0).Visible = False
            End If
    End Select

    If lUsuario(Index).ForeColor = vbRed Then
        lUsuario(Index).ForeColor = vbBlack
    End If
    
End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 3:
            Call Mayusculas(KeyAscii)
        Case 4:
            Call Mayusculas(KeyAscii)
        Case 13:
            Call Numeros(KeyAscii)
        Case 12:
            Call Numeros(KeyAscii)
        Case 5:
            Call Numeros(KeyAscii)
        Case 16:
            txtUsuario(Index).PasswordChar = "*"
            txtUsuario(Index).FontBold = True
            txtUsuario(Index).FontSize = 16
        Case 18:
            txtUsuario(Index).PasswordChar = "*"
            txtUsuario(Index).FontBold = True
            txtUsuario(Index).FontSize = 16
            If txtUsuario(18).Text = txtUsuario(16).Text Then
                Image1(0).Visible = True
                Image1(1).Visible = False
                
            Else
                Image1(1).Visible = True
                Image1(0).Visible = False
            End If
    End Select

End Sub
