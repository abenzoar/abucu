VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Apartados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apartados - Pagos Apartados"
   ClientHeight    =   9585
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   16785
   Icon            =   "FRM_Apartados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   16785
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   16960
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   Listado de Apartados/pagos"
      TabPicture(0)   =   "FRM_Apartados.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2(1)"
      Tab(0).Control(1)=   "Label1(7)"
      Tab(0).Control(2)=   "Line1(2)"
      Tab(0).Control(3)=   "lblClieId(5)"
      Tab(0).Control(4)=   "lblClieId(4)"
      Tab(0).Control(5)=   "lblClieId(3)"
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(7)=   "Label1(6)"
      Tab(0).Control(8)=   "Label1(5)"
      Tab(0).Control(9)=   "Label1(4)"
      Tab(0).Control(10)=   "lBus(4)"
      Tab(0).Control(11)=   "lBus(0)"
      Tab(0).Control(12)=   "lBus(1)"
      Tab(0).Control(13)=   "lBus(2)"
      Tab(0).Control(14)=   "lBus(3)"
      Tab(0).Control(15)=   "Line1(1)"
      Tab(0).Control(16)=   "Label1(8)"
      Tab(0).Control(17)=   "Line1(5)"
      Tab(0).Control(18)=   "Label1(11)"
      Tab(0).Control(19)=   "Label1(12)"
      Tab(0).Control(20)=   "Lista2"
      Tab(0).Control(21)=   "Lista1"
      Tab(0).Control(22)=   "TimeSize"
      Tab(0).Control(23)=   "txtPgo(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmBoton(4)"
      Tab(0).Control(25)=   "txtPgo(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPgo(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmBoton(3)"
      Tab(0).Control(28)=   "txtInfoAprt"
      Tab(0).Control(29)=   "textBus(4)"
      Tab(0).Control(30)=   "textBus(0)"
      Tab(0).Control(31)=   "textBus(1)"
      Tab(0).Control(32)=   "textBus(2)"
      Tab(0).Control(33)=   "cmBoton(5)"
      Tab(0).Control(34)=   "cmbStatus"
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "   Nuevo apartado"
      TabPicture(1)   =   "FRM_Apartados.frx":0B24
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lUsuario(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblUserId(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblUserId(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblUserId(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblDatos(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lUsuario(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblClieId(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblClieId(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblClieId(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lUsuario(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "imgFoto(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblDatos(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "imgFoto(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblDatos(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(13)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(14)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(15)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(16)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Line1(6)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lUsuario(4)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(17)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Line1(7)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Line1(8)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(18)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Line1(9)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(19)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Line1(10)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(20)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lUsuario(6)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lUsuario(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lUsuario(8)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lUsuario(9)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Line1(11)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label1(21)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lUsuario(10)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Label1(0)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label1(23)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Line1(12)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label1(24)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Line1(13)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Label1(25)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "lUsuario(11)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lblDatos(6)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Label1(26)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lUsuario(12)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lista"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "cmBoton(2)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "txtTotalPago"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "txtTotalAnt"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "txtAnticipo(0)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "cmBoton(1)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "txtTotal"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "cmBoton(0)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "txtNumPagos"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "cmdPeriodo"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "cmbDato(0)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "txtClave(0)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "txtClave(2)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "txtAnticipo(1)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "txtAnticipo(2)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "txtAnticipo(3)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "txtAnticipo(4)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "txtAnticipo(5)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "txtDescuento"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "txtSubTotal"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "cmbUser"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "cmdBus(0)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "cmdBus(1)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "CheckPC"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).ControlCount=   70
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FRM_Apartados.frx":10BE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CheckBox CheckPC 
         BackColor       =   &H8000000A&
         Caption         =   "Check1"
         Height          =   375
         Left            =   3360
         TabIndex        =   95
         Top             =   840
         Width           =   255
      End
      Begin VB.ComboBox cmbStatus 
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
         Left            =   -66600
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdBus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   6840
         Picture         =   "FRM_Apartados.frx":10DA
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton cmdBus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   3240
         Picture         =   "FRM_Apartados.frx":1664
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2280
         Width           =   495
      End
      Begin VB.ComboBox cmbUser 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtSubTotal 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8400
         Width           =   2295
      End
      Begin VB.TextBox txtDescuento 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8400
         Width           =   1935
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "0"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   3600
         TabIndex        =   76
         Text            =   "0"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   2400
         TabIndex        =   73
         Text            =   "0"
         Top             =   3585
         Width           =   975
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0"
         Top             =   3585
         Width           =   1815
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   12000
         TabIndex        =   69
         Text            =   "0"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir información"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   -59880
         Picture         =   "FRM_Apartados.frx":1BEE
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox textBus 
         Height          =   405
         Index           =   2
         Left            =   -69840
         TabIndex        =   23
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox textBus 
         Height          =   405
         Index           =   1
         Left            =   -73200
         TabIndex        =   22
         Top             =   960
         Width           =   3135
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
         Left            =   -74880
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox textBus 
         Height          =   285
         Index           =   4
         Left            =   -60000
         TabIndex        =   20
         Text            =   "50"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cmbDato 
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton cmdPeriodo 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   16
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtNumPagos 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13440
         TabIndex        =   15
         Text            =   "1"
         Top             =   3600
         Width           =   855
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
         Height          =   615
         Index           =   0
         Left            =   14520
         Picture         =   "FRM_Apartados.frx":2178
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8400
         Width           =   2775
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realizar operación"
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
         Index           =   1
         Left            =   13200
         Picture         =   "FRM_Apartados.frx":2702
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8040
         Width           =   1695
      End
      Begin VB.TextBox txtAnticipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   10800
         TabIndex        =   10
         Text            =   "0"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtTotalAnt 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8400
         Width           =   2535
      End
      Begin VB.TextBox txtInfoAprt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -66480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   5760
         Width           =   6255
      End
      Begin VB.TextBox txtTotalPago 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   10200
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8400
         Width           =   2775
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
         Height          =   975
         Index           =   2
         Left            =   15000
         Picture         =   "FRM_Apartados.frx":2FCC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8040
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
         Index           =   3
         Left            =   -62040
         Picture         =   "FRM_Apartados.frx":3896
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8520
         Width           =   1695
      End
      Begin VB.TextBox txtPgo 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   2
         Left            =   -67680
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8760
         Width           =   3375
      End
      Begin VB.TextBox txtPgo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   1
         Left            =   -71280
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8760
         Width           =   3375
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realizar pago"
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
         Index           =   4
         Left            =   -63960
         Picture         =   "FRM_Apartados.frx":4160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8520
         Width           =   1695
      End
      Begin VB.TextBox txtPgo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "$0.0"
         Top             =   8760
         Width           =   3375
      End
      Begin VB.Timer TimeSize 
         Interval        =   500
         Left            =   -60840
         Top             =   360
      End
      Begin MSFlexGridLib.MSFlexGrid Lista1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   11
         Top             =   1560
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   37
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Apartados.frx":4A2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Lista2 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   24
         Top             =   5760
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Apartados.frx":4C2D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   3255
         Left            =   240
         TabIndex        =   25
         Top             =   4200
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Apartados.frx":4CC5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
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
         Index           =   12
         Left            =   3360
         TabIndex        =   96
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monedero"
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
         Index           =   26
         Left            =   13200
         TabIndex        =   94
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   13200
         TabIndex        =   93
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
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
         Index           =   11
         Left            =   3600
         TabIndex        =   88
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Cód barra F4"
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
         Index           =   25
         Left            =   5160
         TabIndex        =   87
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   13
         X1              =   5040
         X2              =   6720
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Cód barra F2"
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
         Index           =   24
         Left            =   1560
         TabIndex        =   86
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   12
         X1              =   1560
         X2              =   3240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   240
         TabIndex        =   85
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total descuento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   83
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   10
         Left            =   4800
         TabIndex        =   80
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valores para anticipo"
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
         Index           =   21
         Left            =   6960
         TabIndex        =   78
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   6840
         X2              =   14520
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lUsuario 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Index           =   9
         Left            =   3360
         TabIndex        =   77
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label lUsuario 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   8
         Left            =   2160
         TabIndex        =   75
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
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
         Index           =   7
         Left            =   2400
         TabIndex        =   74
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
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
         Index           =   6
         Left            =   240
         TabIndex        =   72
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario atiende"
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
         Index           =   20
         Left            =   7200
         TabIndex        =   68
         Top             =   600
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   10
         X1              =   7200
         X2              =   10200
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto o Servicio seleccionado"
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
         Index           =   19
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   3015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   9
         X1              =   240
         X2              =   3240
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente seleccionado"
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
         Index           =   18
         Left            =   3840
         TabIndex        =   66
         Top             =   600
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   8
         X1              =   3840
         X2              =   6840
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   240
         X2              =   6720
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valores para el producto"
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
         Index           =   17
         Left            =   240
         TabIndex        =   65
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipo $"
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
         Index           =   4
         Left            =   12000
         TabIndex        =   64
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   240
         X2              =   15120
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valores para realizar pago"
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
         Index           =   16
         Left            =   240
         TabIndex        =   62
         Top             =   7680
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   4680
         TabIndex        =   61
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total anticipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   7560
         TabIndex        =   60
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago a realizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   10200
         TabIndex        =   59
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valores para realizar pago"
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   57
         Top             =   8160
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos del cliente y de la operación"
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
         Index           =   11
         Left            =   -66480
         TabIndex        =   56
         Top             =   5400
         Width           =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   5
         X1              =   -66480
         X2              =   -58320
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagos realizados para el apartado seleccionado"
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
         Index           =   8
         Left            =   -74880
         TabIndex        =   53
         Top             =   5400
         Width           =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   1
         X1              =   -74880
         X2              =   -66720
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   -66600
         TabIndex        =   52
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   -69840
         TabIndex        =   51
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
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
         Left            =   -73200
         TabIndex        =   50
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Folio operación"
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
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   1335
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
         Left            =   -60000
         TabIndex        =   48
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguno"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   47
         Top             =   960
         Width           =   1815
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   2
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguno"
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
         Left            =   1560
         TabIndex        =   46
         Top             =   960
         Width           =   1815
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo *"
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
         Index           =   2
         Left            =   6840
         TabIndex        =   45
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   11880
         TabIndex        =   44
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   11880
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   11880
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagos #"
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
         Left            =   13440
         TabIndex        =   41
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   8400
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   10680
         TabIndex        =   39
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   10680
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   10680
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   36
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipo %"
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
         Left            =   10800
         TabIndex        =   35
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total anticipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -71280
         TabIndex        =   34
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total pago anticipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -67800
         TabIndex        =   33
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago a realizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -67680
         TabIndex        =   32
         Top             =   8520
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total adeudo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -71280
         TabIndex        =   31
         Top             =   8520
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -74880
         TabIndex        =   30
         Top             =   8520
         Width           =   2175
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -74880
         X2              =   -60000
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   3
         Left            =   -59040
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   4
         Left            =   -59040
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   5
         Left            =   -59040
         TabIndex        =   27
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   -74880
         X2              =   -63960
         Y1              =   700
         Y2              =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista"
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
         Index           =   7
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   5175
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   0
         Left            =   0
         Picture         =   "FRM_Apartados.frx":4E32
         Stretch         =   -1  'True
         Top             =   360
         Width           =   17655
      End
      Begin VB.Image Image2 
         Height          =   9855
         Index           =   1
         Left            =   -75000
         Picture         =   "FRM_Apartados.frx":11E72
         Stretch         =   -1  'True
         Top             =   360
         Width           =   17655
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   240
      TabIndex        =   81
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Anticipo %"
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
      Left            =   5520
      TabIndex        =   70
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Anticipo %"
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
      Left            =   6720
      TabIndex        =   63
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagos realizados para el apartado seleccionado"
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
      Index           =   10
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   4
      X1              =   0
      X2              =   8160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagos realizados para el apartado seleccionado"
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
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   3
      X1              =   0
      X2              =   8160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Menu mn 
      Caption         =   "Menu"
      Begin VB.Menu mnSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mn_NuevoAprt 
      Caption         =   "Búsqueda"
      Begin VB.Menu mn_BusProd 
         Caption         =   "Buscar producto"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mn_BusClte 
         Caption         =   "Buscar cliente"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mn_Part 
      Caption         =   "Apartado"
      Begin VB.Menu mn_NewApar 
         Caption         =   "Nuevo apartado"
      End
      Begin VB.Menu mn_RePag 
         Caption         =   "Realizar pago"
      End
      Begin VB.Menu mn_PagoTodos 
         Caption         =   "Realizar pago de los apartados del cliente"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_lineCancel 
         Caption         =   "-"
      End
      Begin VB.Menu mn_CancelAprt 
         Caption         =   "Cancelar apartado"
      End
      Begin VB.Menu mn_lineApatados 
         Caption         =   "-"
      End
      Begin VB.Menu mn_PrintGral 
         Caption         =   "Imprimir ticket general del apartado"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mn_Ctlg 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_Periodo 
         Caption         =   "Periodo"
      End
   End
   Begin VB.Menu mn_PgosAprt 
      Caption         =   "Menu pagos"
      Visible         =   0   'False
      Begin VB.Menu mn_PrintPago 
         Caption         =   "Imprimir ticket pago"
      End
   End
   Begin VB.Menu mn_SubMenu 
      Caption         =   "SubMenu"
      Visible         =   0   'False
      Begin VB.Menu mn_Cancel 
         Caption         =   "Cancelar"
      End
   End
   Begin VB.Menu mn_Options 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Export 
         Caption         =   "Exportar"
      End
   End
End
Attribute VB_Name = "FRM_Apartados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim resProd As Recordset
Dim resClie As Recordset
Dim resLista As Recordset
Dim resPagos As Recordset
Dim resInfoCLie As Recordset
Dim resApart As Recordset
Dim sql1 As String
Dim errorValor As Boolean
Dim valida As Boolean

Private Sub borrarDatosProd()
    For b1 = 0 To 5
        txtAnticipo(b1).Text = "0"
    Next b1

End Sub

Private Sub Check1_Click()
    

    
End Sub

Private Sub cmBoton_Click(Index As Integer)

Dim ques As String

    Select Case Index
        Case 0:
            checkValores
            If errorValor = False Then
                add_Lista
                borrarDatosProd
            Else
                MsgBox "Falta información. Verfifique. ", vbInformation
            End If
        Case 1:
            If lista.Rows > 1 Then
                
                'if val(txtTotalPago.text) <>
                'Validar que si no he dado enter al pago no se ejecute la opción hasta validar
                
                txtTotalPago_KeyPress (13)
                                
                If valida = True Then
                
                    ques = MsgBox("Total: " & txtTotal.Text & vbCrLf & "Monto a pagar: " & txtTotalPago.Text & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
                                    
                    If ques = vbYes Then
                                            
                    Dim usuarios(40) As String
                    Dim ListUsuarios As String
                    Dim encuentra As Boolean
                    
                    ListUsuarios = ""
                    encuentra = False
                                        
                    For b1 = 1 To lista.Rows - 1
                        For c1 = 0 To b1 - 1
                            If usuarios(c1) = lista.TextMatrix(b1, 8) Then
                                encuentra = True
                                Exit For
                            End If
                        Next c1
                        If encuentra = False Then
                            usuarios(b1 - 1) = lista.TextMatrix(b1, 8)
                            ListUsuarios = ListUsuarios & vbCrLf & lista.TextMatrix(b1, 8)
                        End If
                        encuentra = False
                    Next b1
                    MsgBox "Usuarios asignados: " & vbCrLf & ListUsuarios, vbInformation, "Usuarios en apartados"
                        
                        
                        
                        tipoCobro = "APARTADOS1"
                        FRM_Cobro.txtTot.Text = txtTotalPago.Text
                        FRM_Cobro.Show vbModal
                    End If
                End If
                'crearApartado
                
            Else
                MsgBox "No se puede realizar la operación. Verifique.", vbInformation
            End If
        Case 2:
            cancelar
            borrarDatosProd
        Case 3:
            datosPagos ("False")
        Case 4:
            txtPgo(2).Text = FormatCurrency(Val(Format(txtPgo(2).Text, "General Number")))
            If Val(Format(txtPgo(2).Text, "General Number")) > Val(Format(txtPgo(1).Text, "General Number")) Then
                MsgBox "El valor proporcionado no es adecuado. Verfiique.", vbInformation
                txtPgo(2).SetFocus
            Else
                            
                ques = MsgBox("Total adeudo: " & txtPgo(1).Text & vbCrLf & "Monto a pagar: " & txtPgo(2).Text & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
                                
                If ques = vbYes Then
                    tipoCobro = "APARTADOS2"
                    FRM_Cobro.txtTot.Text = txtPgo(2).Text
                    FRM_Cobro.Show vbModal
                End If
            End If
        Case 5:
            ques = MsgBox("Imprimir información del apartado", vbYesNo + vbQuestion)
            If ques = vbYes Then
                infoApartado (txtInfoAprt)
            End If
    End Select
    
End Sub
Private Sub cancelar()
    lista.Rows = 1
    For b1 = 0 To 2
         lblClieId(b1).Caption = ""
    Next b1
    lblDatos(0).Caption = "Ninguno"
    lblDatos(2).Caption = "Ninguno"
    txtTotal.Text = "$0.00"
    txtTotalAnt.Text = "$0.00"
    txtTotalPago.Text = "$0.00"
    txtDescuento.Text = "$0.00"
    txtSubTotal.Text = "$0.00"

    imgFoto(0).Picture = LoadPicture("")
    imgFoto(2).Picture = LoadPicture("")
    txtClave(0).Text = ""
    txtClave(2).Text = ""
    
End Sub
Private Sub checkValores()
    errorValor = False
    
    If lblClieId(0).Caption = "" Then
        errorValor = True
    Else
        If lblUserId(0).Caption = "" Then
            errorValor = True
        Else
            If cmbDato(0).Text = "" Then
                errorValor = True
            Else
                If txtNumPagos.Text = "" Then
                    errorValor = True
                Else
                    If lblDatos(0).Caption = "Ninguno" Then
                        errorValor = True
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub add_Lista()
    If resProd.Fields("PROD_SERV") = "P" Then
'        If resProd.Fields("PROD_CANT") = 0 Then
'            MsgBox "No hay productos en existencia. Verifique.", vbInformation
'            Exit Sub
'        End If
        For b1 = 1 To lista.Rows - 1
            If lista.TextMatrix(b1, 1) = resProd.Fields("PROD_CODIGO") Then
                If resProd.Fields("PROD_CANT") > lista.TextMatrix(b1, 3) Then
                    lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                    'updateVentDet (b1)
                    checkPrecio (b1)
                    Exit Sub
                Else
                    MsgBox "No hay productos en existencia. Verifique.", vbInformation
                    Exit Sub
                End If
            End If
        Next b1
    Else
        If resProd.Fields("PROD_SERV") = "S" Then
            For b1 = 1 To lista.Rows - 1
                If lista.TextMatrix(b1, 1) = resProd.Fields("PROD_CODIGO") Then
                        lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                        'updateVentDet (b1)
                        checkPrecio (b1)
                        Exit Sub
                End If
            Next b1
        End If
    End If
    
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = resProd.Fields("TIPO_PROD")
    lista.TextMatrix(lista.Rows - 1, 1) = resProd.Fields("PROD_CODIGO")
    lista.TextMatrix(lista.Rows - 1, 2) = resProd.Fields("PROD_NOMBRE")
    lista.TextMatrix(lista.Rows - 1, 3) = "1"
    lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(txtAnticipo(2).Text)
    lista.TextMatrix(lista.Rows - 1, 6) = resProd.Fields("PROD_SERV")
    lista.TextMatrix(lista.Rows - 1, 7) = resProd.Fields("PROD_ID")
    lista.TextMatrix(lista.Rows - 1, 8) = lblDatos(1).Caption
    lista.TextMatrix(lista.Rows - 1, 9) = lblUserId(1).Caption
    lista.TextMatrix(lista.Rows - 1, 10) = lblUserId(0).Caption
    lista.TextMatrix(lista.Rows - 1, 11) = txtAnticipo(4).Text
    lista.TextMatrix(lista.Rows - 1, 12) = txtAnticipo(3).Text
    lista.TextMatrix(lista.Rows - 1, 13) = cmbDato(0).Text
    lista.TextMatrix(lista.Rows - 1, 14) = cmbDato(0).ItemData(cmbDato(0).ListIndex)
    lista.TextMatrix(lista.Rows - 1, 15) = Val(txtNumPagos.Text)
    lista.TextMatrix(lista.Rows - 1, 16) = FormatCurrency(Val(Format(txtAnticipo(5).Text, "General Number")) / Val(txtNumPagos.Text))
    lista.TextMatrix(lista.Rows - 1, 17) = txtAnticipo(1).Text
    lista.TextMatrix(lista.Rows - 1, 18) = txtAnticipo(0).Text
    ''''Estos quedan fijos para referencia
    lista.TextMatrix(lista.Rows - 1, 19) = txtAnticipo(1).Text
    lista.TextMatrix(lista.Rows - 1, 20) = txtAnticipo(0).Text
    checkPrecio (lista.Rows - 1)
    
        
    
    'addVentDet
End Sub
Public Sub apartado_crearApartado()
    Dim aprtId As Long
    Dim folioPago As Long
    Dim statusOper As String
    
    
    

    sql1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDTIPOID, VENT_VENDPERID, VENT_VENDTIPO, " & _
    "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO, vent_SubTotal, vent_Descuento, vent_Total, vent_Pagado, vent_Cambio, " & _
    "vent_PagoEfectivo, vent_PagoTarjeta, vent_PagoCheque, vent_FechaHora_Cobro) VALUES " & _
    "(NOW(), 'A', '" & lblUserId(1).Caption & "', '" & lblUserId(0).Caption & "', '" & lblUserId(2).Caption & "', " & _
    "'" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "', '" & Val(Format(txtSubTotal.Text, "General Number")) & "', '" & Val(Format(txtDescuento.Text, "General Number")) & "', '" & Val(Format(txtTotal.Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(4).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtCambio.Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(0).Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(1).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) & "', NOW())"
    'MsgBox SQL1
    con.Execute (sql1)
    
    sql1 = "select last_insert_id() folioId"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        folioPago = RES1.Fields("folioId")
        folioTicket = RES1.Fields("folioId")
    End If
    
    ''''MONEDERO
    If Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) > 0 Then
        sql1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
        "VALUES ('A', '" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "', '" & folioPago & "', " & _
        "'" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', '" & lblUserId(2).Caption & "',  '" & (Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) * (-1)) & "', 'E', NOW() ) "
        con.Execute (sql1)
        
    End If
    
        sql1 = "UPDATE VENTAS SET VENT_PUNTOSUSA = '" & Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) & "' " & _
        "WHERE VENT_IDFOLIO = '" & folioPago & "'"
        con.Execute (sql1)
    
    For b1 = 1 To lista.Rows - 1
        With lista

            sql1 = "INSERT INTO CAT_APARTADOS (aprt_ProdId, aprt_ProdServ, aprt_CliePerId, aprt_ClieId, aprt_CliePerTipo, aprt_ProdPrecio, " & _
            "aprt_ProdCantidad, aprt_FechaHora, aprt_MostId, aprt_MostPerId, aprt_MostPerTipo, aprt_Status, aprt_Periodo, aprt_PagosCant, aprt_Anticipo, " & _
            "aprt_Desc, aprt_userid, aprt_userperid, aprt_userpertipo, aprt_FolioPago, aprt_ProdNombre, aprt_ProdCodigo, aprt_Tipo) VALUES (" & _
            " '" & .TextMatrix(b1, 7) & "', '" & .TextMatrix(b1, 6) & "', '" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "', " & _
            " '" & Format(.TextMatrix(b1, 4), "General Number") & "', '" & .TextMatrix(b1, 3) & "', now(), '" & lblUserId(1).Caption & "', " & _
            " '" & lblUserId(0).Caption & "', '" & lblUserId(2).Caption & "', 'A', '" & .TextMatrix(b1, 14) & "', '" & .TextMatrix(b1, 15) & "', " & _
            "'" & Format(.TextMatrix(b1, 17), "General Number") & "', '" & Format(.TextMatrix(b1, 11), "General Number") & "', '" & .TextMatrix(b1, 9) & "', " & _
            "'" & .TextMatrix(b1, 10) & "', 'U', '" & folioPago & "', '" & .TextMatrix(b1, 2) & "', '" & .TextMatrix(b1, 1) & "', '" & Left(tipoAprt, 1) & "')"
            con.Execute (sql1)
        
            sql1 = "select last_insert_id() aprtId"
            Set RES1 = con.Execute(sql1)
            If Not RES1.EOF Then
                aprtId = RES1.Fields("aprtId")
            End If
            
            sql1 = "INSERT INTO PAGOS_APARTADOS (appg_aprtid, appg_prodid, appg_prodser, appg_clieid, appg_clieperid, appg_cliepertipo, appg_fechahora, " & _
            "appg_pago, appg_mostid, papg_mostperid, appg_mostpertipo, apPg_FolioVenta, appg_folioaprt) VALUES ('" & aprtId & "', " & _
            "'" & .TextMatrix(b1, 7) & "', '" & .TextMatrix(b1, 6) & "', '" & lblClieId(1).Caption & "', '" & lblClieId(0).Caption & "', " & _
            "'" & lblClieId(2).Caption & "', now(), '" & Format(.TextMatrix(b1, 17), "General Number") & "', '" & lblUserId(1).Caption & "' " & _
            ", '" & lblUserId(0).Caption & "', '" & lblUserId(2).Caption & "', '" & folioPago & "', '" & folioPago & "') "
            con.Execute (sql1)
            
            If .TextMatrix(b1, 6) = "P" Then
                sql1 = "UPDATE PRODUCTOS SET PROD_CANT = (PROD_CANT - '" & .TextMatrix(b1, 3) & "') " & _
                "WHERE PROD_ID = '" & .TextMatrix(b1, 7) & "' AND PROD_SERV = 'P'"
                con.Execute (sql1)
            End If
        
        End With
    Next b1

    MsgBox "Información guardada. Verifique.", vbInformation
    cargaValores
    cargaApartados
    
End Sub
Public Sub apartado_crearApartado2()
    Dim aprtId As Long
    Dim folioApartado As Long
    Dim folioPago As Long
    Dim pago As Double
    Dim pagoProd As Double
    Dim adeudo_prod As Double
    
    sql1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDTIPOID, VENT_VENDPERID, VENT_VENDTIPO, " & _
    "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO, vent_SubTotal, vent_Descuento, vent_Total, vent_Pagado, vent_Cambio, " & _
    "vent_PagoEfectivo, vent_PagoTarjeta, vent_PagoCheque, vent_FechaHora_Cobro) VALUES " & _
    "(NOW(), 'A', '" & lblUserId(1).Caption & "', '" & lblUserId(0).Caption & "', '" & lblUserId(2).Caption & "', " & _
    "'" & lblClieId(3).Caption & "', '" & lblClieId(4).Caption & "', '" & lblClieId(5).Caption & "', '" & Val(Format(FRM_Cobro.txtTot.Text, "General Number")) & "', '0', '" & Val(Format(FRM_Cobro.txtTot.Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(4).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtCambio.Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(0).Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(1).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) & "', NOW())"
    'MsgBox SQL1
    con.Execute (sql1)
    
    sql1 = "select last_insert_id() folioId"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        folioPago = RES1.Fields("folioId")
        folioTicket = RES1.Fields("folioId")
    End If
    
    ''''MONEDERO
    If Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) > 0 Then
        sql1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
        "VALUES ('A', '" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "', '" & folioPago & "', " & _
        "'" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', '" & lblUserId(2).Caption & "',  '" & (Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) * (-1)) & "', 'E', NOW() ) "
        con.Execute (sql1)
        
    End If
    
        sql1 = "UPDATE VENTAS SET VENT_PUNTOSUSA = '" & Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) & "' " & _
        "WHERE VENT_IDFOLIO = '" & folioPago & "'"
        con.Execute (sql1)
    
    
    With lista1
        b1 = .Row
        aprtId = lista1.TextMatrix(b1, 24)
        folioApartado = lista1.TextMatrix(b1, 0)
        pago = Val(Format(txtPgo(2).Text, "General Number"))
        pagoProd = 0
        
        sql1 = "SELECT APARTADO, APRT_PRODID PRODID, APRT_PRODSERV PRODSER, APRT_CLIEID CLIEID, APRT_CLIEPERID CLIEPERID, PAGOS_PROD, TOTAL_PROD, ADEUDO_pROD FROM VIEW_APARTADOS " & _
        "WHERE FOLIO = '" & folioApartado & "' "
        Set resApart = con.Execute(sql1)
        
        Do While Not resApart.EOF
            adeudo_prod = resApart.Fields("ADEUDO_PROD")
            If Val(pago) >= Val(adeudo_prod) Then
                pago = Val(pago) - Val(adeudo_prod)
                pagoProd = Val(adeudo_prod)
            Else
                If Val(pago) < Val(adeudo_prod) Then
                    pagoProd = Val(pago)
                End If
            End If
                                
             sql1 = "INSERT INTO PAGOS_APARTADOS (appg_aprtid, appg_prodid, appg_prodser, appg_clieid, appg_clieperid, appg_cliepertipo, appg_fechahora, " & _
             "appg_pago, appg_mostid, papg_mostperid, appg_mostpertipo, apPg_FolioVenta, appg_folioaprt) VALUES ('" & resApart.Fields("APARTADO") & "', " & _
             "'" & resApart.Fields("PRODID") & "', '" & resApart.Fields("PRODSER") & "', '" & resApart.Fields("CLIEID") & "', '" & resApart.Fields("CLIEPERID") & "', " & _
             "'C', now(), '" & pagoProd & "', '" & lblUserId(1).Caption & "' " & _
             ", '" & lblUserId(0).Caption & "', '" & lblUserId(2).Caption & "', '" & folioPago & "', '" & folioApartado & "') "
                'MsgBox SQL1
             con.Execute (sql1)
            
            resApart.MoveNext
        Loop
         
    End With
    
    MsgBox "Información guardada. Verifique.", vbInformation
    cargaValores
    cargaApartados
    datosPagos ("False")
    
End Sub
Public Sub checkPrecio(fila As Long)
    If lista.TextMatrix(fila, 11) = "" Then
        lista.TextMatrix(fila, 11) = FormatCurrency(0)
    End If
    lista.TextMatrix(fila, 5) = lista.TextMatrix(fila, 3) * lista.TextMatrix(fila, 4) - Val(Format(lista.TextMatrix(fila, 11), "General Number"))
    lista.TextMatrix(fila, 5) = FormatCurrency(lista.TextMatrix(fila, 5))
    lista.TextMatrix(fila, 16) = FormatCurrency(Val(Format(lista.TextMatrix(fila, 5), "General Number")) / Val(lista.TextMatrix(fila, 15)))
    checkPrecioFinal
End Sub


Private Sub cmbStatus_Click()
    cargaApartados
End Sub

Private Sub cmbUser_Click()
    If indx = 0 Then
        lista.TextMatrix(lista.Row, lista.Col) = cmbUser.Text
        cmbUser.Visible = False
        
        sql1 = "SELECT T4.PERTP_PER_ID, T4.PERTP_TIPO_ID, " & _
        "CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) USUARIO " & _
        "FROM PERSONA T2, PER_tIPO T4 " & _
        "WHERE T2.PER_ID = T4.PERTP_PER_ID AND T4.PERTP_STATUS = 'A' AND T4.PERTP_PER_TIPO = 'U' " & _
        "AND concat(T4.PERTP_PER_ID, T4.PERTP_TIPO_ID) = '" & cmbUser.ItemData(cmbUser.ListIndex) & "'"
        'MsgBox SQL1
        Set RES1 = con.Execute(sql1)
            
        If Not RES1.EOF Then
            lista.TextMatrix(lista.Row, 9) = RES1.Fields("PERTP_TIPO_ID")
            lista.TextMatrix(lista.Row, 10) = RES1.Fields("PERTP_PER_ID")
            'updateVentDet (lista.Row)
        End If
    End If
End Sub

Private Sub cmdBus_Click(Index As Integer)
    If Index = 0 Then
        mn_BusProd_Click
    Else
        If Index = 1 Then
            mn_BusClte_Click
        End If
    End If
End Sub

Private Sub cmdPeriodo_Click()
    periodoValor = "Apartado"
    CAT_Periodos.Show vbModal

End Sub



Private Sub Form_Load()
    valida = True
    cargaTipo
    cargaTitulos
    cargaValores
    cargaApartados
    datosPagos ("False")
    lista2.Rows = 1
End Sub
Private Sub cargaTipo()
    If tipoAprt = "CRED" Then
        Me.Caption = "Crédistos - Pagos dréditos"
        SSTab1.TabCaption(0) = "Lista de Créditos/Pagos"
        SSTab1.TabCaption(1) = "Nuevo crédito"
        mn_NewApar.Caption = "Nuevo crédito"
        mn_CancelAprt.Caption = "Cancelar crédito"
        mn_Part.Caption = "Crédito"
        Label1(7).Caption = "Lista de créditos"
        Label1(8).Caption = "Pagos realizados para el crédito seleccionado"
    Else
        If tipoAprt = "APRT" Then
            Me.Caption = "Apartados - Pagos apartados"
            SSTab1.TabCaption(0) = "Lista de Apartados/Pagos"
            SSTab1.TabCaption(1) = "Nuevo apartado"
            mn_Part.Caption = "Apartado"
            mn_NewApar.Caption = "Nuevo apartado"
            mn_CancelAprt.Caption = "Cancelar apartado"
            Label1(7).Caption = "Lista de apartados"
            Label1(8).Caption = "Pagos realizados para el apartado seleccionado"
        End If
    End If

    cmbStatus.Clear
    cmbStatus.AddItem "TODOS"
    cmbStatus.AddItem "VIGENTES"
    cmbStatus.AddItem "VENCIDOS"
    cmbStatus.AddItem "CANCELADOS"
    cmbStatus.AddItem "CONCLUIDOS"
    
    
        

End Sub
Private Sub datosPagos(tipo As String)

    For b1 = 0 To 2
        txtPgo(b1).Enabled = tipo
    Next b1

    cmBoton(3).Enabled = tipo
    cmBoton(4).Enabled = tipo
        
    If tipo = "False" Then
        txtPgo(0).Text = FormatCurrency(0)
        txtPgo(1).Text = FormatCurrency(0)
        txtPgo(2).Text = FormatCurrency(0)
        SSTab1.TabEnabled(1) = True
        mn_NewApar.Enabled = True
        
    End If
End Sub
Private Sub cargaTitulos()
    
    SSTab1.TabEnabled(2) = False
    SSTab1.TabCaption(2) = ""
    
    lista1.MergeCells = flexMergeRestrictColumns
    lista1.MergeCol(0) = True
    lista1.MergeCol(1) = True
    lista1.MergeCol(2) = True
    lista1.MergeCol(3) = True
    lista1.MergeCol(4) = True
    lista1.MergeCol(6) = True
    lista1.MergeCol(18) = True
    
    lista1.ColWidth(21) = 0
    lista1.ColWidth(22) = 0
    lista1.ColWidth(23) = 0
    lista1.ColWidth(24) = 0
    lista1.ColWidth(25) = 0
    lista1.ColWidth(26) = 0
    lista1.ColWidth(27) = 0
    lista1.ColWidth(28) = 0

    lista1.ColWidth(29) = 0
    lista1.ColWidth(30) = 0
    lista1.ColWidth(31) = 0
    lista1.ColWidth(32) = 0
    lista1.ColWidth(33) = 0
    lista1.ColWidth(34) = 0
    lista1.ColWidth(35) = 0
    lista1.ColWidth(36) = 0
    
    lista.Row = 1
    lista.WordWrap = True
    
    lista.ColWidth(0) = 0
    lista.ColWidth(6) = 0
    lista.ColWidth(7) = 0
    lista.ColWidth(9) = 0
    lista.ColWidth(10) = 0
    lista.ColWidth(19) = 0
    lista.ColWidth(20) = 0
    
    lista2.ColWidth(4) = 0
    lista2.ColWidth(5) = 0
    lista2.ColWidth(6) = 0
    lista2.ColWidth(7) = 0
    lista2.ColWidth(8) = 0
    lista2.ColWidth(9) = 0
    lista2.ColWidth(10) = 0
    lista2.ColWidth(11) = 0

End Sub
Private Sub cargaApartados()
    Dim filaFolio As String
    Dim tipofila As String
    Dim texto1 As String
    
    texto1 = ""
    If cmbStatus.Text = "CONCLUIDOS" Then
        texto1 = texto1 & "AND upper(STATUS) = 'CONCLUIDO' "
    Else
        If cmbStatus.Text = "VIGENTES" Then
            texto1 = texto1 & "AND upper(STATUS) = 'ACTIVO' AND TRANSCURRIDOS <= DIAS "
        Else
            If cmbStatus.Text = "VENCIDOS" Then
                texto1 = texto1 & "AND upper(STATUS) = 'ACTIVO' AND TRANSCURRIDOS > DIAS "
            Else
                If cmbStatus.Text = "CANCELADOS" Then
                    texto1 = texto1 & "AND upper(STATUS) = 'CANCELADO'"
                End If
            End If
        End If
    End If
               
    sql1 = "SELECT * FROM VIEW_APARTADOS WHERE FOLIO LIKE '%" & textBus(0).Text & "%' " & _
    "AND upper(PRODUCTO) LIKE upper('%" & textBus(1).Text & "%') " & _
    "AND upper(CLIENTE) LIKE upper('%" & textBus(2).Text & "%') " & _
    "AND LEFT(TIPO, 1) = '" & Left(tipoAprt, 1) & "' " & texto1 & _
    " ORDER BY FECHA DESC " & _
    "Limit 0, " & Val(textBus(4).Text) & ""
    
'    MsgBox SQL1
    Set resLista = con.Execute(sql1)
        
    lista1.MergeCells = flexMergeRestrictColumns
    tipofila = "1"
    lista1.Rows = 1
    lista1.Redraw = False
        
    Do While Not resLista.EOF
    
        filaFolio = resLista.Fields("FOLIO")
        If lista1.Rows - 1 > 0 Then
            If filaFolio <> lista1.TextMatrix(lista1.Rows - 1, 0) Then
                lista1.AddItem ""
                lista1.RowHeight(lista1.Rows - 1) = 0
                If tipofila = "1" Then
                    tipofila = "2"
                Else
                    tipofila = "1"
                End If
            End If
        End If
    
        lista1.AddItem ""
        lista1.TextMatrix(lista1.Rows - 1, 0) = resLista.Fields("FOLIO")
        lista1.TextMatrix(lista1.Rows - 1, 1) = Format(resLista.Fields("FECHA"), "Short Date") & " " & Format(resLista.Fields("FECHA"), "Short Time")
        lista1.TextMatrix(lista1.Rows - 1, 2) = FormatCurrency(resLista.Fields("TOTAL"))
        lista1.TextMatrix(lista1.Rows - 1, 3) = FormatCurrency(resLista.Fields("PAGADO"))
        lista1.TextMatrix(lista1.Rows - 1, 4) = FormatCurrency(resLista.Fields("ADEUDO"))
        
        lista1.TextMatrix(lista1.Rows - 1, 5) = resLista.Fields("STATUS")
        lista1.TextMatrix(lista1.Rows - 1, 6) = resLista.Fields("CLIENTE")
        lista1.TextMatrix(lista1.Rows - 1, 7) = resLista.Fields("PRODUCTO")
        lista1.TextMatrix(lista1.Rows - 1, 8) = resLista.Fields("CODIGO")
        lista1.TextMatrix(lista1.Rows - 1, 9) = FormatCurrency(resLista.Fields("PRECIO"))
        lista1.TextMatrix(lista1.Rows - 1, 10) = resLista.Fields("CANTIDAD")
        lista1.TextMatrix(lista1.Rows - 1, 11) = FormatCurrency(resLista.Fields("DESCUENTO"))
        lista1.TextMatrix(lista1.Rows - 1, 12) = FormatCurrency(resLista.Fields("TOTAL_PROD"))
        lista1.TextMatrix(lista1.Rows - 1, 13) = FormatCurrency(resLista.Fields("PAGOS_PROD"))
        lista1.TextMatrix(lista1.Rows - 1, 14) = FormatCurrency(resLista.Fields("ADEUDO_PROD"))
        lista1.TextMatrix(lista1.Rows - 1, 15) = resLista.Fields("PERIODO")
        lista1.TextMatrix(lista1.Rows - 1, 16) = resLista.Fields("DIAS")
        lista1.TextMatrix(lista1.Rows - 1, 17) = resLista.Fields("TRANSCURRIDOS")
        lista1.TextMatrix(lista1.Rows - 1, 18) = Format(resLista.Fields("LIQUIDACION"), "Short Date") & " " & Format(resLista.Fields("LIQUIDACION"), "Short Time")
        lista1.TextMatrix(lista1.Rows - 1, 19) = resLista.Fields("MOSTRADOR")
        lista1.TextMatrix(lista1.Rows - 1, 20) = resLista.Fields("VENDEDOR")
        lista1.TextMatrix(lista1.Rows - 1, 21) = resLista.Fields("APRT_PRODID")
        lista1.TextMatrix(lista1.Rows - 1, 22) = resLista.Fields("APRT_CLIEID")
        lista1.TextMatrix(lista1.Rows - 1, 23) = resLista.Fields("APRT_CLIEPERID")
        lista1.TextMatrix(lista1.Rows - 1, 24) = resLista.Fields("APARTADO")
        lista1.TextMatrix(lista1.Rows - 1, 25) = resLista.Fields("APRT_MOSTID")
        lista1.TextMatrix(lista1.Rows - 1, 26) = resLista.Fields("APRT_MOSTPERID")
        lista1.TextMatrix(lista1.Rows - 1, 27) = resLista.Fields("APRT_USERID")
        lista1.TextMatrix(lista1.Rows - 1, 28) = resLista.Fields("APRT_USERPERID")
        
        lista1.TextMatrix(lista1.Rows - 1, 29) = resLista.Fields("TEL1") & ""
        lista1.TextMatrix(lista1.Rows - 1, 30) = resLista.Fields("TEL2") & ""
        lista1.TextMatrix(lista1.Rows - 1, 31) = resLista.Fields("COLONIA") & ""
        lista1.TextMatrix(lista1.Rows - 1, 32) = resLista.Fields("CP") & ""
        lista1.TextMatrix(lista1.Rows - 1, 33) = resLista.Fields("CALLE") & ""
        lista1.TextMatrix(lista1.Rows - 1, 34) = resLista.Fields("NUME") & ""
        lista1.TextMatrix(lista1.Rows - 1, 35) = resLista.Fields("NUMI") & ""
        lista1.TextMatrix(lista1.Rows - 1, 36) = resLista.Fields("EMAIL") & ""
        
        
        lista1.Row = lista1.Rows - 1
        
        lista1.Col = 0
        lista1.CellFontSize = 11
        lista1.Col = 1
        lista1.CellFontSize = 11
        lista1.Col = 2
        lista1.CellFontSize = 11
        lista1.Col = 3
        lista1.CellFontSize = 11
        lista1.Col = 4
        lista1.CellFontSize = 11
        lista1.Col = 5
        lista1.CellFontSize = 11

        If tipofila = "2" Then
            For b1 = 0 To 36
                lista1.Col = b1
                lista1.CellBackColor = &HFFFFC0
            Next b1
        End If
        
    If resLista.Fields("STATUS") = "CANCELADO" Then
        lista1.Row = lista1.Rows - 1
        lista1.Col = 0
        lista1.CellForeColor = &H40C0&
        lista1.Col = 1
        lista1.CellForeColor = &H40C0&
        lista1.Col = 2
        lista1.CellForeColor = &H40C0&
        lista1.Col = 3
        lista1.CellForeColor = &H40C0&
        lista1.Col = 4
        lista1.CellForeColor = &H40C0&
        lista1.Col = 5
        lista1.CellForeColor = &H40C0&
    Else
        If Val(resLista.Fields("DIAS")) < resLista.Fields("TRANSCURRIDOS") And resLista.Fields("STATUS") <> "CONCLUIDO" Then
            lista1.Row = lista1.Rows - 1
            lista1.Col = 0
            lista1.CellForeColor = vbRed
            lista1.Col = 1
            lista1.CellForeColor = vbRed
            lista1.Col = 2
            lista1.CellForeColor = vbRed
            lista1.Col = 3
            lista1.CellForeColor = vbRed
            lista1.Col = 4
            lista1.CellForeColor = vbRed
            lista1.Col = 5
            lista1.CellForeColor = vbRed
        Else
            If (Val(resLista.Fields("TRANSCURRIDOS")) - Val(resLista.Fields("DIAS"))) <= 5 And (Val(resLista.Fields("TRANSCURRIDOS")) - Val(resLista.Fields("DIAS"))) >= 0 Then
                lista1.Row = lista1.Rows - 1
                lista1.Col = 0
                lista1.CellForeColor = &H808000
                lista1.Col = 1
                lista1.CellForeColor = &H808000
                lista1.Col = 2
                lista1.CellForeColor = &H808000
                lista1.Col = 3
                lista1.CellForeColor = &H808000
                lista1.Col = 4
                lista1.CellForeColor = &H808000
                lista1.Col = 5
                lista1.CellForeColor = &H808000
            Else
                If resLista.Fields("STATUS") = "CONCLUIDO" Then
                    lista1.Row = lista1.Rows - 1
                    lista1.Col = 0
                    lista1.CellForeColor = &H8000&
                    lista1.Col = 1
                    lista1.CellForeColor = &H8000&
                    lista1.Col = 2
                    lista1.CellForeColor = &H8000&
                    lista1.Col = 3
                    lista1.CellForeColor = &H8000&
                    lista1.Col = 4
                    lista1.CellForeColor = &H8000&
                    lista1.Col = 5
                    lista1.CellForeColor = &H8000&
                End If
            End If
        End If
    
        For b1 = 0 To 36
            lista1.MergeCol(b1) = True
        Next b1
    
    End If
        resLista.MoveNext
    Loop
    lista1.Redraw = True


End Sub

Private Sub cargaValores()

    SSTab1.Tab = 0
    mn_NuevoAprt.Enabled = False
    Aprt_cargaPeriodo

    lblDatos(1).Caption = FRM_Menu.menuBarra2.Panels(5).Text
    lblUserId(0).Caption = FRM_Menu.menuBarra2.Panels(7).Text
    lblUserId(1).Caption = FRM_Menu.menuBarra2.Panels(8).Text
    lblUserId(2).Caption = "U"

    cancelar
    
End Sub

Public Sub checkPrecioFinal()
    Dim total, DESCUENTO, ANTICIPO
'    ''''Para los descuentos
'    If lista.Rows = 1 Then
'        descGral = False
'        txtDesc(0).Text = "0"
'        txtDesc(1).Text = "0"
'        txtDesc(0).Locked = False
'        txtDesc(1).Locked = False
'    End If
    
    total = 0
    DESCUENTO = 0
    ANTICIPO = 0
    'DESCUENTO = Val(Format(txtDesc(0).Text, "General Number"))
    'MsgBox DESCUENTO
    
    For b1 = 1 To lista.Rows - 1
        If lista.TextMatrix(b1, 0) <> "Descuento" Then
            total = total + Val(Format(lista.TextMatrix(b1, 4), "General Number"))
            DESCUENTO = DESCUENTO + Val(Format(lista.TextMatrix(b1, 11), "General Number"))
            'MsgBox lista.TextMatrix(b1, 5)
        Else
            DESCUENTO = DESCUENTO + Val(Format(lista.TextMatrix(b1, 11), "General Number"))
            'MsgBox lista.TextMatrix(b1, 5)
        End If
        ANTICIPO = ANTICIPO + Val(Format(lista.TextMatrix(b1, 17), "General Number"))
    Next b1
    'MsgBox DESCUENTO
    txtSubTotal.Text = FormatCurrency(total)
    txtDescuento.Text = FormatCurrency(DESCUENTO)
    txtTotal.Text = FormatCurrency(total - DESCUENTO)
    txtTotalAnt.Text = FormatCurrency(ANTICIPO)
    txtTotalPago.Text = FormatCurrency(ANTICIPO)
    'tOper(0).Text = lista.Rows - 1
    
End Sub


Public Sub Aprt_cargaPeriodo()
    validar = False
    
    sql1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS FROM CAT_PERIODO"
    Set RES1 = con.Execute(sql1)
    cmbDato(0).Clear
    Do While Not RES1.EOF
        cmbDato(0).AddItem RES1.Fields("CTPR_PERIODO")
        cmbDato(0).ItemData(cmbDato(0).ListCount - 1) = RES1.Fields("CTID_PERIODO")
        RES1.MoveNext
    Loop
    
    If cmbDato(0).ListCount > 0 Then
        cmbDato(0).ListIndex = 0
    End If
    
End Sub

Private Sub cargaUsuarios()
    sql1 = "SELECT T4.PERTP_PER_ID, T4.PERTP_TIPO_ID, concat(T4.PERTP_PER_ID, T4.PERTP_TIPO_ID) USERID,  " & _
    "CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) USUARIO " & _
    "FROM PERSONA T2, PER_tIPO T4 " & _
    "WHERE T2.PER_ID = T4.PERTP_PER_ID AND T4.PERTP_STATUS = 'A' AND T4.PERTP_PER_TIPO = 'U' ORDER BY CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO)"
    Set RES1 = con.Execute(sql1)
    cmbUser.Clear
    Do While Not RES1.EOF
        cmbUser.AddItem RES1.Fields("USUARIO")
        cmbUser.ItemData(cmbUser.ListCount - 1) = RES1.Fields("USERID")
        RES1.MoveNext
    Loop
    
    
    
End Sub

Private Sub lista_DblClick()
Select Case lista.Col
    Case 8:
            cargaUsuarios
            cmbUser.Top = lista.CellTop + lista.Top
            cmbUser.Left = lista.CellLeft + lista.Left
            'cmbUser.Height = lista.CellHeight
            cmbUser.width = lista.CellWidth
            cmbUser.Text = lista.TextMatrix(lista.Row, lista.Col)
            cmbUser.Visible = True
            cmbUser.SetFocus

End Select
    
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista.Rows > 1 Then
        If Button = vbRightButton Then
            PopupMenu mn_SubMenu, vbPopupMenuLeftAlign
        End If
    End If
End Sub

Private Sub lista1_Click()
    lista2.Rows = 1
    
    sql1 = "SELECT * FROM VIEW_PAGOS_APARTTOTAL WHERE FOLIO_APRT = '" & lista1.TextMatrix(lista1.Row, 0) & "'"
    Set resPagos = con.Execute(sql1)
    
    Do While Not resPagos.EOF
        lista2.AddItem ""
        lista2.TextMatrix(lista2.Rows - 1, 0) = resPagos.Fields("FOLIO_VENTA")
        lista2.TextMatrix(lista2.Rows - 1, 1) = resPagos.Fields("FECHA")
        lista2.TextMatrix(lista2.Rows - 1, 2) = FormatCurrency(resPagos.Fields("PAGO"))
        lista2.TextMatrix(lista2.Rows - 1, 3) = resPagos.Fields("MOSTRADOR")
        lista2.TextMatrix(lista2.Rows - 1, 4) = resPagos.Fields("SUBTOTAL")
        lista2.TextMatrix(lista2.Rows - 1, 5) = resPagos.Fields("DESCUENTO")
        lista2.TextMatrix(lista2.Rows - 1, 6) = resPagos.Fields("TOTAL")
        lista2.TextMatrix(lista2.Rows - 1, 7) = resPagos.Fields("PAGADO")
        lista2.TextMatrix(lista2.Rows - 1, 8) = resPagos.Fields("CAMBIO")
        lista2.TextMatrix(lista2.Rows - 1, 9) = resPagos.Fields("EFECTIVO")
        lista2.TextMatrix(lista2.Rows - 1, 10) = resPagos.Fields("TARJETA")
        lista2.TextMatrix(lista2.Rows - 1, 11) = resPagos.Fields("CHEQUE")
        
        resPagos.MoveNext
    Loop
    
    If lista1.Rows > 2 Then
        infoClie (lista1.TextMatrix(lista1.Row, 23))
    End If
    
    
    lblClieId(3).Caption = lista1.TextMatrix(lista1.Row, 23)    'resClie.Fields("PER_ID")
    lblClieId(4).Caption = lista1.TextMatrix(lista1.Row, 22)     'resClie.Fields("PERTP_TIPO_ID")
    lblClieId(5).Caption = "C"
    
End Sub
Private Sub infoClie(clieId As Long)
    txtInfoAprt.Text = ""
    
'    SQL1 = "SELECT PER_ID, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_TEL1, PER_TEL2, PER_DIR_COLONIA, PER_DIR_CP, PER_DIR_CALLE, " & _
'    "PER_DIR_NUM_EXT, PER_DIR_NUM_INT, PER_EMAIL FROM PERSONA where PER_ID = '" & clieId & "'"
'    'MsgBox SQL1
'    Set resInfoCLie = con.Execute(SQL1)
    
'    If Not resInfoCLie.EOF Then
'        txtInfoAprt.Text = "Cliente: " & resInfoCLie.Fields("PER_Nombre") & "  " & resInfoCLie.Fields("per_paterno") & "  " & resInfoCLie.Fields("per_materno") & vbCrLf & _
'        "Domicilio: " & resInfoCLie.Fields("PER_DIR_COLONIA") & " " & resInfoCLie.Fields("PER_DIR_CP") & "  " & resInfoCLie.Fields("PER_DIR_CALLE") & "  " & resInfoCLie.Fields("PER_DIR_NUM_EXT") & " " & resInfoCLie.Fields("PER_DIR_NUM_INT") & vbCrLf & _
'        "Teléfono: " & resInfoCLie.Fields("PER_TEL1") & "   " & resInfoCLie.Fields("PER_TEL2") & vbCrLf & _
'        "Email: " & resInfoCLie.Fields("PER_EMAIL")
'    End If
    With lista1
        txtInfoAprt.Text = "Cliente: " & .TextMatrix(.Row, 6) & vbCrLf & _
        "Domicilio: " & .TextMatrix(.Row, 31) & " " & .TextMatrix(.Row, 32) & "  " & .TextMatrix(.Row, 33) & "  " & .TextMatrix(.Row, 34) & " " & .TextMatrix(.Row, 35) & vbCrLf & _
        "Teléfono(s): " & .TextMatrix(.Row, 29) & " " & .TextMatrix(.Row, 30) & vbCrLf & _
        "Email: " & .TextMatrix(.Row, 36)
    End With
    
    txtInfoAprt = txtInfoAprt & vbCrLf & "----- Apartado -----" & vbCrLf & _
    "Fecha/Hora: " & lista1.TextMatrix(lista1.Row, 1) & vbCrLf & _
    "Total apartado: " & lista1.TextMatrix(lista1.Row, 2) & vbCrLf & _
    "Pagos: " & lista1.TextMatrix(lista1.Row, 3) & vbCrLf & _
    "Faltante " & lista1.TextMatrix(lista1.Row, 4)
    
  
End Sub

Private Sub Lista1_GotFocus()
    ConScroll lista1
End Sub

Private Sub Lista1_LostFocus()
    SinScroll lista1
End Sub

Private Sub Lista1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista1.Rows > 1 Then
        If Button = vbRightButton Then
            If lista1.TextMatrix(lista1.Row, 5) = "CANCELADO" Then
                mn_RePag.Enabled = False
                mn_CancelAprt.Enabled = False
            Else
                If lista1.TextMatrix(lista1.Row, 5) = "CONCLUIDO" Then
                    mn_RePag.Enabled = False
                    mn_CancelAprt.Enabled = False
                Else
                    mn_RePag.Enabled = True
                    mn_CancelAprt.Enabled = True
                End If
            End If
            mn_RePag.Caption = "Realizar pago apartado folio: " & lista1.TextMatrix(lista1.MouseRow, 0) & " cliente: " & lista1.TextMatrix(lista1.MouseRow, 6) & ". Faltante: " & lista1.TextMatrix(lista1.Row, 4)
            mn_CancelAprt.Caption = "Cancelar apartado folio: " & lista1.TextMatrix(lista1.MouseRow, 0) & " cliente: " & lista1.TextMatrix(lista1.MouseRow, 6) & ". Faltante: " & lista1.TextMatrix(lista1.Row, 4)
            mn_PrintGral.Caption = "Imprimir ticket general del apartado folio " & lista1.TextMatrix(lista1.MouseRow, 0)

            lista1.Row = lista1.MouseRow
            PopupMenu mn_Part, vbPopupMenuLeftAlign
        End If
    End If
End Sub

Private Sub Lista1_SelChange()
    lista1_Click
End Sub

Private Sub lista2_Click()
''''aaaa
End Sub

Private Sub Lista2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista2.Rows > 1 Then
        If Button = vbRightButton Then
            mn_PrintPago.Caption = "Imprimir ticket de pago apartado folio: " & lista2.TextMatrix(lista2.MouseRow, 0) & " fecha:  " & lista2.TextMatrix(lista2.MouseRow, 1)
            lista2.Row = lista2.MouseRow
            PopupMenu mn_PgosAprt, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_BusClte_Click()
    tipoBusqueda = "C"
    BUSQ_Usuarios.Caption = "Búsqueda de clientes."
    modBusqueda = "Apartado"
    BUSQ_Usuarios.Show vbModal

End Sub

Private Sub mn_BusProd_Click()
    modBusqueda = "Apartado"
    BUSQ_ProdSer.Show vbModal

End Sub

Private Sub mn_Cancel_Click()
    Dim ques As String
    
    ques = MsgBox("Cancelar " & lista.TextMatrix(lista.Row, 2), vbYesNo + vbQuestion)
    
    If ques = vbYes Then
        If lista.Rows <= 2 Then
            lista.Rows = 1
        Else
            lista.RemoveItem (lista.Row)
        End If
            checkPrecioFinal
    End If
    
End Sub

Private Sub mn_CancelAprt_Click()
    Dim ques As String
    
    ques = MsgBox("¿Cancelar apartado folio: " & lista1.TextMatrix(lista1.Row, 0) & " fecha " & lista1.TextMatrix(lista1.Row, 1) & "? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        If Val(Format(lista1.TextMatrix(lista1.Row, 3), "General Number")) > 0 Then
            ques = MsgBox("Existe un abono por " & lista1.TextMatrix(lista1.Row, 3) & vbCrLf & vbCrLf & "Este monto se asignará a monedero electrónico." & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
            If ques = vbNo Then
                Exit Sub
            Else
                sql1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
                "VALUES ('A', '" & lblClieId(3).Caption & "', '" & lblClieId(4).Caption & "', '" & lblClieId(5).Caption & "', '" & lista1.TextMatrix(lista1.Row, 0) & "', " & _
                "'" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', '" & lblUserId(2).Caption & "',  '" & (Val(Format(lista1.TextMatrix(lista1.Row, 3), "General Number"))) & "', 'R', NOW() ) "
                MsgBox sql1
                con.Execute (sql1)
            End If
        End If
        sql1 = "UPDATE CAT_APARTADOS SET APRT_STATUS = 'C' WHERE APRT_FOLIOPAGO = '" & lista1.TextMatrix(lista1.Row, 0) & "'"
        con.Execute (sql1)
        cargaApartados
    End If
    
End Sub

Private Sub mn_Export_Click()
    Dim ques As String
    ques = MsgBox("¿Exportar la lista a excel?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(lista1)
    End If

End Sub

Private Sub mn_NewApar_Click()
    SSTab1.Tab = 1
    txtClave(0).SetFocus
End Sub

Private Sub mn_PagoTodos_Click()
Dim ques As String
Dim textProd As String
Dim resInfo As Recordset

    sql1 = "SELECT CONCAT(PRODUCTO, ' ', CODIGO, ' ', FALTANTE) Info " & _
    "FROM VIEW_APARTADOS WHERE CLIE_ID = '" & lista1.TextMatrix(lista1.Row, 5) & "' AND STATUS = 'ACTIVO'"
    Set resInfo = con.Execute(sql1)
    
    textProd = ""
    
    Do While Not resInfo.EOF
        textProd = textProd & vbCrLf & resInfo.Fields("Info")
        resInfo.MoveNext
    Loop

    ques = MsgBox("Realizar el pago de anticipo por: " & vbCrLf & textProd, vbYesNo + vbQuestion)


End Sub

Private Sub mn_PrintGral_Click()
    Dim ques As String
    ques = MsgBox("¿Imprimir ticket de operación folio " & lista1.TextMatrix(lista1.Row, 0) & vbCrLf & vbCrLf & "Fecha de pago:  " & lista1.TextMatrix(lista1.Row, 1) & "?", vbYesNo + vbQuestion)
    If ques = vbYes Then
            If tipoAprt = "APRT" Then
                notaApartado (lista1.TextMatrix(lista1.Row, 26))
            Else
                If tipoAprt = "CRED" Then
                    notaCredito (lista1.TextMatrix(lista1.Row, 26))
                End If
            End If
    End If

End Sub

Private Sub mn_PrintPago_Click()
    Dim ques As String
    ques = MsgBox("¿Imprimir ticket para el pago realizado folio " & lista2.TextMatrix(lista2.Row, 0) & " fecha: " & lista2.TextMatrix(lista2.Row, 1), vbYesNo + vbQuestion)
    If ques = vbYes Then
            If tipoAprt = "APRT" Then
                notaApartado (lista2.TextMatrix(lista2.Row, 0))
            Else
                If tipoAprt = "CRED" Then
                    notaCredito (lista2.TextMatrix(lista2.Row, 0))
                End If
            End If
    End If
End Sub

Private Sub mn_RePag_Click()
    On Error Resume Next
    
    Dim ques As String
    
    If permEdit = "SI" Then
        If lista1.TextMatrix(lista1.Row, 1) <> "CONCLUIDO" Then
            ques = MsgBox("Realizar pago para la operación folio: " & lista1.TextMatrix(lista1.Row, 0) & vbCrLf & vbCrLf & _
            "Cliente: " & lista1.TextMatrix(lista1.Row, 6) & vbCrLf & vbCrLf & "Faltante: " & lista1.TextMatrix(lista1.Row, 4), vbQuestion + vbYesNo)
            If ques = vbYes Then
                txtPgo(0).Text = lista1.TextMatrix(lista1.Row, 2)
                txtPgo(1).Text = lista1.TextMatrix(lista1.Row, 4)
                txtPgo(2).Text = lista1.TextMatrix(lista1.Row, 4)
                MsgBox "Por favor escribe la cantidad del pago a realizar en la sección de pago en la parte inferior. ", vbInformation
                datosPagos ("True")
                txtPgo(2).SetFocus
                SSTab1.TabEnabled(1) = False
                mn_NewApar.Enabled = False
            End If
        Else
            MsgBox "La operación no se puede realizar. Verfique el estatus. ", vbInformation
        End If
    Else
        MsgBox "Opción no disponible. Verifique", vbInformation
    End If

End Sub

Private Sub NO_Click()
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        If permAdd = "SI" Then
            mn_NuevoAprt.Enabled = True
            mn_Part.Enabled = False
            txtClave(0).SetFocus
        Else
            SSTab1.Tab = 0
            MsgBox "Opción no disponible. Verifique", vbInformation
        End If
    Else
        mn_NuevoAprt.Enabled = False
        mn_Part.Enabled = True
    End If
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaApartados
    End If

End Sub

Private Sub TimeSize_Timer()
    TimeSize.Enabled = False
    SSTab1.width = Me.width - 200
    lista1.width = Me.width - 500
    lista.width = Me.width - 500
    Image2(0).width = Me.width
    Image2(0).height = Me.height
    Image2(1).width = Me.width
    Image2(1).height = Me.height
    
End Sub

Private Sub txtAnticipo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Index = 3 Then
            valorDescuento ("Porcentaje")
        Else
            If Index = 4 Then
                valorDescuento ("Cantidad")
            Else
                If Index = 0 Then
                    valorDescuento ("PorcentajeAnti")
                Else
                    If Index = 1 Then
                        valorDescuento ("CantidadAnti")
                    End If
                End If
            End If
        End If
        valorAnticipo
    End If
End Sub

Private Sub txtClave_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Index = 0 Then
           txtClave(0).Text = Replace(txtClave(0).Text, "'", "-")
           If Left(txtClave(0).Text, 1) = " " Then
                txtClave(0).Text = Right(txtClave(0).Text, (Len(txtClave(0).Text) - 1))
           End If
          aprt_checkProducto
        Else
            If Index = 2 Then
                aprt_checkCliente
            End If
        End If
    End If
End Sub
Public Sub aprt_checkProducto()
    On Error Resume Next
    
    Dim porcentajeCredito As Double
    
    
    If tipoAprt = "CRED" Then
        sql1 = "SELECT SUC_CreditoPorcentaje PORCENTAJE FROM SUCURSAL "
        Set resProd = con.Execute(sql1)
            
        If Not resProd.EOF Then
            porcentajeCredito = Val(resProd.Fields("PORCENTAJE"))
        End If
    Else
        porcentajeCredito = "0"
    End If
    
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRECIO_COSTO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, prod_PrecioDesc, prod_AplicaDesc " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "PROD_CODIGO = '" & txtClave(0).Text & "' AND PROD_STATUS = 'A'"
    Set resProd = con.Execute(sql1)
    Dim b1 As Long
    If Not resProd.EOF Then
        lblDatos(0).Caption = resProd.Fields("PROD_NOMBRE")
        If IsNull(resProd.Fields("PROD_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write resProd.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            imgFoto(0).Picture = LoadPicture("")
        End If
        If CheckPC.value = Checked Then
            If IsNull(resProd.Fields("PROD_PRECIO_COSTO")) = False Then
                txtAnticipo(2).Text = FormatCurrency(Val(resProd("prod_precio_costo")) + (resProd("prod_precio_costo") * (porcentajeCredito / 100)))
            Else
                MsgBox "El producto seleciconado no cuenta con precio de costo. " & vbCrLf & cvbrlf & _
                "Se asignará su precio de venta.", vbExclamation
                txtAnticipo(2).Text = FormatCurrency(Val(resProd("prod_precio")) + (resProd("prod_precio") * (porcentajeCredito / 100)))
            End If
        Else
            txtAnticipo(2).Text = FormatCurrency(Val(resProd("prod_precio")) + (resProd("prod_precio") * (porcentajeCredito / 100)))
        End If
        valorDescuento ("Porcentaje")
        valorAnticipo
    Else
        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation

    End If
    

End Sub
Private Sub valorDescuento(tipo As String)
    If tipo = "Porcentaje" Then
'        txtAnticipo(5).Text = (Val(Format(txtAnticipo(2).Text, "General Number"))) - (Val(Format(txtAnticipo(2).Text, "General Number")) * (Val(txtAnticipo(3).Text) / 100))
'        txtAnticipo(5).Text = FormatCurrency(txtAnticipo(5).Text)
'        txtAnticipo(4).Text = (Val(Format(txtAnticipo(2).Text, "General Number")) * (Val(txtAnticipo(3).Text) / 100))
'        txtAnticipo(4).Text = FormatCurrency(txtAnticipo(4).Text)
        
        If Val(resProd.Fields("prod_preciodesc")) = 0 Then
            txtAnticipo(4).Text = 0
        Else
            txtAnticipo(4).Text = FormatCurrency(resProd("prod_precio") - resProd.Fields("prod_preciodesc"))
        End If
        valorDescuento ("Cantidad")
    Else
        If tipo = "Cantidad" Then
            If Val(Format(txtAnticipo(4).Text, "General Number")) > Val(Format(txtAnticipo(2).Text, "General Number")) Then
                MsgBox "La cantidad no puede ser mayor o igual al precio del producto. Verifique.", vbInformation
            Else
                txtAnticipo(5).Text = (Val(Format(txtAnticipo(2).Text, "General Number"))) - (Val(Format(txtAnticipo(4).Text, "General Number")))
                txtAnticipo(5).Text = FormatCurrency(txtAnticipo(5).Text)
                txtAnticipo(3).Text = ((Val(Format(txtAnticipo(4).Text, "General Number")) * 100) / Val(Format(txtAnticipo(2).Text, "General Number")))
                txtAnticipo(3).Text = Round((txtAnticipo(3).Text), 2)
                txtAnticipo(4).Text = FormatCurrency(txtAnticipo(4).Text)
            End If
        Else
            If tipo = "PorcentajeAnti" Then
                txtAnticipo(1).Text = (Val(Format(txtAnticipo(5).Text, "General Number")) * (Val(txtAnticipo(0).Text) / 100))
                txtAnticipo(1).Text = FormatCurrency(txtAnticipo(1).Text)
            Else
                If tipo = "CantidadAnti" Then
                    If Val(Format(txtAnticipo(1).Text, "General Number")) > Val(Format(txtAnticipo(5).Text, "General Number")) Then
                        MsgBox "La cantidad no puede ser mayor o igual al precio actual del producto. Verifique.", vbInformation
                    Else
                        txtAnticipo(0).Text = ((Val(Format(txtAnticipo(1).Text, "General Number")) * 100) / Val(Format(txtAnticipo(5).Text, "General Number")))
                        'txtAnticipo(0).Text = Round((txtAnticipo(0).Text), 2)
                        txtAnticipo(1).Text = FormatCurrency(txtAnticipo(1).Text)

                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub valorAnticipo()
    txtAnticipo(1).Text = (Val(Format(txtAnticipo(5).Text, "General Number")) * (Val(txtAnticipo(0).Text) / 100))
    
    txtAnticipo(0).Text = Round((txtAnticipo(0).Text), 2)
    txtAnticipo(1).Text = FormatCurrency(txtAnticipo(1).Text)

End Sub
Public Sub aprt_checkCliente()
    
    On Error Resume Next
    
    modBusqueda = "Apartado"
    
    sql1 = "SELECT PERTP_USUARIO,  PERTP_CODIGO_MEMBRESIA, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_PER_TIPO, PERTP_TIPO_ID, CTPT_TIPO, PER_ID, PER_FOTO, (SELECT T4.TOTAL FROM VIEW_MONEDERO_CLIENTES T4 WHERE T1.PER_ID = T4.PER_ID) TOTAL  " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'C' " & _
    "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'C' " & _
    "AND T2.PERTP_CODIGO_MEMBRESIA = '" & txtClave(2).Text & "'"
    Set resClie = con.Execute(sql1)
    
    If Not resClie.EOF Then
        lblDatos(2).Caption = resClie.Fields("PER_NOMBRE") & " " & resClie.Fields("PER_PATERNO") & " " & resClie.Fields("PER_MATERNO")
        lblClieId(0).Caption = resClie.Fields("PER_ID")
        lblClieId(1).Caption = resClie.Fields("PERTP_TIPO_ID")
        lblClieId(2).Caption = resClie.Fields("PERTP_PER_TIPO")
                               
        If IsNull(resClie.Fields("total")) Then
        lblDatos(6).Caption = FormatCurrency(0)
        Else
        lblDatos(6).Caption = FormatCurrency(Val(resClie.Fields("TOTAL")))
        End If
        
        If IsNull(resClie.Fields("PER_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write resClie.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempClie.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(2).Picture = LoadPicture(direccionSistema & "\Temp\TempClie.dat")
        Else
            imgFoto(2).Picture = LoadPicture("")
        End If
    Else
        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation
    End If
        
End Sub

Private Sub txtPgo_GotFocus(Index As Integer)
    
    If Index = 2 Then
        txtPgo(2).SelStart = 0
        txtPgo(2).SelLength = Len(txtPgo(2).Text)
    End If
    
End Sub

Private Sub txtPgo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        If KeyAscii = 13 Then
            txtPgo(2).Text = FormatCurrency(Val(Val(Format(txtPgo(2).Text, "General Number"))))
            cmBoton_Click (4)
        End If
    End If
End Sub

Private Sub txtTotalPago_GotFocus()
    txtTotalPago.SelStart = 0
    txtTotalPago.SelLength = Len(txtTotalPago.Text)
End Sub

Private Sub txtTotalPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTotalAnt.Text = "0"
        pago_Anticipo
    End If
End Sub


Private Sub pago_Anticipo()
valida = True
    If Val(Format(txtTotalPago.Text, "General Number")) < Val(Format(txtTotalAnt.Text, "General Number")) Then
        valida = False
        MsgBox "El pago de anticipo no puede ser menor al establecido como mínimo de anticipo. Verifique.", vbInformation
        txtTotalPago.SetFocus
    Else
        If Val(Format(txtTotalPago.Text, "General Number")) >= Val(Format(txtTotal.Text, "General Number")) Then
            valida = False
            MsgBox "El pago de anticipo no puede ser igual o mayor al pago total. Verifique.", vbInformation
            txtTotalPago.SetFocus
        Else
            valida = True
            ajustarAnticipo
        End If
    End If
End Sub

Private Sub ajustarAnticipo()
    
    Dim pagoFila, pagoDif
    
'    pagoFila = (Val(Val(Format(txtTotalPago.Text, "General Number"))) / (lista.Rows - 1))
    pagoDif = 0
    pagoDif = (Val(Val(Format(txtTotalPago.Text, "General Number"))) - (Val(Format(txtTotalAnt.Text, "General Number"))))
'    MsgBox pagoDif
    For b1 = 1 To lista.Rows - 1
        lista.TextMatrix(b1, 17) = lista.TextMatrix(b1, 19)
        
       ' MsgBox "ok"
        lista.TextMatrix(b1, 17) = Val(Format(lista.TextMatrix(b1, 17), "General Number")) + Val(pagoDif)
       ' MsgBox "Ok1"
        If Val(Format(lista.TextMatrix(b1, 17), "General Number")) > Val(Format(lista.TextMatrix(b1, 4), "General Number")) Then
            pagoDif = Val(Format(lista.TextMatrix(b1, 17), "General Number")) - Val(Format(lista.TextMatrix(b1, 5), "General Number"))
            lista.TextMatrix(b1, 17) = FormatCurrency(lista.TextMatrix(b1, 5))
            lista.TextMatrix(b1, 18) = ((Val(Format(lista.TextMatrix(b1, 17), "General Number")) * 100) / Val(Format(lista.TextMatrix(b1, 5), "General Number")))
            lista.TextMatrix(b1, 18) = Round(Val(lista.TextMatrix(b1, 18)), 2)
        Else
            lista.TextMatrix(b1, 17) = FormatCurrency(lista.TextMatrix(b1, 17))
            lista.TextMatrix(b1, 18) = ((Val(Format(lista.TextMatrix(b1, 17), "General Number")) * 100) / Val(Format(lista.TextMatrix(b1, 5), "General Number")))
            lista.TextMatrix(b1, 18) = Round(Val(lista.TextMatrix(b1, 18)), 2)
            Exit For
        End If
    Next b1

    txtTotalPago.Text = FormatCurrency(Val(Val(Format(txtTotalPago.Text, "General Number"))))
End Sub
