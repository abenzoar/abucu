VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Productos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   10260
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   19155
   Icon            =   "FRM_Productos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   19155
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   -120
      TabIndex        =   36
      Top             =   0
      Width           =   19215
      _ExtentX        =   33893
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   697
      BackColor       =   16777215
      TabCaption(0)   =   "  Lista de productos"
      TabPicture(0)   =   "FRM_Productos.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2(1)"
      Tab(0).Control(1)=   "Shape1(7)"
      Tab(0).Control(2)=   "lBus(3)"
      Tab(0).Control(3)=   "lBus(2)"
      Tab(0).Control(4)=   "lBus(1)"
      Tab(0).Control(5)=   "lBus(0)"
      Tab(0).Control(6)=   "lInfo(10)"
      Tab(0).Control(7)=   "lBus(4)"
      Tab(0).Control(8)=   "Borde(15)"
      Tab(0).Control(9)=   "Borde(16)"
      Tab(0).Control(10)=   "Borde(17)"
      Tab(0).Control(11)=   "Borde(18)"
      Tab(0).Control(12)=   "Shape1(6)"
      Tab(0).Control(13)=   "lProd(16)"
      Tab(0).Control(14)=   "lBus(5)"
      Tab(0).Control(15)=   "ListaUsers"
      Tab(0).Control(16)=   "textBus(1)"
      Tab(0).Control(17)=   "textBus(0)"
      Tab(0).Control(18)=   "Picture1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textBus(4)"
      Tab(0).Control(20)=   "cmdAll"
      Tab(0).Control(21)=   "cmbProd(5)"
      Tab(0).Control(22)=   "cmbProd(6)"
      Tab(0).Control(23)=   "Check1"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "  Datos generales"
      TabPicture(1)   =   "FRM_Productos.frx":0E64
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Borde(12)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lProd(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lProd(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lProd(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lProd(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lProd(51)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lProd(61)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lProd(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lProd(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lProd(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lProd(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lProd(11)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lUsuario(25)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbStatus"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lProd(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lProd(10)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Borde(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Borde(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Borde(3)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lProd(12)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Shape1(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Borde(4)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Borde(5)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Borde(6)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Borde(7)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Borde(8)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lProd(13)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Shape1(3)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lProd(14)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Shape1(4)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Borde(9)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Borde(10)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Borde(11)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Borde(13)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lProd(15)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "lProd(9)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Borde(14)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Borde(0)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "lProd(17)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "lProd(18)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lProd(19)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Borde(19)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lProd(20)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Shape1(8)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Shape1(10)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Borde(20)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lProd(21)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "lProd(22)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Borde(21)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Shape1(9)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lProd(23)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Borde(22)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Borde(23)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "lProd(24)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "lProd(26)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Borde(24)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "lProd(27)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Shape1(11)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "lProd(28)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Borde(25)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "cMd1"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "txtProd(0)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "txtProd(1)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "txtProd(7)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "txtProd(2)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "txtProd(3)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "cmbProd(0)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "cmbProd(1)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "txtProd(6)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "cmbProd(2)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "txtProd(4)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "txtProd(5)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "cmbProd(3)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "cmBoton(2)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "cmBoton(1)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "cmBoton(0)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "cmbProd(4)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "cmd_Marca"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "cmdTipo"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "cmBoton(3)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "cmBoton(4)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "cmBoton(5)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "cmBoton(6)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "time1"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "txtProd(8)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Option1(0)"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "Option1(1)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "Option1(2)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "txtProd(9)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "txtProd(10)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "cmBoton(7)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "cmdProveed"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "cmdPresentacion"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "cmbProd(7)"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "listDependiente"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "cmBoton(8)"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "cmBoton(9)"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "iFoto"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "cmbProd(8)"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "cmbProd(9)"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "txtProd(11)"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "txtProd(12)"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "cmbProd(10)"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).ControlCount=   106
      TabCaption(2)   =   "    Productos imágenes"
      TabPicture(2)   =   "FRM_Productos.frx":13FE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Timer1"
      Tab(2).Control(1)=   "ListProd1"
      Tab(2).Control(2)=   "listprod2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "    Dependientes - Recetas"
      TabPicture(3)   =   "FRM_Productos.frx":1998
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lista_Receta"
      Tab(3).Control(1)=   "Shape1(5)"
      Tab(3).Control(2)=   "lProd(25)"
      Tab(3).Control(3)=   "Image2(2)"
      Tab(3).ControlCount=   4
      Begin VB.ComboBox cmbProd 
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
         Index           =   10
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3720
         Width           =   2055
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "0"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12360
         TabIndex        =   93
         Top             =   -5000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   -61080
         TabIndex        =   89
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtProd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   14760
         MaxLength       =   9
         TabIndex        =   88
         Text            =   "0"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   9
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   5040
         Width           =   2055
      End
      Begin VB.ComboBox cmbProd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   14880
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   5400
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.PictureBox iFoto2 
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   10800
         ScaleHeight     =   1515
         ScaleWidth      =   1395
         TabIndex        =   85
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox iFoto 
         Height          =   2415
         Left            =   16920
         ScaleHeight     =   2355
         ScaleWidth      =   2115
         TabIndex        =   84
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -62400
         Top             =   360
      End
      Begin MSFlexGridLib.MSFlexGrid ListProd1 
         Height          =   9015
         Left            =   -74760
         TabIndex        =   82
         Top             =   720
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   15901
         _Version        =   393216
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         GridColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   13680
         Picture         =   "FRM_Productos.frx":2272
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   12840
         Picture         =   "FRM_Productos.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid listDependiente 
         Height          =   2415
         Left            =   7920
         TabIndex        =   78
         Top             =   6120
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         WordWrap        =   -1  'True
         FormatString    =   $"FRM_Productos.frx":2D86
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
      Begin VB.ComboBox cmbProd 
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
         Index           =   7
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdPresentacion 
         Caption         =   "Presentacion"
         Height          =   255
         Left            =   17280
         TabIndex        =   74
         Top             =   8640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdProveed 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   17280
         TabIndex        =   73
         Top             =   8880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrar presentaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   14760
         Picture         =   "FRM_Productos.frx":2E0F
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   2055
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
         Left            =   3360
         MaxLength       =   17
         TabIndex        =   2
         Top             =   2400
         Width           =   2775
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
         Left            =   7920
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   11640
         MaskColor       =   &H0000FFFF&
         TabIndex        =   20
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   10080
         MaskColor       =   &H0000FFFF&
         TabIndex        =   19
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7965
         MaskColor       =   &H0000FFFF&
         TabIndex        =   18
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   6
         Left            =   -66120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         ToolTipText     =   "Selecciona la marca a la que pertenece el producto, o agrega o edita las existentes"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   5
         Left            =   -69240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Selecciona el tipo de clasificación a la que pertenece el producto, o agrega o edita los existentes"
         Top             =   1080
         Width           =   2895
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
         Index           =   8
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "0"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Timer time1 
         Interval        =   500
         Left            =   6840
         Top             =   120
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrar provedores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   5160
         Picture         =   "FRM_Productos.frx":3399
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7920
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrar tipos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   5160
         Picture         =   "FRM_Productos.frx":3923
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7080
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrar marcas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   5160
         Picture         =   "FRM_Productos.frx":3EAD
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ver todos"
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
         Left            =   -59040
         Picture         =   "FRM_Productos.frx":4437
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox textBus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -60000
         TabIndex        =   34
         Text            =   "0"
         ToolTipText     =   "Número de productos que se mostraran en la lista"
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -63240
         ScaleHeight     =   1095
         ScaleWidth      =   4815
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   8880
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar y agregar otro producto"
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
         Index           =   3
         Left            =   3840
         Picture         =   "FRM_Productos.frx":49C1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   8760
         Width           =   3255
      End
      Begin VB.CommandButton cmdTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   17280
         TabIndex        =   58
         Top             =   9480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Marca 
         Caption         =   "Marca"
         Height          =   375
         Left            =   17280
         TabIndex        =   57
         Top             =   9120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   4
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   9240
         Width           =   3975
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar e ir a la lista"
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
         Index           =   0
         Left            =   240
         Picture         =   "FRM_Productos.frx":528B
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   8760
         Width           =   3375
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
         Index           =   1
         Left            =   16440
         Picture         =   "FRM_Productos.frx":5B55
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   8760
         Width           =   2055
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
         Left            =   16920
         Picture         =   "FRM_Productos.frx":641F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4320
         Width           =   1815
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   3
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   8040
         Width           =   4575
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
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   11
         Text            =   "0"
         Top             =   5040
         Width           =   855
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "0"
         Top             =   5040
         Width           =   855
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   2
         Left            =   10560
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2160
         Width           =   3975
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
         Left            =   7920
         MaxLength       =   25
         TabIndex        =   21
         Text            =   "0"
         ToolTipText     =   "Ingresa la medida o presentación actual del producto"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   1
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   7200
         Width           =   4575
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   0
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   6360
         Width           =   4575
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
         Left            =   360
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "0"
         Top             =   3720
         Width           =   1455
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
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   9
         Text            =   "0"
         Top             =   5040
         Width           =   1215
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
         Height          =   2055
         Index           =   7
         Left            =   7920
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   3120
         Width           =   6615
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
         Left            =   360
         MaxLength       =   17
         TabIndex        =   1
         Top             =   2400
         Width           =   2775
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
         Left            =   2880
         MaxLength       =   65
         TabIndex        =   0
         Top             =   1560
         Width           =   4575
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
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -72480
         TabIndex        =   31
         Top             =   1080
         Width           =   3015
      End
      Begin MSFlexGridLib.MSFlexGrid ListaUsers 
         Height          =   7935
         Left            =   -74640
         TabIndex        =   37
         Top             =   1680
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   22
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Productos.frx":69A9
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
      Begin MSComDlg.CommonDialog cMd1 
         Left            =   15360
         Top             =   9600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid ListaSel 
         Height          =   1335
         Left            =   6840
         TabIndex        =   81
         Top             =   -5000
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Productos.frx":6B65
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
      Begin MSFlexGridLib.MSFlexGrid listprod2 
         Height          =   9015
         Left            =   -67680
         TabIndex        =   83
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   15901
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         GridColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Lista_Receta 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   91
         Top             =   960
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Productos.frx":6D1C
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
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   25
         Left            =   5160
         Top             =   3720
         Width           =   2085
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicar descuento"
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
         Index           =   28
         Left            =   5160
         TabIndex        =   96
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   11
         Left            =   360
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Precios"
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
         Index           =   27
         Left            =   360
         TabIndex        =   95
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   24
         Left            =   3600
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento *"
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
         Left            =   3600
         TabIndex        =   94
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   5
         Left            =   -74880
         Top             =   840
         Width           =   11655
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Recetas"
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
         Index           =   25
         Left            =   -74880
         TabIndex        =   92
         Top             =   600
         Width           =   2895
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
         Left            =   -61920
         TabIndex        =   90
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicado en venta"
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
         Index           =   24
         Left            =   360
         TabIndex        =   87
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   23
         Left            =   360
         Top             =   5040
         Width           =   2085
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   2475
         Index           =   22
         Left            =   7920
         Top             =   6120
         Width           =   10845
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Dependencia de productos"
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
         Index           =   23
         Left            =   7920
         TabIndex        =   77
         Top             =   5640
         Width           =   5535
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   9
         Left            =   7920
         Top             =   5880
         Width           =   10815
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   21
         Left            =   360
         Top             =   1560
         Width           =   2355
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de producto *"
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
         Left            =   360
         TabIndex        =   75
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Código / Clave del Proveedor"
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
         Left            =   3360
         TabIndex        =   72
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   20
         Left            =   3360
         Top             =   2400
         Width           =   2805
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   10
         Left            =   7920
         Top             =   9000
         Width           =   3975
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   8
         Left            =   12720
         Top             =   5880
         Width           =   15
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Mayoreo *"
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
         Left            =   7920
         TabIndex        =   71
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   19
         Left            =   8040
         Top             =   3720
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Talla especifica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   11880
         TabIndex        =   70
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Talla gral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   10320
         TabIndex        =   69
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad medida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   8280
         TabIndex        =   68
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   2880
         Top             =   1545
         Width           =   4605
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
         TabIndex        =   67
         Top             =   480
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   -74640
         Top             =   720
         Width           =   11655
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   18
         Left            =   -66120
         Top             =   1080
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   17
         Left            =   -69240
         Top             =   1080
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   16
         Left            =   -72480
         Top             =   1080
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   15
         Left            =   -74640
         Top             =   1080
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   14
         Left            =   2040
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo *"
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
         Left            =   2040
         TabIndex        =   66
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Index           =   15
         Left            =   7920
         TabIndex        =   65
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   2115
         Index           =   13
         Left            =   7920
         Top             =   3120
         Width           =   6645
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   11
         Left            =   360
         Top             =   8040
         Width           =   4605
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   10
         Left            =   360
         Top             =   7200
         Width           =   4605
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   9
         Left            =   360
         Top             =   6360
         Width           =   4605
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   4
         Left            =   360
         Top             =   5880
         Width           =   6975
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Características"
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
         Index           =   14
         Left            =   360
         TabIndex        =   64
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   3
         Left            =   7920
         Top             =   1080
         Width           =   11175
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Presentación"
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
         Index           =   13
         Left            =   7920
         TabIndex        =   63
         Top             =   840
         Width           =   1695
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   8
         Left            =   7920
         Top             =   9240
         Width           =   4005
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   7
         Left            =   10560
         Top             =   2160
         Width           =   4005
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   6
         Left            =   7920
         Top             =   2160
         Width           =   2445
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   5
         Left            =   5160
         Top             =   5040
         Width           =   885
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   4080
         Top             =   5040
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   2
         Left            =   360
         Top             =   1080
         Width           =   7215
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Index           =   12
         Left            =   360
         TabIndex        =   62
         Top             =   840
         Width           =   2895
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   2640
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   360
         Top             =   3720
         Width           =   1485
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   1
         Left            =   360
         Top             =   2400
         Width           =   2805
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario Sotck"
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
         Index           =   10
         Left            =   360
         TabIndex        =   61
         Top             =   4320
         Width           =   2055
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
         TabIndex        =   60
         Top             =   840
         Width           =   735
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
         Left            =   -74520
         TabIndex        =   56
         Top             =   9840
         Width           =   15375
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Status *"
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
         Index           =   6
         Left            =   7920
         TabIndex        =   55
         Top             =   8760
         Width           =   2415
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
         Left            =   240
         TabIndex        =   54
         Top             =   9840
         Width           =   5535
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen o foto "
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
         Left            =   16920
         TabIndex        =   53
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor *"
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
         Left            =   360
         TabIndex        =   52
         Top             =   7680
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Máximo *"
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
         Left            =   5160
         TabIndex        =   51
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Mínimo *"
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
         Left            =   4080
         TabIndex        =   50
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de presentación"
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
         Left            =   10560
         TabIndex        =   49
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor presentación"
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
         Left            =   7920
         TabIndex        =   48
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo*"
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
         Index           =   61
         Left            =   360
         TabIndex        =   47
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca *"
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
         Index           =   51
         Left            =   360
         TabIndex        =   46
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Venta *"
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
         TabIndex        =   45
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual *"
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
         Left            =   2640
         TabIndex        =   44
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Código / Clave de venta *"
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
         TabIndex        =   43
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del producto *"
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
         Left            =   2880
         TabIndex        =   42
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Código producto"
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
         Left            =   -74640
         TabIndex        =   41
         Top             =   840
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
         Height          =   255
         Index           =   1
         Left            =   -72480
         TabIndex        =   40
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
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
         Left            =   -69240
         TabIndex        =   39
         Top             =   840
         Width           =   1815
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
         Index           =   3
         Left            =   -66120
         TabIndex        =   38
         Top             =   840
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   360
         Top             =   4560
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   240
         Top             =   9840
         Width           =   18255
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   2475
         Index           =   12
         Left            =   16920
         Top             =   1680
         Width           =   2205
      End
      Begin VB.Image Image2 
         Height          =   9735
         Index           =   0
         Left            =   0
         Picture         =   "FRM_Productos.frx":6E12
         Stretch         =   -1  'True
         Top             =   480
         Width           =   19215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   420
         Index           =   7
         Left            =   -74640
         Top             =   9720
         Width           =   15615
      End
      Begin VB.Image Image2 
         Height          =   9735
         Index           =   1
         Left            =   -75000
         Picture         =   "FRM_Productos.frx":13E52
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
      Begin VB.Image Image2 
         Height          =   9735
         Index           =   2
         Left            =   -75120
         Picture         =   "FRM_Productos.frx":20E92
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
   End
   Begin VB.Menu mn_Prod 
      Caption         =   "Productos"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mn_Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mn_lineAdd 
         Caption         =   "-"
      End
      Begin VB.Menu mn_AddSame 
         Caption         =   "Agregar uno igual"
      End
   End
   Begin VB.Menu mn_Catalogos 
      Caption         =   "Catálogos"
      Begin VB.Menu mn_Marca 
         Caption         =   "Marca"
      End
      Begin VB.Menu mn_TipoProd 
         Caption         =   "Tipo de Producto"
      End
      Begin VB.Menu mn_Proveedor 
         Caption         =   "Proveedor"
      End
      Begin VB.Menu mn_CatTipoPresen 
         Caption         =   "Tipo de presentación"
      End
      Begin VB.Menu mn_Etiquetas 
         Caption         =   "Etiquetas"
      End
   End
   Begin VB.Menu mn_Options 
      Caption         =   "Opciones"
      Begin VB.Menu mn_PrintLust 
         Caption         =   "Exportar lista"
         Begin VB.Menu mn_PrintGroup 
            Caption         =   "Exportar lista agrupado por cantidad"
         End
         Begin VB.Menu mn_PrintAll 
            Caption         =   "Exportar lista por cada producto"
         End
      End
      Begin VB.Menu mn_PrintCodigos 
         Caption         =   "Imprimir etiquetas"
      End
      Begin VB.Menu mn_Seleccion 
         Caption         =   "Seleccionar lista"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FRM_Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sql1 As String
    Dim RES1 As Recordset
    Dim RES2 As Recordset
    Dim RES3 As Recordset
    Dim RESTIPO_PROD As Recordset
    Dim RES_PROD As Recordset
    Dim checkError As Boolean
    Dim prodId As String
    Dim Id As Long
    Dim save As Boolean
    Dim mayus As Boolean
    Dim activaSeleccion As Boolean
    Dim tipoId(90, 3)
    Dim tipoValor(90, 3)
    
Private Sub CargaGeneral()
    SSTab1.Tab = 0

    listDependiente.Rows = 1
    listDependiente.WordWrap = True
    'listDependiente.ColAlignment(2) = flexAlignLeftTop
'    listDependiente.ColWidth(5) = 0
'    listDependiente.RowHeight(0) = 500
    
    cmbProd(7).Clear
    cmbProd(7).AddItem "UNICO"
    cmbProd(7).ItemData(cmbProd(7).ListCount - 1) = 1
    cmbProd(7).AddItem "DEPENDIENTE"
    cmbProd(7).ItemData(cmbProd(7).ListCount - 1) = 2
    cmbProd(7).ListIndex = 0
    
    cmbProd(9).Clear
    cmbProd(9).AddItem "SI"
    cmbProd(9).AddItem "NO"
    cmbProd(9).ListIndex = 0
    
    cmbProd(10).Clear
    cmbProd(10).AddItem "SI"
    cmbProd(10).AddItem "NO"
    cmbProd(10).ListIndex = 1
    
    If mesas = True Then
        SSTab1.TabCaption(3) = "   Recetas"
        Lista_Receta.TextMatrix(0, 1) = "Receta"
        Lista_Receta.TextMatrix(0, 3) = "Ingrediente"
        Lista_Receta.TextMatrix(0, 6) = "Costo Ingrediente"
        Lista_Receta.TextMatrix(0, 7) = "Costo equivalente"
        Lista_Receta.TextMatrix(0, 8) = "Código Ingrediente"
    Else
        SSTab1.TabCaption(3) = "   Dependencias"
        Lista_Receta.TextMatrix(0, 1) = "Producto"
        Lista_Receta.TextMatrix(0, 3) = "Producto dependiente"
        Lista_Receta.TextMatrix(0, 6) = "Costo producto"
        Lista_Receta.TextMatrix(0, 7) = "Costo equivalente"
        Lista_Receta.TextMatrix(0, 8) = "Codigo dependiente"
    End If
    
    Option1(0).value = True
    Option1_Click (0)

    iFoto.Picture = LoadPicture("")
    SSTab1.TabEnabled(1) = False
    cargaMarca
    cargaTipoProd

    cargaProveedor
    For b1 = 4 To 7
        cmBoton(b1).Visible = False
    Next b1

    cmbProd(4).Clear
    cmbProd(4).AddItem "ACTIVO"
    cmbProd(4).AddItem "INACTIVO"
    cmbProd(4).ListIndex = 0

    sql1 = "SELECT SUC_PRODLISTA FROM SUCURSAL "
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        textBus(4).Text = RES1.Fields("SUC_PRODLISTA")
    End If
 
End Sub
Private Sub cargaPresentacion(tipoPresen As String)

    sql1 = ("SELECT CTPS_ID, CTPS_NOMBRE FROM CAT_PRESENTACION WHERE CTPS_TIPO = '" & tipoPresen & "' ORDER BY CTPS_NOMBRE ASC")
    Set RES1 = con.Execute(sql1)
    
    cmbProd(2).Clear
    Do While Not RES1.EOF
        cmbProd(2).AddItem RES1.Fields("CTPS_NOMBRE")
        cmbProd(2).ItemData(cmbProd(2).ListCount - 1) = RES1.Fields("CTPS_ID")
        RES1.MoveNext
    Loop

End Sub
Private Sub cancelar()
    limpiarCampos
    CargaGeneral
'    cargaTipoUsuario
    cargaLista

End Sub
Private Sub cargaTipoProd()

    sql1 = "SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'P' ORDER BY CTPT_TIPO"
    Set RES1 = con.Execute(sql1)
    
    cmbProd(1).Clear
    cmbProd(6).Clear
    cmbProd(6).AddItem "TODOS"
    Do While Not RES1.EOF
        cmbProd(1).AddItem RES1.Fields("CTPT_TIPO")
        cmbProd(1).ItemData(cmbProd(1).ListCount - 1) = RES1.Fields("CTPT_ID")
        cmbProd(6).AddItem RES1.Fields("CTPT_TIPO")
        cmbProd(6).ItemData(cmbProd(1).ListCount - 1) = RES1.Fields("CTPT_ID")
        RES1.MoveNext
    Loop


End Sub
Private Sub cargaProveedor()

    sql1 = "SELECT PER_ID, CONCAT(PER_ALIAS, ' - ', PER_NOMBRE, ' ', PER_PATERNO, ' ', PER_MATERNO) PROVEEDOR " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'V'  "
    Set RES1 = con.Execute(sql1)
    
    cmbProd(3).Clear
    Do While Not RES1.EOF
        cmbProd(3).AddItem RES1.Fields("PROVEEDOR")
        cmbProd(3).ItemData(cmbProd(3).ListCount - 1) = RES1.Fields("PER_ID")
        RES1.MoveNext
    Loop
    If cmbProd(3).ListCount > 0 Then
        cmbProd(3).ListIndex = 0
    End If

End Sub


Private Sub cargaMarca()
    sql1 = "SELECT CTMR_ID, CTMR_MARCA FROM CAT_MARCA ORDER BY CTMR_MARCA"
    Set RES1 = con.Execute(sql1)
    
    cmbProd(0).Clear
    cmbProd(5).Clear
    cmbProd(5).AddItem "TODOS"
    
    Do While Not RES1.EOF
        cmbProd(0).AddItem RES1.Fields("CTMR_MARCA")
        cmbProd(0).ItemData(cmbProd(0).ListCount - 1) = RES1.Fields("CTMR_ID")
        
        cmbProd(5).AddItem RES1.Fields("CTMR_MARCA")
        cmbProd(5).ItemData(cmbProd(5).ListCount - 1) = RES1.Fields("CTMR_ID")
        RES1.MoveNext
    Loop
    
End Sub

Private Sub limpiarCampos()
    
    For b1 = 0 To 10
        txtProd(b1).Text = ""
        If b1 > 1 And b1 <> 7 Then
            txtProd(b1).Text = "0"
        End If
    Next b1
    
    For b1 = 0 To 4
        cmbProd(b1).Clear
    Next b1

    

End Sub


Private Sub Check1_Click()
    cargaLista
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            checarCampos
            If checkError = False Then
                crearProducto
            Else
                MsgBox "Falta información. Por favor verifique. ", vbExclamation
            End If
    
        Case 1:
            Dim ques As String
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        Case 2:
            buscarImagen
        Case 3:
            checarCampos
            If checkError = False Then
                crearProducto
                agregarNuevo
                
                sql1 = "SELECT MAX(PROD_ID) + 1 CLAVE FROM PRODUCTOS"
                Set RES1 = con.Execute(sql1)
                
                If Not RES1.EOF Then
                    txtProd(1).Text = "P" & Format(RES1.Fields("CLAVE"), "000000000")
                End If
                
                
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
            End If
        Case 4:
            mn_Marca_Click
        Case 5:
            
            mn_TipoProd_Click
                        
        Case 6:
            mn_Proveedor_Click
        Case 7:
            mn_CatTipoPresen_Click
        Case 8:
            If listDependiente.TextMatrix(listDependiente.Rows - 1, 1) <> "" Then
                listDependiente.AddItem ""
                listDependiente.Row = listDependiente.Rows - 1
                listDependiente.Col = 6
                listDependiente.CellFontName = "Wingdings"
                listDependiente.CellFontBold = True
                listDependiente.CellFontSize = 12
                listDependiente.TextMatrix(listDependiente.Rows - 1, 6) = Chr(168)
            Else
                MsgBox "Por favor concluya con el producto anterior", vbInformation
            End If
        Case 9:
    
        If lbStatus.Caption = "Estatus: Agregando producto" Then
    
            num = 0
            For b1 = 1 To listDependiente.Rows - 1
                num = num + 1
                If listDependiente.TextMatrix(num, 6) = Chr(254) Then
                    If listDependiente.Rows > 2 Then
                        listDependiente.RemoveItem (num)
                        num = num - 1
                    Else
                        listDependiente.Rows = 1
                        b1 = 1
                    End If
                End If
            Next b1
        Else
            num = 0
            For b1 = 1 To listDependiente.Rows - 1
                num = num + 1
                If listDependiente.TextMatrix(num, 6) = Chr(254) Then
                    If listDependiente.Rows > 2 Then
                        sql1 = "DELETE FROM PRODUCTO_DEPENDIENTE WHERE PROD_ID = '" & Id & "' " & _
                        "AND PROD_DEPEN_ID = '" & listDependiente.TextMatrix(num, 4) & "'"
                        'MsgBox sql1
                        con.Execute (sql1)
                        listDependiente.RemoveItem (num)
                        num = num - 1
                    Else
                        listDependiente.Rows = 1
                        b1 = 1
                    End If
                End If
            Next b1
            
        End If
    
    End Select

End Sub
Private Sub buscarImagen()
    cMd1.DialogTitle = "Buscando imagen..."
    cMd1.Filter = "Archivos de Imagenes|*.jpg*||*.bmp*||*.gif*||*.wmf*||*.emf*|"
    cMd1.FileName = ""
    cMd1.ShowOpen
    If cMd1.FileName <> "" Then
        mostrarImagen
    End If
End Sub
Private Sub mostrarImagen()
    With cMd1
        iFoto2.Picture = LoadPicture(.FileName)
        iFoto.AutoRedraw = True
        iFoto.PaintPicture iFoto2.Picture, _
            iFoto.ScaleLeft, iFoto.ScaleTop, _
                iFoto.ScaleWidth, iFoto.ScaleHeight, _
            iFoto2.ScaleLeft, iFoto2.ScaleTop, _
                iFoto2.ScaleWidth, iFoto2.ScaleHeight
        iFoto.Picture = iFoto.Image
        
    End With
End Sub



Private Sub crearProducto()

    'On Error Resume Next
    
    Dim status As String
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    Dim uniMed As String
    Dim prodPerId As String
    Dim prodTipoId As String
    Dim prodTipo As String
    
    prodPerId = "NULL"
    prodTipoId = "NULL"
    prodTipo = "NULL"
    
    If cmbProd(3).Text <> "" Then
        sql1 = "SELECT PERTP_TIPO_ID, PERTP_PER_ID, PERTP_PER_TIPO FROM PER_TIPO WHERE PERTP_PER_ID = '" & cmbProd(3).ItemData(cmbProd(3).ListIndex) & "' "
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
            prodPerId = RES1.Fields("PERTP_PER_ID")
            prodTipoId = RES1.Fields("PERTP_TIPO_ID")
            prodTipo = RES1.Fields("PERTP_PER_TIPO")
        End If
    End If
    
    
    If cmbProd(2).Text = "" Then
        uniMed = "null"
    Else
        uniMed = cmbProd(2).ItemData(cmbProd(2).ListIndex)
    End If
        
    status = Left(cmbProd(4).Text, 1)
    
    If lbStatus.Caption = "Estatus: Agregando producto" Then
        
    
        Err.Clear
        sql1 = "INSERT INTO PRODUCTOS " & _
        "(PROD_NOMBRE, PROD_CODIGO, PROD_DESCRIPCION, PROD_SERV, PROD_CANT, PROD_PRECIO, PROD_MARCA,  " & _
        "PROD_TIPO, PROD_SUBTIPO, PROD_STATUS, PROD_PRESENTACION, PROD_UNIMED_PRESENT, PROD_STOCK_MIN, " & _
        "PROD_STOCK_MAX, PROD_PERALTA_ID, PROD_PERALTA_TIPOID, PROD_PERALTA_TIPO, PROD_ALTA_FECHA, prod_proveedor, " & _
        "prod_provtipo, prod_provsubtipo, PROD_PRECIO_COSTO, PROD_PRECIO_MAY, PROD_CODIGO_PROV, prod_Dependiente, PROD_INVENTARIO, prod_PrecioDesc, prod_AplicaDesc) VALUES (" & _
        "'" & txtProd(0).Text & "', '" & txtProd(1).Text & "', '" & txtProd(7).Text & "', 'P', '" & txtProd(2).Text & "', '" & txtProd(3).Text & "', '" & cmbProd(0).ItemData(cmbProd(0).ListIndex) & "',  " & _
        "'" & cmbProd(1).ItemData(cmbProd(1).ListIndex) & "', 'P', '" & status & "', '" & txtProd(6).Text & "', " & uniMed & ", '" & txtProd(4).Text & "', '" & txtProd(5).Text & "', " & _
        "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW(), '" & prodPerId & "', '" & prodTipoId & "', '" & prodTipo & "', '" & txtProd(8).Text & "', '" & txtProd(9).Text & "', '" & txtProd(10).Text & "', '" & Left(cmbProd(7).Text, 1) & "', '" & Left(cmbProd(9).Text, 1) & "', '" & txtProd(12).Text & "', '" & Left(cmbProd(10).Text, 1) & "' )"
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() prodid"
        Set RES1 = con.Execute(sql1)
        If Not RES1.EOF Then
            prodId = RES1.Fields("prodid")
        End If
                
        With listDependiente
            For b1 = 1 To .Rows - 1
                sql1 = "INSERT INTO PRODUCTO_DEPENDIENTE (PROD_ID, PROD_SERV, PROD_CANT_EQUI, PROD_DEPEN_ID, PROD_dEPEN_SERV) VALUES " & _
                " ('" & prodId & "', 'P', '" & .TextMatrix(b1, 2) & "', '" & .TextMatrix(b1, 5) & "', 'P' )"
                con.Execute (sql1)
            Next b1
        End With
                
        prodId = txtProd(1).Text
    Else
        sql1 = "UPDATE PRODUCTOS SET PROD_NOMBRE = '" & txtProd(0).Text & "', " & _
        "PROD_CODIGO = '" & txtProd(1).Text & "', " & _
        "PROD_DESCRIPCION = '" & txtProd(7).Text & "', " & _
        "PROD_CANT = '" & txtProd(2).Text & "', " & _
        "PROD_PRECIO = '" & txtProd(3).Text & "', " & _
        "PROD_MARCA = " & cmbProd(0).ItemData(cmbProd(0).ListIndex) & ", " & _
        "PROD_TIPO = " & cmbProd(1).ItemData(cmbProd(1).ListIndex) & ", " & _
        "PROD_STATUS = '" & status & "', " & _
        "PROD_PRESENTACION = '" & txtProd(6).Text & "', " & _
        "PROD_UNIMED_PRESENT = '" & cmbProd(2).ItemData(cmbProd(2).ListIndex) & "', " & _
        "PROD_STOCK_MIN = '" & txtProd(4).Text & "', " & _
        "PROD_STOCK_MAX = '" & txtProd(5).Text & "', " & _
        "PROD_PRECIO_COSTO = '" & txtProd(8).Text & "', " & _
        "PROD_PROVEEDOR = '" & prodPerId & "', " & _
        "PROD_PROVTIPO = '" & prodTipoId & "', " & _
        "PROD_PROVSUBTIPO = '" & prodTipo & "', " & _
        "PROD_PRECIO_MAY = '" & txtProd(9).Text & "', " & _
        "prod_Dependiente = '" & Left(cmbProd(7).Text, 1) & "', PROD_INVENTARIO =  '" & Left(cmbProd(9).Text, 1) & "', " & _
        "PROD_CODIGO_PROV = '" & txtProd(10).Text & "', prod_PrecioDesc = '" & txtProd(12).Text & "', prod_AplicaDesc = '" & Left(cmbProd(10).Text, 1) & "' " & _
        "WHERE PROD_CODIGO = '" & prodId & "'"
        con.Execute (sql1)
                
        sql1 = "DELETE FROM PRODUCTO_DEPENDIENTE WHERE PROD_ID = '" & Id & "'"
        con.Execute (sql1)
        
        With listDependiente
            For b1 = 1 To .Rows - 1
                sql1 = "INSERT INTO PRODUCTO_DEPENDIENTE (PROD_ID, PROD_SERV, PROD_CANT_EQUI, PROD_DEPEN_ID, PROD_dEPEN_SERV, PROD_DEPEN_TIPO) VALUES " & _
                " ('" & Id & "', 'P', '" & .TextMatrix(b1, 2) & "', '" & .TextMatrix(b1, 5) & "', 'P', 'P'   )"
                'MsgBox sql1
                con.Execute (sql1)
            Next b1
        End With
                        
        
                    
    End If

        If Err.Number = -2147217900 Then
            MsgBox "El código que quiere registrar ya existe para un producto y no puede duplicarse. " & vbCrLf & vbCrLf & "Por favor verifique.", vbCritical
            Exit Sub
        End If
    
    'Para la fotoi
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempProd.dat")
        'If Not RES1.EOF Then
            res.Open "SELECT * FROM Productos WHERE prod_codigo = '" & prodId & "'", con, adOpenStatic, adLockOptimistic
            'MsgBox "-" & prodId & "-"
            If Not res.EOF Then
                '''NO DEBE
            'Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempProd.dat")
                res.Fields("prod_Foto") = Imagen1.Read
                res.Update
            Else
                ''''
            End If
        'End If
    End If
    
       
    
    MsgBox "Información guardada.", vbInformation
    save = True
    cancelar


End Sub
Private Sub cargaProductos()

    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO ORDER BY NOMBRE"
    Set RES1 = con.Execute(sql1)

    cmbProd(8).Clear
    
    Do While Not RES1.EOF
        cmbProd(8).AddItem RES1.Fields("nombre")
        cmbProd(8).ItemData(cmbProd(8).ListCount - 1) = RES1.Fields("PROD_ID")
        RES1.MoveNext
    Loop
End Sub


Private Sub cargaLista()
    Dim texto1 As String
    Dim cantTotal As Double
    Dim cantAgotados As Double
    
    texto1 = ""
    If cmbProd(5).Text <> "TODOS" Then
        texto1 = texto1 & "AND upper(MARCA) LIKE upper('%" & cmbProd(5).Text & "%') "
    End If
    If cmbProd(6).Text <> "TODOS" Then
        texto1 = texto1 & "AND upper(TIPO) LIKE upper('%" & cmbProd(6).Text & "%') "
    End If
    
    If Check1.value = Checked Then
        texto1 = texto1 & " AND id_status = 'A' "
    End If
    
    texto1 = texto1 & " order by FECHA DESC "
    
    If Val(textBus(4).Text) > 0 Then
        texto1 = texto1 & "Limit 0, " & Val(textBus(4).Text) & ""
    End If
    
    
    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE  SUBTIPO = 'PRODUCTO' AND " & _
    "CODIGO LIKE '%" & textBus(0).Text & "%' " & _
    "AND upper(NOMBRE) LIKE upper('%" & textBus(1).Text & "%') " & texto1
    Set RES1 = con.Execute(sql1)
    
    ListaUsers.Rows = 1
    ListaUsers.Redraw = False
    cantTotal = 0
    cantAgotados = 0
    Do While Not RES1.EOF
        ListaUsers.AddItem ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 0) = RES1.Fields("CODIGO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 1) = RES1.Fields("NOMBRE")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 2) = RES1.Fields("TIPO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 3) = RES1.Fields("MARCA")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 4) = RES1.Fields("CANTIDAD")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 5) = FormatCurrency(RES1.Fields("PRECIO_VENTA"))
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 6) = RES1.Fields("STATUS")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = RES1.Fields("STOCK_MIN")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 8) = RES1.Fields("STOCK_MAX")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 9) = RES1.Fields("PRESENTACION") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 10) = RES1.Fields("UNIDAD_mEDIDA")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 11) = RES1.Fields("PROVEEDOR") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 12) = RES1.Fields("FECHA")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 13) = RES1.Fields("USUARIO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 15) = RES1.Fields("CODIGO_PROV") & ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 16) = FormatCurrency(RES1.Fields("PRECIO_COSTO"))
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 17) = FormatCurrency(RES1.Fields("PRECIO_DESC"))
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 18) = RES1.Fields("TIPO_DEPEN")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 19) = RES1.Fields("PROD_ID")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 20) = RES1.Fields("DESCRIPCION")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 21) = RES1.Fields("INVENTARIO")
        
        Dim encontrado As Boolean
        encontrado = False
        ListaUsers.Row = ListaUsers.Rows - 1
        ListaUsers.Col = 14
        ListaUsers.CellFontName = "Wingdings"
        ListaUsers.CellFontBold = True
        ListaUsers.CellFontSize = 16
        If mn_Seleccion.Checked = True Then
            ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = Chr(254)
        Else
            If ListaSel.Rows > 1 Then
                For b1 = 1 To ListaSel.Rows - 1
                    If ListaSel.TextMatrix(b1, 0) = ListaUsers.TextMatrix(ListaUsers.Rows - 1, 0) Then
                        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = Chr(254)
                        encontrado = True
                        Exit For
                    End If
                Next b1
                If encontrado = False Then
                    ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = Chr(168)
                Else
                    encontrado = False
                End If
            Else
                ListaUsers.TextMatrix(ListaUsers.Rows - 1, 14) = Chr(168)
            End If
        End If
        If RES1.Fields("ID_STATUS") = "I" Or RES1.Fields("CANTIDAD") <= 0 Then
            If RES1.Fields("inventario") = "SI" Then
                cantAgotados = cantAgotados + 1
                ListaUsers.Row = ListaUsers.Rows - 1
                For b1 = 0 To ListaUsers.Cols - 1
                    ListaUsers.Col = b1
                    ListaUsers.CellForeColor = vbRed
                Next b1
            End If
        Else
        cantTotal = cantTotal + RES1.Fields("CANTIDAD")
            If RES1.Fields("CANTIDAD") <= RES1.Fields("STOCK_MIN") Then
                ListaUsers.Row = ListaUsers.Rows - 1
                For b1 = 0 To ListaUsers.Cols - 1
                    ListaUsers.Col = b1
                    ListaUsers.CellForeColor = &H80FF&
                Next b1
            End If
        End If
        
    
        
        RES1.MoveNext
    Loop
    lInfo(10).Caption = "Productos en lista: " & ListaUsers.Rows - 1 & "  Total productos existentes:  " & cantTotal & "  Productos agotados: " & cantAgotados
    ListaUsers.Redraw = True

End Sub


Private Sub checarCampos()
    checkError = False
    
    For b1 = 0 To 5
        If txtProd(b1).Text = "" Then
            checkError = True
            lProd(b1).ForeColor = vbRed
            Exit For
        End If
    Next b1
    
        If txtProd(8).Text = "" Then
            checkError = True
            lProd(9).ForeColor = vbRed
        End If
    
    If txtProd(9).Text = "" Then
        txtProd(9).Text = "0"
    End If
    If txtProd(12).Text = "" Then
        txtProd(12).Text = "0"
    End If
    
    If checkError = False Then
        If cmbProd(0).Text = "" Then
            checkError = True
            lProd(51).ForeColor = vbRed
        Else
            If cmbProd(1).Text = "" Then
                checkError = True
                lProd(61).ForeColor = vbRed
            Else
                If cmbProd(4).Text = "" Then
                    checkError = True
                    lProd(6).ForeColor = vbRed
                Else
                    If cmbProd(3).Text = "" Then
                        checkError = True
                        lProd(11).ForeColor = vbRed
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub cmbProd_Click(Index As Integer)
    Select Case Index
        Case 4:
            If lProd(6).ForeColor = vbRed Then
                lProd(6).ForeColor = vbBlack
            End If
        Case 0:
            If lProd(51).ForeColor = vbRed Then
                lProd(51).ForeColor = vbBlack
            End If
        Case 1:
            If lProd(61).ForeColor = vbRed Then
                lProd(61).ForeColor = vbBlack
            End If
        Case 3:
            If lProd(11).ForeColor = vbRed Then
                lProd(11).ForeColor = vbBlack
            End If
        Case 5:
            cargaLista
        Case 6:
            cargaLista
        Case 7:
            If cmbProd(7).ItemData(cmbProd(7).ListIndex) = 1 Then
                txtProd(2).Enabled = True
                cmBoton(8).Enabled = False
            Else
                If cmbProd(7).ItemData(cmbProd(7).ListIndex) = 2 Then
                    txtProd(2).Enabled = False
                    checkCantDependiente
                    cmBoton(8).Enabled = True
                End If
            End If
        Case 9:
            If cmbProd(9).Text = "NO" Then
                txtProd(2).Text = "0"
                txtProd(4).Text = "0"
                txtProd(5).Text = "0"
                txtProd(2).Enabled = False
                txtProd(4).Enabled = False
                txtProd(5).Enabled = False
            Else
                txtProd(2).Enabled = True
                txtProd(4).Enabled = True
                txtProd(5).Enabled = True
            End If
'        Case 8:
'            enviar_Productos
    
    
    End Select
End Sub
Private Sub enviar_Productos()
    
    'MsgBox cmbProd(8).ItemData(cmbProd(8).ListIndex)
    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE PROD_ID = '" & cmbProd(8).ItemData(cmbProd(8).ListIndex) & "'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        listDependiente.TextMatrix(listDependiente.Row, 0) = RES1.Fields("CODIGO")
        listDependiente.TextMatrix(listDependiente.Row, 1) = RES1.Fields("NOMBRE")
        listDependiente.TextMatrix(listDependiente.Row, 2) = ""
        listDependiente.TextMatrix(listDependiente.Row, 3) = RES1.Fields("UNIDAD_MEDIDA")
        listDependiente.TextMatrix(listDependiente.Row, 4) = RES1.Fields("TIPO")
        listDependiente.TextMatrix(listDependiente.Row, 5) = RES1.Fields("PROD_ID")
    End If

    cmbProd(8).Visible = False
    

End Sub
Private Sub checkCantDependiente()
    ''''''''Para mostra la cantidad proporcional al producto que dependa
End Sub
Private Sub cmbProd_GotFocus(Index As Integer)
    If Index = 0 Then
        cmBoton(4).Visible = True
        cmBoton(5).Visible = False
        cmBoton(6).Visible = False
        cmBoton(7).Visible = False
    Else
        If Index = 1 Then
            cmBoton(5).Visible = True
            cmBoton(4).Visible = False
            cmBoton(6).Visible = False
            cmBoton(7).Visible = False
        Else
            If Index = 3 Then
                cmBoton(6).Visible = True
                cmBoton(5).Visible = False
                cmBoton(4).Visible = False
                cmBoton(7).Visible = False
            Else
                If Index = 2 Then
                    cmBoton(7).Visible = True
                    cmBoton(6).Visible = False
                    cmBoton(5).Visible = False
                    cmBoton(4).Visible = False
                Else
                    cmBoton(7).Visible = False
                    cmBoton(6).Visible = False
                    cmBoton(5).Visible = False
                    cmBoton(4).Visible = False
                End If
            End If
        End If
    End If
        
End Sub

Private Sub cmbProd_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 8 Then
        If KeyAscii = 13 Then
        enviar_Productos
        End If
    End If
End Sub

Private Sub cmbProd_LostFocus(Index As Integer)
    If Index = 8 Then
        cmbProd(8).Visible = False
    End If
End Sub

Public Sub cmd_Marca_Click()
    cargaMarca
End Sub

Private Sub cmdAll_Click()
Dim ques As String
ques = MsgBox("¿Ver todos? Esto puede tardar dependiendo la cantidad de información. " & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        sql1 = "SELECT COUNT(*) NUM FROM PRODUCTOS"
        Set RES1 = con.Execute(sql1)
        textBus(4).Text = Val(RES1.Fields("Num")) + 1
        cargaLista
        textBus(4).Text = "50"
    End If
    
End Sub

Public Sub cmdPresentacion_Click()
    Dim tipo As String
'    Select Case Option1(Index)
'        Case 0: Tipo = "U"
'        Case 1: Tipo = "M"
'        Case 2: Tipo = "T"
'
'    End Select
    For b1 = 0 To 2
        If Option1(b1).value = True Then
            Select Case b1
                Case 0: tipo = "U"
                Case 1: tipo = "T"
                Case 2: tipo = "M"
            End Select
            cargaPresentacion (tipo)
            Exit For
        End If
    Next b1
End Sub

Public Sub cmdProveed_Click()
    cargaProveedor
End Sub

Public Sub cmdTipo_Click()
    cargaTipoProd
End Sub

Private Sub Form_Load()
'    bordesProductos
'    ListaUsers.Rows = 1
    
    mn_Seleccion.Checked = False
    activaSeleccion = False
    ListaSel.Rows = 1
    ListaSel.Cols = ListaUsers.Cols
    'listDependiente.ColWidth(5) = 0
    'listDependiente.ColWidth(4) = 0
    CargaGeneral
    cargaLista
    cargaReceta
    checkMayus
    cargaToolTips
    
    'cargaLista_TipoImagen
End Sub


Private Sub checkMayus()
    sql1 = "SELECT SUC_MAYUSCULAS FROM SUCURSAL"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        If RES1.Fields("SUC_MAYUSCULAS") = "1" Then
            mayus = True
        Else
            mayus = False
        End If
    End If
    
End Sub
Private Sub cargaReceta()
    Lista_Receta.Rows = 1
    
    sql1 = "SELECT * FROM VIEW_PRODUCTO_DEPENDIENTES ORDER BY PRODUCTO, PRODUCTO_DEPEN ASC"
    Set RES1 = con.Execute(sql1)
    
    
    Lista_Receta.MergeCells = flexMergeRestrictColumns
    
    Do While Not RES1.EOF
        Lista_Receta.AddItem ""
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 0) = RES1.Fields("CODIGO")
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 1) = RES1.Fields("PRODUCTO")
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 2) = FormatCurrency(RES1.Fields("COSTO_RECETA"))
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 3) = RES1.Fields("PRODUCTO_DEPEN")
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 4) = RES1.Fields("CANTIDAD_EQUI")
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 5) = RES1.Fields("PRESENTACION")
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 6) = FormatCurrency(RES1.Fields("COSTO_INDIVIDUAL"))
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 7) = FormatCurrency(RES1.Fields("COSTO_EQUIVALENTE"))
        Lista_Receta.TextMatrix(Lista_Receta.Rows - 1, 8) = RES1.Fields("CODIGO_DEPEN")
        
        RES1.MoveNext
    Loop
    
    
    Lista_Receta.MergeCol(0) = True
    Lista_Receta.MergeCol(1) = True
    Lista_Receta.MergeCol(2) = True
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If SSTab1.Tab = 1 And save = False Then
        a = MsgBox("Perderá la información. ¿Salir?", vbQuestion + vbYesNo)
        If a = vbYes Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
    
TT1.Destroy
TT2.Destroy
TT3.Destroy
TT4.Destroy
TT5.Destroy
TT6.Destroy
TT7.Destroy
TT8.Destroy
TT9.Destroy
TT10.Destroy
TT11.Destroy
TT12.Destroy
TT13.Destroy
TT14.Destroy
TT15.Destroy
TT16.Destroy
End Sub

Private Sub ListaUsers_Click()
'    muestraInfo (ListaUsers.TextMatrix(ListaUsers.Row, 0))
End Sub
Private Sub muestraInfo(prodCodigo As String)

    fotoProd.Picture = LoadPicture("")
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, CTMR_MARCA, if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, " & _
    "PROD_CANT, CTPT_TIPO, prod_Stock_Min, PROD_STOCK_MAX, PROD_PRESENTACION, PROD_FOTO, CTPS_NOMBRE " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL)   AND T1.PROD_CODIGO  = '" & prodCodigo & "' "
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        If IsNull(RES1.Fields("PROD_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            fotoProd.Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            fotoProd.Picture = LoadPicture("")
        End If
        lInfo(0).Caption = "Producto: " & RES1.Fields("PROD_NOMBRE")
        lInfo(1).Caption = "Código: " & RES1.Fields("PROD_CODIGO")
        lInfo(2).Caption = "Marca: " & RES1.Fields("CTMR_MARCA")
        lInfo(3).Caption = "Tipo: " & RES1.Fields("CTPT_TIPO")
        lInfo(4).Caption = "Cantidad: " & RES1.Fields("PROD_CANT")
        lInfo(5).Caption = "Precio: " & FormatCurrency(RES1.Fields("PROD_PRECIO"))
        lInfo(6).Caption = "Stock min: " & RES1.Fields("PROD_STOCK_MIN")
        lInfo(7).Caption = "Stock max: " & RES1.Fields("PROD_STOCK_MAX")
        lInfo(8).Caption = "Presentación: " & RES1.Fields("PROD_PRESENTACION") & " " & RES1.Fields("CTPS_NOMBRE")
        lInfo(9).Caption = "Proveedor: "
        
    Else
        fotoProd.Picture = LoadPicture("")
        lInfo(0).Caption = "Producto: "
        lInfo(1).Caption = "Código: "
        lInfo(2).Caption = "Marca: "
        lInfo(3).Caption = "Tipo: "
        lInfo(4).Caption = "Cantidad: "
        lInfo(5).Caption = "Precio: "
        lInfo(6).Caption = "Stock min: "
        lInfo(7).Caption = "Stock max: "
        lInfo(8).Caption = "Presentación: "
        lInfo(9).Caption = "Proveedor: "
    End If
End Sub

Private Sub ListaUsers_DblClick()
    Dim ques As String
    
    If ListaUsers.MouseRow = 0 Then
        Call ordenarLista(ListaUsers)
    Else
        If ListaUsers.Col = 5 Then
            
            
    
            Call checarPermisos("FRM_PRODUCTOS", FRM_Menu.menuBarra2.Panels(8).Text)
            
            If permEdit = "SI" Then
                txtPrecio.Top = ListaUsers.CellTop + ListaUsers.Top
                txtPrecio.Left = ListaUsers.CellLeft + ListaUsers.Left
                txtPrecio.height = ListaUsers.CellHeight
                txtPrecio.width = ListaUsers.CellWidth
                txtPrecio.Text = Format(ListaUsers.TextMatrix(ListaUsers.Row, ListaUsers.Col), "General Number")
                txtPrecio.Visible = True
                txtPrecio.SelStart = 0
                txtPrecio.SelLength = Len(txtPrecio.Text)
                txtPrecio.SetFocus
            
            Else
                MsgBox "Opción no disponible. Verifique", vbInformation
            End If
        Else
            If ListaUsers.Col = 4 Then
                Call checarPermisos("FRM_PRODUCTOS", FRM_Menu.menuBarra2.Panels(8).Text)
                
                If permEdit = "SI" Then
                    txtPrecio.Top = ListaUsers.CellTop + ListaUsers.Top
                    txtPrecio.Left = ListaUsers.CellLeft + ListaUsers.Left
                    txtPrecio.height = ListaUsers.CellHeight
                    txtPrecio.width = ListaUsers.CellWidth
                    txtPrecio.Text = Format(ListaUsers.TextMatrix(ListaUsers.Row, ListaUsers.Col), "General Number")
                    txtPrecio.Visible = True
                    txtPrecio.SelStart = 0
                    txtPrecio.SelLength = Len(txtPrecio.Text)
                    txtPrecio.SetFocus
                
                Else
                    MsgBox "Opción no disponible. Verifique", vbInformation
                End If
                
        
            Else
            
                Dim b1 As Long
                b1 = ListaUsers.Row
                
                ListaUsers.Row = b1
                ListaUsers.Col = 14
                If ListaUsers.TextMatrix(b1, 14) = Chr(168) Then
                    ListaUsers.TextMatrix(b1, 14) = Chr(254)
                    enviaProductoSel (ListaUsers.Row)
                Else
                    ListaUsers.TextMatrix(b1, 14) = Chr(168)
                End If
            End If
        End If
    End If
End Sub
Private Sub enviaProductoSel(fila As Integer)
    ListaSel.AddItem ""
    'MsgBox ListaSel.Cols & "  " & ListaUsers.Cols
    ListaSel.Redraw = False
    For b1 = 0 To ListaUsers.Cols - 1
        ListaSel.TextMatrix(ListaSel.Rows - 1, b1) = ListaUsers.TextMatrix(fila, b1)
    Next b1
    ListaSel.Redraw = True

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
            mn_Eliminar.Enabled = True
            PopupMenu mn_Prod, vbPopupMenuLeftAlign
        End If
    Else
            mn_Add.Enabled = True
            mn_Edit.Enabled = False
            mn_Eliminar.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_Prod, vbPopupMenuLeftAlign
        End If
    End If

End Sub
Private Sub agregarNuevo()
                        
            Option1_Click (0)
            lbStatus.Caption = "Estatus: Agregando producto"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            txtProd(0).SetFocus
            txtProd(6).Text = "0"
            cmbProd(2).Text = "OTRO"
            txtProd(4).Text = "0"
            txtProd(5).Text = "0"
            cmbProd(4).Text = "ACTIVO"
            
            sql1 = "SELECT MAX(PROD_ID) + 1 CLAVE FROM PRODUCTOS"
            Set RES1 = con.Execute(sql1)
            
            If Not RES1.EOF Then
                If IsNull(RES1.Fields("CLAVE")) = True Then
                    txtProd(1).Text = "P000000001"
                Else
                    txtProd(1).Text = "P" & Format(RES1.Fields("CLAVE"), "000000000")
                End If
            End If
            save = False


End Sub

Private Sub listDependiente_DblClick()
        
    If listDependiente.Col = 6 Then
        If listDependiente.TextMatrix(listDependiente.Row, 6) = Chr(168) Then
            listDependiente.TextMatrix(listDependiente.Row, 6) = Chr(254)
        Else
            listDependiente.TextMatrix(listDependiente.Row, 6) = Chr(168)
        End If
        
    Else
            
        If listDependiente.TextMatrix(listDependiente.Row, 2) = "" Then
            
            Dim b1 As Long
            b1 = listDependiente.Row
            
            If listDependiente.Col = 1 Then
                cargaProductos
                cmbProd(8).Top = listDependiente.CellTop + listDependiente.Top
                cmbProd(8).Left = listDependiente.CellLeft + listDependiente.Left
                cmbProd(8).width = listDependiente.CellWidth
                'cmbProd(8).Text = listDependiente.TextMatrix(listDependiente.Row, listDependiente.Col)
                cmbProd(8).Visible = True
                cmbProd(8).SetFocus
            Else
                If listDependiente.Col = 2 Then
                    txtProd(11).Top = listDependiente.CellTop + listDependiente.Top
                    txtProd(11).Left = listDependiente.CellLeft + listDependiente.Left
                    txtProd(11).height = listDependiente.CellHeight
                    txtProd(11).width = listDependiente.CellWidth
                    txtProd(11).Text = listDependiente.TextMatrix(listDependiente.Row, listDependiente.Col)
                    txtProd(11).Visible = True
                    txtProd(11).SelStart = 0
                    txtProd(11).SelLength = Len(txtProd(11).Text)
                    txtProd(11).SetFocus
                End If
            End If
        End If
    End If
    

End Sub

Private Sub listDependiente_KeyPress(KeyAscii As Integer)
'    Dim valor As Long
'    Dim fila As Long
'    Dim codigo As String
'        listDependiente.WordWrap = False
'        If listDependiente.Col = 0 Then
'            If KeyAscii = 13 Then
'
'                codigo = listDependiente.TextMatrix(listDependiente.Row, 0)
'                fila = listDependiente.Row
'                For b1 = 1 To listDependiente.Rows - 1
'                    If UCase(listDependiente.TextMatrix(b1, 0)) = UCase(codigo) And fila <> b1 Then
'                        MsgBox "Este código ya se encuentra en la lista. Verifique.", vbInformation
'                        Exit Sub
'                    End If
'                Next b1
'                checkProd_Depen (codigo)
'            Else
'                listDependiente.Text = listDependiente.Text & Chr(KeyAscii)
'            End If
'        Else
'            If listDependiente.Col = 2 Then
'                If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
'                    listDependiente.Text = listDependiente.Text & Chr(KeyAscii)
'                    If Val(listDependiente.TextMatrix(listDependiente.Row, 2)) > 0 Then
'                        txtProd(2).Text = Round(Val(listDependiente.TextMatrix(listDependiente.Row, 5)) / Val(listDependiente.TextMatrix(listDependiente.Row, 2)), 2)
'                    Else
'                        txtProd(2).Text = "0"
'                    End If
'                End If
'            End If
'        End If

End Sub

Private Sub checkProd_Depen(codigo As String)

    sql1 = "SELECT PROD_ID, PROD_CODIGO, PROD_NOMBRE, PROD_CANT FROM PRODUCTOS WHERE PROD_CODIGO = '" & codigo & "'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        'listDependiente.AddItem ""
        listDependiente.TextMatrix(listDependiente.Row, 0) = RES1.Fields("prod_codigo")
        listDependiente.TextMatrix(listDependiente.Row, 1) = RES1.Fields("prod_nombre")
        listDependiente.TextMatrix(listDependiente.Row, 2) = "1"
        listDependiente.TextMatrix(listDependiente.Row, 4) = RES1.Fields("PROD_ID")
        listDependiente.TextMatrix(listDependiente.Row, 5) = RES1.Fields("PROD_CANT")
        If listDependiente.Rows = 2 Then
            listDependiente.TextMatrix(listDependiente.Row, 3) = "PRINCIPAL"
        Else
            If listDependiente.Rows > 2 Then
                listDependiente.TextMatrix(listDependiente.Row, 3) = "SECUNDARIO"
            End If
        End If
        
        listDependiente.Row = listDependiente.Rows - 1
        listDependiente.Col = 6
        listDependiente.CellFontName = "Wingdings"
        listDependiente.CellFontBold = True
        listDependiente.CellFontSize = 16
        listDependiente.TextMatrix(listDependiente.Rows - 1, 6) = Chr(254)
                
    Else
        MsgBox "Codigo de producto no encontrado. Verifique", vbInformation
    End If
End Sub

Private Sub listDependiente_KeyUp(KeyCode As Integer, Shift As Integer)
        If listDependiente.Col = 0 Or listDependiente.Col = 2 Then
            Select Case KeyCode
                Case vbKeyDelete
                    listDependiente.Text = ""
                    If listDependiente.Col = 0 Then
                        listDependiente.TextMatrix(listDependiente.Row, 0) = ""
                        listDependiente.TextMatrix(listDependiente.Row, 1) = ""
                        listDependiente.TextMatrix(listDependiente.Row, 2) = ""
                        listDependiente.TextMatrix(listDependiente.Row, 3) = ""
                        listDependiente.TextMatrix(listDependiente.Row, 4) = ""
                        listDependiente.TextMatrix(listDependiente.Row, 5) = ""
                    End If
                Case vbKeyBack
                    If Len(listDependiente.Text) > 0 Then
                        listDependiente.Text = Val(Left(listDependiente.Text, Len(listDependiente.Text) - 1))
                    Else
                        listDependiente.Text = ""
                        If listDependiente.Col = 0 Then
                            listDependiente.TextMatrix(listDependiente.Row, 0) = ""
                            listDependiente.TextMatrix(listDependiente.Row, 1) = ""
                            listDependiente.TextMatrix(listDependiente.Row, 2) = ""
                            listDependiente.TextMatrix(listDependiente.Row, 3) = ""
                            listDependiente.TextMatrix(listDependiente.Row, 4) = ""
                            listDependiente.TextMatrix(listDependiente.Row, 5) = ""
                        
                        End If
                    End If
            End Select
        End If

End Sub


Private Sub ListProd1_Click()
    
    'MsgBox tipoId(ListProd1.Row, ListProd1.Col) & "   " & tipoValor(ListProd1.Row, ListProd1.Col)
    cargaLista_ProdImagen
End Sub

Private Sub ListProd1_GotFocus()
    ConScroll ListProd1
End Sub

Private Sub ListProd1_LostFocus()
    SinScroll ListProd1
End Sub

Private Sub listprod2_GotFocus()
    ConScroll listprod2
    
End Sub

Private Sub listprod2_LostFocus()
    SinScroll listprod2
End Sub

Private Sub mn_Add_Click()
    Dim ques As String
    
    Call checarPermisos("FRM_PRODUCTOS", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAdd = "SI" Then
        ques = MsgBox("¿Desea agregar un producto?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                agregarNuevo
            End If
    Else
        MsgBox "Opción no disponible. Verifique", vbInformation
    End If
End Sub

Private Sub mn_AddSame_Click()
    Dim ques As String
    
    ques = MsgBox("Desea agrega un producto similar al: " & ListaUsers.TextMatrix(ListaUsers.Row, 0) & vbCrLf & _
    ListaUsers.TextMatrix(ListaUsers.Row, 1) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 2), vbYesNo + vbQuestion)
        If ques = vbYes Then
            prodId = ListaUsers.TextMatrix(ListaUsers.Row, 0)
            lbStatus.Caption = "Estatus: Agregando producto"
            cargaEdit
            

            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            sql1 = "SELECT MAX(PROD_ID) + 1 CLAVE FROM PRODUCTOS"
            Set RES1 = con.Execute(sql1)
            
            If Not RES1.EOF Then
                If IsNull(RES1.Fields("CLAVE")) = True Then
                    txtProd(1).Text = "P000000001"
                Else
                    txtProd(1).Text = "P" & Format(RES1.Fields("CLAVE"), "000000000")
                End If
            End If
            save = False
        End If
    
End Sub

Private Sub mn_CatTipoPresen_Click()
    CAT_TipoPresetacion.Show vbModal
End Sub

Private Sub mn_Edit_Click()
    Dim ques As String
    
    Call checarPermisos("FRM_PRODUCTOS", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permEdit = "SI" Then
        ques = MsgBox("Desea editar el producto: " & ListaUsers.TextMatrix(ListaUsers.Row, 0) & vbCrLf & _
            ListaUsers.TextMatrix(ListaUsers.Row, 1) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 2), vbYesNo + vbQuestion)
        If ques = vbYes Then
            prodId = ListaUsers.TextMatrix(ListaUsers.Row, 0)
            Id = ListaUsers.TextMatrix(ListaUsers.Row, 19)
            lbStatus.Caption = "Estatus: Editando producto"
            cargaEdit
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            save = False
        End If
    Else
        MsgBox "Opción no disponible. Verifique", vbInformation
    End If
    
End Sub
Private Sub cargaEdit()
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
'    pFoto.Visible = False
    iFoto.Visible = True
    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE CODIGO = '" & prodId & "' "
    Set RES2 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES2.EOF Then
        Select Case RES2.Fields("TIPO_PRESEN")
            Case "U":
                Option1(0).value = True
                Option1_Click (0)
            Case "T":
                Option1(1).value = True
                Option1_Click (1)
            Case "M":
                Option1(2).value = True
                Option1_Click (2)
        End Select
        
        txtProd(0).Text = RES2.Fields("NOMBRE")
        txtProd(1).Text = RES2.Fields("CODIGO")
        txtProd(2).Text = RES2.Fields("CANTIDAD")
        txtProd(3).Text = RES2.Fields("PRECIO_VENTA")
        
        cmbProd(7).Text = RES2.Fields("TIPO_DEPEN")
        'cmbProd(2).Text = "" & RES2.Fields("PRESENTACION")
        
        txtProd(6).Text = "" & RES2.Fields("PRESENTACION")
        txtProd(4).Text = RES2.Fields("STOCK_MIN")
        txtProd(5).Text = RES2.Fields("STOCK_MAX")
        txtProd(7).Text = RES2.Fields("DESCRIPCION")
        txtProd(8).Text = RES2.Fields("precio_costo") & ""
        txtProd(9).Text = RES2.Fields("precio_may") & ""
        txtProd(10).Text = RES2.Fields("codigo_prov") & ""
        txtProd(12).Text = RES2.Fields("PRECIO_DESC") & ""
        
        
        If IsNull(RES2.Fields("MARCA")) Then
        Else
            cmbProd(0).Text = RES2.Fields("MARCA")
        End If
        If IsNull(RES2.Fields("TIPO")) Then
        Else
            cmbProd(1).Text = RES2.Fields("TIPO")
        End If
        'MsgBox RES2.Fields("PROVEEDOR")
        If IsNull(RES2.Fields("PROVEEDOR")) Then
        Else
            cmbProd(3).Text = RES2.Fields("PROVEEDOR")
        End If
        
        If IsNull(RES2.Fields("STATUS")) Then
        Else
            cmbProd(4).Text = RES2.Fields("STATUS")
        End If
        
        If IsNull(RES2.Fields("UNIDAD_mEDIDA")) = True Then
        Else
            cmbProd(2).Text = RES2.Fields("UNIDAD_MEDIDA")
        End If
        If IsNull(RES2.Fields("INVENTARIO")) Then
        Else
            cmbProd(9).Text = RES2.Fields("INVENTARIO")
        End If
        If IsNull(RES2.Fields("APLICA_DESC")) Then
        Else
            cmbProd(10).Text = RES2.Fields("APLICA_dESC")
        End If
        
        
        If IsNull(RES2.Fields("fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES2.Fields("FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            iFoto.Picture = LoadPicture("")
        End If
        
        ''' PARA CARGAR EL DETALLE DE LOS DEPENDIENTES (FALTA LA VISTA
        sql1 = "SELECT * FROM VIEW_PRODUCTO_dEPENDIENTES WHERE CODIGO = '" & prodId & "'"
        Set RES1 = con.Execute(sql1)
        'MsgBox SQL1
        
        listDependiente.Rows = 1
        Do While Not RES1.EOF
            listDependiente.AddItem ""
            listDependiente.TextMatrix(listDependiente.Rows - 1, 0) = RES1.Fields("CODIGO_DEPEN")
            listDependiente.TextMatrix(listDependiente.Rows - 1, 1) = RES1.Fields("PRODUCTO_DEPEN")
            listDependiente.TextMatrix(listDependiente.Rows - 1, 2) = RES1.Fields("CANTIDAD_EQUI")
            listDependiente.TextMatrix(listDependiente.Rows - 1, 3) = RES1.Fields("PRESENTACION")
            'listDependiente.TextMatrix(listDependiente.Rows - 1, 4) = res1.Fields("CANTIDAD_DEPEN")
            listDependiente.TextMatrix(listDependiente.Rows - 1, 5) = RES1.Fields("ID_DEPEN")
            
            listDependiente.Row = listDependiente.Rows - 1
            listDependiente.Col = 6
            listDependiente.CellFontName = "Wingdings"
            listDependiente.CellFontBold = True
            listDependiente.CellFontSize = 12
'            listDependiente.TextMatrix(listDependiente.Rows - 1, 6) = Chr(254)
            listDependiente.TextMatrix(listDependiente.Rows - 1, 6) = Chr(168)
            
            RES1.MoveNext
        Loop
        
    End If
    
End Sub

Private Sub mn_Etiquetas_Click()
    CAT_Etiquetas.Show vbModal
End Sub

Private Sub mn_Marca_Click()
    Call checarPermisos("CAT_MARCA", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        CAT_Marca.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
End Sub

Private Sub mn_PrintAll_Click()
    Dim ques As String
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel2(ListaUsers)
    End If

End Sub

Private Sub mn_PrintCodigos_Click()
'    imprimirCodigos
    Dim b1 As Long
'    For b1 = 1 To ListaSel.Rows - 1
'        If ListaUsers.TextMatrix(b1, 14) = Chr(254) Then
'            num1 = num1 + (1 * ListaUsers.TextMatrix(b1, 4))
'        End If
'    Next b1
    num1 = ListaSel.Rows - 1
    
    If num1 > 0 Then
        PRINT_Etiquetas.txtInfo(1).Text = num1
        PRINT_Etiquetas.Show vbModal
    Else
        MsgBox "No hay selección para imprimir. Verifique.", vbInformation
    End If
End Sub
Private Sub imprimirCodigos()
Dim ques As String
Dim valorx As Long
Dim valory As Long
Dim num As Long
Dim num1 As Long


        For b1 = 1 To ListaUsers.Rows - 1
            If ListaUsers.TextMatrix(b1, 14) = Chr(254) Then
                num1 = num1 + 1
            End If
        Next b1
        
        ques = MsgBox("Va a imprimir " & num1 & " codigos. " & vbCrLf & vbCrLf & _
        "¿Continuar?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            valorx = 400
            valory = 250
            num = 0
            num1 = 0
            Printer.KillDoc
            For b1 = 1 To ListaUsers.Rows - 1
                If ListaUsers.TextMatrix(b1, 14) = Chr(254) Then
                    num = num + 1
                    num1 = num1 + 1
                    Call DrawBarcode(ListaUsers.TextMatrix(b1, 0), Picture1)
                    Printer.Font = "Sans Serif"
                    Printer.FontBold = True
                    Printer.FontSize = 14
    '                Printer.PaintPicture Imagen2, 10, 25, 10250, 15550
        
                    Alto = Picture1.height
                    Ancho = Picture1.width
                    Picture1.Picture = Picture1.Image
                    Picture1.height = Alto
                    Picture1.width = Ancho
                        
                    Picture1.Picture = Picture1.Image
                    Printer.Font = "Courier New"
                    Printer.FontSize = 8
                    Printer.FontBold = False
                    Printer.CurrentX = valorx + 150
                    Printer.CurrentY = valory
                    If Len(ListaUsers.TextMatrix(b1, 1)) > 25 Then
                        Printer.Print Left(ListaUsers.TextMatrix(b1, 1), 25)
                    Else
                        Printer.Print ListaUsers.TextMatrix(b1, 1)
                    End If
                    Printer.PaintPicture Picture1, 150 + valorx, 100 + valory + 120, 2800, 800
                    'Printer.PaintPicture Picture1, 7250, 6000, 2500, 800
                    If num1 = 56 Then
                        Printer.NewPage
                        num = 0
                        valorx = 0
                        valory = 0
                        num1 = 0
                    Else
                        If num = 3 Then
                            num = 0
                            valorx = 400
                            valory = valory + 2850
                        Else
                            valorx = valorx + 4000
                        End If
                    End If
                End If
            Next b1
        Printer.EndDoc
        End If

End Sub


Private Sub mn_PrintGroup_Click()
    Dim ques As String
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(ListaUsers)
    End If
    

End Sub

Private Sub mn_Proveedor_Click()
    tipoPersona = "PROVEEDOR"
    ADD_Cliente.Show vbModal
End Sub

Private Sub mn_Seleccion_Click()
    ListaUsers.Redraw = False
    If mn_Seleccion.Checked = True Then
        For b1 = 1 To ListaUsers.Rows - 1
            ListaUsers.Col = 5
            ListaUsers.Row = b1
            ListaUsers.TextMatrix(b1, 14) = Chr(168)
        Next b1
        mn_Seleccion.Checked = False
    Else
        For b1 = 1 To ListaUsers.Rows - 1
            ListaUsers.Col = 5
            ListaUsers.Row = b1
            ListaUsers.TextMatrix(b1, 14) = Chr(254)
        Next b1
        
        mn_Seleccion.Checked = True
        activaSeleccion = False
    End If
    ListaUsers.Redraw = True
    
    
    If ListaSel.Rows > 1 Then
        ques = MsgBox("Actualmente tiene productos seleccionados." & vbCrLf & vbCrLf & "¿Desea mandtener la selección?", vbYesNo + vbQuestion)
        If ques = vbNo Then
            ListaSel.Rows = 1
            MsgBox "Selección desecha", vbInformation
        End If
    End If
    

End Sub

Private Sub mn_TipoProd_Click()
    
            Call checarPermisos("CAT_TIPO", FRM_Menu.menuBarra2.Panels(8).Text)
            
            If permAcceso = "SI" Then
                tipoCatTipo = "P"
                CAT_Tipo.Show vbModal
            Else
                MsgBox "Opción no disponible.", vbInformation
            End If
    
End Sub

Private Sub Option1_Click(Index As Integer)

    If Index = 0 Then
        If Option1(Index).value = True Then
            'txtProd(6).Text = "N/A"
            txtProd(6).Enabled = True
             cargaPresentacion ("U")
        Else
            txtProd(6).Text = "N/A"
            txtProd(6).Enabled = False
        End If
    Else
        If Index = 1 Then
            If Option1(Index).value = True Then
                cargaPresentacion ("T")
                txtProd(6).Text = "N/A"
                txtProd(6).Enabled = False
            End If
        Else
            If Index = 2 Then
                If Option1(Index).value = True Then
                    cargaPresentacion ("M")
                    txtProd(6).Text = "N/A"
                    txtProd(6).Enabled = False
                End If
            End If
        End If
    End If

End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
           textBus(0).Text = Replace(textBus(0).Text, "'", "-")
           If Left(textBus(0).Text, 1) = " " Then
                textBus(0).Text = Right(textBus(0).Text, (Len(textBus(0).Text) - 1))
           End If
        End If
        
        cargaLista
    End If

End Sub


Private Sub time1_Timer()
    time1.Enabled = False
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    Image2(0).width = Me.width
    Image2(0).height = Me.height
    Image2(1).width = Me.width
    Image2(1).height = Me.height
    Image2(2).height = Me.height
    Image2(2).width = Me.width
    ListaUsers.width = Me.width - 500
    listprod2.width = Me.width - 7600
    Lista_Receta.width = Me.width - 500
    'ListaUsers.height = Me.height - 3700

End Sub

Private Sub cargaLista_ProdImagen()
Dim Ancho As Long, Alto As Long
Dim contaFila As Long
Dim contaCasillas As Long
Dim contaTipos As Long
Dim columnas As Long
Dim Imagen1 As Stream
Set Imagen1 = New Stream

Ancho = 2175
Alto = 2415

'MsgBox "ANcho  " & listprod2.width & " Entran cols: " & (listprod2.width) / Ancho
columnas = Round(((listprod2.width) / (Ancho)), 0)


sql1 = "SELECT * fROM VIEW_PRODUCTOS_INVENTARIO WHERE TIPO_ID = '" & tipoId(ListProd1.Row, ListProd1.Col) & "' ORDER BY NOMBRE ASC"
Set RES_PROD = con.Execute(sql1)

listprod2.Rows = 0

If RES_PROD.RecordCount >= columnas Then
    listprod2.Cols = columnas
    For b1 = 1 To columnas
        listprod2.ColWidth(b1 - 1) = Ancho
    Next b1
Else
    listprod2.Cols = RES_PROD.RecordCount
    For b1 = 1 To RES_PROD.RecordCount
        listprod2.ColWidth(b1 - 1) = Ancho
    Next b1
    
End If

contaFila = 0
contaTipos = 0
contaCasillas = columnas

Do While Not RES_PROD.EOF
    If contaCasillas = columnas Then
        listprod2.AddItem ""
        listprod2.RowHeight(listprod2.Rows - 1) = 2415
        contaCasillas = 0
    End If

    If RES_PROD.Fields("FOTO_SN") = "SI" Then
        If IsNull(RES_PROD.Fields("FOTO")) = False Then
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES_PROD.Fields("FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\Prod" & contaTipos & ".jpg", adSaveCreateOverWrite
            Imagen1.Close

            listprod2.Row = listprod2.Rows - 1
            listprod2.Col = contaCasillas
            Set listprod2.CellPicture = LoadPicture(direccionSistema & "\Temp\Prod" & contaTipos & ".jpg")
            listprod2.CellAlignment = 2
            listprod2.TextMatrix(listprod2.Rows - 1, contaCasillas) = RES_PROD.Fields("NOMBRE")
        End If
    Else
        listprod2.Row = listprod2.Rows - 1
        listprod2.Col = contaCasillas
        listprod2.CellAlignment = 2
        listprod2.TextMatrix(listprod2.Rows - 1, contaCasillas) = RES_PROD.Fields("NOMBRE")
    End If
'    tipoId(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("CLAVE")
'    tipoValor(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
    contaCasillas = contaCasillas + 1
    contaTipos = contaTipos + 1

    RES_PROD.MoveNext
Loop
'
listprod2.WordWrap = True
End Sub


Private Sub cargaLista_TipoImagen()
Dim Ancho As Long, Alto As Long
Dim contaFila As Long
Dim contaCasillas As Long
Dim contaTipos As Long

Dim Imagen1 As Stream
Set Imagen1 = New Stream


Ancho = 2175
Alto = 2415

sql1 = "SELECT * fROM VIEW_TIPOPRODUCTOS ORDER BY TIPO ASC"
Set RESTIPO_PROD = con.Execute(sql1)

ListProd1.Rows = 0
ListProd1.ColWidth(0) = 2175
ListProd1.ColWidth(1) = 2175
ListProd1.ColWidth(2) = 2175
contaFila = 0
contaTipos = 0
contaCasillas = 3

Do While Not RESTIPO_PROD.EOF
    If contaCasillas = 3 Then
        ListProd1.AddItem ""
        ListProd1.RowHeight(ListProd1.Rows - 1) = 2415
        contaCasillas = 0
    End If
    
    If RESTIPO_PROD.Fields("FOTO_SN") = "SI" Then
        If IsNull(RESTIPO_PROD.Fields("fOTO")) = False Then
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RESTIPO_PROD.Fields("FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\" & contaTipos & ".jpg", adSaveCreateOverWrite
            Imagen1.Close

            ListProd1.Row = ListProd1.Rows - 1
            ListProd1.Col = contaCasillas
            Set ListProd1.CellPicture = LoadPicture(direccionSistema & "\Temp\" & contaTipos & ".jpg")
            ListProd1.CellAlignment = 2
            ListProd1.TextMatrix(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
            
        End If
    Else
        ListProd1.Row = ListProd1.Rows - 1
        ListProd1.Col = contaCasillas
        ListProd1.CellAlignment = 2
        ListProd1.TextMatrix(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
    End If
    tipoId(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("CLAVE")
    tipoValor(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
    contaCasillas = contaCasillas + 1
    contaTipos = contaTipos + 1
    
    RESTIPO_PROD.MoveNext
Loop

ListProd1.WordWrap = True

End Sub


Private Sub Timer1_Timer()
    
    Timer1.Enabled = False
    cargaLista_TipoImagen

End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)

    NumerosPunto (txtPrecio.Text)
    If KeyAscii = 27 Then
        txtPrecio.Text = ""
        txtPrecio.Visible = False
    Else
        If KeyAscii = 13 Then
            If ListaUsers.Col = 5 Then
                sql1 = "update productos set prod_precio = '" & Val(txtPrecio.Text) & "' " & _
                "where prod_codigo = '" & ListaUsers.TextMatrix(ListaUsers.Row, 0) & "' "
                con.Execute (sql1)
                
                ListaUsers.TextMatrix(ListaUsers.Row, 5) = FormatCurrency(Val(txtPrecio.Text))
                txtPrecio.Text = ""
                txtPrecio.Visible = False
            Else
                If ListaUsers.Col = 4 Then
                    sql1 = "update productos set prod_cant = '" & Val(txtPrecio.Text) & "' " & _
                    "where prod_codigo = '" & ListaUsers.TextMatrix(ListaUsers.Row, 0) & "' "
                    con.Execute (sql1)
                    
                    ListaUsers.TextMatrix(ListaUsers.Row, 4) = (Val(txtPrecio.Text))
                    txtPrecio.Text = ""
                    txtPrecio.Visible = False
                End If
            End If
            'cargaLista
        End If
    End If
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.Visible = False
    txtPrecio.Text = ""
End Sub

Private Sub txtProd_Change(Index As Integer)
    If lProd(Index).ForeColor = vbRed Then
        lProd(Index).ForeColor = vbBlack
    End If

End Sub
Private Sub cargaToolTips()

Dim titulo As String
Dim Descripcion As String

    titulo = "Nombre del producto"
    Descripcion = "Escribe el nombre del producto o una breve descripción para identificarlo"
    TT1.Title = titulo
    TT1.TipText = Descripcion
    TT1.Style = TTBalloon
    TT1.Icon = TTIconInfo
    TT1.ForeColor = vbWhite
    TT1.BackColor = &HCE7110
    TT1.PopupOnDemand = False
    TT1.VisibleTime = 6000
    TT1.CreateToolTip txtProd(0).hWnd
                    
    titulo = "Código del producto"
    Descripcion = "Escribe la clave o código del producto. " & vbCrLf & vbclrf & _
            "Con esta será identificado para su venta tecleandola o desde un lector de código de barras"
    TT2.Title = titulo
    TT2.TipText = Descripcion
    TT2.Style = TTBalloon
    TT2.Icon = TTIconInfo
    TT2.ForeColor = vbWhite
    TT2.BackColor = &HCE7110
    TT2.PopupOnDemand = False
    TT2.VisibleTime = 6000
    TT2.CreateToolTip txtProd(2).hWnd
        
    titulo = "Código del proveedor"
    Descripcion = "Escribe la clave o código con el que el proveedor identfica este propducto. " & vbCrLf & vbCrLf & _
            "Con esta será identificado futuros pedidos."
    TT3.Title = titulo
    TT3.TipText = Descripcion
    TT3.Style = TTBalloon
    TT3.Icon = TTIconInfo
    TT3.ForeColor = vbWhite
    TT3.BackColor = &HCE7110
    TT3.PopupOnDemand = False
    TT3.VisibleTime = 6000
    TT3.CreateToolTip txtProd(10).hWnd
                    
    titulo = "Precio de venta"
    Descripcion = "Escribe el precio de venta del producto" & vbCrLf & vbCrLf & _
            "Este preció se utilizará al realizar una operación de venta"
    TT4.Title = titulo
    TT4.TipText = Descripcion
    TT4.Style = TTBalloon
    TT4.Icon = TTIconInfo
    TT4.ForeColor = vbWhite
    TT4.BackColor = &HCE7110
    TT4.PopupOnDemand = False
    TT4.VisibleTime = 6000
    TT4.CreateToolTip txtProd(3).hWnd
        
    titulo = "Precio de costo"
    Descripcion = "Escribe el precio de costo del producto" & vbCrLf & vbCrLf & _
            "Este preció se utilizará solo para realizar cálculos administrativos"
    TT5.Title = titulo
    TT5.TipText = Descripcion
    TT5.Style = TTBalloon
    TT5.Icon = TTIconInfo
    TT5.ForeColor = vbWhite
    TT5.BackColor = &HCE7110
    TT5.PopupOnDemand = False
    TT5.VisibleTime = 6000
    TT5.CreateToolTip txtProd(8).hWnd
        
    titulo = "Precio de mayoreo"
    Descripcion = "Escribe el precio de mayoreo del producto" & vbCrLf & vbCrLf & _
        "Este preció se utilizará solo para cliente con venta de mayoreo"
    TT6.Title = titulo
    TT6.TipText = Descripcion
    TT6.Style = TTBalloon
    TT6.Icon = TTIconInfo
    TT6.ForeColor = vbWhite
    TT6.BackColor = &HCE7110
    TT6.PopupOnDemand = False
    TT6.VisibleTime = 6000
    TT6.CreateToolTip txtProd(9).hWnd
        
    titulo = "Cantidad"
    Descripcion = "Escribe la cantidad de productos" & vbCrLf & vbCrLf & _
        "Debe de ser el total de productos en existencia"
    TT7.Title = titulo
    TT7.TipText = Descripcion
    TT7.Style = TTBalloon
    TT7.Icon = TTIconInfo
    TT7.ForeColor = vbWhite
    TT7.BackColor = &HCE7110
    TT7.PopupOnDemand = False
    TT7.VisibleTime = 6000
    TT7.CreateToolTip txtProd(2).hWnd
        
    titulo = "Sotkc mínimo"
    Descripcion = "Escribe la cantidad mínima de productos que deben haber en existencia" & vbCrLf & vbCrLf & _
        "Una vez que los productos lleguen a esta cantidad se marcará para que puedan ser identificados"
    TT8.Title = titulo
    TT8.TipText = Descripcion
    TT8.Style = TTBalloon
    TT8.Icon = TTIconInfo
    TT8.ForeColor = vbWhite
    TT8.BackColor = &HCE7110
    TT8.PopupOnDemand = False
    TT8.VisibleTime = 6000
    TT8.CreateToolTip txtProd(4).hWnd
        
    titulo = "Sotkc máximo"
    Descripcion = "Escribe la cantidad máxima de productos que deben haber en existencia" & vbCrLf & vbCrLf & _
        "Una vez que los productos lleguen a esta cantidad se marcará para que puedan ser identificados al momento de dar de alta"
    TT9.Title = titulo
    TT9.TipText = Descripcion
    TT9.Style = TTBalloon
    TT9.Icon = TTIconInfo
    TT9.ForeColor = vbWhite
    TT9.BackColor = &HCE7110
    TT9.PopupOnDemand = False
    TT9.VisibleTime = 6000
    TT9.CreateToolTip txtProd(5).hWnd
            
    titulo = "Valor de presentación"
    Descripcion = "Escribe la forma de presentar el producto" & vbCrLf & vbCrLf & _
        "Puede ser una unidad de medidad, talla o porción"
    TT10.Title = titulo
    TT10.TipText = Descripcion
    TT10.Style = TTBalloon
    TT10.Icon = TTIconInfo
    TT10.ForeColor = vbWhite
    TT10.BackColor = &HCE7110
    TT10.PopupOnDemand = False
    TT10.VisibleTime = 6000
    TT10.CreateToolTip txtProd(6).hWnd
        
    titulo = "Descripción"
        Descripcion = "Escribe una descripción general del producto"
    TT11.Title = titulo
    TT11.TipText = Descripcion
    TT11.Style = TTBalloon
    TT11.Icon = TTIconInfo
    TT11.ForeColor = vbWhite
    TT11.BackColor = &HCE7110
    TT11.PopupOnDemand = False
    TT11.VisibleTime = 6000
    TT11.CreateToolTip txtProd(7).hWnd

    titulo = "Marca"
    Descripcion = "Selecciona la marca del producto" & vbCrLf & vbCrLf & _
        "Si la marca no se encuentra da clic en -Administrar marcas- (botón de la derecha) para agregar o modificar. >"
    TT12.Title = titulo
    TT12.TipText = Descripcion
    TT12.Style = TTBalloon
    TT12.Icon = TTIconInfo
    TT12.ForeColor = vbWhite
    TT12.BackColor = &HCE7110
    TT12.PopupOnDemand = False
    TT12.VisibleTime = 6000
    TT12.CreateToolTip cmbProd(0).hWnd

    titulo = "Tipo"
    Descripcion = "Selecciona el tipo de clasificación para este producto" & vbCrLf & vbCrLf & _
        "Si el tipo no se encuentra da clic en -Administrar tipos- (botón de la derecha) para agregar o modificar. >"
    TT13.Title = titulo
    TT13.TipText = Descripcion
    TT13.Style = TTBalloon
    TT13.Icon = TTIconInfo
    TT13.ForeColor = vbWhite
    TT13.BackColor = &HCE7110
    TT13.PopupOnDemand = False
    TT13.VisibleTime = 6000
    TT13.CreateToolTip cmbProd(1).hWnd

    titulo = "Proveedor"
    Descripcion = "Selecciona el proveedor del producto" & vbCrLf & vbCrLf & _
        "Si el proveedor no se encuentra da clic en -Administrar proveedores- (botón de la derecha) para agregar o modificar. >"
    TT13.Title = titulo
    TT13.TipText = Descripcion
    TT13.Style = TTBalloon
    TT13.Icon = TTIconInfo
    TT13.ForeColor = vbWhite
    TT13.BackColor = &HCE7110
    TT13.PopupOnDemand = False
    TT13.VisibleTime = 6000
    TT13.CreateToolTip cmbProd(3).hWnd

    titulo = "Tipo de presentación"
    Descripcion = "Selecciona la presentación del producto" & vbCrLf & vbCrLf & _
        "Si la presentación no se encuentra da clic en -Administrar presentacions- (botón de la derecha) para agregar o modificar. >"
    TT13.Title = titulo
    TT13.TipText = Descripcion
    TT13.Style = TTBalloon
    TT13.Icon = TTIconInfo
    TT13.ForeColor = vbWhite
    TT13.BackColor = &HCE7110
    TT13.PopupOnDemand = False
    TT13.VisibleTime = 6000
    TT13.CreateToolTip cmbProd(2).hWnd
    
    titulo = "Status"
    Descripcion = "Status del producto" & vbCrLf & vbCrLf & _
        "Si el status es inactivo el producto no podrá ser utilizado para operaciones"
    TT13.Title = titulo
    TT13.TipText = Descripcion
    TT13.Style = TTBalloon
    TT13.Icon = TTIconInfo
    TT13.ForeColor = vbWhite
    TT13.BackColor = &HCE7110
    TT13.PopupOnDemand = False
    TT13.VisibleTime = 6000
    TT13.CreateToolTip cmbProd(4).hWnd
    
    titulo = "Tipo de producto"
    Descripcion = "Tipo del producto" & vbCrLf & vbCrLf & _
        "Si el tipo es -dependiente- dependerá de uno o varios productos y afectará las cantidades de dichos productos. " & vbclrf & vbCrLf & _
        "La cantidad será la proporcional al producto principal."
    TT14.Title = titulo
    TT14.TipText = Descripcion
    TT14.Style = TTBalloon
    TT14.Icon = TTIconInfo
    TT14.ForeColor = vbWhite
    TT14.BackColor = &HCE7110
    TT14.PopupOnDemand = False
    TT14.VisibleTime = 9000
    TT14.CreateToolTip cmbProd(7).hWnd

End Sub
Private Sub txtProd_GotFocus(Index As Integer)
    cmBoton(5).Visible = False
    cmBoton(6).Visible = False
    cmBoton(4).Visible = False
    cmBoton(7).Visible = False

    txtProd(Index).SelStart = 0
    txtProd(Index).SelLength = Len(txtProd(Index))

End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
         
    Select Case Index
        Case 4: Call Numeros(KeyAscii)
        Case 5: Call Numeros(KeyAscii)
        'Case 6: Call Numeros(KeyAscii)
        Case 2: Call Numeros(KeyAscii)
        Case 8: Call NumerosPunto(KeyAscii)
        Case 3: Call NumerosPunto(KeyAscii)
        Case 11: Call NumerosPunto(KeyAscii)
                    If KeyAscii = 13 Then
                        listDependiente.TextMatrix(listDependiente.Row, 2) = txtProd(11).Text
                        txtProd(11).Text = ""
                        txtProd(11).Visible = False
                    Else
                        If KeyAscii = 27 Then
                            txtProd(11).Text = ""
                            txtProd(11).Visible = False
                            Exit Sub
                        End If
                    End If
    End Select

     If mayus = True Then
        Call Mayusculas(KeyAscii)
     End If


End Sub

Private Sub txtProd_LostFocus(Index As Integer)
    If Index = 11 Then
        txtProd(11).Visible = False
    Else
        If Index = 1 Then
            sql1 = "SELECT prod_nombre, prod_codigo FROM productos where prod_codigo = '" & txtProd(1).Text & "'"
            Set RES1 = con.Execute(sql1)
            
            If Not RES1.EOF Then
                MsgBox "El codigo ya existe en el producto:  " & vbCrLf & "Código: " & RES1.Fields("Prod_Codigo") & vbCrLf & "Producto: " & RES1.Fields("Prod_nombre") & " ", vbInformation
                txtProd(1).SelStart = 0
                txtProd(1).SelLength = Len(txtProd(1).Text)
                txtProd(1).SetFocus
            End If
        End If
    End If
    TT1.Destroy
End Sub
