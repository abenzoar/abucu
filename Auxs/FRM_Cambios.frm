VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_Cambios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios - Devoluciones"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   18150
   Icon            =   "FRM_Cambios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRM_Cambios.frx":058A
   ScaleHeight     =   10125
   ScaleWidth      =   18150
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   10095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   17806
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Cambios/Devoluciones realizados"
      TabPicture(0)   =   "FRM_Cambios.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cambio/Devolución"
      TabPicture(1)   =   "FRM_Cambios.frx":10AE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtEntrega(2)"
      Tab(1).Control(1)=   "txtEntrega(1)"
      Tab(1).Control(2)=   "txtProd(3)"
      Tab(1).Control(3)=   "txtProd(2)"
      Tab(1).Control(4)=   "txtProd(1)"
      Tab(1).Control(5)=   "txtProd(0)"
      Tab(1).Control(6)=   "txtDif"
      Tab(1).Control(7)=   "txtEntrega(0)"
      Tab(1).Control(8)=   "txtDevo"
      Tab(1).Control(9)=   "txtInfoAprt(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtClave(2)"
      Tab(1).Control(11)=   "cmBoton(2)"
      Tab(1).Control(12)=   "cmBoton(1)"
      Tab(1).Control(13)=   "txtInfoAprt(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmBoton(0)"
      Tab(1).Control(15)=   "txtClave(1)"
      Tab(1).Control(16)=   "txtInfoAprt(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtClave(0)"
      Tab(1).Control(18)=   "Lista2"
      Tab(1).Control(19)=   "Label1(17)"
      Tab(1).Control(20)=   "Line1(18)"
      Tab(1).Control(21)=   "Label1(16)"
      Tab(1).Control(22)=   "Line1(17)"
      Tab(1).Control(23)=   "Line1(16)"
      Tab(1).Control(24)=   "Label1(15)"
      Tab(1).Control(25)=   "Line1(15)"
      Tab(1).Control(26)=   "Label1(14)"
      Tab(1).Control(27)=   "Label1(13)"
      Tab(1).Control(28)=   "Line1(14)"
      Tab(1).Control(29)=   "Label1(12)"
      Tab(1).Control(30)=   "Line1(13)"
      Tab(1).Control(31)=   "Label1(11)"
      Tab(1).Control(32)=   "Line1(12)"
      Tab(1).Control(33)=   "Label1(10)"
      Tab(1).Control(34)=   "Line1(11)"
      Tab(1).Control(35)=   "Line1(10)"
      Tab(1).Control(36)=   "Label1(9)"
      Tab(1).Control(37)=   "Line1(8)"
      Tab(1).Control(38)=   "Label1(8)"
      Tab(1).Control(39)=   "Line1(7)"
      Tab(1).Control(40)=   "Label1(7)"
      Tab(1).Control(41)=   "Label1(6)"
      Tab(1).Control(42)=   "Line1(6)"
      Tab(1).Control(43)=   "Line1(5)"
      Tab(1).Control(44)=   "Label1(5)"
      Tab(1).Control(45)=   "Label1(4)"
      Tab(1).Control(46)=   "Line1(4)"
      Tab(1).Control(47)=   "lInfo(2)"
      Tab(1).Control(48)=   "Label1(3)"
      Tab(1).Control(49)=   "Line1(3)"
      Tab(1).Control(50)=   "imgFoto(1)"
      Tab(1).Control(51)=   "lblDatos(1)"
      Tab(1).Control(52)=   "Line1(2)"
      Tab(1).Control(53)=   "Label1(2)"
      Tab(1).Control(54)=   "Line1(1)"
      Tab(1).Control(55)=   "Label1(1)"
      Tab(1).Control(56)=   "Line1(0)"
      Tab(1).Control(57)=   "Label1(0)"
      Tab(1).Control(58)=   "imgFoto(0)"
      Tab(1).Control(59)=   "lblDatos(0)"
      Tab(1).Control(60)=   "Line1(9)"
      Tab(1).Control(61)=   "Label1(19)"
      Tab(1).ControlCount=   62
      Begin VB.TextBox txtEntrega 
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
         Index           =   2
         Left            =   -61800
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "$0.0"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtEntrega 
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
         Left            =   -61800
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "$0.0"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   -69000
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "$0.0"
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   -70920
         TabIndex        =   35
         Text            =   "0"
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   -72840
         TabIndex        =   33
         Text            =   "0"
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "$0.0"
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox txtDif 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   585
         Left            =   -65280
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "$0.0"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtEntrega 
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
         Left            =   -61800
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "$0.0"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtDevo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -65280
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "$0.0"
         Top             =   1920
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
         Height          =   1455
         Index           =   2
         Left            =   -65280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4800
         Width           =   6015
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
         Left            =   -73440
         TabIndex        =   0
         Top             =   2160
         Width           =   1575
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
         Height          =   1095
         Index           =   2
         Left            =   -60000
         Picture         =   "FRM_Cambios.frx":1648
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar cambios y realizar devolución"
         Enabled         =   0   'False
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
         Left            =   -64200
         Picture         =   "FRM_Cambios.frx":1F12
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6600
         Width           =   2655
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
         Height          =   2175
         Index           =   1
         Left            =   -71400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4080
         Width           =   5775
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar cambio para devolución"
         Enabled         =   0   'False
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
         Left            =   -66960
         Picture         =   "FRM_Cambios.frx":27DC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6600
         Width           =   2655
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   -73440
         TabIndex        =   9
         Top             =   5280
         Width           =   1575
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
         Height          =   2775
         Index           =   0
         Left            =   -71400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   5775
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
         Left            =   -73440
         TabIndex        =   1
         Top             =   3000
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid Lista1 
         Height          =   7215
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   18
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Cambios.frx":2D66
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
         Height          =   1815
         Left            =   -74760
         TabIndex        =   29
         Top             =   8040
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Cambios.frx":2F18
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
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
         Left            =   -61800
         TabIndex        =   42
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   18
         X1              =   -61800
         X2              =   -59280
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total devolución"
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
         Left            =   -61800
         TabIndex        =   40
         Top             =   480
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   17
         X1              =   -61800
         X2              =   -59280
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   16
         X1              =   -69000
         X2              =   -67320
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   15
         Left            =   -69000
         TabIndex        =   38
         Top             =   6360
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   15
         X1              =   -70920
         X2              =   -69240
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento $"
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
         Index           =   14
         Left            =   -70920
         TabIndex        =   36
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento %"
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
         Index           =   13
         Left            =   -72840
         TabIndex        =   34
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   14
         X1              =   -72840
         X2              =   -71160
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio producto"
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
         Left            =   -74760
         TabIndex        =   32
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   13
         X1              =   -74760
         X2              =   -73080
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de cambios para devolución"
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   7680
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   12
         X1              =   -74760
         X2              =   -58680
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total (diferencia)"
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
         Left            =   -65280
         TabIndex        =   28
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   11
         X1              =   -65280
         X2              =   -62880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   10
         X1              =   -61800
         X2              =   -59280
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total devolución"
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
         Left            =   -61800
         TabIndex        =   27
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   8
         X1              =   -65280
         X2              =   -62760
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total venta anterior"
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
         Left            =   -65280
         TabIndex        =   26
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   7
         X1              =   -73440
         X2              =   -71760
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Cód prod."
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
         Left            =   -73440
         TabIndex        =   22
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Index           =   6
         Left            =   -65280
         TabIndex        =   21
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   6
         X1              =   -65280
         X2              =   -59280
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   5
         X1              =   -73440
         X2              =   -71760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Folio Ticket"
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
         Index           =   5
         Left            =   -73440
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Cód prod."
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
         Left            =   -73440
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   4
         X1              =   -73440
         X2              =   -71760
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Abierto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   2
         Left            =   -65280
         TabIndex        =   15
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus del producto para devolución"
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
         Left            =   -65280
         TabIndex        =   14
         Top             =   480
         Width           =   3255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   3
         X1              =   -65280
         X2              =   -62040
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   1
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   4080
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
         Index           =   1
         Left            =   -73440
         TabIndex        =   10
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   -71400
         X2              =   -65640
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Información de la operación del producto"
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
         Left            =   -71400
         TabIndex        =   8
         Top             =   3720
         Width           =   4215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   1
         X1              =   -74760
         X2              =   -71760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto que se entrega al cliente"
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
         Left            =   -74760
         TabIndex        =   7
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   0
         X1              =   -71400
         X2              =   -65640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Información de la operación del producto"
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
         Left            =   -71400
         TabIndex        =   6
         Top             =   480
         Width           =   4215
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   0
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   840
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
         Left            =   -73440
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   9
         X1              =   -74760
         X2              =   -71760
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto que devuelve el cliente"
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
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Opciones"
      Begin VB.Menu mn_NewDevo 
         Caption         =   "Realizar una devolución"
      End
      Begin VB.Menu mn_Cliente 
         Caption         =   "Agregar un cliente a la devolución"
      End
      Begin VB.Menu mn_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Cancelar 
         Caption         =   "Cancelar cambio"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn_Reprint 
         Caption         =   "Reimprimir ticket"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FRM_Cambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resProdVenta As Recordset
Dim resProdDevo As Recordset
Dim RES1 As Recordset
Dim SQL1 As String
Dim resMaxId As Recordset

Private Sub cmBoton_Click(Index As Integer)
    If Index = 2 Then
        cancelar
    Else
        If Index = 0 Then
            If txtInfoAprt(0).Text <> "" And txtInfoAprt(1).Text <> "" And lInfo(2).Caption = "Permitido" Then
                agregarLista
            Else
                MsgBox "Opción no dispoinible. Verifique.", vbInformation
            End If
        Else
            If Index = 1 Then
                If txtInfoAprt(1).Text <> "" Then
                    cmBoton_Click (0)
                End If
                
                If Val(Format(txtDif.Text, "General Number")) < 0 Then
                    Dim ques As String
                    ques = MsgBox("La diferencia de la devolución del precio del producto actual es menor al precio de compra del producto original." & vbCrLf & vbCrLf & _
                    "La diferencia de : " & txtDif.Text & " se abonará a la cuenta del cliente" & vbCrLf & vbCrLf & _
                    "Verifique que el cliente seleccionado sea el adecuado. " & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
                    If ques = vbNo Then
                        Exit Sub
                    End If
                End If
                    tipoCobro = "CAMBIOS"
                    FRM_Cobro.txtTot.Text = txtDif.Text
                    FRM_Cobro.Show vbModal
            End If
        End If
    End If
End Sub
Public Sub realizarCambios()
    Dim idCambio As Long
    Dim folioPago As Long
    
    SQL1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDPERID, VENT_VENDTIPOID, VENT_VENDTIPO, " & _
    "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO, vent_SubTotal, vent_Descuento, vent_Total, vent_Pagado, vent_Cambio, " & _
    "vent_PagoEfectivo, vent_PagoTarjeta, vent_PagoCheque, vent_FechaHora_Cobro) VALUES " & _
    "(NOW(), 'B', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', " & _
    "'" & resProdVenta.Fields("CLIE_PERID") & "', '" & resProdVenta.Fields("CLIE_ID") & "', '" & resProdVenta.Fields("CLIE_TIPO") & "', '" & Val(Format(txtEntrega(1).Text, "General Number")) & "', '" & Val(Format(txtEntrega(2).Text, "General Number")) & "', '" & Val(Format(txtEntrega(0).Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(4).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtCambio.Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(0).Text, "General Number")) & "', " & _
    "'" & Val(Format(FRM_Cobro.txtPago(1).Text, "General Number")) & "', '" & Val(Format(FRM_Cobro.txtPago(2).Text, "General Number")) & "', NOW())"
    con.Execute (SQL1)
    
    SQL1 = "select last_insert_id() folioId"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        folioPago = RES1.Fields("folioId")
    End If
    
    If Val(Format(txtDif.Text, "General Number")) < 0 Then
        SQL1 = "UPDATE PER_TIPO SET TEMP_MONEDERO = TEMP_MONEDERO + '" & (Val(Format(txtDif.Text, "General Number")) * -1) & "' WHERE PERTP_TIPO_ID = '" & resProdVenta.Fields("CLIE_ID") & "' AND PERTP_PER_ID = '" & resProdVenta.Fields("CLIE_PERID") & "'"
        con.Execute (SQL1)
        
'        SQL1 = "SELECT PERTP_MEMBRESIA  FROM PER_TIPO WHERE PERTP_PER_ID = '" & resProdVenta.Fields("CLIE_PERID") & "' AND  PERTP_TIPO_ID =  '" & resProdVenta.Fields("CLIE_ID") & "' AND  PERTP_PER_TIPO =  '" & resProdVenta.Fields("CLIE_TIPO") & "'"
'        Set res1 = con.Execute(SQL1)
'
'        If res1.Fields("PERTP_MEMBRESIA") = "S" Then
    
            SQL1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
            "VALUES ('D', '" & resProdVenta.Fields("CLIE_PERID") & "', '" & resProdVenta.Fields("CLIE_ID") & "', '" & resProdVenta.Fields("CLIE_TIPO") & "', '" & folioPago & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', " & _
            "'" & (Val(Format(txtDif.Text, "General Number")) * -1) & "', 'R', NOW() )"
            
            con.Execute (SQL1)
            
'        End If
        
    End If
           
    SQL1 = "INSERT INTO CAMBIOS (CMB_FOLIOVENTA, CMB_FECHA, CMB_CLIE, CMB_CLIEPERID, CMB_CLIETIPO, CMB_USUARIO, CMB_USUPERID, CMB_USUTIPO, CMB_OBSERVACIONES, cmb_FolioVentCambio, cmd_tipo) values " & _
    "('" & resProdVenta.Fields("FOLIO") & "', now(), '" & resProdVenta.Fields("CLIE_ID") & "', '" & resProdVenta.Fields("CLIE_PERID") & "', '" & resProdVenta.Fields("CLIE_TIPO") & "', '" & resProdVenta.Fields("VEND_ID") & "', '" & resProdVenta.Fields("VEND_PERID") & "', '" & resProdVenta.Fields("VEND_TIPO") & "', '" & txtInfoAprt(2).Text & "', '" & folioPago & "', '" & resProdVenta.Fields("TIPO_CMB") & "') "
    con.Execute (SQL1)
    
    SQL1 = "select last_insert_id() idCambio"
    Set resMaxId = con.Execute(SQL1)
    If Not resMaxId.EOF Then
        idCambio = resMaxId.Fields("idCambio")
        folioCambio = resMaxId.Fields("idCambio")
    End If

    For b1 = 1 To Lista2.Rows - 1
        SQL1 = "INSERT INTO CAMBIOS_DETALLE (CMBD_ID, CMBD_FOLIOVENTA, CMBD_PRODID_DEV, CMBD_PRODSER_DEV, CMBD_PRODID_CAM, CMBD_PRODSER_CAM, CMBD_PRODSER_CAMNOMBRE, CMBD_PRODSER_CAMPRECIO, CMBD_PRODSER_CAMDESCUENTO, CMBD_PRODSER_CAMTOTAL, CMBD_PRODSER_DEVNOMBRE, CMBD_PRODSER_DEVPRECIO, cmbd_ProdSer_DevCodigo, cmbd_ProdSer_CamCodigo) " & _
        "VALUES ('" & idCambio & "', '" & resProdVenta.Fields("FOLIO") & "', '" & Lista2.TextMatrix(b1, 13) & "', '" & Lista2.TextMatrix(b1, 14) & "', '" & Lista2.TextMatrix(b1, 15) & "', '" & Lista2.TextMatrix(b1, 16) & "', '" & Lista2.TextMatrix(b1, 4) & "', " & _
        "'" & Val(Format(Lista2.TextMatrix(b1, 6), "General Number")) & "',  '" & Val(Format(Lista2.TextMatrix(b1, 8), "General Number")) & "', '" & Val(Format(Lista2.TextMatrix(b1, 9), "General Number")) & "', '" & Lista2.TextMatrix(b1, 1) & "', '" & Lista2.TextMatrix(b1, 3) & "', '" & Lista2.TextMatrix(b1, 2) & "', '" & Lista2.TextMatrix(b1, 5) & "')"
        con.Execute (SQL1)
    
        SQL1 = "UPDATE PRODUCTOS SET PROD_CANT = (PROD_CANT + 1) WHERE PROD_ID = '" & Lista2.TextMatrix(b1, 13) & "' AND PROD_SERV = '" & Lista2.TextMatrix(b1, 14) & "' "
        con.Execute (SQL1)
        
        SQL1 = "UPDATE PRODUCTOS SET PROD_CANT = (PROD_CANT - 1) WHERE PROD_ID = '" & Lista2.TextMatrix(b1, 15) & "' AND PROD_SERV = '" & Lista2.TextMatrix(b1, 16) & "' "
        con.Execute (SQL1)
        
    Next b1
    
    
    
    limpiar
    cargaCambios
    
    MsgBox "Información de cambios realizada. ", vbInformation
    
    SSTab1.Tab = 0
    
End Sub

Private Sub agregarLista()

    For b1 = 1 To Lista2.Rows - 1
        If Lista2.TextMatrix(b1, 2) = resProdVenta.Fields("CODIGO") Then
            MsgBox "El producto que desea cambiar ya se encuentra en la lista. Verifique. ", vbInformation
            Exit Sub
        End If
    Next b1

    Lista2.AddItem ""
    
    Lista2.TextMatrix(Lista2.Rows - 1, 0) = resProdVenta.Fields("FOLIO")
    Lista2.TextMatrix(Lista2.Rows - 1, 1) = resProdVenta.Fields("PRODUCTO")
    Lista2.TextMatrix(Lista2.Rows - 1, 2) = resProdVenta.Fields("CODIGO")
    Lista2.TextMatrix(Lista2.Rows - 1, 3) = resProdVenta.Fields("PRECIO")

    Lista2.TextMatrix(Lista2.Rows - 1, 4) = resProdDevo.Fields("PROD_NOMBRE")
    Lista2.TextMatrix(Lista2.Rows - 1, 5) = resProdDevo.Fields("PROD_CODIGO")
    Lista2.TextMatrix(Lista2.Rows - 1, 6) = FormatCurrency(resProdDevo.Fields("PROD_PRECIO"))
    
    Lista2.TextMatrix(Lista2.Rows - 1, 7) = FormatCurrency(txtProd(1).Text)
    Lista2.TextMatrix(Lista2.Rows - 1, 8) = txtProd(2).Text
    Lista2.TextMatrix(Lista2.Rows - 1, 9) = txtProd(3).Text
    
    Lista2.TextMatrix(Lista2.Rows - 1, 10) = resProdVenta.Fields("FECHA_HORA")
    Lista2.TextMatrix(Lista2.Rows - 1, 11) = resProdVenta.Fields("USUARIO")
    Lista2.TextMatrix(Lista2.Rows - 1, 12) = resProdVenta.Fields("CLIENTE")
    
    Lista2.TextMatrix(Lista2.Rows - 1, 13) = resProdVenta.Fields("PROD_ID")
    Lista2.TextMatrix(Lista2.Rows - 1, 14) = resProdVenta.Fields("PROD_SER")
    
    Lista2.TextMatrix(Lista2.Rows - 1, 15) = resProdDevo.Fields("PROD_ID")
    Lista2.TextMatrix(Lista2.Rows - 1, 16) = resProdDevo.Fields("PROD_SERV")
    
    

    Lista2.Row = Lista2.Rows - 1
    Lista2.Col = 1
    Lista2.CellBackColor = &HC0E0FF
    Lista2.Col = 2
    Lista2.CellBackColor = &HC0E0FF
    Lista2.Col = 3
    Lista2.CellBackColor = &HC0E0FF

    Lista2.Col = 4
    Lista2.CellBackColor = &HFFC0C0
    Lista2.Col = 5
    Lista2.CellBackColor = &HFFC0C0
    Lista2.Col = 6
    Lista2.CellBackColor = &HFFC0C0
    Lista2.Col = 7
    Lista2.CellBackColor = &HFFC0C0
    Lista2.Col = 8
    Lista2.CellBackColor = &HFFC0C0
    Lista2.Col = 9
    Lista2.CellBackColor = &HFFC0C0

    despuesAgregar
    
End Sub
Private Sub checkPrecio()
    Dim totVenta As Double
    Dim totDevo As Double
    Dim totDif As Double
    Dim totDesc As Double
    
    totVenta = 0
    totDevo = 0
    totDif = 0
    totDesc = 0
    For b1 = 1 To Lista2.Rows - 1
        totVenta = totVenta + Val(Format(Lista2.TextMatrix(b1, 3), "General Number"))
        totDevo = totDevo + Val(Format(Lista2.TextMatrix(b1, 6), "General Number"))
        totDesc = totDesc + Val(Format(Lista2.TextMatrix(b1, 8), "General Number"))
    Next b1

    txtDevo.Text = FormatCurrency(totVenta)
    
    txtEntrega(1).Text = FormatCurrency(totDevo)
    txtEntrega(2).Text = FormatCurrency(totDesc)
    txtEntrega(0).Text = FormatCurrency(totDevo - totDesc)
    totDif = totDevo - totDesc - totVenta
        
    txtDif.Text = FormatCurrency(totDif)

End Sub
Private Sub cancelar()
Dim ques As String
    
    ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        limpiar
    End If
End Sub
Private Sub limpiar()
        cmBoton(1).Enabled = False
        cmBoton(0).Enabled = False
        For b1 = 0 To 2
            txtClave(b1).Text = ""
            txtInfoAprt(b1).Text = ""
        Next b1
        lblDatos(0).Caption = ""
        lblDatos(1).Caption = ""
        
        imgFoto(0).Picture = LoadPicture("")
        imgFoto(1).Picture = LoadPicture("")
        
        Lista2.Rows = 1
        txtClave(2).Enabled = True
        
        lblDatos(0).Caption = ""
        lblDatos(1).Caption = ""
                
        txtDevo.Text = ""
        txtDif.Text = ""
        txtEntrega(0).Text = ""
        txtEntrega(1).Text = ""
        txtEntrega(2).Text = ""
        
        For b1 = 0 To 3
            txtProd(b1).Text = "0"
        Next b1
        
        imgFoto(0).Picture = LoadPicture("")
        imgFoto(1).Picture = LoadPicture("")
        
        mn_NewDevo.Enabled = True
        mn_Cliente.Enabled = True
        mn_Cancelar.Enabled = False
        mn_Reprint.Enabled = False

End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    lInfo(2).Caption = ""
    Lista2.Rows = 1
    Lista2.ColWidth(13) = 0
    Lista2.ColWidth(14) = 0
    Lista2.ColWidth(15) = 0
    Lista2.ColWidth(16) = 0
        
    cargaCambios
End Sub
Private Sub cargaCambios()
    
    Dim resLista As Recordset
    SQL1 = "SELECT * fROM VIEW_CAMBIOS"
    Set resLista = con.Execute(SQL1)
    
    Lista1.Rows = 1
    
    Do While Not resLista.EOF
        Lista1.AddItem ""
        Lista1.TextMatrix(Lista1.Rows - 1, 0) = resLista.Fields("FECHA_dEVO")
        Lista1.TextMatrix(Lista1.Rows - 1, 1) = resLista.Fields("CLIENTE")
        Lista1.TextMatrix(Lista1.Rows - 1, 2) = resLista.Fields("DEV_PRODUCTO")
        Lista1.TextMatrix(Lista1.Rows - 1, 3) = resLista.Fields("PROD_DEVCODIGO")
        
        Lista1.TextMatrix(Lista1.Rows - 1, 4) = FormatCurrency(resLista.Fields("DEV_PRECIO"))
        Lista1.TextMatrix(Lista1.Rows - 1, 5) = resLista.Fields("CAM_PRODUCTO")
        Lista1.TextMatrix(Lista1.Rows - 1, 6) = resLista.Fields("PROD_CAMCODIGO")
        
        Lista1.TextMatrix(Lista1.Rows - 1, 7) = FormatCurrency(resLista.Fields("CAM_PRECIO"))
        Lista1.TextMatrix(Lista1.Rows - 1, 8) = FormatCurrency(resLista.Fields("CAM_DESC"))
        Lista1.TextMatrix(Lista1.Rows - 1, 9) = FormatCurrency(resLista.Fields("CAM_TOTAL"))
        Lista1.TextMatrix(Lista1.Rows - 1, 10) = FormatCurrency(resLista.Fields("TOT_DIF"))
        Lista1.TextMatrix(Lista1.Rows - 1, 11) = resLista.Fields("USUARIO_dEVO")
        Lista1.TextMatrix(Lista1.Rows - 1, 12) = resLista.Fields("FECHA_VENTA")
        Lista1.TextMatrix(Lista1.Rows - 1, 13) = resLista.Fields("DIAS")
        Lista1.TextMatrix(Lista1.Rows - 1, 14) = resLista.Fields("FOLIO_VENTA")
        Lista1.TextMatrix(Lista1.Rows - 1, 15) = resLista.Fields("Tipo")
        Lista1.TextMatrix(Lista1.Rows - 1, 16) = resLista.Fields("CMB_OBSERVACIONES")
        Lista1.TextMatrix(Lista1.Rows - 1, 17) = resLista.Fields("CMB_ID")
        
        If Val(resLista.Fields("TOT_DIF")) < 0 Then
            Lista1.Row = Lista1.Rows - 1
            Lista1.Col = 10
            Lista1.CellBackColor = &H8080FF
            Lista1.Col = 0
            Lista1.CellBackColor = &H8080FF
        End If
        
        resLista.MoveNext
    Loop
    
    
End Sub

Private Sub Lista1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Lista1.Rows > 1 Then
        If Button = vbRightButton Then
            mn_NewDevo.Enabled = False
            mn_Cliente.Enabled = False
            mn_Cancelar.Enabled = False
            mn_Reprint.Enabled = True
            PopupMenu mn_Menu, vbPopupMenuLeftAlign
        End If
    End If
End Sub

Private Sub Lista2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Lista2.Rows > 1 Then
        If Button = vbRightButton Then
            mn_NewDevo.Enabled = False
            mn_Cliente.Enabled = False
            mn_Cancelar.Enabled = True
            mn_Reprint.Enabled = False
            PopupMenu mn_Menu, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Cancelar_Click()
    Dim ques As String
    
    ques = MsgBox("¿Cancelar cambio producto " & vbCrLf & vbCrLf & Lista2.TextMatrix(Lista2.Row, 1) & vbCrLf & vbCrLf & " por " & vbCrLf & vbCrLf & Lista2.TextMatrix(Lista2.Row, 4) & "?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        If Lista2.Rows = 2 Then
            Lista2.Rows = 1
        Else
            If Lista2.Rows > 2 Then
                Lista2.RemoveItem (Lista2.Row)
            End If
        End If
        checkPrecio
        
        despuesAgregar
    End If
End Sub
Private Sub despuesAgregar()
    txtClave(2).Enabled = False
    
    txtInfoAprt(0).Text = ""
    txtInfoAprt(1).Text = ""
    txtClave(0).Text = ""
    txtClave(1).Text = ""
    'txtClave(2).Text = ""
    lInfo(2).Caption = ""
    lblDatos(0).Caption = ""
    lblDatos(1).Caption = ""

    checkPrecio

    imgFoto(0).Picture = LoadPicture("")
    imgFoto(1).Picture = LoadPicture("")
    
    For b1 = 0 To 3
        txtProd(b1).Text = ""
    Next b1
    
        mn_NewDevo.Enabled = True
        mn_Cliente.Enabled = True
        mn_Cancelar.Enabled = False
End Sub
Private Sub mn_Cliente_Click()
    tipoPersona = "CLIENTE_DEVO"
    ADD_Cliente.Show vbModal
End Sub

Private Sub mn_Reprint_Click()
    Dim ques As String
    ques = MsgBox("Imprimir ticket cambio clave: " & Lista1.TextMatrix(Lista1.Row, 17), vbYesNo + vbQuestion)
    
    If ques = vbYes Then
        notaCambio (Lista1.TextMatrix(Lista1.Row, 17))
    End If
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        txtClave(2).SetFocus
    End If
End Sub

Private Sub txtClave_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
           txtClave(Index).Text = Replace(txtClave(Index).Text, "'", "-")
           If Left(txtClave(Index).Text, 1) = " " Then
                txtClave(Index).Text = Right(txtClave(Index).Text, (Len(txtClave(Index).Text) - 1))
           End If
        If Index = 0 Then
          dev_checkProducto
        Else
            If Index = 2 Then
                dev_checkProducto
            Else
                If Index = 1 Then
                    checkProducto
                End If
            End If
        End If
    End If
End Sub


Private Sub checkProducto()
    'On Error Resume Next
    SQL1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "PROD_CODIGO = '" & txtClave(1).Text & "' AND PROD_STATUS = 'A'"
    Set resProdDevo = con.Execute(SQL1)
    Dim b1 As Long

    If Not resProdDevo.EOF Then
        lblDatos(1).Caption = resProdDevo.Fields("PROD_NOMBRE")
        If IsNull(resProdDevo.Fields("PROD_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write resProdDevo.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(1).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            imgFoto(1).Picture = LoadPicture("")
        End If
        
        txtInfoAprt(1).Text = "Producto" & resProdDevo.Fields("PROD_NOMBRE") & vbCrLf & _
        "Codigo: " & resProdDevo.Fields("PROD_CODIGO") & vbCrLf & "Precio: " & resProdDevo.Fields("Prod_precio") & vbCrLf & _
        "Tipo: " & resProdDevo.Fields("Prod_Tipo") & vbCrLf & "Marca: " & resProdDevo.Fields("Prod_marca") & vbCrLf & _
        "Cantidad: " & resProdDevo.Fields("PROD_Cant") & vbCrLf & "Descripción: " & resProdDevo.Fields("PROD_DESCRIPCION") '& vbCrLf & "Descuento " & resProdDevo.Fields("descuento")
        
        txtClave(1).Enabled = True
        cmBoton(1).Enabled = True
        cmBoton(0).Enabled = True
        txtProd(0).Text = FormatCurrency(resProdDevo.Fields("PROD_PRECIO"))
        txtProd(1).Text = "0"
        txtProd(2).Text = FormatCurrency(0)
        txtProd(3).Text = FormatCurrency(resProdDevo.Fields("PROD_PRECIO"))
        'txtEntrega.Text = FormatCurrency(resProdDevo.Fields("PROD_PRECIO"))
        'txtDif.Text = Val(Format(txtEntrega.Text, "General Number")) - Val(Format(txtDevo.Text, "General Number"))
        If Val(resProdDevo.Fields("PROD_CANT")) = 0 Then
            MsgBox "El artículo está agotado. Verifique. ", vbInformation
            lInfo(2).Caption = "No permitido"
            txtClave(2).Enabled = False
            For b1 = 0 To 1
                cmBoton(b1).Enabled = False
            Next b1
        Else
            lInfo(2).Caption = "Permitido"
            txtClave(1).Enabled = True
            txtClave(1).SetFocus

        End If
        
    Else
        'txtClave(1).Enabled = False
        txtInfoAprt(1).Text = ""
        'txtEntrega.Text = ""
        'txtDif.Text = ""
        cmBoton(0).Enabled = False
        'cmBoton(0).Enabled = False
        
        MsgBox "Producto no encontrado", vbInformation
    End If
    
    
End Sub

Private Sub dev_checkProducto()
    lInfo(2).Caption = ""
    txtInfoAprt(0).Text = ""
    
    SQL1 = "SELECT * FROM VIEW_CAMBIOS WHERE FOLIO_VENTA = '" & txtClave(2).Text & "' AND PROD_dEVCODIGO = '" & txtClave(0).Text & "'"
    Set resMaxId = con.Execute(SQL1)
    
    If Not resMaxId.EOF Then
        MsgBox "El producto que quiere devolver ya se encuentra registrado como devuelto. Verifique. ", vbInformation
        Exit Sub
    Else
        '''''
    End If
    
    SQL1 = "SELECT FECHA_HORA, CLIENTE, USUARIO, FOLIO, CODIGO, PRODUCTO, PRECIO, DESCUENTO, (((PRECIO * CANTIDAD) - DESCUENTO)/CANTIDAD) PRECIO_TOTAL,  DIAS,  CLIE_PERID, CLIE_ID, CLIE_TIPO, " & _
    "vend_id , vend_perid, vend_tipo, prod_id, prod_ser, 'V' TIPO_CMB " & _
    "From VIEW_VENTASDETALLE " & _
    "WHERE CODIGO = '" & txtClave(0).Text & "' AND FOLIO = '" & txtClave(2).Text & "' " & _
    "UNION ALL " & _
    "SELECT FECHA FECHA_HORA, CLIENTE, MOSTRADOR USUARIO, FOLIO, CODIGO, PRODUCTO, PRECIO, DESCUENTO, TOTAL_PROD PRECIO_TOTAL,  (TO_dAYS(NOW()) -TO_DAYS(FECHA))  DIAS, " & _
    "APRT_CLIEPERID CLIE_PERID, APRT_CLIEID CLIE_ID, 'C' CLIE_TIPO, APRT_USERID VEND_ID, APRT_USERPERID VEND_PERID, 'U' VEND_TIPO, APRT_PRODID PROD_ID, APRT_PRODSERV PROD_SER, 'A' TIPO_CMB  " & _
    "From VIEW_APARTADOS " & _
    "WHERE CODIGO = '" & txtClave(0).Text & "' AND FOLIO = '" & txtClave(2).Text & "' "
    
    Set resProdVenta = con.Execute(SQL1)
    
    If Not resProdVenta.EOF Then
        txtInfoAprt(0).Text = "Fecha/Hora Compra: " & resProdVenta.Fields("FECHA_HORA") & vbCrLf & _
        "Cliente: " & resProdVenta.Fields("CLIENTE") & vbCrLf & "Usuario venta: " & resProdVenta.Fields("USUARIO") & vbCrLf & _
        "Folio venta: " & resProdVenta.Fields("FOLIO") & vbCrLf & "PRODUCTO CODIGO " & resProdVenta.Fields("CODIGO") & vbCrLf & _
        "Producto: " & resProdVenta.Fields("PRODUCTO") & vbCrLf & "Precio: " & resProdVenta.Fields("PRECIO") & vbCrLf & "Descuento " & resProdVenta.Fields("descuento") & _
        vbCrLf & "Dias: " & resProdVenta.Fields("dias")
        
        lblDatos(0).Caption = resProdVenta.Fields("PRODUCTO")
        
        If Val(resProdVenta.Fields("dias")) > 30 Then
            lInfo(2).Caption = "No permitido"
            txtClave(1).Enabled = False
            For b1 = 0 To 1
                cmBoton(b1).Enabled = False
            Next b1
        Else
            lInfo(2).Caption = "Permitido"
            txtClave(1).Enabled = True
            txtClave(1).SetFocus
'            For b1 = 0 To 1
'                cmBoton(b1).Enabled = True
'            Next b1
        End If
    Else
'        txtDevo.Text = ""
'        txtDif.Text = ""
        MsgBox "No se encontró información. Verifique. ", vbInformation
    End If

End Sub
Private Sub valorDescuento(tipo As String)
    If tipo = "Porcentaje" Then
        txtProd(3).Text = (Val(resProdDevo.Fields("prod_precio"))) - (Val(resProdDevo.Fields("prod_precio")) * (Val(txtProd(1).Text) / 100))
        txtProd(3).Text = FormatCurrency(txtProd(3).Text)
        txtProd(2).Text = (Val(resProdDevo.Fields("prod_precio")) * (Val(txtProd(1).Text) / 100))
        txtProd(2).Text = FormatCurrency(txtProd(2).Text)
    Else
        If tipo = "Cantidad" Then
            If Val(Format(txtProd(2).Text, "General Number")) > Val(Format(txtProd(0).Text, "General Number")) Then
                MsgBox "La cantidad no puede ser mayor o igual al precio del producto. Verifique.", vbInformation
            Else
                txtProd(3).Text = (Val(resProdDevo.Fields("prod_precio"))) - (Val(Format(txtProd(2).Text, "General Number")))
                txtProd(3).Text = FormatCurrency(txtProd(3).Text)
                txtProd(1).Text = ((Val(Format(txtProd(2).Text, "General Number")) * 100) / Val(Format(txtProd(0).Text, "General Number")))
                txtProd(1).Text = Round((txtProd(1).Text), 2)
                txtProd(2).Text = FormatCurrency(txtProd(2).Text)
            End If
        End If
    End If
End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 1 Then
            valorDescuento ("Porcentaje")
        Else
            If Index = 2 Then
                valorDescuento ("Cantidad")
            End If
        End If
    End If
End Sub
