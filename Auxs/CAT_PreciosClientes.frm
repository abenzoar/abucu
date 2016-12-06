VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CAT_PreciosClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación precios - clientes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   15975
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   688
      TabCaption(0)   =   "Lista clientes - productos"
      TabPicture(0)   =   "CAT_PreciosClientes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListaUsers"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle general"
      TabPicture(1)   =   "CAT_PreciosClientes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Borde(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line1(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblDatos(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "imgFoto(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtClave(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtClave 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         TabIndex        =   1
         Top             =   1980
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid ListaUsers 
         Height          =   7935
         Left            =   -75000
         TabIndex        =   5
         Top             =   960
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   21
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"CAT_PreciosClientes.frx":0038
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
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   2
         Left            =   240
         Stretch         =   -1  'True
         Top             =   900
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
         Height          =   735
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Top             =   900
         Width           =   1815
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
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   3
         X1              =   240
         X2              =   3240
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Código   F4"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H00004080&
         BorderWidth     =   4
         Height          =   435
         Index           =   7
         Left            =   1680
         Top             =   1980
         Width           =   1725
      End
   End
End
Attribute VB_Name = "CAT_PreciosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
