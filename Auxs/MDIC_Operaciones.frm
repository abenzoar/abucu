VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MDIC_Operaciones 
   BackColor       =   &H80000004&
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17265
   Icon            =   "MDIC_Operaciones.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   17265
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtImpuesto 
      BackColor       =   &H00E9E9E9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   8040
      Width           =   3975
   End
   Begin VB.ComboBox cmbMesa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Timer Time_listaRapida 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   14520
      Top             =   0
   End
   Begin VB.ListBox lista_rapida 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   1560
      TabIndex        =   61
      Top             =   -5000
      Width           =   8175
   End
   Begin VB.ComboBox cmbEstado 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   15360
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   480
      Width           =   2895
   End
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtObservacion 
      BackColor       =   &H00E9E9E9&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Timer TTime 
      Interval        =   250
      Left            =   14040
      Top             =   9480
   End
   Begin VB.TextBox textDesc 
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
      Left            =   12480
      TabIndex        =   26
      Top             =   -500
      Width           =   615
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
      Left            =   12720
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   10560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdOperCheck 
      Caption         =   "Command1"
      Height          =   255
      Index           =   3
      Left            =   15000
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   -500
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCant 
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
      Left            =   11880
      TabIndex        =   23
      Top             =   -500
      Width           =   615
   End
   Begin VB.CommandButton cmdOperCheck 
      Caption         =   "Command1"
      Height          =   255
      Index           =   2
      Left            =   14400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   -500
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOperCheck 
      Caption         =   "Command1"
      Height          =   255
      Index           =   1
      Left            =   13920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   -500
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdOperCheck 
      Caption         =   "Command1"
      Height          =   255
      Index           =   0
      Left            =   13320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   -500
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtTotal 
      BackColor       =   &H00E9E9E9&
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   8760
      Width           =   3975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1095
      Left            =   1320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483627
      TabCaption(0)   =   "Cantidad moneda"
      TabPicture(0)   =   "MDIC_Operaciones.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label10(0)"
      Tab(0).Control(1)=   "Borde(3)"
      Tab(0).Control(2)=   "txtDesc(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cantidad porcentaje"
      TabPicture(1)   =   "MDIC_Operaciones.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Borde(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtDesc(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00E9E9E9&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Index           =   1
         Left            =   480
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00E9E9E9&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   -74520
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H8000000D&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   480
         Top             =   480
         Width           =   3285
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H8000000D&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   -74520
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   255
      End
   End
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
      Left            =   8160
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   3615
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   20
      FixedCols       =   0
      BackColorFixed  =   9520683
      ForeColorFixed  =   16777215
      BackColorBkg    =   15329769
      GridColor       =   16711680
      AllowUserResizing=   1
      FormatString    =   $"MDIC_Operaciones.frx":05C2
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
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtSub 
      BackColor       =   &H00E9E9E9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   6000
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A."
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
      Left            =   240
      TabIndex        =   68
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   480
      Index           =   8
      Left            =   1320
      Top             =   8025
      Width           =   4005
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
      Index           =   10
      Left            =   240
      TabIndex        =   66
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   15
      X1              =   12240
      X2              =   14040
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa"
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
      Left            =   12240
      TabIndex        =   65
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblDatos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   12240
      TabIndex        =   64
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   14
      X1              =   12240
      X2              =   15120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa"
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
      Left            =   12240
      TabIndex        =   63
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub estado"
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
      Left            =   15360
      TabIndex        =   60
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   12
      X1              =   15360
      X2              =   18240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Left            =   10080
      TabIndex        =   58
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Servicios"
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
      Left            =   7560
      TabIndex        =   57
      Top             =   6360
      Width           =   855
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
      Index           =   10
      Left            =   7560
      TabIndex        =   56
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
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
      Left            =   6480
      TabIndex        =   55
      Top             =   6360
      Width           =   855
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
      Index           =   9
      Left            =   6480
      TabIndex        =   54
      Top             =   6600
      Width           =   735
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
      Index           =   24
      Left            =   5640
      TabIndex        =   53
      Top             =   6360
      Width           =   855
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
      Index           =   8
      Left            =   5640
      TabIndex        =   52
      Top             =   6600
      Width           =   735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      X1              =   11160
      X2              =   11160
      Y1              =   7200
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      X1              =   10080
      X2              =   16440
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizados"
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
      Index           =   23
      Left            =   12960
      TabIndex        =   51
      Top             =   7560
      Width           =   855
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
      Index           =   7
      Left            =   12960
      TabIndex        =   50
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
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
      Index           =   22
      Left            =   11280
      TabIndex        =   49
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones del producto seleccionado"
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
      Left            =   5640
      TabIndex        =   48
      Top             =   7080
      Width           =   3975
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
      Left            =   11280
      TabIndex        =   47
      Top             =   7800
      Width           =   1095
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
      Index           =   20
      Left            =   11280
      TabIndex        =   46
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lblDatos 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
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
      Index           =   5
      Left            =   10080
      TabIndex        =   45
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Membresia"
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
      Left            =   10080
      TabIndex        =   44
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lblDatos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno"
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
      Index           =   4
      Left            =   10080
      TabIndex        =   43
      Top             =   6240
      Width           =   6375
   End
   Begin VB.Label Label1 
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
      Index           =   18
      Left            =   10080
      TabIndex        =   42
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   435
      Index           =   7
      Left            =   8160
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   435
      Index           =   6
      Left            =   4800
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   435
      Index           =   5
      Left            =   1560
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   675
      Index           =   1
      Left            =   1320
      Top             =   8760
      Width           =   4005
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   1875
      Index           =   2
      Left            =   5640
      Top             =   7440
      Width           =   4005
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      Height          =   555
      Index           =   0
      Left            =   1320
      Top             =   6000
      Width           =   4005
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
      Height          =   255
      Index           =   3
      Left            =   10080
      TabIndex        =   41
      Top             =   8640
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto/Servicio"
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
      Left            =   5640
      TabIndex        =   39
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   5640
      X2              =   16440
      Y1              =   5880
      Y2              =   5880
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
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
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
      Left            =   240
      TabIndex        =   35
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   9
      X1              =   240
      X2              =   5280
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   8
      X1              =   10200
      X2              =   12000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado operación"
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
      Left            =   10200
      TabIndex        =   34
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   7
      X1              =   10200
      X2              =   12000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
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
      Index           =   7
      Left            =   10200
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   6
      X1              =   8160
      X2              =   9840
      Y1              =   1440
      Y2              =   1440
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
      Left            =   8160
      TabIndex        =   32
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   5
      X1              =   4800
      X2              =   6480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave/Código   F3"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   4
      X1              =   1560
      X2              =   3240
      Y1              =   1440
      Y2              =   1440
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
      Index           =   4
      Left            =   1560
      TabIndex        =   30
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   3
      X1              =   6720
      X2              =   9720
      Y1              =   360
      Y2              =   360
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
      Left            =   6720
      TabIndex        =   29
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   2
      X1              =   3480
      X2              =   6480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario seleccionado"
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
      Left            =   3480
      TabIndex        =   28
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   240
      X2              =   3240
      Y1              =   360
      Y2              =   360
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
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   14280
      TabIndex        =   22
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   14280
      TabIndex        =   21
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   14280
      TabIndex        =   20
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   14160
      TabIndex        =   19
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   14160
      TabIndex        =   18
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   14160
      TabIndex        =   17
      Top             =   2000
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   10200
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   10320
      TabIndex        =   10
      Top             =   480
      Width           =   1695
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
      Left            =   8160
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   480
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
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   480
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
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   11655
      Index           =   1
      Left            =   0
      Picture         =   "MDIC_Operaciones.frx":0735
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19095
   End
End
Attribute VB_Name = "MDIC_Operaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim folio As Long
Dim userId As Long
Dim descGral As Boolean
Dim monedero As Boolean
'Dim vendetId As Long
Private Sub cmbEstado_Click()
    updateVenta (Val(lInfo(1).Caption))
End Sub

Private Sub cmbMesa_Click()
    updateVenta (Val(lInfo(1).Caption))
    lblDatos(11).Caption = cmbMesa.ItemData(cmbMesa.ListIndex)
End Sub
Private Sub cmbUser_Click()
    lista.TextMatrix(lista.Row, lista.Col) = cmbUser.Text
    cmbUser.Visible = False
    
    sql1 = "SELECT T4.PERTP_PER_ID, T4.PERTP_TIPO_ID, " & _
    "CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) USUARIO " & _
    "FROM PERSONA T2, PER_tIPO T4 " & _
    "WHERE T2.PER_ID = T4.PERTP_PER_ID AND T4.PERTP_STATUS = 'A' AND T4.PERTP_PER_TIPO = 'U' " & _
    "AND concat(T4.PERTP_PER_ID, T4.PERTP_TIPO_ID) = '" & cmbUser.ItemData(cmbUser.ListIndex) & "'"
    Set RES1 = con.Execute(sql1)
        
    If Not RES1.EOF Then
        lista.TextMatrix(lista.Row, 9) = RES1.Fields("PERTP_TIPO_ID")
        lista.TextMatrix(lista.Row, 10) = RES1.Fields("PERTP_PER_ID")
        updateVentDet (lista.Row)
    End If
End Sub

Private Sub cmbUser_LostFocus()
    cmbUser.Visible = False
End Sub

Public Sub cmdOperCheck_Click(Index As Integer)
    Select Case Index
        Case 2: checkCliente
        Case 0:
            If tipoBusqueda = "P" Then
                checkProducto
            Else
                If tipoBusqueda = "S" Then
                    checkServicio
                End If
            End If
        Case 1: checkUsuario
        'Case 3: checkPrecio
    End Select

End Sub

Private Sub Form_GotFocus()
    Set FrmFocus = Me
End Sub

Private Sub Form_Load()
    
    lista_rapida.Visible = False
    'checkDatos
    Set FrmFocus = Me
    numFrmOper = numFrmOper + 1
    lista.ColWidth(6) = 0
    lista.ColWidth(7) = 0
    'lista.ColWidth(17) = 0
    lista.ColWidth(18) = 0
    lblDatos(3).Caption = ""
    
    SSTab1.Tab = 1
    
    lista.ColWidth(9) = 0
'    lista.ColWidth(10) = 0
'    lista.ColWidth(13) = 0
'    lInfo(0).Caption = "0"
    
    lista.Rows = 1
    txtCant.Visible = False
    textDesc.Visible = False
    descGral = False
    txtObservacion.Locked = False
        
    cmbEstado.Clear
    cmbEstado.AddItem "SIN ATENDER"
    cmbEstado.AddItem "ATENDIDO"
    cmbEstado.AddItem "RECIBIDO"
    cmbEstado.AddItem "NINGUNO"
        
    If tikcet = False Then
        txtClave(1).Text = FRM_Menu.menuBarra2.Panels(7).Text
        txtClave(2).Text = "CLTE"
        lblDatos(11).Caption = ""
        checkUsuario
        checkCliente
        crearFolio
    Else
        If tikcet = True Then
            tikcet = False
            lInfo(1).Caption = folioTicket
            cargaTicket
        End If
    End If

    If FRM_Menu.menuBarra2.Panels(14).Text = "A" Then
        cmbEstado.Visible = True
        Label1(12).Visible = True
        Line1(12).Visible = True
    Else
        cmbEstado.Visible = False
        Label1(12).Visible = False
        Line1(12).Visible = False
    End If
      
    If mesas = True Then
        cmbMesa.Visible = True
        Label1(14).Visible = True
        Line1(14).Visible = True
        Line1(15).Visible = True
        Label1(17).Visible = True
        lblDatos(11).Visible = True
        carga_Mesa
        If lblDatos(11).Caption <> "" Then
            cmbMesa.Visible = False
            Label1(14).Visible = False
            Line1(14).Visible = False
            Line1(15).Visible = False
        End If
    Else
        cmbMesa.Visible = False
        Label1(14).Visible = False
        Line1(14).Visible = False
        Line1(15).Visible = False
        Label1(17).Visible = False
        lblDatos(11).Visible = False
    End If
        
End Sub
Private Sub carga_Mesa()
    sql1 = "SELECT * FROM VIEW_MESAS_ESTADO WHERE ESTADO = 'DISPONIBLE' ORDER BY MESA_ID"
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        cmbMesa.AddItem RES1.Fields("MESA_NOMBRE")
        cmbMesa.ItemData(cmbMesa.ListCount - 1) = RES1.Fields("mesa_id")
        RES1.MoveNext
    Loop
    
End Sub
Private Sub cargaTicket()
   On Error Resume Next
    sql1 = "SELECT * fROM VIEW_VENTAS WHERE FOLIO = '" & folioTicket & "'"
    Set RES1 = con.Execute(sql1)
    lista.Rows = 1
    If Not RES1.EOF Then
        
        lblDatos(2).Caption = RES1.Fields("CLIENTE")
        lblDatos(4).Caption = RES1.Fields("CLIENTE")
        lblClieId(0).Caption = RES1.Fields("CLIE_PERID")
        lblClieId(1).Caption = RES1.Fields("CLIE_TIPOID")
        lblClieId(2).Caption = RES1.Fields("CLIE_TIPO")
        lblDatos(3).Caption = RES1.Fields("CLIE_EMAIL")
        lblDatos(11).Caption = RES1.Fields("MESA") & ""
        If RES1.Fields("puntos_Mem") = "S" Then
            lblDatos(5).Caption = "SI"
        Else
            lblDatos(5).Caption = "NO"
        End If
        lblDatos(6).Caption = FormatCurrency(RES1.Fields("PUNTOS_TOT"))
        lblDatos(7).Caption = FormatCurrency(RES1.Fields("PUNTOS_USA"))
        txtObservacion.Text = RES1.Fields("OBSERVACIONES") & ""
'        If IsNull(RES1.Fields("MONEDERO")) = True Then
'            lblDatos(6).Caption = "$0.00"
'        Else
'            lblDatos(6).Caption = FormatCurrency(RES1.Fields("MONEDERO"))
'        End If
        
        lblDatos(1).Caption = RES1.Fields("USUARIO")
        lblUserId(0).Caption = RES1.Fields("USU_PERID")
        lblUserId(1).Caption = RES1.Fields("USU_TIPOID")
        lblUserId(2).Caption = RES1.Fields("USU_TIPO")
        
        txtSub.Text = FormatCurrency(RES1.Fields("SUBTOTAL1"))
        txtDesc(0).Text = FormatCurrency(RES1.Fields("DESCUENTO1"))

        txtTotal = FormatCurrency(RES1.Fields("TOTAL1"))
        If Val(RES1.Fields("DESCUENTO1")) > 0 Then
            txtDesc(1).Text = Round((Val(Format(txtDesc(0).Text, "General Number")) * 100) / Val(Format(txtSub.Text, "General Number")), 2)
        Else
            txtDesc(1).Text = "0"
        End If
        'descuentoPorcentaje
        Me.Caption = "Operación Ticket " & folioTicket & " Clte: " & lblDatos(2).Caption
        
        'MsgBox RES1.Fields("SUB_STATUS")
        cmbEstado.Text = RES1.Fields("SUB_STATUS")
        
        
    End If
        
    sql1 = "SELECT * fROM VIEW_VENTASDETALLE WHERE FOLIO = '" & folioTicket & "' AND STATUS = 'A' "
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("TIPO_PROD")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("CODIGO")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("producto")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("cantidad")
        lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("PRECIO"))
        lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("PROD_SER")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("PROD_ID")
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("USUARIO")
        lista.TextMatrix(lista.Rows - 1, 9) = RES1.Fields("VEND_ID")
        lista.TextMatrix(lista.Rows - 1, 10) = RES1.Fields("VEND_PERID")
        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(RES1.Fields("descuento"))
        lista.TextMatrix(lista.Rows - 1, 13) = RES1.Fields("DESCRIPCION") & ""
        lista.TextMatrix(lista.Rows - 1, 17) = RES1.Fields("vendet_id") & ""
        lista.TextMatrix(lista.Rows - 1, 18) = RES1.Fields("dependiente_id") & ""
        lista.TextMatrix(lista.Rows - 1, 19) = RES1.Fields("dependiente") & ""
        checkPrecio (lista.Rows - 1)
        If Val(RES1.Fields("descuento")) > 0 Then
            textDesc.Visible = True
            valor = Val(Format(lista.TextMatrix(lista.Rows - 1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(lista.Rows - 1, 5), "General Number"))))
            lista.TextMatrix(lista.Rows - 1, 12) = Round(valor, 2)
            checkDescuentoInd
        Else
            lista.TextMatrix(lista.Rows - 1, 12) = Round(0, 2)
        End If
        lista.Row = lista.Rows - 1
        lista.Col = 16
        lista.CellFontName = "Wingdings"
        lista.CellFontBold = True
        lista.CellFontSize = 16
        If RES1.Fields("Seguimiento") = "SI" Then
            lista.TextMatrix(lista.Rows - 1, 16) = Chr(254)
        Else
            lista.TextMatrix(lista.Rows - 1, 16) = Chr(168)
        End If
            
        textDesc.Visible = False

        RES1.MoveNext
    Loop
    
    
End Sub
Private Sub crearFolio()
'    On Error Resume Next
    sql1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDPERID, VENT_VENDTIPOID, VENT_VENDTIPO, " & _
    "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO) VALUES " & _
    "('" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', 'G', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', " & _
    "'" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "')"
    'MsgBox SQL1
    con.Execute (sql1)
    
    
    sql1 = "select last_insert_id() folioId"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        folio = RES1.Fields("folioId")
    End If

    lInfo(1).Caption = folio
    lInfo(2).Caption = "Abierto"
    
End Sub
Private Sub checkDatos()
'    lInfo.Caption = Format(Date, "Short date") & " " & Format(Time, "Short time")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim ques As String
    Dim formularios As Form

    
    
    If UCase(lInfo(2).Caption) = UCase("Abierto") Then
        ques = MsgBox("¿Salir?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            sql1 = "SELECT   SUC_PrintPreticket ticket, SUC_ticket_copia copias FROM SUCURSAL "
            Set RES1 = con.Execute(sql1)
            
            If RES1.Fields("ticket") = "A" Then
                ques = MsgBox("¿Imprimir pre-ticket?", vbYesNo + vbQuestion)
                If ques = vbYes Then
                    For b1 = 1 To Val(RES1.Fields("copias"))
                        notaPreTicket (Val(lInfo(1).Caption))
                        MsgBox "Impresión pre-ticket " & b1 & " de " & RES1.Fields("copias"), vbInformation
                    Next b1
                End If
            End If
            
        
            Cancel = 0
'            numFrmOper = numFrmOper - 1
        Else
            Cancel = 1
        End If
    Else
    End If
    
    If Val(MDI_Operaciones.StatusBar1.Panels(4).Text) = 0 Then
        Set FrmTickets = New MDIC_OperTickets
        FrmTickets.Show
    Else
        If Val(MDI_Operaciones.StatusBar1.Panels(4).Text) = 1 Then
            For Each formularios In Forms
                If formularios.Name = "MDIC_OperTickets" Then
                    formularios.Show
                    formularios.cargaTickets
                    Exit For
                End If
            Next
        End If
    End If
    
    
End Sub



Private Sub Lista_Click()
'    txtObservacion.Text = Lista.TextMatrix(Lista.Row, 13)
    

    
End Sub


Private Sub lista_DblClick()
    If UCase(lInfo(2).Caption) <> "ABIERTO" Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
Select Case lista.Col
    Case 3:
        ''''Para la cantidad
            If lista.TextMatrix(lista.Row, 6) <> "M" Then
                txtCant.Top = lista.CellTop + lista.Top
                txtCant.Left = lista.CellLeft + lista.Left
                txtCant.height = lista.CellHeight
                txtCant.width = lista.CellWidth
                txtCant.Text = lista.TextMatrix(lista.Row, lista.Col)
                txtCant.Visible = True
                txtCant.SelStart = 0
                txtCant.SelLength = Len(txtCant.Text)
                txtCant.SetFocus
            Else
                MsgBox "Para membresías solo aplica 1 vez por venta.", vbInformation
            End If
    Case 11:
        ''''Para el descuento moneda
            If descGral = False Then
                textDesc.Top = lista.CellTop + lista.Top
                textDesc.Left = lista.CellLeft + lista.Left
                textDesc.height = lista.CellHeight
                textDesc.width = lista.CellWidth
                textDesc.Text = lista.TextMatrix(lista.Row, lista.Col)
                textDesc.Visible = True
                textDesc.SelStart = 0
                textDesc.SelLength = Len(textDesc.Text)
                textDesc.SetFocus
            Else
                MsgBox "No se puede asignar un descuento individual si ya existe un descuento general. Verfiqiue.", vbInformation
            End If
    Case 12:
        ''''Para el descuento porcentaje
            If descGral = False Then
                textDesc.Top = lista.CellTop + lista.Top
                textDesc.Left = lista.CellLeft + lista.Left
                textDesc.height = lista.CellHeight
                textDesc.width = lista.CellWidth
                textDesc.Text = lista.TextMatrix(lista.Row, lista.Col)
                textDesc.Visible = True
                textDesc.SelStart = 0
                textDesc.SelLength = Len(textDesc.Text)
                textDesc.SetFocus
            Else
                MsgBox "No se puede asignar un descuento individual si ya existe un descuento general. Verfiqiue.", vbInformation
            End If
    Case 14:
        ''''Para el TOTAL a quedar
            If descGral = False Then
                textDesc.Top = lista.CellTop + lista.Top
                textDesc.Left = lista.CellLeft + lista.Left
                textDesc.height = lista.CellHeight
                textDesc.width = lista.CellWidth
                textDesc.Text = lista.TextMatrix(lista.Row, lista.Col)
                textDesc.Visible = True
                textDesc.SelStart = 0
                textDesc.SelLength = Len(textDesc.Text)
                textDesc.SetFocus
            Else
                MsgBox "No se puede asignar un descuento individual si ya existe un descuento general. Verfiqiue.", vbInformation
            End If
    Case 8:
            cargaUsuarios
            cmbUser.Top = lista.CellTop + lista.Top
            cmbUser.Left = lista.CellLeft + lista.Left
            'cmbUser.Height = lista.CellHeight
            cmbUser.width = lista.CellWidth
            cmbUser.Text = lista.TextMatrix(lista.Row, lista.Col)
            cmbUser.Visible = True
            cmbUser.SetFocus
    Case 16:
        Dim b1 As Long
        b1 = lista.Row
        
        lista.Row = b1
        lista.Col = 16
        lista.CellFontName = "Wingdings"
        lista.CellFontBold = True
        lista.CellFontSize = 16
        
        If lista.TextMatrix(b1, 16) = Chr(168) Then
            lista.TextMatrix(b1, 16) = Chr(254)
        Else
            lista.TextMatrix(b1, 16) = Chr(168)
        End If
        updateVentDet (lista.Row)
        


End Select

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

Private Sub lista_GotFocus()
    Set FrmFocus = Me
    ConScroll lista
    
End Sub

Private Sub lista_KeyPress(KeyAscii As Integer)
    ''''as
End Sub

Private Sub lista_LostFocus()
    SinScroll lista
    txtObservacion.Text = ""
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If UCase(lInfo(2).Caption) <> "ABIERTO" Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
    
    If lista.Rows > 1 Then
        If Button = vbRightButton Then
            If lista.Rows > 2 Then
                If Val(Format(lista.TextMatrix(lista.Row, 11), "General Number")) > 0 Then
                    MDI_Operaciones.mn_AddDesc_Other.Enabled = True
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Visible = True
                    MDI_Operaciones.mn_AddDesc_Other.Caption = "Agregar el valor de descuento de " & lista.TextMatrix(lista.Row, 11) & " a los demas registros"
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Enabled = True
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Caption = "Agregar el valor de descuento de " & lista.TextMatrix(lista.Row, 12) & "% a los demas registros"
                Else
                    MDI_Operaciones.mn_AddDesc_Other.Enabled = False
                    MDI_Operaciones.mn_AddDesc_Other.Caption = "Agregar el valor de descuento selecionado a los demas registros"
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Enabled = False
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Visible = False
                    MDI_Operaciones.mn_AddDesc_OtherProcen.Caption = "Agregar el valor de descuento selecionado a los demas registros"
                End If
            End If
            MDI_Operaciones.mn_CancelAll.Visible = False
            PopupMenu MDI_Operaciones.mn_menu, vbPopupMenuLeftAlign
        
        End If
    End If


End Sub

Private Sub mn_Menu_Click()
    Dim question As String
    
    question = MsgBox("Cancelar: " & lista.TextMatrix(lista.Row, 0) & "  " & lista.TextMatrix(lista.Row, 2) & "¿Continuar?", vbYesNo)
    If question = vbYes Then
        MsgBox "Cancela"
    End If
End Sub

Private Sub cargaFrom_ListaRapida()
   On Error Resume Next
    
    If tipoBusqueda = "P" Then
        If Left(lista_rapida.Text, 1) = "*" Then
            txtClave(0).Text = lista_rapida.ItemData(lista_rapida.ListIndex)
            lista_rapida.Visible = False
            checkMembresia
        Else
            sql1 = "select prod_codigo from productos where prod_id = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
            Set RES1 = con.Execute(sql1)
            
            txtClave(0).Text = RES1.Fields("PROD_CODIGO")
            lista_rapida.Visible = False
            checkProducto
        End If
        txtClave(0).SetFocus
    Else
        If tipoBusqueda = "U" Then
            'txtClave(1).Text = a
            sql1 = "select PERTP_CODIGO_MEMBRESIA from PER_TIPO where PERTP_PER_ID = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
            Set RES1 = con.Execute(sql1)
            
            txtClave(1).Text = RES1.Fields("PERTP_CODIGO_MEMBRESIA")
            lista_rapida.Visible = False
            checkUsuario
            txtClave(1).SetFocus
        Else
            If tipoBusqueda = "C" Then
                'txtClave(2).Text = a
                sql1 = "select PERTP_CODIGO_MEMBRESIA from PER_TIPO where PERTP_PER_ID = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
                Set RES1 = con.Execute(sql1)
                
                txtClave(2).Text = RES1.Fields("PERTP_CODIGO_MEMBRESIA")
                lista_rapida.Visible = False
                checkCliente
                txtClave(2).SetFocus
            
            End If
        End If
    End If
    
    
    tipoBusqueda = ""

End Sub

Private Sub lista_rapida_DblClick()
    cargaFrom_ListaRapida
End Sub

Private Sub lista_rapida_GotFocus()
    Time_listaRapida.Enabled = False
End Sub

Private Sub lista_rapida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaFrom_ListaRapida
    End If
End Sub

Private Sub lista_rapida_LostFocus()
    lista_rapida.Visible = False
End Sub

Private Sub lista_SelChange()
    Lista_Click
End Sub



Public Sub textDesc_KeyPress(KeyAscii As Integer)
    Call NumerosPunto(KeyAscii)
    Dim valor As Double
    If KeyAscii = 27 Then
        textDesc.Visible = False
        textDesc.Text = ""
    Else
        If KeyAscii = 13 Then
            If lista.Col = 11 Then
                If Val(textDesc.Text) <= Val(Format(lista.TextMatrix(lista.Row, 5), "General Number")) And Val(textDesc.Text) > 0 Then
                    lista.TextMatrix(lista.Row, 11) = FormatCurrency(textDesc.Text)
                    valor = Val(Format(lista.TextMatrix(lista.Row, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(lista.Row, 5), "General Number"))))
                    lista.TextMatrix(lista.Row, 12) = Round(valor, 2)
                    updateVentDet (lista.Row)
                    checkPrecio (lista.Row)
                    checkDescuentoInd
                    lista.Col = 13
                        If Val(textDesc.Text) > 0 Then
                            lista.CellBackColor = &H40C0&
                            lista.CellForeColor = vbWhite
                        Else
                            lista.CellBackColor = vbWhite
                            lista.CellForeColor = vbWhite
                        End If
                    textDesc.Text = ""
                    textDesc.Visible = False
                    txtDesc(0).Locked = True
                    txtDesc(1).Locked = True
                    Exit Sub
                Else
                    MsgBox "El descuento no puede ser mayor al total. Verifique.", vbInformation
                    textDesc.SelStart = 0
                    textDesc.SelLength = Len(textDesc.Text)
                    textDesc.SetFocus
                    Exit Sub
                End If
            Else
                If lista.Col = 12 Then
                    If Val(textDesc.Text) <= 100 Then
                        lista.TextMatrix(lista.Row, 12) = Round(Val(textDesc.Text), 2)
                        valor = Val(Format(lista.TextMatrix(lista.Row, 5), "General Number")) * (Val(lista.TextMatrix(lista.Row, 12)) / 100)
                        lista.TextMatrix(lista.Row, 11) = FormatCurrency(valor)
                        updateVentDet (lista.Row)
                        checkPrecio (lista.Row)
                        checkDescuentoInd
                        lista.Col = 12
                        If Val(textDesc.Text) > 0 Then
                            lista.CellBackColor = &H40C0&
                            lista.CellForeColor = vbWhite
                        Else
                            lista.CellBackColor = vbWhite
                            lista.CellForeColor = vbWhite
                        End If
                        textDesc.Text = ""
                        textDesc.Visible = False
                        txtDesc(0).Locked = True
                        txtDesc(1).Locked = True
                        Exit Sub
                    Else
                        MsgBox "El descuento no puede ser mayor al total. Verifique.", vbInformation
                        textDesc.SelStart = 0
                        textDesc.SelLength = Len(textDesc.Text)
                        textDesc.SetFocus
                        Exit Sub
                    End If
                Else
                '''''Cuando se coloca el monto total (Para calcular descuento)
                    If lista.Col = 14 Then
                        If lista.TextMatrix(lista.Row, 1) = "MND" Then
                                 textDesc.Text = (Val(textDesc.Text) * (-1))
                            If Val(textDesc.Text) >= Val(Format(lista.TextMatrix(lista.Row, 5), "General Number")) And Val(textDesc.Text) < 0 Then
                                lista.TextMatrix(lista.Row, 4) = FormatCurrency(textDesc.Text)
                                lista.TextMatrix(lista.Row, 5) = FormatCurrency(textDesc.Text)
                                lista.TextMatrix(lista.Row, 14) = FormatCurrency(textDesc.Text)
                                updateVentDet (lista.Row)
                                checkPrecio (lista.Row)
                                updateVenta (Val(lInfo(1).Caption))
                                textDesc.Text = ""
                                textDesc.Visible = False
                                Exit Sub
                            End If
                        Else
                            If Val(textDesc.Text) <= Val(Format(lista.TextMatrix(lista.Row, 5), "General Number")) And Val(textDesc.Text) >= 0 Then
                                lista.TextMatrix(lista.Row, 11) = (Val(Format(lista.TextMatrix(lista.Row, 5), "General Number")) - FormatCurrency(textDesc.Text))
                                valor = Val(Format(lista.TextMatrix(lista.Row, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(lista.Row, 5), "General Number"))))
                                lista.TextMatrix(lista.Row, 12) = Round(valor, 2)
                                updateVentDet (lista.Row)
                                checkDescuentoInd
                                checkPrecio (lista.Row)
                                textDesc.Text = ""
                                textDesc.Visible = False
                                txtDesc(0).Locked = True
                                txtDesc(1).Locked = True
                                Exit Sub
                            Else
                                MsgBox "El cantidad total no puede ser mayor al subtotal ni menor a cero. Verifique.", vbInformation
                                textDesc.SelStart = 0
                                textDesc.SelLength = Len(textDesc.Text)
                                textDesc.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Sub
Private Sub checkDescuentoInd()
    Dim desc As Double
    desc = 0
    For b1 = 1 To lista.Rows - 1
        desc = desc + Val(Format(lista.TextMatrix(b1, 11), "General Number"))
    Next b1
    
    txtDesc(0).Text = desc
    Call txtDesc_KeyPress(0, 13)
    
    
End Sub
Private Sub textDesc_LostFocus()
    textDesc.Visible = False
    textDesc.Text = ""
    
End Sub


Private Sub Time_listaRapida_Timer()
Time_listaRapida.Enabled = False
lista_rapida.Visible = False
End Sub

Private Sub TTime_Timer()
TTime.Enabled = False

lista.width = Me.width - 500
Image2(1).width = Me.width
Image2(1).height = Me.height
Me.Caption = "Operación Ticket " & lInfo(1).Caption & " Clte: " & lblDatos(2).Caption
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    NumerosPunto (txtCant.Text)
    If KeyAscii = 13 Then
        If Val(txtCant.Text) <> Val(lista.TextMatrix(lista.Row, 3)) And Val(txtCant.Text) > 0 Then
            If lista.TextMatrix(lista.Row, 6) = "P" Then
                sql1 = "SELECT PROD_CANT, PROD_INVENTARIO, prod_dependiente FROM PRODUCTOS WHERE PROD_CODIGO = '" & lista.TextMatrix(lista.Row, 1) & "'"
                Set RES1 = con.Execute(sql1)
                If Not RES1.EOF Then
                    If RES1.Fields("prod_inventario") = "N" Then
                        lista.TextMatrix(lista.Row, 3) = txtCant.Text
                        updateVentDet (lista.Row)
                        checkPrecio (lista.Row)
                        txtCant.Text = ""
                        txtCant.Visible = False
                        tPro = Val(tPro) + lista.TextMatrix(lista.Row, 3)
                        Exit Sub
                    Else
                        If RES1.Fields("PROD_CANT") < Val(txtCant.Text) Then
                            MsgBox "Quedan " & RES1.Fields("PROD_CANT") & " en existencia, Verifique.", vbInformation
                            Exit Sub
                        End If
                        
                        lista.TextMatrix(lista.Row, 3) = txtCant.Text
                        updateVentDet (lista.Row)
                        checkPrecio (lista.Row)
                        txtCant.Text = ""
                        txtCant.Visible = False
                        tPro = Val(tPro) + lista.TextMatrix(lista.Row, 3)
                        Exit Sub
                        
                    End If
'                    Else
'                        MsgBox "La cantidad supera los productos en existencia. Verifique.", vbInformation
'                        txtCant.SelStart = 0
'                        txtCant.SelLength = Len(txtCant.Text)
'                        txtCant.SetFocus
'                        Exit Sub
'                    End If
                End If
            Else
                If lista.TextMatrix(lista.Row, 6) = "S" Then
                    lista.TextMatrix(lista.Row, 3) = txtCant.Text
                    updateVentDet (lista.Row)
                    checkPrecio (lista.Row)
                    txtCant.Text = ""
                    txtCant.Visible = False
                    tSer = Val(tSer) + lista.TextMatrix(lista.Row, 3)
                    Exit Sub
                Else
                    MsgBox "Operación no permitida. Verifique.", vbInformation
                End If
            End If
        Else
            txtCant.Text = ""
            txtCant.Visible = False
            Exit Sub
        End If
    Else
        If KeyAscii = 27 Then
            txtCant.Text = ""
            txtCant.Visible = False
        End If
    End If
End Sub

Private Sub txtCant_LostFocus()
    txtCant.Text = ""
    txtCant.Visible = False
    
End Sub

Private Sub txtClave_Change(Index As Integer)
'''Colocar para que al momento de teclear la clave aparezcan 3 opciones cercanas
If Len(txtClave(Index).Text) > 0 Then
    lista_rapida.Top = txtClave(Index).Top + 375
    lista_rapida.Left = txtClave(Index).Left
    lista_rapida.Visible = True
    Time_listaRapida.Enabled = False
    Time_listaRapida.Enabled = True
    
    cargaLista_General (Index)
Else
    If Len(txtClave(0).Text) <= 0 Then
        lista_rapida.Visible = False
    End If
End If

End Sub
Private Sub cargaLista_General(Index As Integer)
    On Error Resume Next
    Dim textoLista As String
    Dim idTexto As String
    
    If Len(txtClave(Index).Text) > 0 Then
        If Index = 0 Then
            tipoBusqueda = "P"

            sql1 = "SELECT concat(PROD_CODIGO, '          ', PROD_NOMBRE, '         ', '$', ROUND(PROD_PRECIO, 2)) PRODUCTOS, PROD_ID FROM PRODUCTOS " & _
            "WHERE PROD_STATUS = 'A' AND UPPER(concat(PROD_CODIGO, ' ', PROD_NOMBRE)) LIKE UPPER('%" & txtClave(0).Text & "%') LIMIT 8 " & _
            " Union All " & _
            "SELECT CONCAT('* ', CTMB_NOMBRE, '          DIAS: ', CTMB_DIAS, '          ', CTMB_PRECIO) PRODUCTOS, CTMB_ID PROD_ID FROM CAT_MEMBRESIAS " & _
            "WHERE CTMB_STATUS = 'A'  AND UPPER(CONCAT(CTMB_NOMBRE)) LIKE UPPER('%" & txtClave(0).Text & "%') LIMIT 8"
            
            textoLista = "Productos"
            idTexto = "Prod_id"
'            MsgBox txtClave(0).Text
'            MsgBox SQL1
            Set RES1 = con.Execute(sql1)
        Else
            If Index = 1 Then
                tipoBusqueda = "U"
                sql1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA,  " & _
                "CONCAT(T2.PER_NOMBRE, '  ', T2.PER_PATERNO, '  ', T2.PER_MATERNO) USUARIO, T2.PER_ID " & _
                "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
                "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
                "AND upper(concat(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO)) LIKE UPPER('%" & txtClave(1).Text & "%') " & _
                "AND T4.PERTP_PER_TIPO = 'U' AND T4.PERTP_STATUS = 'A'" & _
                "ORDER BY T2.PER_NOMBRE ASC"
                textoLista = "Usuario"
                idTexto = "Per_Id"
                'MsgBox SQL1
                Set RES1 = con.Execute(sql1)
            Else
                If Index = 2 Then
                    tipoBusqueda = "C"
                    sql1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA,  " & _
                    "CONCAT(T2.PER_NOMBRE, '  ', T2.PER_PATERNO, '  ', T2.PER_MATERNO) CLIENTE, T2.PER_ID " & _
                    "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
                    "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
                    "AND upper(concat(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO)) LIKE UPPER('%" & txtClave(2).Text & "%') " & _
                    "AND T4.PERTP_PER_TIPO = 'C' AND T4.PERTP_STATUS = 'A'" & _
                    "ORDER BY T2.PER_NOMBRE ASC"
                    textoLista = "Cliente"
                    idTexto = "Per_Id"
                    'MsgBox SQL1
                    Set RES1 = con.Execute(sql1)
                End If
            End If
        End If
        
                                           
        lista_rapida.Clear
        Do While Not RES1.EOF
            lista_rapida.AddItem RES1.Fields(textoLista)
            lista_rapida.ItemData(lista_rapida.NewIndex) = RES1.Fields(idTexto)
            RES1.MoveNext
        Loop
    End If
End Sub
Private Sub txtClave_Click(Index As Integer)
        txtClave(Index).SelStart = 0
        txtClave(Index).SelLength = Len(txtClave(Index).Text)

End Sub

Private Sub txtClave_GotFocus(Index As Integer)
    
     If Index = 0 Then
        TT1.Title = "Código/Clave de Producto o Servicio"
        TT1.TipText = "Escribe o escanea el código o clave del producto o servicio"
    Else
        If Index = 1 Then
            TT1.Title = "Código de Usuario"
            TT1.TipText = "Escribe el código del usuario que registrará la operación"
        Else
            If Index = 2 Then
                TT1.Title = "Código/Membresia del cliente"
                TT1.TipText = "Escribe el código de membresia del cliente"
            End If
        End If
    End If
        TT1.Style = TTBalloon
        TT1.Icon = TTIconError
        TT1.ForeColor = vbWhite
        TT1.BackColor = &HCE7110
        TT1.PopupOnDemand = False
        TT1.CreateToolTip txtClave(Index).hWnd
        'TT1.Show 0, txtUsuario(Index).height / Screen.TwipsPerPixelX - 1
   
    
    Set FrmFocus = Me
        txtClave(Index).SelStart = 0
        txtClave(Index).SelLength = Len(txtClave(Index).Text)

        
        
        
        
End Sub

Private Sub txtClave_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        If lista_rapida.Visible = True Then
            lista_rapida.SetFocus
        End If
    End If
End Sub

Private Sub txtClave_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
     
        Select Case Index
           Case 1: checkUsuario
           Case 0:
            lista_rapida.Visible = False
            txtClave(0).Text = Replace(txtClave(0).Text, "'", "-")
            If Left(txtClave(0).Text, 1) = " " Then
                 txtClave(0).Text = Right(txtClave(0).Text, (Len(txtClave(0).Text) - 1))
            End If
           checkProducto
           Case 2: checkCliente
        End Select
        
        txtClave(Index).SelStart = 0
        txtClave(Index).SelLength = Len(txtClave(Index).Text)
          
     End If
End Sub
Private Sub checkMembresia()
    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If

    sql1 = "SELECT 'MEMBRESIA' TIPO_PROD, ID PROD_CODIGO, MEMBRESIA PROD_NOMBRE, PRECIO PROD_PRECIO, 'M' PROD_SERV, " & _
    "DIAS_MEMBRESIA, DIAS_PERIODO, PERIODO, TIPO, ID PROD_ID, DIAS_PERIODO PROD_DESCRIPCION, prod_AplicaDesc, prod_PrecioDesc   " & _
    "FROM VIEW_MEMBRESIAS WHERE " & _
    "ID = '" & txtClave(0).Text & "' AND STATUS = 'ACTIVO'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        addLista
    Else
        MsgBox "No se ha encontrado información con la clave proporcionada. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
    End If

End Sub

Private Sub checkServicio()
'   On Error Resume Next
    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
    
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_TIPO, " & _
    "PROD_FOTO, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, prod_AplicaDesc, prod_PrecioDesc " & _
    "FROM PRODUCTOS T1, CAT_TIPO T3 " & _
    "WHERE T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND PROD_CODIGO = '" & txtClave(0).Text & "' AND PROD_STATUS = 'A'"
    Set RES1 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES1.EOF Then
        lblDatos(0).Caption = RES1.Fields("PROD_NOMBRE")
        
        If IsNull(RES1.Fields("PROD_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            imgFoto(0).Picture = LoadPicture("")
        End If
    
    addLista
    Else
        checkMembresia
        'MsgBox "No se ha encontrado información con la clave proporcionada. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
    End If
    

End Sub
Private Sub checkProducto()

'    On Error Resume Next
        
    lista_rapida.Visible = False
     monedero = False
    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
    
    If Left(txtClave(0).Text, 2) = "MD" Then
        If Val(Format(lblDatos(6).Caption, "General Number")) > 0 Then
            monedero = True
            'addMonedero
            txtClave(0).Text = Right(txtClave(0).Text, (Len(txtClave(0).Text) - 2))
        Else
            MsgBox "No se puede asignar monedero a la cuenta del cliente seleccionado. Verifique.", vbInformation
            Exit Sub
        End If
    End If
    
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, PROD_PRECIODESC, PROD_APLICADESC, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, PROD_INVENTARIO, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, PROD_DEPENDIENTE, IF(PROD_DEPENDIENTE = 'D', 'DEPENDIENTE', 'UNICO') SUBTIPO " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "PROD_CODIGO = '" & txtClave(0).Text & "' AND PROD_STATUS = 'A'"
    Set RES1 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES1.EOF Then
        lblDatos(0).Caption = RES1.Fields("PROD_NOMBRE")
        If IsNull(RES1.Fields("PROD_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            imgFoto(0).Picture = LoadPicture("")
        End If
        
        addLista
        
        If monedero = True Then
            Call addMonedero(0, 0)
        End If
    Else
        
        checkServicio
    End If
    
    
End Sub
Public Sub addMonedero(mone As Double, moneus As Double)
'Dim Mone As Double, MoneUs As Double

Dim total As Double
    
    total = 0
    
    'EL MAXIMO TOTAL QUE TIENE EL USUARIO
    mone = Val(Format(lblDatos(6).Caption, "General Number"))
    'LO QUE SE ESTA DESCONTANDO
    If moneus = 0 Then
        moneus = Val(Format(lblDatos(7).Caption, "General Number"))
    Else
        '''
    End If
    If moneus >= mone Then
        Exit Sub
    Else
        For b1 = 1 To lista.Rows - 1
            If lista.TextMatrix(b1, 1) <> "MND" And lista.TextMatrix(b1, 15) = "MND" Then
                total = total + Val(Format(lista.TextMatrix(b1, 14), "General Number"))
                If Val(total) >= Val(mone) Then
                    total = mone
                    Exit For
                End If
            End If
        Next b1
                
        For b1 = 1 To lista.Rows - 1
            If lista.TextMatrix(b1, 1) = "MND" Then
                If Val(total) = 0 Then
                    deleteVentDet (b1)
                    If lista.Rows = 2 And b1 = 1 Then
                        lista.Rows = 1
                    Else
                        If lista.Rows > 2 Then
                            lista.RemoveItem (b1)
                        End If
                    End If
                Else
                    lista.TextMatrix(b1, 4) = FormatCurrency(total * (-1))
                    lista.TextMatrix(b1, 5) = FormatCurrency(total * (-1))
                    lista.TextMatrix(b1, 14) = FormatCurrency(total * (-1))
                    updateVentDet (b1)
                End If
                checkPrecioFinal
                Exit Sub
            End If
        Next b1
    
    
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = "MONEDERO"
        lista.TextMatrix(lista.Rows - 1, 1) = "MND"
        lista.TextMatrix(lista.Rows - 1, 2) = "DESCUENTO MONEDERO"
        lista.TextMatrix(lista.Rows - 1, 3) = "1"
        If moneus > 0 Then
            lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(moneus * -1)  'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
            lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(moneus * -1)  'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
            lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(moneus * -1)  'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
        Else
            lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(total * -1)  'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
            lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(total * -1) 'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
            lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(total * -1) 'FormatCurrency(Val(Format(Lista.TextMatrix(Lista.Rows - 2, 5), "General Number")) * (-1))
        End If
        lista.TextMatrix(lista.Rows - 1, 6) = "R"
        lista.TextMatrix(lista.Rows - 1, 7) = "1"
        lista.TextMatrix(lista.Rows - 1, 8) = lblDatos(1).Caption
        lista.TextMatrix(lista.Rows - 1, 9) = lblUserId(1).Caption
        lista.TextMatrix(lista.Rows - 1, 10) = lblUserId(0).Caption
        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(0)
        lista.TextMatrix(lista.Rows - 1, 12) = Round(0, 2)
        lista.TextMatrix(lista.Rows - 1, 13) = "DESCUENTO MONEDERO"
    
        checkPrecioFinal
        updateVenta (Val(lInfo(1).Caption))
        addVentDet
    End If

End Sub
Private Sub addLista()
    Dim valor As Double
    ''''Para productos
    vendetId = "0"
    If RES1.Fields("PROD_SERV") = "P" Then
        If RES1.Fields("prod_inventario") = "N" Then
            For b1 = 1 To lista.Rows - 1
                If lista.TextMatrix(b1, 1) = RES1.Fields("PROD_CODIGO") Then
                    If Val(RES1.Fields("PROD_CANT")) > Val(lista.TextMatrix(b1, 3)) Then
                        lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                        updateVentDet (b1)
                        checkPrecio (b1)
                        tPro = Val(tPro) + 1
                        Exit Sub
                    End If
                End If
            Next b1
        Else
            If RES1.Fields("PROD_CANT") <= 0 And RES1.Fields("PROD_DEPENDIENTE") = "U" Then
                MsgBox "Quedan " & RES1.Fields("PROD_CANT") & " en existencia, Verifique.", vbInformation
                Exit Sub
            End If
            
            For b1 = 1 To lista.Rows - 1
                If lista.TextMatrix(b1, 1) = RES1.Fields("PROD_CODIGO") Then
                    If Val(RES1.Fields("PROD_CANT")) > Val(lista.TextMatrix(b1, 3)) Then
                        lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                        updateVentDet (b1)
                        checkPrecio (b1)
                        tPro = Val(tPro) + 1
                        Exit Sub
                    Else
                        MsgBox "Quedan " & RES1.Fields("PROD_CANT") & " en existencia, Verifique.", vbInformation
                        Exit Sub
                    End If
                End If
            Next b1
        
        End If
    Else
        '''PARA SERVICIOS
        If RES1.Fields("PROD_SERV") = "S" Then
            For b1 = 1 To lista.Rows - 1
                If lista.TextMatrix(b1, 1) = RES1.Fields("PROD_CODIGO") Then
                        lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                        updateVentDet (b1)
                        checkPrecio (b1)
                        'tSer.Text = Val(ttSer.Text) + 1
                        Exit Sub
                End If
            Next b1
        Else
            '''PARA MEMBRESIAS
            If RES1.Fields("PROD_SERV") = "M" Then
                For b1 = 1 To lista.Rows - 1
                    If lista.TextMatrix(b1, 1) = RES1.Fields("PROD_CODIGO") Then
                            lista.TextMatrix(b1, 3) = lista.TextMatrix(b1, 3) + 1
                            updateVentDet (b1)
                            checkPrecio (b1)
                            'tSer.Text = Val(ttSer.Text) + 1
                            Exit Sub
                    End If
                Next b1
            End If
        End If
    End If
        
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("TIPO_PROD")
    lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("PROD_CODIGO")
    lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("PROD_NOMBRE")
    lista.TextMatrix(lista.Rows - 1, 3) = "1"
    lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
    lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("PROD_SERV")
    lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("PROD_ID")
    lista.TextMatrix(lista.Rows - 1, 8) = lblDatos(1).Caption
    lista.TextMatrix(lista.Rows - 1, 9) = lblUserId(1).Caption
    lista.TextMatrix(lista.Rows - 1, 10) = lblUserId(0).Caption
    If RES1.Fields("PROD_APLICADESC") = "S" Then
        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(Val(RES1.Fields("PROD_PRECIO")) - (Val(RES1.Fields("PROD_PRECIODESC"))))
        valor = Val(Format(lista.TextMatrix(lista.Rows - 1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(lista.Rows - 1, 4), "General Number"))))
        lista.TextMatrix(lista.Rows - 1, 12) = Round(valor, 2)
'        lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
        
    Else
        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(0)
        lista.TextMatrix(lista.Rows - 1, 12) = Round(0, 2)
        lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
    End If
    lista.TextMatrix(lista.Rows - 1, 13) = RES1.Fields("prod_DESCRIPCION")
    If RES1.Fields("PROD_SERV") = "P" Then
        lista.TextMatrix(lista.Rows - 1, 18) = RES1.Fields("PROD_DEPENDIENTE")
        lista.TextMatrix(lista.Rows - 1, 19) = RES1.Fields("SUBTIPO")
    Else
        lista.TextMatrix(lista.Rows - 1, 18) = "N"
        lista.TextMatrix(lista.Rows - 1, 19) = "N"
    End If
    
    
    
    mone = Val(Format(lblDatos(6).Caption, "General Number"))
    moneus = Val(Format(lblDatos(7).Caption, "General Number"))
    
    If monedero = True Then
        If mone - moneus > 0 Then
            lista.TextMatrix(lista.Rows - 1, 15) = "MND"
        Else
            lista.TextMatrix(lista.Rows - 1, 15) = ""
        End If
    End If
    
    lista.Row = lista.Rows - 1
    lista.Col = 16
    lista.CellFontName = "Wingdings"
    lista.CellFontBold = True
    lista.CellFontSize = 16
    lista.TextMatrix(lista.Rows - 1, 16) = Chr(168)
'    lista.TextMatrix(lista.Rows - 1, 14) = Chr(254)
       
    checkPrecio (lista.Rows - 1)
    checkDescuentoInd
    addVentDet
        
    lista.TextMatrix(lista.Rows - 1, 17) = vendetId
    
        
End Sub
Public Sub addVentDet()
    With lista
        
        Dim seg As String
        If lista.TextMatrix(.Rows - 1, 16) = Chr(254) Then
            seg = "S"
        Else
            If lista.TextMatrix(.Rows - 1, 16) = Chr(168) Then
                seg = "N"
            End If
        End If
    
        sql1 = "INSERT INTO VENTA_DETALLE (VENDET_FOLIO, VENDET_PRODUCTOID, VENDET_PRODSERV, VENDET_PRODUCTONOMBRE, " & _
        "VENDET_CANTIDAD, VENDET_PRECIO, VENDET_TIPO, VENDET_PRODCODIGO, VENDET_VENDPERID, VENDET_VENDTIPOID, VENDET_VENDTIPO, venDet_Descuento, vendet_descripcion, venDet_Seguimiento, vendet_ProdDepen, vendet_FechaHora) " & _
        "VALUES (" & _
        "'" & lInfo(1).Caption & "', '" & .TextMatrix(.Rows - 1, 7) & "', '" & .TextMatrix(.Rows - 1, 6) & "', " & _
        "'" & .TextMatrix(.Rows - 1, 2) & "', '" & .TextMatrix(.Rows - 1, 3) & "', " & _
        "'" & Val(Format(.TextMatrix(.Rows - 1, 4), "General Number")) & "', 'V', '" & .TextMatrix(.Rows - 1, 1) & "', " & _
        "'" & .TextMatrix(lista.Rows - 1, 10) & "', '" & .TextMatrix(lista.Rows - 1, 9) & "', 'U', '" & Val(Format(.TextMatrix(.Rows - 1, 11), "General Number")) & "', '" & .TextMatrix(.Rows - 1, 13) & "', '" & seg & "', '" & .TextMatrix(.Rows - 1, 18) & "', now())"
    '    MsgBox sql1
        con.Execute (sql1)
    End With
    
    
    sql1 = "select max(vendet_id) vendet_id from venta_Detalle"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        vendetId = RES1.Fields("vendet_id")
    End If
    
End Sub
Private Sub updateVentDet(fila As Long)


''''
        Dim seg As String
        If lista.TextMatrix(fila, 16) = Chr(254) Then
            seg = "S"
        Else
            If lista.TextMatrix(fila, 16) = Chr(168) Then
                seg = "N"
            End If
        End If
    
    
    sql1 = "UPDATE VENTA_DETALLE SET VENDET_CANTIDAD = '" & lista.TextMatrix(fila, 3) & "', " & _
    "VENDET_VENDPERID = '" & lista.TextMatrix(fila, 10) & "', VENDET_VENDTIPOID = '" & lista.TextMatrix(fila, 9) & "', " & _
    "venDet_Descuento = '" & Val(Format(lista.TextMatrix(fila, 11), "General Number")) & "', " & _
    "venDet_PRECIO = '" & Val(Format(lista.TextMatrix(fila, 4), "General Number")) & "', venDet_Seguimiento = '" & seg & "' " & _
    "WHERE VENDET_FOLIO = '" & lInfo(1).Caption & "' AND VENDET_PRODUCTOID = '" & lista.TextMatrix(fila, 7) & "' AND VENDET_PRODSERV = '" & lista.TextMatrix(fila, 6) & "' AND VENDET_ID = '" & lista.TextMatrix(fila, 17) & "'   " 'AND VENDET_VENDPERID = '" & lista.TextMatrix(fila, 10) & "'"
    'MsgBox SQL1
    con.Execute (sql1)
    
End Sub
Private Sub updateVenta(numFolio As Long)
    Dim MESA As String
    
    If cmbMesa.Text <> "" Then
        MESA = cmbMesa.ItemData(cmbMesa.ListIndex)
    Else
        MESA = ""
    End If
    
    If numFolio <> 0 Then
        If MESA = "" Then
            sql1 = "UPDATE VENTAS SET VENT_CLIEPERID = '" & lblClieId(0).Caption & "', vent_StatusOper = '" & Left(cmbEstado.Text, 1) & "', " & _
            "VENT_CLIETIPOID = '" & lblClieId(1).Caption & "', VENT_CLIETIPO = '" & lblClieId(2).Caption & "', " & _
            "VENT_VENDPERID='" & lblUserId(0).Caption & "', VENT_VENDTIPOID='" & lblUserId(1).Caption & "', VENT_VENDTIPO = 'U', vent_observaciones = '" & txtObservacion.Text & "', " & _
            "vent_membresia = '" & Left(lblDatos(5).Caption, 1) & "', vent_PuntosTot = '" & Val(Format(lblDatos(6).Caption, "General Number")) & "', vent_PuntosUsa = '" & Val(Format(lblDatos(7).Caption, "General Number")) & "' " & _
            "WHERE VENT_IDFOLIO = '" & numFolio & "'"
        Else
            sql1 = "UPDATE VENTAS SET VENT_CLIEPERID = '" & lblClieId(0).Caption & "', vent_StatusOper = '" & Left(cmbEstado.Text, 1) & "', " & _
            "VENT_CLIETIPOID = '" & lblClieId(1).Caption & "', VENT_CLIETIPO = '" & lblClieId(2).Caption & "', " & _
            "VENT_VENDPERID='" & lblUserId(0).Caption & "', VENT_VENDTIPOID='" & lblUserId(1).Caption & "', VENT_VENDTIPO = 'U', vent_observaciones = '" & txtObservacion.Text & "', VENT_MESA = '" & MESA & "', " & _
            "vent_membresia = '" & Left(lblDatos(5).Caption, 1) & "', vent_PuntosTot = '" & Val(Format(lblDatos(6).Caption, "General Number")) & "', vent_PuntosUsa = '" & Val(Format(lblDatos(7).Caption, "General Number")) & "' " & _
            "WHERE VENT_IDFOLIO = '" & numFolio & "'"
        End If
    
    con.Execute (sql1)
    End If

End Sub

Private Sub updateVentaTotales(numFolio As Long)
    If numFolio <> 0 Then
        sql1 = "UPDATE VENTAS SET VENT_SUBTOTAL = '" & Val(Format(txtSub.Text, "General Number")) & "', " & _
    "VENT_DESCUENTO = '" & Val(Format(txtDesc(0).Text, "General Number")) & "', " & _
    "VENT_TOTAL = '" & Val(Format(txtTotal.Text, "General Number")) & "' " & _
        "WHERE VENT_IDFOLIO = '" & numFolio & "'"
        con.Execute (sql1)
    End If


End Sub

Public Sub deleteVentDet(fila As Long)

    sql1 = "UPDATE VENTA_DETALLE " & _
    "SET vendet_Status = 'C', vendet_CancelMotivo = '" & FRM_Cancelar.txtMotivo.Text & "', vendet_FechaHoraCancel = now(), " & _
    "vendet_AutorizaPerId = '" & FRM_Cancelar.lblAutoriza(0).Caption & "', vendet_AutorizaTipoId = '" & FRM_Cancelar.lblAutoriza(1).Caption & "', vendet_AutorizaTipo = '" & FRM_Cancelar.lblAutoriza(2).Caption & "'  " & _
    "WHERE VENDET_FOLIO = '" & lInfo(1).Caption & "' AND VENDET_PRODUCTOID = '" & lista.TextMatrix(fila, 7) & "' " & _
    " AND VENDET_VENDPERID = '" & lista.TextMatrix(fila, 10) & "' AND VENDET_ID = '" & lista.TextMatrix(fila, 17) & "' "
    con.Execute (sql1)

    Call nota_Cocina(lInfo(1).Caption, "CANCEL")

    sql1 = "UPDATE PRODUCTOS SET PROD_CANT = PROD_CANT + " & lista.TextMatrix(fila, 3) & "        " & _
    "WHERE PROD_ID = '" & lista.TextMatrix(fila, 7) & "' "
    con.Execute (sql1)
    
End Sub



Public Sub deleteVentDetAll()
''''
'    tPro.Text = "0"
 '   tSer.Text = "0"
    
    sql1 = "DELETE FROM VENTA_DETALLE " & _
    "WHERE VENDET_FOLIO = '" & lInfo(1).Caption & "'"
    con.Execute (sql1)
    
    
End Sub

Public Sub checkPrecio(fila As Long)
    If lista.TextMatrix(fila, 11) = "" Then
        lista.TextMatrix(fila, 11) = FormatCurrency(0)
    End If
    lista.TextMatrix(fila, 5) = lista.TextMatrix(fila, 3) * lista.TextMatrix(fila, 4) '- Val(Format(lista.TextMatrix(fila, 11), "General Number"))
    lista.TextMatrix(fila, 14) = FormatCurrency(Val(lista.TextMatrix(fila, 5)) - Val(Format(lista.TextMatrix(fila, 11), "General Number")))
    
    lista.TextMatrix(fila, 5) = FormatCurrency(lista.TextMatrix(fila, 5))
    
    
    checkPrecioFinal
End Sub
Public Sub checkPrecioFinal()
    Dim total, DESCUENTO, monedero
    ''''Para los descuentos
    If lista.Rows = 1 Then
        descGral = False
        txtDesc(0).Text = "0"
        txtDesc(1).Text = "0"
        txtDesc(0).Locked = False
        txtDesc(1).Locked = False
    End If
    
    total = 0
    monedero = 0
    DESCUENTO = Val(Format(txtDesc(0).Text, "General Number"))
    'MsgBox DESCUENTO
    
    For b1 = 1 To lista.Rows - 1
        If UCase(lista.TextMatrix(b1, 0)) = UCase("Descuento") Then
            DESCUENTO = DESCUENTO + Val(Format(lista.TextMatrix(b1, 5), "General Number"))
        Else
            If UCase(lista.TextMatrix(b1, 0)) = UCase("Monedero") Then
                monedero = monedero + (Val(Format(lista.TextMatrix(b1, 5), "General Number")) * (-1))
                total = total + Val(Format(lista.TextMatrix(b1, 5), "General Number"))
            Else
                total = total + Val(Format(lista.TextMatrix(b1, 5), "General Number"))
            End If
        End If
    Next b1
    'MsgBox DESCUENTO
    txtSub.Text = FormatCurrency(total)
    txtDesc(0).Text = FormatCurrency(DESCUENTO)
    txtTotal = FormatCurrency(total - DESCUENTO)
    lblDatos(7).Caption = FormatCurrency(monedero)
    'tOper(0).Text = lista.Rows - 1
    
    updateVentaTotales (Val(lInfo(1).Caption))
    
End Sub

Private Sub checkUsuario()
On Error Resume Next

    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If

    sql1 = "SELECT PERTP_USUARIO, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_TIPO_ID, PERTP_PER_TIPO, CTPT_TIPO, PER_ID, PER_FOTO " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' " & _
    "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'U' " & _
    "AND T2.PERTP_CODIGO_MEMBRESIA = '" & txtClave(1).Text & "'"
    'MsgBox SQL1
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        userId = txtClave(1).Text
            
        lblDatos(1).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
        lblUserId(0).Caption = RES1.Fields("PER_ID")
        lblUserId(1).Caption = RES1.Fields("PERTP_TIPO_ID")
        lblUserId(2).Caption = RES1.Fields("PERTP_PER_TIPO")
        If IsNull(RES1.Fields("PER_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(1).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            imgFoto(1).Picture = LoadPicture("")
        End If
        updateVenta (Val(lInfo(1).Caption))
'        txtClave(0).SetFocus
    Else
        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation
    End If
    
End Sub
Private Sub checkCliente()
    On Error Resume Next
    
    lista_rapida.Visible = False
    
    If UCase(lInfo(2).Caption) <> "ABIERTO" Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
    
    sql1 = "SELECT PERTP_USUARIO, IF(PERTP_MEMBRESIA ='S', 'SI', 'NO') MEMBRESIA, PERTP_CODIGO_MEMBRESIA, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_PER_TIPO, PERTP_TIPO_ID, CTPT_TIPO, T1.PER_ID, PER_FOTO, PER_EMAIL, t2.TEMP_MONEDERO, (SELECT T4.TOTAL FROM VIEW_MONEDERO_CLIENTES T4 WHERE T1.PER_ID = T4.PER_ID) TOTAL " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'C' " & _
    "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'C' AND " & _
    "T2.PERTP_CODIGO_MEMBRESIA = '" & txtClave(2).Text & "'"
    'MsgBox SQL1
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
            
'       PARA CHECAR QUE SOLO UNA VENTA POR CLIENTA EXISTA
        sql1 = "SELECT COUNT(*) NUM  fROM VENTAS " & _
        "WHERE VENT_CLIEPERID = '" & RES1.Fields("PER_ID") & "' AND VENT_STATUS = 'G' AND VENT_IDFOLIO <> '" & Val(lInfo(1).Caption) & "'"

        Set RES2 = con.Execute(sql1)

        If Not RES2.EOF Then
            If Val(RES2.Fields("num")) > 0 And RES1.Fields("PERTP_CODIGO_MEMBRESIA") <> "CLTE" Then
'                    lInfo(0).Caption = "0"
'                    lblDatos(3).Caption = ""
                    MsgBox "El cliente que quiere agregar tiene " & RES2.Fields("NUM") & " operaciones generadas.  " & vbCrLf & vbCrLf & _
                    "Verifique para cerrar o cancelar la venta del cliente. Verfique.", vbInformation
                    'Exit Sub

            End If
        End If
        
        
        lblDatos(2).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
        lblDatos(4).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
        lblDatos(5).Caption = RES1.Fields("MEMBRESIA")
        
        If IsNull(RES1.Fields("total")) Then
        lblDatos(6).Caption = FormatCurrency(0)
        Else
        lblDatos(6).Caption = FormatCurrency(Val(RES1.Fields("TOTAL")))
        End If
        lblClieId(0).Caption = RES1.Fields("PER_ID")
        lblClieId(1).Caption = RES1.Fields("PERTP_TIPO_ID")
        lblClieId(2).Caption = RES1.Fields("PERTP_PER_TIPO")
        lblDatos(3).Caption = RES1.Fields("PER_EMAIL")
        'lInfo(0).Caption = FormatCurrency(RES1.Fields("TEMP_MONEDERO"))
        Me.Caption = "Operación Ticket " & lInfo(1).Caption & " Clte: " & lblDatos(2).Caption
        
        If IsNull(RES1.Fields("PER_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempClie.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(2).Picture = LoadPicture(direccionSistema & "\Temp\TempClie.dat")
        Else
            imgFoto(2).Picture = LoadPicture("")
        End If
        
        updateVenta (Val(lInfo(1).Caption))
        
    Else
        lInfo(0).Caption = "0"
        lblDatos(3).Caption = ""
        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation
    End If
    
End Sub


Private Sub txtClave_LostFocus(Index As Integer)
'    lista_rapida.Visible = False
End Sub

Private Sub txtDesc_DblClick(Index As Integer)
    If txtDesc(Index).Locked = True Then
        MsgBox "No se puede aplicar un descuento general si existen descuentos individuales. Verifique.", vbInformation
    End If
End Sub

Private Sub txtDesc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            If Index = 0 Then
                If Val(Format(txtDesc(0).Text, "General number")) <= Val(Format(txtSub.Text, "General number")) And Val(Format(txtDesc(0).Text, "General number")) >= 0 Then
                    descuentoGral
                    'descuentoGral2
                Else
                    MsgBox "No se puede realizar la operación. Verifique la cantidad.", vbInformation
                End If
            Else
                If Index = 1 Then
                    If Val(Format(txtDesc(1).Text, "General number")) <= 100 Then
                        descuentoPorcentaje
                        'descuentoPorcentaje2
                    Else
                        MsgBox "No se puede realizar la operación. Verifique la cantidad.", vbInformation
                    End If
                End If
            End If
    End If
End Sub
Private Sub descuentoPorcentaje2()
Dim valor As Double

    For b1 = 1 To lista.Rows - 1
        If lista.TextMatrix(b1, 0) = "PRODUCTO" And lista.TextMatrix(b1, 15) <> "MND" Then
            lista.Row = b1
            lista.TextMatrix(b1, 12) = Round(Val(txtDesc(1).Text), 2)
            valor = Val(Format(lista.TextMatrix(b1, 5), "General Number")) * (Val(lista.TextMatrix(b1, 12)) / 100)
            lista.TextMatrix(b1, 11) = FormatCurrency(valor)
            updateVentDet (b1)
            checkPrecio (b1)
            'checkDescuentoInd
            txtDesc(0).Text = (Val(Format(txtDesc(1).Text, "General Number")) / 100) * Val(Format(txtSub.Text, "General Number"))
            'checkPrecioFinal
        End If
    Next b1
End Sub
Private Sub descuentoGral2()
Dim valor As Double
Dim descGral2  As Double
    
    descGral2 = Round(Val(txtDesc(0).Text), 2)
    txtDesc(1).Text = Round((Val(Format(txtDesc(0).Text, "General Number")) * 100) / Val(Format(txtSub.Text, "General Number")), 2)
    
    For b1 = 1 To lista.Rows - 1
        If lista.TextMatrix(b1, 0) = "PRODUCTO" And lista.TextMatrix(b1, 15) <> "MND" Then
            If Val(descGral2) > Val(Format(lista.TextMatrix(b1, 5), "General Number")) Then
                lista.TextMatrix(b1, 11) = lista.TextMatrix(b1, 5)
                valor = Val(Format(lista.TextMatrix(b1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(b1, 5), "General Number"))))
                lista.TextMatrix(b1, 12) = Round(valor, 2)
                updateVentDet (b1)
                checkPrecio (b1)
                'checkDescuentoInd
            Else
                If Val(descGral2) > 0 Then
                    lista.TextMatrix(b1, 11) = FormatCurrency(descGral2)
                    valor = Val(Format(lista.TextMatrix(b1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(b1, 5), "General Number"))))
                    lista.TextMatrix(b1, 12) = Round(valor, 2)
                    updateVentDet (b1)
                    checkPrecio (b1)
                Else
                    If Val(descGral2) <= 0 Then
                        lista.TextMatrix(b1, 11) = FormatCurrency(0)
                        'valor = Val(Format(lista.TextMatrix(b1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(b1, 5), "General Number"))))
                        lista.TextMatrix(b1, 12) = 0
                        updateVentDet (b1)
                        checkPrecio (b1)
                    End If
                End If
            End If
            descGral2 = descGral2 - Val(Format(lista.TextMatrix(b1, 5), "General Number"))
        End If
    Next b1
End Sub
Private Sub descuentoGral()
    txtDesc(1).Text = Round((Val(Format(txtDesc(0).Text, "General Number")) * 100) / Val(Format(txtSub.Text, "General Number")), 2)
    If textDesc.Visible = False Then
        descGral = True
    End If
    checkPrecioFinal

End Sub

Private Sub descuentoPorcentaje()
    txtDesc(0).Text = (Val(Format(txtDesc(1).Text, "General Number")) / 100) * Val(Format(txtSub.Text, "General Number"))
    checkPrecioFinal
End Sub

Private Sub txtObservacion_GotFocus()
'MsgBox Len(txtObservacion.Text)
'    If Len(txtObservacion.Text) <= 1 Then
'        txtObservacion.Text = vbCrLf & ""
'    End If
End Sub

'Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
'End Sub

Private Sub txtObservacion_LostFocus()
    updateVenta (Val(lInfo(1).Caption))
End Sub

Private Sub txtSub_Change()
    On Error Resume Next
    txtImpuesto = FormatCurrency((Format(txtSub.Text, "General Number")) * (0.16))
End Sub
