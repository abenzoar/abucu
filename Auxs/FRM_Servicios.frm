VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Servicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicios"
   ClientHeight    =   9975
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   16050
   Icon            =   "FRM_Servicios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   16050
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   697
      TabCaption(0)   =   "   Lista de servicios"
      TabPicture(0)   =   "FRM_Servicios.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2(1)"
      Tab(0).Control(1)=   "Shape1(7)"
      Tab(0).Control(2)=   "lInfo(10)"
      Tab(0).Control(3)=   "fotoProd"
      Tab(0).Control(4)=   "lUsuario(20)"
      Tab(0).Control(5)=   "lInfo(0)"
      Tab(0).Control(6)=   "lInfo(1)"
      Tab(0).Control(7)=   "lInfo(3)"
      Tab(0).Control(8)=   "lInfo(5)"
      Tab(0).Control(9)=   "lInfo(8)"
      Tab(0).Control(10)=   "lBus(0)"
      Tab(0).Control(11)=   "lBus(1)"
      Tab(0).Control(12)=   "lBus(2)"
      Tab(0).Control(13)=   "lProd(16)"
      Tab(0).Control(14)=   "Shape1(6)"
      Tab(0).Control(15)=   "Borde(15)"
      Tab(0).Control(16)=   "Borde(0)"
      Tab(0).Control(17)=   "Borde(1)"
      Tab(0).Control(18)=   "lInfo(2)"
      Tab(0).Control(19)=   "lBus(5)"
      Tab(0).Control(20)=   "ListaUsers"
      Tab(0).Control(21)=   "textBus(0)"
      Tab(0).Control(22)=   "textBus(1)"
      Tab(0).Control(23)=   "Timer1"
      Tab(0).Control(24)=   "cmbProd(5)"
      Tab(0).Control(25)=   "Check2"
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "   Datos generales"
      TabPicture(1)   =   "FRM_Servicios.frx":11A4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lProd(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbStatus"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lUsuario(25)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "iFoto"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lProd(61)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lProd(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lProd(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lProd(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lProd(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lUsuario(21)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Borde(21)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Borde(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Borde(3)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Borde(4)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Borde(5)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Borde(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Borde(8)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cMd1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdTipo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmd_Marca"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmbProd(4)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmBoton(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmbProd(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtProd(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtProd(3)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtProd(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtProd(0)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmbUser(7)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Check1"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmBoton(3)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmBoton(0)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmBoton(1)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).ControlCount=   34
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
         Left            =   16320
         Picture         =   "FRM_Servicios.frx":1A7E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   8160
         Width           =   2055
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
         Left            =   120
         Picture         =   "FRM_Servicios.frx":2348
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   8160
         Width           =   3375
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
         Left            =   3720
         Picture         =   "FRM_Servicios.frx":2C12
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   8160
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   -64320
         TabIndex        =   37
         Top             =   1560
         Value           =   1  'Checked
         Width           =   255
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
         Left            =   -69360
         Style           =   2  'Dropdown List
         TabIndex        =   35
         ToolTipText     =   "Selecciona el tipo de clasificación a la que pertenece el producto, o agrega o edita los existentes"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   -66000
         Top             =   840
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asociar servicio a un periodo de tiempo"
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
         Left            =   11880
         MaskColor       =   &H00808080&
         TabIndex        =   32
         Top             =   4080
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
         Index           =   7
         Left            =   11880
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5400
         Visible         =   0   'False
         Width           =   3495
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
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
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
         TabIndex        =   10
         Top             =   1440
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
         Index           =   0
         Left            =   360
         MaxLength       =   65
         TabIndex        =   1
         Top             =   960
         Width           =   4095
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
         TabIndex        =   2
         Top             =   1920
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
         Height          =   2535
         Index           =   3
         Left            =   5400
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   960
         Width           =   5415
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
         Left            =   360
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2880
         Width           =   2655
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
         TabIndex        =   4
         Top             =   3840
         Width           =   3975
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
         Height          =   735
         Index           =   2
         Left            =   14280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
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
         Index           =   4
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5040
         Width           =   3975
      End
      Begin VB.CommandButton cmd_Marca 
         Caption         =   "Marca"
         Height          =   375
         Left            =   9960
         TabIndex        =   9
         Top             =   6960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   9960
         TabIndex        =   8
         Top             =   7560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid ListaUsers 
         Height          =   6975
         Left            =   -74880
         TabIndex        =   12
         Top             =   2040
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   12303
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         GridLines       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Servicios.frx":34DC
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
         Left            =   7080
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   2835
         Index           =   8
         Left            =   11760
         Top             =   960
         Width           =   2475
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   2595
         Index           =   6
         Left            =   5400
         Top             =   960
         Width           =   5475
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   5
         Left            =   360
         Top             =   5040
         Width           =   4035
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   360
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   360
         Top             =   2880
         Width           =   2715
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   360
         Top             =   1920
         Width           =   3315
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   21
         Left            =   360
         Top             =   960
         Width           =   4155
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
         Left            =   -65160
         TabIndex        =   38
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Servicios en lista:"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   9480
         Width           =   15375
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   1
         Left            =   -69360
         Top             =   1440
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   -72600
         Top             =   1440
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   15
         Left            =   -74880
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   -74760
         Top             =   960
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
         Left            =   -74760
         TabIndex        =   34
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
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
         Left            =   11880
         TabIndex        =   33
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
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
         Index           =   2
         Left            =   -69360
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio"
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
         Left            =   -72600
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave servicio"
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   8
         Left            =   -62640
         TabIndex        =   27
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
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
         Left            =   -62640
         TabIndex        =   26
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   -62640
         TabIndex        =   25
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -62640
         TabIndex        =   24
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio:"
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
         Left            =   -62640
         TabIndex        =   23
         Top             =   2280
         Width           =   3735
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
         Left            =   -62640
         TabIndex        =   22
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Image fotoProd 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   -62880
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Servicio *"
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
         TabIndex        =   21
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Código *"
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
         TabIndex        =   20
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción general"
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
         Left            =   5400
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio*"
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
         TabIndex        =   18
         Top             =   2520
         Width           =   2415
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
         TabIndex        =   17
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Image iFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   11760
         Stretch         =   -1  'True
         Top             =   960
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
         Left            =   11760
         TabIndex        =   16
         Top             =   720
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
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   9360
         Width           =   4695
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
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   14
         Top             =   4680
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
         Left            =   -74880
         TabIndex        =   13
         Top             =   8760
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   420
         Index           =   7
         Left            =   -74880
         Top             =   9360
         Width           =   15615
      End
      Begin VB.Image Image2 
         Height          =   9735
         Index           =   1
         Left            =   -75000
         Picture         =   "FRM_Servicios.frx":357F
         Stretch         =   -1  'True
         Top             =   480
         Width           =   17655
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   0
         Top             =   9360
         Width           =   18255
      End
      Begin VB.Image Image2 
         Height          =   9735
         Index           =   0
         Left            =   -120
         Picture         =   "FRM_Servicios.frx":105BF
         Stretch         =   -1  'True
         Top             =   480
         Width           =   19215
      End
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   2595
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   5475
   End
   Begin VB.Menu mn_Serv 
      Caption         =   "Servicios"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mn_Eliminar 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu mn_Cat 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_TipServ 
         Caption         =   "Tipo de servicio"
      End
   End
End
Attribute VB_Name = "FRM_Servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sql1 As String
    Dim res1 As Recordset
    Dim RES2 As Recordset
    Dim RES3 As Recordset
    Dim checkError As Boolean
    Dim prodId As String
    Dim save As Boolean
Private Sub cargaPeriodo()
    sql1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS FROM CAT_PERIODO"
    Set res1 = con.Execute(sql1)
    
    Do While Not res1.EOF
        cmbUser(7).AddItem res1.Fields("CTPR_PERIODO")
        cmbUser(7).ItemData(cmbUser(7).ListCount - 1) = res1.Fields("CTID_PERIODO")
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

Private Sub Check2_Click()
    cargaLista
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            checarCampos
            If checkError = False Then
                crearServicio
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
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
                crearServicio
                lbStatus.Caption = "Estatus: Agregando servicio"
                SSTab1.TabEnabled(1) = True
                SSTab1.Tab = 1
                SSTab1.TabEnabled(0) = False
                txtProd(0).SetFocus
                save = False
            Else
                MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
            End If
           
    End Select

End Sub

Private Sub cmbProd_Click(Index As Integer)
    Select Case Index
        Case 4:
            If lProd(6).ForeColor = vbRed Then
                lProd(6).ForeColor = vbBlack
            End If
        Case 1:
            If lProd(61).ForeColor = vbRed Then
                lProd(61).ForeColor = vbBlack
            End If
        Case 5: cargaLista
    End Select

End Sub

Public Sub cmdTipo_Click()
    cargaTipoServ
End Sub

Private Sub ListaUsers_Click()
    muestraInfo (ListaUsers.TextMatrix(ListaUsers.Row, 0))
End Sub
Private Sub muestraInfo(prodCodigo As String)

    fotoProd.Picture = LoadPicture("")
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_DESCRIPCION, " & _
    "CTPT_TIPO, PROD_PRESENTACION, PROD_FOTO " & _
    "FROM PRODUCTOS T1, CAT_TIPO T3 " & _
    "WHERE T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND T1.PROD_CODIGO = '" & prodCodigo & "' "
    Set res1 = con.Execute(sql1)
    If Not res1.EOF Then
        If IsNull(res1.Fields("PROD_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write res1.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempServ.dat", adSaveCreateOverWrite
            Imagen1.Close
            fotoProd.Picture = LoadPicture(direccionSistema & "\Temp\TempServ.dat")
        Else
            fotoProd.Picture = LoadPicture("")
        End If
        lInfo(0).Caption = "Servicio: " & res1.Fields("PROD_NOMBRE")
        lInfo(1).Caption = "Código: " & res1.Fields("PROD_CODIGO")
        lInfo(3).Caption = "Tipo: " & res1.Fields("CTPT_TIPO")
        lInfo(5).Caption = "Precio: " & FormatCurrency(res1.Fields("PROD_PRECIO"))
        lInfo(8).Caption = "Descripcion: " & res1.Fields("PROD_DESCRIPCION")
    Else
        'MsgBox "OK"
        fotoProd.Picture = LoadPicture("")
        lInfo(0).Caption = "Servicio: "
        lInfo(1).Caption = "Código: "
        lInfo(3).Caption = "Tipo: "
        lInfo(5).Caption = "Precio: "
        lInfo(8).Caption = "Descripción: "
    End If
End Sub

Private Sub crearServicio()
'    Dim status As String
'    Dim cp As String
'    Dim tel1 As String
'    Dim tel2 As String
    On Error Resume Next
    
    Dim presentacion As String
    Dim status As String
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
        
    If txtProd(3).Text = "" Then
        txtProd(3).Text = "Ninguna"
    End If

    status = Left(cmbProd(4).Text, 1)
    
    If lbStatus.Caption = "Estatus: Agregando servicio" Then
        Err.Clear
        sql1 = "INSERT INTO PRODUCTOS " & _
        "(PROD_NOMBRE, PROD_CODIGO, PROD_DESCRIPCION, PROD_SERV, PROD_PRECIO, " & _
        "PROD_TIPO, PROD_SUBTIPO, PROD_STATUS,  " & _
        "PROD_PERALTA_ID, PROD_PERALTA_TIPOID, PROD_PERALTA_TIPO, PROD_ALTA_FECHA, PROD_CANT, PROD_dEPENDIENTE) VALUES (" & _
        "'" & txtProd(0).Text & "', '" & txtProd(1).Text & "', '" & txtProd(3).Text & "', 'S', " & _
        "'" & txtProd(2).Text & "',  " & _
        "'" & cmbProd(1).ItemData(cmbProd(1).ListIndex) & "', 'S', '" & status & "', " & _
        "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW(), '0', 'U')"
'        MsgBox sql1
        con.Execute (sql1)
        'MsgBox Err.Number
        If Err.Number = -2147217900 Then
            MsgBox "El código que quiere registrar ya existe para un servicio y no puede duplicarse. Por favor verifique.", vbCritical
            Exit Sub
        End If
        sql1 = "select last_insert_id() prodid"
        Set res1 = con.Execute(sql1)
        If Not res1.EOF Then
            prodId = res1.Fields("prodid")
        End If
    
    
    Else
        sql1 = "UPDATE PRODUCTOS SET PROD_NOMBRE = '" & txtProd(0).Text & "', " & _
        "PROD_CODIGO = '" & txtProd(1).Text & "', " & _
        "PROD_DESCRIPCION = '" & txtProd(3).Text & "', " & _
        "PROD_PRECIO = '" & txtProd(2).Text & "', " & _
        "PROD_TIPO = " & cmbProd(1).ItemData(cmbProd(1).ListIndex) & ", " & _
        "PROD_STATUS = '" & status & "' " & _
        "WHERE PROD_CODIGO = '" & prodId & "'"
        con.Execute (sql1)
                
    End If
    'Para la fotoi
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempSer.dat")
        If Not res1.EOF Then
            res.Open "SELECT * FROM Productos WHERE prod_codigo = '" & prodId & "'", con, adOpenStatic, adLockOptimistic
            If res.EOF Then
            Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempSer.dat")
                res.Fields("prod_Foto") = Imagen1.Read
                res.Update
            End If
        End If
    End If
    
    MsgBox "Información guardada.", vbInformation
    save = True
    cancelar


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
    
End Sub

Private Sub limpiarCampos()
    
    For b1 = 0 To 3
        txtProd(b1).Text = ""
    Next b1
    
    cmbProd(1).Clear
    cmbProd(4).Clear

End Sub

Private Sub cancelar()
    limpiarCampos
    CargaGeneral
    cargaTipoServ
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


Private Sub checarCampos()
    checkError = False
    
    For b1 = 0 To 2
        If txtProd(b1).Text = "" Then
            checkError = True
            lProd(b1).ForeColor = vbRed
            Exit For
        End If
    Next b1
    
    If checkError = False Then
        If cmbProd(1).Text = "" Then
            checkError = True
            lProd(61).ForeColor = vbRed
        Else
            If cmbProd(4).Text = "" Then
                checkError = True
                lProd(6).ForeColor = vbRed
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    CargaGeneral
    cargaLista
    cargaPeriodo
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
            PopupMenu mn_Serv, vbPopupMenuLeftAlign
        End If
    Else
            mn_Add.Enabled = True
            mn_Edit.Enabled = False
            mn_Eliminar.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_Serv, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Add_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea agregar un servicio?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Agregando servicio"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            txtProd(0).SetFocus
            save = False
        End If

End Sub

Private Sub mn_Edit_Click()
    Dim ques As String
    
    ques = MsgBox("Desea editar el servicio: " & ListaUsers.TextMatrix(ListaUsers.Row, 0) & vbCrLf & _
            ListaUsers.TextMatrix(ListaUsers.Row, 1) & " " & ListaUsers.TextMatrix(ListaUsers.Row, 2), vbYesNo + vbQuestion)
        If ques = vbYes Then
            prodId = ListaUsers.TextMatrix(ListaUsers.Row, 0)
            lbStatus.Caption = "Estatus: Editando servicio"
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
    iFoto.Visible = True
   sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_STATUS, PROD_PRECIO, PROD_DESCRIPCION, " & _
    "CTPT_TIPO, PROD_PRESENTACION, PROD_FOTO " & _
    "FROM PRODUCTOS T1, CAT_TIPO T3 " & _
    "WHERE T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND PROD_CODIGO = '" & prodId & "' "
    'MsgBox SQL1
    Set RES2 = con.Execute(sql1)
    Dim b1 As Long
    If Not RES2.EOF Then
        txtProd(0).Text = RES2.Fields("PROD_NOMBRE")
        txtProd(1).Text = RES2.Fields("PROD_CODIGO")
        txtProd(2).Text = RES2.Fields("PROD_PRECIO")
        txtProd(3).Text = RES2.Fields("PROD_DESCRIPCION")
        If IsNull(RES2.Fields("CTPT_TIPO")) Then
        Else
            cmbProd(1).Text = RES2.Fields("CTPT_TIPO")
        End If
        
        If RES2.Fields("PROD_STATUS") = "A" Then
            cmbProd(4).Text = "ACTIVO"
        Else
            cmbProd(4).Text = "INACTIVO"
        End If
                       
        If IsNull(RES2.Fields("PROD_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES2.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempServ.dat", adSaveCreateOverWrite
            Imagen1.Close
            iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempServ.dat")
        Else
            iFoto.Picture = LoadPicture("")
        End If
        
    End If
    
End Sub

Private Sub mn_TipServ_Click()
    tipoCatTipo = "S"
    CAT_Tipo.Show vbModal

End Sub
Private Sub CargaGeneral()
    SSTab1.Tab = 0
    fotoProd.Picture = LoadPicture("")
    iFoto.Picture = LoadPicture("")
    SSTab1.TabEnabled(1) = False
    cargaTipoServ
End Sub
Private Sub cargaLista()
    ListaUsers.Rows = 1
    Dim texto1 As String
    

    texto1 = ""
    If cmbProd(5).Text <> "TODOS" Then
        texto1 = texto1 & " AND upper(CTPT_TIPO) LIKE upper('%" & cmbProd(5).Text & "%') "
    End If
    
    If Check2.value = Checked Then
        texto1 = texto1 & " AND PROD_STATUS = 'A' "
    End If
    
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_DESCRIPCION, " & _
    "CTPT_TIPO, PROD_PRESENTACION, PROD_FOTO " & _
    "FROM PRODUCTOS T1, CAT_TIPO T3 " & _
    "WHERE T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND T1.PROD_SERV = 'S' " & _
    "AND T1.PROD_CODIGO LIKE '" & textBus(0).Text & "%' " & _
    "AND upper(T1.PROD_NOMBRE) LIKE upper('%" & textBus(1).Text & "%') " & _
    " " & texto1

    Set res1 = con.Execute(sql1)
        
    Do While Not res1.EOF
        ListaUsers.AddItem ""
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 0) = res1.Fields("PROD_CODIGO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 1) = res1.Fields("PROD_NOMBRE")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 2) = res1.Fields("CTPT_TIPO")
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 3) = FormatCurrency(res1.Fields("PROD_PRECIO"))
        ListaUsers.TextMatrix(ListaUsers.Rows - 1, 4) = res1.Fields("STATUS")
            If res1.Fields("STATUS") = "INACTIVO" Then
                ListaUsers.Row = ListaUsers.Rows - 1
                For b1 = 0 To ListaUsers.Cols - 1
                    ListaUsers.Col = b1
                    ListaUsers.CellForeColor = &H80FF&
                Next b1
            End If
        
        res1.MoveNext
    Loop
    lInfo(2).Caption = "Servicios en lista: " & ListaUsers.Rows - 1

End Sub

Private Sub cargaTipoServ()

    sql1 = ("SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'S' ORDER BY CTPT_TIPO")
    Set res1 = con.Execute(sql1)
    
    cmbProd(1).Clear
    cmbProd(5).Clear
    cmbProd(5).AddItem "TODOS"
    Do While Not res1.EOF
        cmbProd(1).AddItem res1.Fields("CTPT_TIPO")
        cmbProd(1).ItemData(cmbProd(1).ListCount - 1) = res1.Fields("CTPT_ID")
        cmbProd(5).AddItem res1.Fields("CTPT_TIPO")
        cmbProd(5).ItemData(cmbProd(5).ListCount - 1) = res1.Fields("CTPT_ID")
        res1.MoveNext
    Loop

    cmbProd(4).Clear
    cmbProd(4).AddItem "ACTIVO"
    cmbProd(4).AddItem "INACTIVO"
    cmbProd(4).ListIndex = 0

End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaLista
    End If
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    Image2(0).width = Me.width
    Image2(0).height = Me.height
    Image2(1).width = Me.width
    Image2(1).height = Me.height

    ListaUsers.width = Me.width - 500

End Sub
