VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_TrasLados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslados de productos"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   18480
   Icon            =   "FRM_TrasLados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   18480
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   18018
      _Version        =   393216
      TabHeight       =   697
      TabCaption(0)   =   "  Traslados"
      TabPicture(0)   =   "FRM_TrasLados.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListTraslados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListDetalle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "  Datos de traslados"
      TabPicture(1)   =   "FRM_TrasLados.frx":0B24
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1(2)"
      Tab(1).Control(1)=   "Shape1(7)"
      Tab(1).Control(2)=   "Shape1(5)"
      Tab(1).Control(3)=   "lProd(15)"
      Tab(1).Control(4)=   "Borde(13)"
      Tab(1).Control(5)=   "Shape1(0)"
      Tab(1).Control(6)=   "lProd(0)"
      Tab(1).Control(7)=   "Shape1(1)"
      Tab(1).Control(8)=   "lProd(1)"
      Tab(1).Control(9)=   "Borde(16)"
      Tab(1).Control(10)=   "Borde(15)"
      Tab(1).Control(11)=   "lBus(0)"
      Tab(1).Control(12)=   "lBus(1)"
      Tab(1).Control(13)=   "lInfo(0)"
      Tab(1).Control(14)=   "lInfo(1)"
      Tab(1).Control(15)=   "lBus(2)"
      Tab(1).Control(16)=   "lBus(3)"
      Tab(1).Control(17)=   "Borde(0)"
      Tab(1).Control(18)=   "Borde(1)"
      Tab(1).Control(19)=   "Borde(17)"
      Tab(1).Control(20)=   "lBus(4)"
      Tab(1).Control(21)=   "ListProd(1)"
      Tab(1).Control(22)=   "txtProd(7)"
      Tab(1).Control(23)=   "cmBoton(0)"
      Tab(1).Control(24)=   "cmBoton(1)"
      Tab(1).Control(25)=   "ListProd(0)"
      Tab(1).Control(26)=   "TimeIni"
      Tab(1).Control(27)=   "cmBoton(2)"
      Tab(1).Control(28)=   "cmBoton(3)"
      Tab(1).Control(29)=   "cmBoton(4)"
      Tab(1).Control(30)=   "cmBoton(5)"
      Tab(1).Control(31)=   "textBus(0)"
      Tab(1).Control(32)=   "textBus(1)"
      Tab(1).Control(33)=   "textBus(2)"
      Tab(1).Control(34)=   "textBus(3)"
      Tab(1).Control(35)=   "cmBoton(6)"
      Tab(1).Control(36)=   "Check1(0)"
      Tab(1).Control(37)=   "Check1(1)"
      Tab(1).Control(38)=   "cmbDat(0)"
      Tab(1).ControlCount=   39
      TabCaption(2)   =   "  Datos de ingreso"
      TabPicture(2)   =   "FRM_TrasLados.frx":10BE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSComDlg.CommonDialog cmd1 
         Left            =   480
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cmbDat 
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
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Selecciona el tipo de clasificación a la que pertenece el producto, o agrega o edita los existentes"
         Top             =   840
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -58680
         TabIndex        =   27
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -66840
         TabIndex        =   26
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
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
         Index           =   6
         Left            =   -57480
         Picture         =   "FRM_TrasLados.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3000
         Width           =   735
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
         Left            =   -64560
         TabIndex        =   22
         Top             =   3000
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
         Index           =   2
         Left            =   -62400
         TabIndex        =   21
         Top             =   3000
         Width           =   3015
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
         Left            =   -72600
         TabIndex        =   16
         Top             =   2160
         Width           =   3015
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
         Left            =   -74760
         TabIndex        =   15
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   -65400
         Picture         =   "FRM_TrasLados.frx":1BE2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   -65400
         Picture         =   "FRM_TrasLados.frx":216C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   -65400
         Picture         =   "FRM_TrasLados.frx":26F6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5520
         Width           =   495
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   -65400
         Picture         =   "FRM_TrasLados.frx":2C80
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4800
         Width           =   495
      End
      Begin VB.Timer TimeIni 
         Interval        =   250
         Left            =   -66120
         Top             =   360
      End
      Begin MSFlexGridLib.MSFlexGrid ListProd 
         Height          =   6855
         Index           =   0
         Left            =   -74760
         TabIndex        =   9
         Top             =   2760
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   12091
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_TrasLados.frx":320A
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
         Left            =   -58920
         Picture         =   "FRM_TrasLados.frx":32A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Guardar y generar el traslado"
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
         Left            =   -58920
         Picture         =   "FRM_TrasLados.frx":3B6E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
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
         Height          =   1335
         Index           =   7
         Left            =   -64560
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "Escriba alguna observación o descripción del producto"
         Top             =   960
         Width           =   5295
      End
      Begin MSFlexGridLib.MSFlexGrid ListDetalle 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   4800
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   9128
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_TrasLados.frx":4438
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
      Begin MSFlexGridLib.MSFlexGrid ListTraslados 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         WordWrap        =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_TrasLados.frx":45C4
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
      Begin MSFlexGridLib.MSFlexGrid ListProd 
         Height          =   5655
         Index           =   1
         Left            =   -64680
         TabIndex        =   10
         Top             =   3480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9975
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_TrasLados.frx":4724
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
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal de destino"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   17
         Left            =   -74760
         Top             =   840
         Width           =   4005
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   1
         Left            =   -62400
         Top             =   3000
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   -64560
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave producto"
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
         Left            =   -64560
         TabIndex        =   24
         Top             =   2760
         Width           =   1335
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
         Index           =   2
         Left            =   -62400
         TabIndex        =   23
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Productos en lista:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   -64560
         TabIndex        =   20
         Top             =   9240
         Width           =   7575
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   19
         Top             =   9720
         Width           =   5775
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
         Left            =   -72600
         TabIndex        =   18
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave producto"
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
         Left            =   -74760
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   15
         Left            =   -74760
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   16
         Left            =   -72600
         Top             =   2160
         Width           =   3045
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de productos seleccionados para trasladar"
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
         Index           =   1
         Left            =   -64560
         TabIndex        =   8
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   1
         Left            =   -64560
         Top             =   2640
         Width           =   6615
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de productos (Origen)"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   -74760
         Top             =   1800
         Width           =   6615
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   1395
         Index           =   13
         Left            =   -64560
         Top             =   960
         Width           =   5325
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
         Left            =   -64560
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   5
         Left            =   -64560
         Top             =   840
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   -74760
         Top             =   9720
         Width           =   9255
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   2
         Left            =   -64680
         Top             =   9240
         Width           =   7935
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Traslados"
      Begin VB.Menu mn_CrearTraslado 
         Caption         =   "Generar un registro de traslado"
      End
      Begin VB.Menu mn_GenIngreso 
         Caption         =   "Generar un ingreso por traslado"
      End
   End
   Begin VB.Menu mn_Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu mn_CloseTraslado 
         Caption         =   "Cerrar traslado"
      End
      Begin VB.Menu mn_PrintTicket 
         Caption         =   "Impresión de ticket de traslado"
      End
      Begin VB.Menu mn_ExportDatos 
         Caption         =   "Exportar información (Excel)"
      End
      Begin VB.Menu mn_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_ExpImp 
         Caption         =   "Exportar / Importar"
         Begin VB.Menu mn_ExportSucur 
            Caption         =   "Exportar información para envío a sucursal"
         End
         Begin VB.Menu mn_ImportSucur 
            Caption         =   "Importar información de una sucursal"
         End
      End
      Begin VB.Menu mn_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_editarTras 
         Caption         =   "Editar detalle de traslado"
      End
   End
End
Attribute VB_Name = "FRM_TrasLados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resSuc As Recordset
Dim sqlSuc As String
Dim RES1 As Recordset
Dim SQL1 As String
Dim SQL2 As String
Dim RES2 As Recordset
Dim validacion As Boolean
Dim textoError As String

Private Sub Check1_Click(Index As Integer)

    listProd(Index).Redraw = False
    If Check1(Index).value = Unchecked Then
        For b1 = 1 To listProd(Index).Rows - 1
            listProd(Index).Col = 0
            listProd(Index).Row = b1
            listProd(Index).TextMatrix(b1, 0) = Chr(168)
        Next b1
    Else
        For b1 = 1 To listProd(Index).Rows - 1
            listProd(Index).Col = 0
            listProd(Index).Row = b1
            listProd(Index).TextMatrix(b1, 0) = Chr(254)
        Next b1
    End If
    listProd(Index).Redraw = True

End Sub
Private Sub checkValores()
    validacion = False
    
    If cmbDat(0).Text = "" Then
        validacion = False
        textoError = "No se seleccionó una sucursal de destino"
    Else
        If listProd(1).Rows <= 1 Then
            validacion = False
            textoError = "No se ha ingresado ningún producto a la lista de traslado"
        Else
            validacion = True
        End If
    End If
End Sub

Private Sub cancelar()
Dim ques As String


    txtProd(7).Text = ""
    listProd(1).Rows = 1
    cargaLisProd
    cargaSucur
    For b1 = 0 To 3
        textBus(b1).Text = ""
    Next b1
    Check1(0).value = Checked
    Check1(1).value = Checked
End Sub
Private Sub cmBoton_Click(Index As Integer)
    Dim ques As String
    
    listProd(0).Redraw = False
    listProd(1).Redraw = False
    validacion = False
    
    Select Case Index
        Case 0:
            checkValores
            If validacion = True Then
                generarTraslado
            Else
                MsgBox "No se puede realizar la operación. Se ha detectado el siguiente error: " & vbCrLf & vbCrLf & textoError & vbCrLf & vbCrLf & "Verifique.", vbInformation
            End If
        Case 1:
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        Case 2:
            pasarIndividual
        Case 4:
            pasarTodo
        Case 6:
            ques = MsgBox("Dejar la lista vacia. ¿Continuar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                listProd(1).Rows = 1
            End If
        Case 3:
            regresarIndividual
            
            
    End Select
    
    listProd(0).Redraw = True
    listProd(1).Redraw = True
        
    lInfo(0).Caption = "Productos en lista: " & listProd(0).Rows - 1
    lInfo(1).Caption = "Productos en lista: " & listProd(1).Rows - 1
    
    checkInfo

End Sub
Private Sub generarTraslado()
    Dim ques As String
    Dim trasladoId As Long
    Dim sucuOrigen As Integer
    
    ques = MsgBox("Realizar el guardado de traslado con la siguiente información: " & vbCrLf & vbCrLf & lInfo(1).Caption, vbYesNo + vbQuestion)
    If ques = vbYes Then
    
        SQL1 = "SELECT SUC_ID FROM SUCURSAL"
        Set RES1 = con.Execute(SQL1)
        
        If Not RES1.EOF Then
            sucuOrigen = RES1.Fields("suc_id")
        Else
            MsgBox "No se ha podido identificar la sucursal origen. Verifique.", vbInformation
        End If
        
        SQL1 = "INSERT INTO TRASLADOS (TRP_FECHA_CREACION, trp_userid_org, trp_userperid_org, trp_usertipo_org, trp_observaciones, trp_sucid_origen, trp_sucid_destino, trp_status)" & _
        " VALUES (NOW(), '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', '" & txtProd(7).Text & "', '" & sucuOrigen & "', " & _
        "'" & cmbDat(0).ItemData(cmbDat(0).ListIndex) & "', 'G') "
        'MsgBox SQL1
        con.Execute (SQL1)
        
        SQL1 = "select last_insert_id() trasladoId"
        Set RES1 = con.Execute(SQL1)
        If Not RES1.EOF Then
            trasladoId = RES1.Fields("trasladoId")
        End If
        
        With listProd(1)
            For b1 = 1 To .Rows - 1
                SQL1 = "INSERT INTO TRASLADOS_DETALLE (trd_id, trd_prodid, trd_prodserv, trd_prodcant) VALUES (" & _
                "'" & trasladoId & "', '" & .TextMatrix(b1, 8) & "', '" & .TextMatrix(b1, 9) & "', '" & .TextMatrix(b1, 3) & "')"
                con.Execute (SQL1)
            Next b1
        End With
        MsgBox "La información se ha generado satisfactoriamente. " & vbCrLf & vbCrLf & "Verifique", vbInformation
        cancelar
        cargaTraslados
        SSTab1.Tab = 0
        
    End If


End Sub


Private Sub regresarIndividual()
On Error Resume Next
Dim Texto As String
Dim ques As String
Dim num As Long

    Texto = ""
    For b1 = 1 To listProd(1).Rows - 1
        If listProd(1).TextMatrix(b1, 0) = Chr(254) Then
            Texto = Texto & vbCrLf & "Código: " & listProd(1).TextMatrix(b1, 1) & "  Producto: " & listProd(1).TextMatrix(b1, 2)
        End If
    Next b1
    
    If Texto <> "" Then
        ques = MsgBox("Los siguientes productos serán eliminados de la lista ¿Continuar? " & vbCrLf & vbclrf & Texto, vbYesNo + vbInformation)
        If ques = vbYes Then
            num = 0
            For b1 = 1 To listProd(1).Rows - 1
                num = num + 1
                If listProd(1).TextMatrix(num, 0) = Chr(254) Then
                    If listProd(1).Rows > 2 Then
                        listProd(1).RemoveItem (num)
                        num = num - 1
                    Else
                        listProd(1).Rows = 1
                        b1 = 1
                    End If
                End If
            Next b1
        End If
    Else
        MsgBox "No se encontraron elementos. Verifique.", vbInformation
    End If
    
End Sub
Private Sub checkInfo()
Dim costo As Double
Dim venta As Double
Dim cantidad As Long
    
    costo = 0
    venta = 0
    cantidad = 0
    listProd(1).Redraw = False
    For b1 = 1 To listProd(1).Rows - 1
        costo = costo + Val(Val(Format(listProd(1).TextMatrix(b1, 6), "General Number")) * Val(listProd(1).TextMatrix(b1, 3)))
        venta = venta + Val(Val(Format(listProd(1).TextMatrix(b1, 7), "General Number")) * Val(listProd(1).TextMatrix(b1, 3)))
        cantidad = cantidad + Val(listProd(1).TextMatrix(b1, 3))
    Next b1

    lInfo(1).Caption = "Productos en lista:    " & listProd(1).Rows - 1 & "    Cantidad total:        " & cantidad & vbCrLf & vbCrLf & _
                        "Total costo:          " & FormatCurrency(costo) & "    Total venta:           " & FormatCurrency(venta)

    listProd(1).Redraw = True

End Sub

Private Sub pasarIndividual()

Dim Texto As String
Dim encontro As Boolean
listProd(1).Redraw = False
    Texto = ""
    For b1 = 1 To listProd(0).Rows - 1
        If listProd(0).TextMatrix(b1, 0) = Chr(254) Then
            If Val(listProd(0).TextMatrix(b1, 3)) <> 0 Then
                encontro = False
            
                For c1 = 1 To listProd(1).Rows - 1
                    If listProd(1).TextMatrix(c1, 1) = listProd(0).TextMatrix(b1, 1) Then
                        Texto = Texto & vbCrLf & "Código: " & listProd(1).TextMatrix(c1, 1) & "  Producto: " & listProd(1).TextMatrix(c1, 2)
                        encontro = True
                        Exit For
                    End If
                Next c1
            
                If encontro = False Then
                    listProd(1).AddItem ""
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 0) = listProd(0).TextMatrix(b1, 0)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 1) = listProd(0).TextMatrix(b1, 1)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 2) = listProd(0).TextMatrix(b1, 2)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 3) = listProd(0).TextMatrix(b1, 3)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 4) = listProd(0).TextMatrix(b1, 4)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 5) = listProd(0).TextMatrix(b1, 5)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 6) = listProd(0).TextMatrix(b1, 6)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 7) = listProd(0).TextMatrix(b1, 7)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 8) = listProd(0).TextMatrix(b1, 8)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 9) = listProd(0).TextMatrix(b1, 9)
                
                    listProd(1).Row = listProd(1).Rows - 1
                    listProd(1).Col = 0
                    listProd(1).CellFontName = "Wingdings"
                    listProd(1).CellFontBold = True
                    listProd(1).CellFontSize = 16
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 0) = Chr(254)
                End If
                
            End If
        End If
    Next b1
    
listProd(1).Redraw = True

    If Texto <> "" Then
        MsgBox "Los siguientes productos ya se encuentran en la lista y no se consideraron: " & vbCrLf & vbclrf & Texto, vbInformation
    End If

End Sub
Private Sub pasarTodo()

Dim Texto As String
Dim encontro As Boolean
listProd(1).Redraw = False
    Texto = ""
    For b1 = 1 To listProd(0).Rows - 1
'        If ListProd(0).TextMatrix(b1, 0) = Chr(254) Then
            If Val(listProd(0).TextMatrix(b1, 3)) <> 0 Then
                encontro = False
            
                For c1 = 1 To listProd(1).Rows - 1
                    If listProd(1).TextMatrix(c1, 1) = listProd(0).TextMatrix(b1, 1) Then
                        Texto = Texto & vbCrLf & "Código: " & listProd(1).TextMatrix(c1, 1) & "  Producto: " & listProd(1).TextMatrix(c1, 2)
                        encontro = True
                        Exit For
                    End If
                Next c1
            
                If encontro = False Then
                    listProd(1).AddItem ""
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 0) = listProd(0).TextMatrix(b1, 0)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 1) = listProd(0).TextMatrix(b1, 1)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 2) = listProd(0).TextMatrix(b1, 2)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 3) = listProd(0).TextMatrix(b1, 3)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 4) = listProd(0).TextMatrix(b1, 4)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 5) = listProd(0).TextMatrix(b1, 5)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 6) = listProd(0).TextMatrix(b1, 6)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 7) = listProd(0).TextMatrix(b1, 7)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 8) = listProd(0).TextMatrix(b1, 8)
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 9) = listProd(0).TextMatrix(b1, 9)
                
                    listProd(1).Row = listProd(1).Rows - 1
                    listProd(1).Col = 0
                    listProd(1).CellFontName = "Wingdings"
                    listProd(1).CellFontBold = True
                    listProd(1).CellFontSize = 16
                    listProd(1).TextMatrix(listProd(1).Rows - 1, 0) = Chr(254)
                End If
                
            End If
'        End If
    Next b1
listProd(1).Redraw = True
    If Texto <> "" Then
        MsgBox "Los siguientes productos ya se encuentran en la lista y no se consideraron: " & vbCrLf & vbclrf & Texto, vbInformation
    End If

End Sub
Private Sub Form_Load()

    cargaInicial
    cargaLisProd
    cargaTraslados
    
    
End Sub
Private Sub cargaTraslados()
    
    ListTraslados.Redraw = False
    SQL1 = "SELECT * FROM VIEW_TRASLADOS ORDER BY FECHA DESC"
    Set RES1 = con.Execute(SQL1)
    
    ListTraslados.Rows = 1
    
    Do While Not RES1.EOF
        ListTraslados.AddItem ""
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 0) = RES1.Fields("TRP_ID")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 1) = RES1.Fields("FECHA")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 2) = RES1.Fields("USUARIO_GENERA")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 3) = RES1.Fields("STATUS")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 4) = RES1.Fields("SUC_ORIGEN")
        
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 5) = RES1.Fields("PRODUCTOS")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 6) = RES1.Fields("TOTAL")
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 7) = FormatCurrency(RES1.Fields("TOT_PRECIO_VENTA"))
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 8) = FormatCurrency(RES1.Fields("TOT_PRECIO_COSTO"))
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 9) = FormatCurrency(RES1.Fields("TOT_PRECIO_MAY"))
        
        
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 10) = RES1.Fields("SUC_DESTINO") & ""
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 11) = RES1.Fields("FECHA_ENTRADA_DESTINO") & ""
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 12) = RES1.Fields("USUARIO_DESTINO") & ""
        ListTraslados.TextMatrix(ListTraslados.Rows - 1, 13) = RES1.Fields("OBSERVACIONES") & ""
        
        If RES1.Fields("STATUS") = "CERRADO" Then
            ListTraslados.Row = ListTraslados.Rows - 1
            ListTraslados.Col = 0
            ListTraslados.Col = 0
            ListTraslados.CellForeColor = &H808000
            ListTraslados.Col = 3
            ListTraslados.CellForeColor = &H808000
        Else
            If RES1.Fields("STATUS") = "CANCELADO" Then
                ListTraslados.Row = ListTraslados.Rows - 1
                ListTraslados.Col = 0
                ListTraslados.CellForeColor = &H40C0&
                ListTraslados.Col = 3
                ListTraslados.CellForeColor = &H40C0&
            End If
        End If
    
               
        
        RES1.MoveNext
    Loop
     ListTraslados.Redraw = True
   
End Sub
Private Sub cargaInicial()
    SSTab1.Tab = 0
    listProd(1).Rows = 1
    cargaSucur
    listProd(0).ColWidth(8) = 0
    listProd(0).ColWidth(9) = 0
End Sub
Private Sub cargaSucur()
    
SQL1 = "SELECT IDSUCURSAL, SUCURSAL FROM SUCURSALES ORDER BY SUCURSAL"
Set RES1 = con.Execute(SQL1)

cmbDat(0).Clear
Do While Not RES1.EOF
    cmbDat(0).AddItem RES1.Fields("SUCURSAL")
    cmbDat(0).ItemData(cmbDat(0).ListCount - 1) = RES1.Fields("IDSUCURSAL")
    RES1.MoveNext
Loop


End Sub

Private Sub ListProd_Click(Index As Integer)
'''''
End Sub

Private Sub ListProd_DblClick(Index As Integer)
    If listProd(Index).MouseRow = 0 Then
        Call ordenarLista(listProd(Index))
    Else
        If listProd(Index).Col = 0 Then
            Dim b1 As Long
            b1 = listProd(Index).Row
            
            listProd(Index).Row = b1
            listProd(Index).Col = 0
            If listProd(Index).TextMatrix(b1, 0) = Chr(168) Then
                listProd(Index).TextMatrix(b1, 0) = Chr(254)
            Else
                listProd(Index).TextMatrix(b1, 0) = Chr(168)
            End If
        End If
    End If
End Sub

Private Sub ListProd_GotFocus(Index As Integer)
    ConScroll listProd(Index)
End Sub

Private Sub ListProd_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim valor As Long
    If Index = 1 Then
        If listProd(Index).Col = 3 Then
            valor = listProd(Index).TextMatrix(listProd(Index).Row, 3)
            If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Then
                listProd(Index).Text = listProd(Index).Text & Chr(KeyAscii)
                listProd(Index).Text = Val(listProd(Index).Text)
                
                SQL1 = "SELECT CANTIDAD FROM VIEW_PRODUCTOS_INVENTARIO WHERE CODIGO = '" & listProd(Index).TextMatrix(listProd(Index).Row, 1) & "'"
                'MsgBox SQL1
                Set RES1 = con.Execute(SQL1)
                
                If Not RES1.EOF Then
                    If Val(listProd(Index).TextMatrix(listProd(Index).Row, 3)) > Val(RES1.Fields("cantidad")) Then
                        MsgBox "No se puede agregar una cantidad mayor a la existente. Verifique", vbInformation
                        listProd(Index).TextMatrix(listProd(Index).Row, 3) = valor
                    Else
                        checkInfo
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub ListProd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 Then
        If listProd(Index).Col = 3 Then
            Select Case KeyCode
                Case vbKeyDelete
                    listProd(Index).Text = "0"
                Case vbKeyBack
                    If Len(listProd(Index).Text) > 0 Then
                        listProd(Index).Text = Val(Left(listProd(Index).Text, Len(listProd(Index).Text) - 1))
                        If listProd(Index).Text = "" Then
                            listProd(Index).Text = "0"
                        End If
                    End If
            End Select
        End If
    End If
End Sub

Private Sub ListProd_LostFocus(Index As Integer)
    SinScroll listProd(Index)
End Sub

Private Sub ListTraslados_Click()
    cargaListDetalle (ListTraslados.TextMatrix(ListTraslados.Row, 0))
    
End Sub

Private Sub cargaListDetalle(clave As String)
    'On Error Resume Next
    ListDetalle.ColWidth(15) = 0
    ListDetalle.ColWidth(16) = 0
    ListDetalle.ColWidth(17) = 0
    ListDetalle.ColWidth(18) = 0
    ListDetalle.ColWidth(19) = 0
    
    ListDetalle.Redraw = False
    SQL1 = "SELECT * fROM VIEW_TRASLADOS_DETALLE WHERE ID_TRASLADO = '" & clave & "'"
    Set RES1 = con.Execute(SQL1)
    
    ListDetalle.Rows = 1
    
    Do While Not RES1.EOF
        ListDetalle.AddItem ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 0) = RES1.Fields("STATUS")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 1) = RES1.Fields("CODIGO")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 2) = RES1.Fields("NOMBRE")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 3) = RES1.Fields("TIPO")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 4) = RES1.Fields("MARCA")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 5) = RES1.Fields("CANTIDAD_TRASLADO")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 6) = FormatCurrency(RES1.Fields("PRECIO_VENTA"))
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 7) = FormatCurrency(RES1.Fields("PRECIO_COSTO"))
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 8) = FormatCurrency(RES1.Fields("PRECIO_MAY"))
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 9) = RES1.Fields("STOCK_MIN")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 10) = RES1.Fields("STOCK_MAX")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 11) = RES1.Fields("PRESENTACION")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 12) = RES1.Fields("UNIDAD_MEDIDA")
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 13) = RES1.Fields("PROVEEDOR") & ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 14) = RES1.Fields("CODIGO_PROV") & ""
        'DATOS_DE_TRASLADO
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 15) = RES1.Fields("PROD_ID") & ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 16) = RES1.Fields("PROD_SERV") & ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 17) = RES1.Fields("MARCA_ID") & ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 18) = RES1.Fields("TIPO_ID") & ""
        ListDetalle.TextMatrix(ListDetalle.Rows - 1, 19) = RES1.Fields("TIPO_SUBTIPO") & ""
        
        If RES1.Fields("STATUS") = "NO ENVIADO" Then
            ListDetalle.Row = ListDetalle.Rows - 1
            For b1 = 0 To 4
                ListDetalle.Col = b1
                ListDetalle.CellForeColor = vbRed
            Next b1
        End If
        
        
        RES1.MoveNext
    Loop
    ListDetalle.Redraw = True
    
End Sub

Private Sub ListTraslados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ListTraslados.Rows > 1 Then
        ListTraslados_Click
        If Button = vbRightButton Then
            If ListTraslados.TextMatrix(ListTraslados.Row, 3) = "GENERADO" Then
                mn_PrintTicket.Enabled = False
                mn_editarTras.Enabled = True
                mn_ExportSucur.Enabled = False
            Else
                mn_PrintTicket.Enabled = True
                mn_editarTras.Enabled = False
                mn_ExportSucur.Enabled = True
            End If
                PopupMenu mn_Opciones, vbPopupMenuLeftAlign
        End If
    End If
End Sub

Private Sub mn_CloseTraslado_Click()
    Dim ques As String
    
    ques = MsgBox("Al cerrar el traslado ya no podrá realizar ninguna modifcación al detalle del traslado." & vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        SQL1 = "UPDATE TRASLADOS SET TRP_STATUS = 'C' WHERE TRP_ID = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
        con.Execute (SQL1)
        
        cargaTraslados
    End If
End Sub

Private Sub mn_editarTras_Click()
    Dim ques As String
    
    ques = MsgBox("Editar", vbYesNo + vbQuestion)
    If ques = vbYes Then
        cargaEdit
    End If

End Sub
Private Sub cargaEdit()
    cancelar
    listProd(1).Redraw = False
    SQL1 = "SELECT * fROM VIEW_TRASLADOS WHERE TRP_ID = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        
        cmbDat(0).Text = RES1.Fields("SUC_ORIGEN")
        txtProd(7).Text = RES1.Fields("OBSERVACIONES")
        
        SSTab1.Tab = 1
        
        SQL1 = "SELECT * fROM VIEW_TRASLADOS_DETALLE WHERE ID_TRASLADO = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
        Set RES1 = con.Execute(SQL1)
            
        Do While Not RES1.EOF
            listProd(1).AddItem ""
            listProd(1).TextMatrix(listProd(1).Rows - 1, 1) = RES1.Fields("CODIGO")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 2) = RES1.Fields("NOMBRE")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 3) = RES1.Fields("CANTIDAD_TRASLADO")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 4) = RES1.Fields("TIPO")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 5) = RES1.Fields("MARCA")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 7) = FormatCurrency(RES1.Fields("PRECIO_VENTA"))
            listProd(1).TextMatrix(listProd(1).Rows - 1, 6) = FormatCurrency(RES1.Fields("PRECIO_COSTO"))
            listProd(1).TextMatrix(listProd(1).Rows - 1, 8) = RES1.Fields("PROD_ID")
            listProd(1).TextMatrix(listProd(1).Rows - 1, 9) = RES1.Fields("PROD_SERV")
                                        
            listProd(1).Row = listProd(1).Rows - 1
            listProd(1).Col = 0
            listProd(1).CellFontName = "Wingdings"
            listProd(1).CellFontBold = True
            listProd(1).CellFontSize = 16
            listProd(1).TextMatrix(listProd(1).Rows - 1, 0) = Chr(254)
            
            RES1.MoveNext
        Loop
    Else
        MsgBox "No se puede cargar la información. Intente nuevamente.", vbInformation
    End If
    listProd(1).Redraw = True
 
    
End Sub
Private Sub mn_ExportDatos_Click()
    
    Call exportExcel_MD(ListTraslados, ListDetalle)

End Sub

Private Sub mn_ExportSucur_Click()
    Dim ques As String
    If ListTraslados.TextMatrix(ListTraslados.Row, 3) = "CERRADO" Then
        ques = MsgBox("¿Exportar datos del traslado " & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            MsgBox "Esta operación puede tardar dependiendo de la cantidad de información a trasladar.", vbInformation
            check_Info_conDestino
            'exportarInfo_Sucur
            'GuardarINFO
        End If
    End If

End Sub
Private Sub check_Info_conDestino()
On Error Resume Next
    Dim marcaId As String
    Dim tipoId As String
    Dim prod As String
    Dim listaProdu_error As String
    Dim hayError As Boolean
    Dim claveTrd As String
    
    SQL1 = "SELECT * fROM VIEW_TRASLADOS WHERE TRP_ID = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
    Set RES1 = con.Execute(SQL1)

    If Not RES1.EOF Then
        Call ConexionDB_Suc(RES1.Fields("SERVIDOR"), RES1.Fields("DB"), RES1.Fields("USUARIO"), RES1.Fields("PASS"), RES1.Fields("PUERTO"))

        sqlSuc = "SELECT COUNT(*) valor FROM PRODUCTOS"
        Set resSuc = conSuc.Execute(sqlSuc)

        If Not resSuc.EOF Then
        
        hayError = False
        
'        SQL1 = "SELECT ID_TRASLADO, PROD_ID, CODIGO, NOMBRE,  MARCA_ID, TIPO_ID, PROD_UNIMED_PRESENT, CONCAT(PROD_PERALTA_ID, '-',  PROD_PERALTA_TIPOID) USER_ALTA, CONCAT(PROD_PROVEEDOR, '-', PROD_PROVTIPO) PROVEEDOR "
        SQL1 = "SELECT * " & _
        "FROM VIEW_TRASLADOS_DETALLE WHERE ID_TRASLADO = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "' AND STATUS = 'NO ENVIADO' "
        Set RES1 = con.Execute(SQL1)
                
        listaProdu_error = ""
        Do While Not RES1.EOF
                claveTrd = RES1.Fields("ID_TRASLADO") & "-" & RES1.Fields("PROD_ID")
                sqlSuc = "SELECT PROD_NOMBRE, PROD_CODIGO FROM PRODUCTOS WHERE PROD_CODIGO = '" & RES1.Fields("CODIGO") & "'"
                Set resSuc = conSuc.Execute(sqlSuc)
                
                If Not resSuc.EOF Then
                    listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " Ya existe el código en la sucursal destino." & vbCrLf & vbCrLf
                    hayError = True
                    'GoTo siguiente
                Else
                    
                    sqlSuc = "SELECT CTMR_ID, CTMR_MARCA FROM CAT_MARCA WHERE CTMR_ID = '" & RES1.Fields("MARCA_ID") & "'"
                    Set resSuc = conSuc.Execute(sqlSuc)
                    If Not resSuc.EOF Then
                    Else
                        listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " No existe la marca en la sucursal destino." & vbCrLf & vbCrLf
                        hayError = True
'                        GoTo siguiente
                    End If
                    
                    sqlSuc = "SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_ID = '" & RES1.Fields("TIPO_ID") & "' AND CTPT_SUBTIPO = '" & RES1.Fields("TIPO_SUBTIPO") & "'"
                    Set resSuc = conSuc.Execute(sqlSuc)
                    If Not resSuc.EOF Then
                    Else
                        listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " No existe el tipo en la sucursal destino." & vbCrLf & vbCrLf
                        hayError = True
'                        GoTo siguiente
                    End If
                    
                    sqlSuc = "SELECT CTPS_ID, CTPS_NOMBRE FROM CAT_pRESENTACION WHERE CTPS_ID = '" & RES1.Fields("PROD_UNIMED_PRESENT") & "'"
                    Set resSuc = conSuc.Execute(sqlSuc)
                    If Not resSuc.EOF Then
                    Else
                        listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " No existe el tipo de presentación en la sucursal destino." & vbCrLf & vbCrLf
                        hayError = True
                        'GoTo siguiente
                    End If
                    
                    sqlSuc = "SELECT PERTP_PER_ID, PERTP_TIPO_ID FROM PER_TIPO WHERE CONCAT(PERTP_PER_ID, '-',  PERTP_TIPO_ID) = '" & RES1.Fields("PROD_PERALTA_ID") & "-" & RES1.Fields("PROD_PERALTA_TIPOID") & "'"
                    Set resSuc = conSuc.Execute(sqlSuc)
                    If Not resSuc.EOF Then
                    Else
                        listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " No existe el usuario de alta en la sucursal destino." & vbCrLf & vbCrLf
                        hayError = True
'                        GoTo siguiente
                    End If
                    
                    sqlSuc = "SELECT PERTP_PER_ID, PERTP_TIPO_ID  FROM PER_TIPO WHERE CONCAT(PERTP_PER_ID, '-',  PERTP_TIPO_ID) = '" & RES1.Fields("PROD_PROVEEDOR") & "-" & RES1.Fields("PROD_PROVTIPO") & "'"
                    Set resSuc = conSuc.Execute(sqlSuc)
                    If Not resSuc.EOF Then
                    Else
                        listaProdu_error = listaProdu_error & RES1.Fields("CODIGO") & " " & RES1.Fields("NOMBRE") & " No existe el proveedor en la sucursal destino." & vbCrLf & vbCrLf
                        hayError = True
'                        GoTo siguiente
                    End If
                                        
                                        
                                        
                    If hayError = False Then
                    
                        sqlSuc = "INSERT INTO PRODUCTOS " & _
                        "(PROD_NOMBRE, PROD_CODIGO, PROD_DESCRIPCION, PROD_SERV, PROD_CANT, PROD_PRECIO, PROD_MARCA,  " & _
                        "PROD_TIPO, PROD_SUBTIPO, PROD_STATUS, PROD_PRESENTACION, PROD_UNIMED_PRESENT, PROD_STOCK_MIN, " & _
                        "PROD_STOCK_MAX, PROD_PERALTA_ID, PROD_PERALTA_TIPOID, PROD_PERALTA_TIPO, PROD_ALTA_FECHA, prod_proveedor, " & _
                        "prod_provtipo, prod_provsubtipo, PROD_PRECIO_COSTO, PROD_PRECIO_MAY, PROD_CODIGO_PROV, prod_Dependiente) VALUES (" & _
                        "'" & RES1.Fields("NOMBRE") & "', '" & RES1.Fields("CODIGO") & "', '" & RES1.Fields("DESCRIPCION") & "', 'P', " & _
                        "'" & RES1.Fields("CANTIDAD") & "', '" & RES1.Fields("PRECIO_VENTA") & "', '" & RES1.Fields("MARCA_ID") & "',  " & _
                        "'" & RES1.Fields("TIPO_ID") & "', 'P', 'A', '" & RES1.Fields("PRESENTACION") & "', " & _
                        "" & RES1.Fields("PROD_UNIMED_PRESENT") & ", '" & RES1.Fields("STOCK_MIN") & "', '" & RES1.Fields("STOCK_MAX") & "', " & _
                        "'" & RES1.Fields("PROD_PERALTA_ID") & "', '" & RES1.Fields("PROD_PERALTA_TIPOID") & "', '" & RES1.Fields("PROD_PERALTA_TIPO") & "', NOW(), '" & RES1.Fields("prod_proveedor") & "', '" & RES1.Fields("prod_provtipo") & "', '" & RES1.Fields("prod_provsubtipo") & "', '" & RES1.Fields("PRECIO_COSTO") & "', '" & RES1.Fields("PRECIO_MAY") & "', '" & RES1.Fields("CODIGO_PROV") & "', '" & RES1.Fields("PROD_DEPENDIENTE") & "')"

                        conSuc.Execute (sqlSuc)
                        
                        If Err.Number <> 0 Then
                            SQL2 = "UPDATE TRASLADOS_DETALLE SET TRD_STATUS = 'S' WHERE CONCAT(TRD_ID, '-', TRD_PRODID) = '" & claveTrd & "' "
                            con.Execute (SQL2)
                        End If
                    Else
                        SQL2 = "UPDATE TRASLADOS_DETALLE SET TRD_STATUS = 'N' WHERE CONCAT(TRD_ID, '-', TRD_PRODID) = '" & claveTrd & "' "
                        con.Execute (SQL2)
                    End If
                    
                End If
                
                                
                hayError = False
                
                RES1.MoveNext
            Loop
            
            If listaProdu_error = "" Then
            
                MsgBox "La información ha sido enviada.", vbInformation
            Else
                MsgBox "Se encontraron unas incosistencias en la información seleccionada para enviar con la de información de destino." & _
                "Verifica en la lista cuales han sido enviados y cuales no.", vbInformation
                MsgBox listaProdu_error
            End If
            
            cargaListDetalle (ListTraslados.TextMatrix(ListTraslados.Row, 0))
        
        End If
    
    End If
    

End Sub
Private Sub exportarInfo_Sucur()
    
    SQL1 = "SELECT * fROM VIEW_TRASLADOS WHERE TRP_ID = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
    Set RES1 = con.Execute(SQL1)
    
    
    If Not RES1.EOF Then
        Call ConexionDB_Suc(RES1.Fields("SERVIDOR"), RES1.Fields("DB"), RES1.Fields("USUARIO"), RES1.Fields("PASS"), RES1.Fields("PUERTO"))
    
        sqlSuc = "SELECT COUNT(*) valor FROM PRODUCTOS"
        Set resSuc = con.Execute(sqlSuc)
        
        If Not resSuc.EOF Then
        
            SQL1 = "SELECT * FROM VIEW_TRASLADOS_DETALLE WHERE ID_TRASLADO = '" & ListTraslados.TextMatrix(ListTraslados.Row, 0) & "'"
            Set RES1 = con.Execute(SQL1)
            
            Do While Not RES1.EOF
                sqlSuc = "INSERT INTO PRODUCTOS " & _
                "(PROD_NOMBRE, PROD_CODIGO, PROD_DESCRIPCION, PROD_SERV, PROD_CANT, PROD_PRECIO, PROD_MARCA,  " & _
                "PROD_TIPO, PROD_SUBTIPO, PROD_STATUS, PROD_PRESENTACION, PROD_UNIMED_PRESENT, PROD_STOCK_MIN, " & _
                "PROD_STOCK_MAX, PROD_PERALTA_ID, PROD_PERALTA_TIPOID, PROD_PERALTA_TIPO, PROD_ALTA_FECHA, prod_proveedor, " & _
                "prod_provtipo, prod_provsubtipo, PROD_PRECIO_COSTO, PROD_PRECIO_MAY, PROD_CODIGO_PROV, prod_Dependiente) VALUES (" & _
                "'" & RES1.Fields("NOMBRE") & "', '" & RES1.Fields("CODIGO") & "', '" & RES1.Fields("DESCRIPCION") & "', 'P', " & _
                "'" & RES1.Fields("CANTIDAD") & "', '" & RES1.Fields("PRECIO_VENTA") & "', '" & RES1.Fields("MARCA_ID") & "',  " & _
                "'" & RES1.Fields("TIPO_ID") & "', 'P', '" & RES1.Fields("STATUS") & "', '" & RES1.Fields("PRESENTACION") & "', " & _
                "" & RES1.Fields("PROD_UNIMED_PRESENT") & ", '" & RES1.Fields("STOCK_MIN") & "', '" & RES1.Fields("STOCK_MAX") & "', " & _
                "'" & RES1.Fields("PROD_PERALTA_ID") & "', '" & RES1.Fields("PROD_PERALTA_TIPOID") & "', '" & RES1.Fields("PROD_PERALTA_TIPO") & "', NOW(), '" & RES1.Fields("prod_proveedor") & "', '" & RES1.Fields("prod_provtipo") & "', '" & RES1.Fields("prod_provsubtipo") & "', '" & RES1.Fields("PRECIO_COSTO") & "', '" & RES1.Fields("PRECIO_MAY") & "', '" & RES1.Fields("CODIGO_PROV") & "', '" & RES1.Fields("PROD_DEPENDIENTE") & "')"
                'MsgBox sqlSuc
                conSuc.Execute (sqlSuc)
                RES1.MoveNext
            Loop
            
            
            MsgBox "La información ha sido enviada.", vbInformation
        
        End If
    
    End If


End Sub

Private Sub GuardarINFO()
    
    cMd1.DialogTitle = "Guardar archivo de importación"
    cMd1.Filter = "Archivos de exportación/importación AUXS auxdt|*.auxdt"
    cMd1.FileName = ""
    cMd1.ShowSave
    If cMd1.FileName <> "" Then
        generarDoc (cMd1.FileName)
    Else
        Exit Sub
    End If
End Sub
Private Sub generarDoc(ruta_Nombre As String)

    Dim fila As Integer
    Dim columna As Integer
    Dim Free_File As Integer

    Free_File = FreeFile
    Open ruta_Nombre For Output As #Free_File
    
        With ListDetalle
            For fila = 1 To .Rows - 1
                .Row = fila
                For columna = 0 To .Cols - 1
                    .Col = columna
                    If columna > 0 Then
                        Print #Free_File, vbTab;
                    End If
                    Print #Free_File, vbNullString & .Text & vbNullString;
                Next
                Print #Free_File, ""
            Next
        End With
    Close

    MsgBox "Archivo generado. Verifique.", vbInformation



'MsgBox
'
'
'Dim carp As String
'carp = Dir(ruta_Nombre, vbDirectory)
'Shell "explorer " & carp


'    Open ruta_nombre For Output As #1
'    Print #1, "PRUEBA DE ARCHIVO EN RUTA SELECCIONADA"
'    Close

End Sub

Private Sub mn_ImportSucur_Click()
    Dim ques As String

    ques = MsgBox("¿Importar datos del traslado?" & vbCrLf & vbclrf & _
    "Una vez realizada esta acción no podrá deshacerse.", vbYesNo + vbQuestion)
    If ques = vbYes Then
        MsgBox "Esta operación puede tardar dependiendo de la cantidad de información a trasladar.", vbInformation
        BuscarINFO
    End If

End Sub
Private Sub BuscarINFO()
    cMd1.DialogTitle = "Abrir archivo para importación"
    cMd1.Filter = "Archivos de exportación/importación AUXS auxdt|*.auxdt"
    cMd1.FileName = ""
    cMd1.ShowOpen
    If cMd1.FileName <> "" Then
        importINFO (cMd1.FileName)
    Else
        Exit Sub
    End If
End Sub


Private Sub importINFO(ruta_Nombre As String)
Dim linea As String
Dim datos() As String

With ListDetalle
    '.Cols = 3
    .Rows = 1
    Open ruta_Nombre For Input As #1
        While Not EOF(1)
        Line Input #1, linea
            If InStr(1, linea, vbTab) > 1 Then
                datos = Split(linea, vbTab)
                .AddItem datos(0) & vbTab & datos(1) & vbTab & datos(2) & vbTab & datos(3) & vbTab & datos(4) & _
                datos(5) & vbTab & datos(6) & vbTab & datos(7) & vbTab & datos(8) & vbTab & datos(9) & _
                datos(10) & vbTab & datos(11) & vbTab & datos(12) & vbTab & datos(13) & vbTab & datos(14) & _
                datos(15) & vbTab & datos(16) & vbTab & datos(17) & vbTab & datos(18)
             End If
        Wend
    Close #1
End With

End Sub

Private Sub mn_Opciones_Click()
            If ListTraslados.TextMatrix(ListTraslados.Row, 3) = "GENERADO" Then
                mn_PrintTicket.Enabled = False
                mn_editarTras.Enabled = True
                mn_ExportSucur.Enabled = False
            Else
                mn_PrintTicket.Enabled = True
                mn_editarTras.Enabled = False
                mn_ExportSucur.Enabled = True
            End If
End Sub

Private Sub mn_PrintTicket_Click()
    If ListTraslados.TextMatrix(ListTraslados.Row, 3) = "GENERADO" Then
        MsgBox "No se puede realizar la acción. El traslado no se ha cerrado", vbInformation
    Else
        ''''''''
    End If
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index >= 0 And Index <= 1 Then
        If KeyAscii = 13 Then
            cargaLisProd
        End If
    End If
End Sub

Private Sub TimeIni_Timer()
    TimeIni.Enabled = False
    If Me.width > SSTab1.width Then
        listProd(1).width = listProd(1).width + (Me.width - SSTab1.width)
    End If
    
    SSTab1.width = Me.width - 50
    ListTraslados.width = Me.width - 500
    ListDetalle.width = Me.width - 500

End Sub

Private Sub cargaLisProd()
    
    SQL1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE  SUBTIPO = 'PRODUCTO' AND " & _
    "CODIGO LIKE '%" & textBus(0).Text & "%' " & _
    "AND upper(NOMBRE) LIKE upper('%" & textBus(1).Text & "%') "


    Set RES1 = con.Execute(SQL1)
        
    listProd(0).Rows = 1
    listProd(0).Redraw = False
    Do While Not RES1.EOF
        listProd(0).AddItem ""
        listProd(0).TextMatrix(listProd(0).Rows - 1, 1) = RES1.Fields("CODIGO")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 2) = RES1.Fields("NOMBRE")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 3) = RES1.Fields("CANTIDAD")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 4) = RES1.Fields("TIPO")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 5) = RES1.Fields("MARCA")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 6) = FormatCurrency(RES1.Fields("PRECIO_COSTO"))
        listProd(0).TextMatrix(listProd(0).Rows - 1, 7) = FormatCurrency(RES1.Fields("PRECIO_VENTA"))
        listProd(0).TextMatrix(listProd(0).Rows - 1, 8) = RES1.Fields("PROD_ID")
        listProd(0).TextMatrix(listProd(0).Rows - 1, 9) = RES1.Fields("PROD_SERV")
        
        
        listProd(0).Row = listProd(0).Rows - 1
        listProd(0).Col = 0
        listProd(0).CellFontName = "Wingdings"
        listProd(0).CellFontBold = True
        listProd(0).CellFontSize = 16
        listProd(0).TextMatrix(listProd(0).Rows - 1, 0) = Chr(254)
        
        If RES1.Fields("ID_STATUS") = "I" Or RES1.Fields("CANTIDAD") = 0 Then
            listProd(0).Row = listProd(0).Rows - 1
            For b1 = 0 To listProd(0).Cols - 1
                listProd(0).Col = b1
                listProd(0).CellForeColor = vbRed
            Next b1
        End If
        
        RES1.MoveNext
    Loop
    lInfo(0).Caption = "Productos en lista: " & listProd(0).Rows - 1
    lInfo(1).Caption = "Productos en lista: " & listProd(1).Rows - 1
    listProd(0).Redraw = True

End Sub
