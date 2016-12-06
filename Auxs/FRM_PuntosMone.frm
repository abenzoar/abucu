VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_PuntosMone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de promociones para Puntos - Monedero electrónico"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16710
   Icon            =   "FRM_PuntosMone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   16710
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   697
      TabCaption(0)   =   "  Lista de promociones"
      TabPicture(0)   =   "FRM_PuntosMone.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "listPromo(0)"
      Tab(0).Control(1)=   "listPromo(1)"
      Tab(0).Control(2)=   "listPromo(2)"
      Tab(0).Control(3)=   "lProd(0)"
      Tab(0).Control(4)=   "Shape1(0)"
      Tab(0).Control(5)=   "lProd(16)"
      Tab(0).Control(6)=   "Shape1(6)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   " Datos generales de los puntos"
      TabPicture(1)   =   "FRM_PuntosMone.frx":0B24
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Borde(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Borde(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lUsuario(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lUsuario(130)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Borde(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lUsuario(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Borde(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lUsuario(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Borde(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lUsuario(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Borde(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lUsuario(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lUsuario(5)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Borde(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Borde(18)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Borde(17)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lUsuario(6)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lUsuario(7)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lUsuario(8)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Borde(6)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "listProd(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "listPromo(3)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtPromo(0)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmbPromo(0)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmbPromo(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtPromo(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmbPromo(2)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtPromo(2)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmbPromo(3)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmBoton(9)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cmBoton(8)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmBoton(0)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmBoton(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmbProd(6)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "cmbProd(5)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Check1(0)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "cmbPromo(4)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).ControlCount=   37
      Begin VB.ComboBox cmbPromo 
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
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   5160
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Selección"
         Enabled         =   0   'False
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
         Left            =   15000
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbProd 
         Enabled         =   0   'False
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
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Selecciona el tipo de clasificación a la que pertenece el producto, o agrega o edita los existentes"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cmbProd 
         Enabled         =   0   'False
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
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selecciona la marca a la que pertenece el producto, o agrega o edita las existentes"
         Top             =   1200
         Width           =   3015
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
         Left            =   14280
         Picture         =   "FRM_PuntosMone.frx":10BE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6960
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
         Left            =   8400
         Picture         =   "FRM_PuntosMone.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6960
         Width           =   3375
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
         Left            =   6600
         Picture         =   "FRM_PuntosMone.frx":2252
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   615
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
         Left            =   7320
         Picture         =   "FRM_PuntosMone.frx":27DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.ComboBox cmbPromo 
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
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtPromo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   2
         Left            =   4440
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   6120
         Width           =   3735
      End
      Begin VB.ComboBox cmbPromo 
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
         TabIndex        =   5
         Top             =   6120
         Width           =   3375
      End
      Begin VB.TextBox txtPromo 
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
         TabIndex        =   3
         Top             =   4080
         Width           =   2535
      End
      Begin VB.ComboBox cmbPromo 
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox cmbPromo 
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtPromo 
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
         Top             =   1200
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid listPromo 
         Height          =   3135
         Index           =   0
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_PuntosMone.frx":2D66
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
      Begin MSFlexGridLib.MSFlexGrid listPromo 
         Height          =   3615
         Index           =   1
         Left            =   -74760
         TabIndex        =   16
         Top             =   4320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Día                                | Promo    "
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
      Begin MSFlexGridLib.MSFlexGrid listPromo 
         Height          =   3615
         Index           =   2
         Left            =   -71880
         TabIndex        =   17
         Top             =   4320
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Código                | Producto                                  | Tipo de producto  |  Valor puntos  "
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
      Begin MSFlexGridLib.MSFlexGrid listPromo 
         Height          =   3615
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Día                                        | dia  | Promo  "
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
      Begin MSFlexGridLib.MSFlexGrid listProd 
         Height          =   4695
         Index           =   0
         Left            =   8400
         TabIndex        =   29
         Top             =   1920
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_PuntosMone.frx":2E58
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
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   6
         Left            =   360
         Top             =   5160
         Width           =   3405
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes que aplica"
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
         Left            =   360
         TabIndex        =   32
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   11520
         TabIndex        =   31
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
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
         Left            =   8400
         TabIndex        =   30
         Top             =   840
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   17
         Left            =   8400
         Top             =   1200
         Width           =   2925
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   18
         Left            =   11520
         Top             =   1200
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   5
         Left            =   4440
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Días que aplica"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   1995
         Index           =   3
         Left            =   4440
         Top             =   6120
         Width           =   3765
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus"
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
         TabIndex        =   25
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   360
         Top             =   6120
         Width           =   3405
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor a aplicar"
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
         TabIndex        =   24
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   1
         Left            =   360
         Top             =   4080
         Width           =   2565
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de valor a aplicar"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   360
         Top             =   3120
         Width           =   3405
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplica a"
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
         Left            =   360
         TabIndex        =   22
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la promoción *"
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
         Top             =   840
         Width           =   2775
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   360
         Top             =   1200
         Width           =   3525
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H0000C000&
         BorderWidth     =   4
         Height          =   435
         Index           =   9
         Left            =   360
         Top             =   2160
         Width           =   3405
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Productos que otrogan los puntos"
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
         Left            =   -71880
         TabIndex        =   20
         Top             =   3840
         Width           =   3615
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   -71880
         Top             =   4080
         Width           =   4815
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Días que aplica"
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
         TabIndex        =   19
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   -74760
         Top             =   4080
         Width           =   2775
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "FRM_PuntosMone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo As String
Dim valDatos As Boolean
Dim sql1 As String
Dim res1 As Recordset

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 8:
        addDia
        Case 9:
        supDay
        Case 0:
        crearPromo
        Case 1
        cancelar
    End Select
End Sub
Private Sub checarDatos()
    valDatos = False
    For b1 = o To 2
        If txtPromo(b1).Text = "" Then
            valDatos = True
            Exit Sub
        Else
            If cmbPromo(0).Text = "" Then
                valDatos = True
                Exit Sub
            Else
                If cmbPromo(1).Text = "" Then
                    valDatos = True
                    Exit Sub
                Else
                    If listPromo(3).Rows = 1 Then
                        valDatos = True
                        Exit Sub
                    End If
                End If
            End If
        End If
        listProd(0).Rows = 1
    Next b1
    
End Sub
Private Sub crearPromo()
    Dim promoId As Long
    
    checarDatos
    
    If valDatos = False Then
        sql1 = "INSERT INTO CAT_PUNTOS (PNT_PROMOCION, PNT_dESCRIPCION, PNT_FECHAHORA, PNT_TIPO, PNT_TIPOVALOR, PNT_VALOR, PNT_STATUS, PNT_APLICA) VALUES (" & _
        "'" & txtPromo(0).Text & "', '" & txtPromo(2).Text & "', NOW(), '" & Left(cmbPromo(0).Text, 1) & "', '" & Left(cmbPromo(1).Text, 1) & "', '" & txtPromo(1).Text & "', '" & Left(cmbPromo(2).Text, 1) & "', '" & Left(cmbPromo(4).Text, 1) & "') "
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() promoId"
        Set res1 = con.Execute(sql1)
        If Not res1.EOF Then
            promoId = res1.Fields("promoId")
        End If
        
        With listPromo(3)
            For b1 = 1 To .Rows - 1
                sql1 = "INSERT INTO CAT_PUNTOS_DIAS (pntds_pntid, pntds_dia) values ('" & promoId & "', '" & .TextMatrix(b1, 1) & "'   ) "
                con.Execute (sql1)
            Next b1
        End With
        
        MsgBox "Promoción de puntos agregada. Verifique.", vbInformation
        borrarDatos
        SSTab1.Tab = 0
        cargaPromos
    Else
        MsgBox "Se ha detectado un error al momento de guardar. Verfiique la información", vbInformation
    End If
    
End Sub
Private Sub cargaPromos()
    '''''
    listPromo(0).Rows = 1
    listPromo(1).Rows = 1
    listPromo(2).Rows = 1
    
    sql1 = "SELECT * fROM VIEW_PUNTOS ORDER BY FECHAHORA DESC"
    Set res1 = con.Execute(sql1)
    Do While Not res1.EOF
         listPromo(0).AddItem ""
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 0) = res1.Fields("ID")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 1) = res1.Fields("PROMOCION")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 2) = res1.Fields("TIPO")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 3) = res1.Fields("TIPO_VALOR")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 4) = res1.Fields("VALOR")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 5) = res1.Fields("APLICA")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 6) = res1.Fields("STATUS")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 7) = "0"
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 8) = res1.Fields("DIAS")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 9) = res1.Fields("FECHAHORA")
         listPromo(0).TextMatrix(listPromo(0).Rows - 1, 10) = res1.Fields("DESCRIPCION")
        res1.MoveNext
    Loop
    
    If listPromo(0).Rows > 2 Then
        listPromo(0).Row = 2
        listPromo_Click (0)
    End If

    
    
End Sub
Private Sub supDay()
    listPromo(3).RemoveItem (listPromo(3).Row)
End Sub
Private Sub addDia()
Dim valida As Boolean
valida = False
    If cmbPromo(3).ListIndex = 0 Then
        For b1 = 1 To 7
            valida = False
            For c1 = 1 To listPromo(3).Rows - 1
                If listPromo(3).TextMatrix(c1, 1) = b1 Then
                    valida = True
                    Exit For
                End If
            Next c1
            If valida = False Then
                listPromo(3).AddItem ""
                listPromo(3).TextMatrix(listPromo(3).Rows - 1, 0) = Format(b1, "dddd")
                listPromo(3).TextMatrix(listPromo(3).Rows - 1, 1) = b1
            End If
        Next b1
    Else
        valida = False
        For c1 = 1 To listPromo(3).Rows - 1
            If listPromo(3).TextMatrix(c1, 1) = cmbPromo(3).ListIndex Then
                valida = True
                Exit For
            End If
        Next c1
        If valida = False Then
            listPromo(3).AddItem ""
            listPromo(3).TextMatrix(listPromo(3).Rows - 1, 0) = cmbPromo(3).Text
            listPromo(3).TextMatrix(listPromo(3).Rows - 1, 1) = cmbPromo(3).ListIndex
        End If
    End If
    
End Sub

Private Sub cmbPromo_Click(Index As Integer)
    If cmbPromo(0).ListIndex = 0 Then
    Else
        If cmbPromo(0).ListIndex = 1 Then
            MsgBox "Opción no disponible.Verifique. ", vbInformation
        End If
    End If
End Sub

Private Sub Form_Load()
    cargaInicial
    cargaPromos
End Sub

Private Sub listPromo_Click(Index As Integer)
    listPromo(1).Rows = 1
    
    sql1 = "SELECT PNTDS_DIA, PNTDS_PNTID FROM CAT_PUNTOS_DIAS WHERE PNTDS_PNTID = '" & listPromo(0).TextMatrix(listPromo(0).Row, 0) & "'"
    Set res1 = con.Execute(sql1)
    
'    MsgBox SQL1
    Do While Not res1.EOF
        listPromo(1).AddItem ""
        listPromo(1).TextMatrix(listPromo(1).Rows - 1, 0) = Format(res1.Fields("PNTDS_DIA"), "dddd")
        listPromo(1).TextMatrix(listPromo(1).Rows - 1, 1) = res1.Fields("PNTDS_PNTID")
        res1.MoveNext
    Loop
End Sub

Private Sub mn_Add_Click()
    tipo = "Add"
    'crearNuevo
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    borrarDatos
    
End Sub
Private Sub borrarDatos()
    For b1 = o To 2
        txtPromo(b1).Text = ""
        listPromo(3).Rows = 1
        listProd(0).Rows = 1
    Next b1
End Sub
Private Sub cancelar()
    Dim ques As String
    ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        borrarDatos
         SSTab1.Tab = 0
         SSTab1.TabEnabled(1) = False
    End If
End Sub
Private Sub cargaInicial()
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(0) = True
    
    borrarDatos
    cargaCombos
End Sub
Private Sub crearNuevo()
    borrarDatos
    SSTab1.Tab = 1

End Sub
Private Sub cargaCombos()
    
    cmbPromo(0).Clear
    cmbPromo(0).AddItem "Total de la venta"
    cmbPromo(0).AddItem "Productos especificos"
    
    cmbPromo(1).Clear
    cmbPromo(1).AddItem "Porcentaje descuento"
    cmbPromo(1).AddItem "Valor específico"
    
    cmbPromo(2).Clear
    cmbPromo(2).AddItem "Activo"
    cmbPromo(2).AddItem "Inactivo"
    cmbPromo(2).ListIndex = 0
    
    cmbPromo(4).Clear
    cmbPromo(4).AddItem "Membresia (activa)"
    cmbPromo(4).AddItem "Todos"
    
    cmbPromo(3).Clear
    cmbPromo(3).AddItem "Todos (L-D)"
    
    For b1 = 1 To 7
        cmbPromo(3).AddItem Format(b1, "dddd")
    Next b1
    

End Sub

Private Sub txtPromo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        Call NumerosPunto(KeyAscii)
    End If
End Sub
