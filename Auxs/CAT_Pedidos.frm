VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CAT_Pedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de inventario/almacen por pedido"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16770
   Icon            =   "CAT_Pedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CAT_Pedidos.frx":058A
   ScaleHeight     =   10350
   ScaleWidth      =   16770
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   697
      MouseIcon       =   "CAT_Pedidos.frx":0B14
      TabCaption(0)   =   "  Lista de pedidos"
      TabPicture(0)   =   "CAT_Pedidos.frx":0B30
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "timeCarga"
      Tab(0).Control(1)=   "ListaPed(0)"
      Tab(0).Control(2)=   "ListaPed(1)"
      Tab(0).Control(3)=   "Shape1(5)"
      Tab(0).Control(4)=   "Shape1(4)"
      Tab(0).Control(5)=   "lProd(2)"
      Tab(0).Control(6)=   "lProd(0)"
      Tab(0).Control(7)=   "lInfo(4)"
      Tab(0).Control(8)=   "Shape1(3)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "  Datos generales"
      TabPicture(1)   =   "CAT_Pedidos.frx":10CA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "imgFoto(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lInfo(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "SSTab2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin TabDlg.SSTab SSTab2 
         Height          =   9375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   16536
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Paso 1: Datos generales del pedido"
         TabPicture(0)   =   "CAT_Pedidos.frx":1664
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmBoton(5)"
         Tab(0).Control(1)=   "cmBoton(0)"
         Tab(0).Control(2)=   "txtProd(4)"
         Tab(0).Control(3)=   "cmBoton(6)"
         Tab(0).Control(4)=   "cmbProd(3)"
         Tab(0).Control(5)=   "txtProd(3)"
         Tab(0).Control(6)=   "dtFecha1(0)"
         Tab(0).Control(7)=   "Borde(4)"
         Tab(0).Control(8)=   "lProd(4)"
         Tab(0).Control(9)=   "Borde(11)"
         Tab(0).Control(10)=   "Borde(1)"
         Tab(0).Control(11)=   "Borde(2)"
         Tab(0).Control(12)=   "lProd(11)"
         Tab(0).Control(13)=   "lProd(1)"
         Tab(0).Control(14)=   "lProd(3)"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Paso 2: Productos del pedido"
         TabPicture(1)   =   "CAT_Pedidos.frx":1680
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Shape1(2)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Shape1(7)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Borde(6)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Borde(5)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Borde(0)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lInfo(2)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Shape1(6)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lProd(16)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lProd(7)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lInfo(3)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Shape1(0)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "ListaProd(0)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "ListaProd(1)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtProd(5)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtProd(1)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtProd(0)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Check1(1)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "cmBoton(3)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Check1(0)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "cmBoton(7)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "cmBoton(8)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "cmBoton(9)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "cmBoton(10)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "cmBoton(4)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "cmBoton(11)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).ControlCount=   25
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
            Index           =   11
            Left            =   15840
            Picture         =   "CAT_Pedidos.frx":169C
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   720
            Width           =   2055
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
            Index           =   5
            Left            =   -67320
            Picture         =   "CAT_Pedidos.frx":1F66
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   5760
            Width           =   2055
         End
         Begin VB.CommandButton cmBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar / Finalizar "
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
            Left            =   12840
            Picture         =   "CAT_Pedidos.frx":2830
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   720
            Width           =   2655
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
            Height          =   375
            Index           =   10
            Left            =   8160
            Picture         =   "CAT_Pedidos.frx":30FA
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4920
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
            Height          =   375
            Index           =   9
            Left            =   7560
            Picture         =   "CAT_Pedidos.frx":3684
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   4920
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
            Height          =   375
            Index           =   8
            Left            =   6960
            Picture         =   "CAT_Pedidos.frx":3C0E
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   4920
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
            Height          =   375
            Index           =   7
            Left            =   6360
            Picture         =   "CAT_Pedidos.frx":4198
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   4920
            Width           =   495
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
            Left            =   3840
            TabIndex        =   30
            Top             =   4920
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
            Height          =   375
            Index           =   3
            Left            =   9720
            Picture         =   "CAT_Pedidos.frx":4722
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4920
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
            Index           =   1
            Left            =   9480
            TabIndex        =   26
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtProd 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   420
            Index           =   0
            Left            =   240
            MaxLength       =   30
            TabIndex        =   25
            Text            =   "CLAVE CODIGO"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtProd 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   420
            Index           =   1
            Left            =   6840
            MaxLength       =   30
            TabIndex        =   24
            Text            =   "CLAVE PROVEEDOR"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtProd 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   420
            Index           =   5
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   23
            Text            =   "PRODUCTO"
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar y continuar"
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
            Left            =   -74520
            Picture         =   "CAT_Pedidos.frx":4CAC
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   5760
            Width           =   2655
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
            Height          =   975
            Index           =   4
            Left            =   -74640
            MaxLength       =   3500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   4320
            Width           =   9255
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
            Left            =   -69840
            Picture         =   "CAT_Pedidos.frx":5576
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3000
            UseMaskColor    =   -1  'True
            Width           =   2175
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
            Left            =   -74640
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3240
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
            Left            =   -74640
            MaxLength       =   25
            TabIndex        =   15
            Top             =   1080
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker dtFecha1 
            Height          =   375
            Index           =   0
            Left            =   -74640
            TabIndex        =   17
            Top             =   2280
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   101711873
            CurrentDate     =   40829
         End
         Begin MSFlexGridLib.MSFlexGrid ListaProd 
            Height          =   3015
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   1920
            Width           =   16455
            _ExtentX        =   29025
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   16
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"CAT_Pedidos.frx":5B00
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
         Begin MSFlexGridLib.MSFlexGrid ListaProd 
            Height          =   3615
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   5640
            Width           =   16575
            _ExtentX        =   29236
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   16
            FixedCols       =   0
            AllowUserResizing=   1
            FormatString    =   $"CAT_Pedidos.frx":5C36
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
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   0
            Left            =   120
            Top             =   5400
            Width           =   14535
         End
         Begin VB.Label lInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Productos en pedido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   11400
            TabIndex        =   36
            Top             =   5040
            Width           =   4455
         End
         Begin VB.Label lProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de productos para pedido"
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
            Index           =   7
            Left            =   240
            TabIndex        =   35
            Top             =   5085
            Width           =   3975
         End
         Begin VB.Label lProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de productos en inventario"
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
            Left            =   240
            TabIndex        =   28
            Top             =   1440
            Width           =   3975
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   6
            Left            =   120
            Top             =   1680
            Width           =   12375
         End
         Begin VB.Label lInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Productos en inventario:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   7320
            TabIndex        =   27
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   435
            Index           =   0
            Left            =   240
            Top             =   720
            Width           =   2205
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   435
            Index           =   5
            Left            =   6840
            Top             =   720
            Width           =   2205
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   435
            Index           =   6
            Left            =   2640
            Top             =   720
            Width           =   4005
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   975
            Index           =   4
            Left            =   -74640
            Top             =   4320
            Width           =   9285
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
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -74640
            TabIndex        =   21
            Top             =   3960
            Width           =   3135
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   375
            Index           =   11
            Left            =   -74640
            Top             =   3240
            Width           =   4605
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   375
            Index           =   1
            Left            =   -74640
            Top             =   2280
            Width           =   2205
         End
         Begin VB.Shape Borde 
            BorderColor     =   &H000080FF&
            BorderWidth     =   4
            Height          =   435
            Index           =   2
            Left            =   -74640
            Top             =   1080
            Width           =   2565
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
            Left            =   -74640
            TabIndex        =   14
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label lProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha del pedido"
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
            TabIndex        =   13
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label lProd 
            BackStyle       =   0  'Transparent
            Caption         =   "Código/Clave del pedido"
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
            TabIndex        =   12
            Top             =   720
            Width           =   3135
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   300
            Index           =   7
            Left            =   7200
            Top             =   1320
            Width           =   4935
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   300
            Index           =   2
            Left            =   11160
            Top             =   5040
            Width           =   4935
         End
      End
      Begin VB.Timer timeCarga 
         Interval        =   200
         Left            =   -74160
         Top             =   240
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agregar a la lista"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   1
         Left            =   13800
         Picture         =   "CAT_Pedidos.frx":5D6C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   -5000
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtProd 
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
         Index           =   2
         Left            =   11040
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "0"
         Top             =   -5000
         Width           =   1935
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   13080
         Picture         =   "CAT_Pedidos.frx":62F6
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   -5000
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid ListaPed 
         Height          =   3135
         Index           =   0
         Left            =   -74880
         TabIndex        =   6
         Top             =   1320
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"CAT_Pedidos.frx":6880
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
      Begin MSFlexGridLib.MSFlexGrid ListaPed 
         Height          =   5055
         Index           =   1
         Left            =   -74880
         TabIndex        =   7
         Top             =   4920
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"CAT_Pedidos.frx":69C1
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
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   5
         Left            =   -74880
         Top             =   4800
         Width           =   12375
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   4
         Left            =   -74880
         Top             =   1200
         Width           =   12375
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de productos del pedido"
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
         Index           =   2
         Left            =   -74880
         TabIndex        =   10
         Top             =   4560
         Width           =   3975
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de pedidos"
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
         Left            =   -74880
         TabIndex        =   9
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Productos en pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -63600
         TabIndex        =   8
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos del producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2655
         Index           =   0
         Left            =   12480
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Index           =   0
         Left            =   15120
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   1
         Left            =   15120
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   -63720
         Top             =   4560
         Width           =   4935
      End
   End
   Begin VB.Menu mn_Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Nuevo 
         Caption         =   "Nuevo pedido"
      End
      Begin VB.Menu mn_Editar1 
         Caption         =   "Editar datos generales del pedido"
      End
      Begin VB.Menu mn_Editar2 
         Caption         =   "Editar detalle del pedido"
      End
      Begin VB.Menu mn_LineaOpciones1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Cerrar 
         Caption         =   "Cerrar pedido"
      End
      Begin VB.Menu mn_Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CAT_Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim res1 As Recordset
Dim sql1 As String
Dim tipoPedido As String
Dim valDatos As Boolean
Dim campo As String
Dim prodId As String
Dim pedId As Long


Private Sub checkProducto()

    'On Error Resume Next
                   
               
    sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, PROD_DEPENDIENTE " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "T1.PROD_ID = '" & ListaProd(1).TextMatrix(ListaProd(1).Row, 1) & "' "
    Set res1 = con.Execute(sql1)
    
    If Not res1.EOF Then
        lInfo(0).Caption = "Datos del producto: " & vbCrLf & vbCrLf & "CODIGO:    " & res1.Fields("PROD_CODIGO") & vbCrLf & "PRODUCTO:    " & res1.Fields("PROD_NOMBRE") & vbCrLf & _
        "MARCA:   " & res1.Fields("CTMR_MARCA") & vbCrLf & "ESTATUS:   " & res1.Fields("STATUS") & vbCrLf & _
        "PRECIO:   " & FormatCurrency(res1.Fields("PROD_PRECIO")) & "  CANTIDAD:  " & res1.Fields("PROD_CANT") & vbCrLf & _
        "TIPO: " & res1.Fields("CTPT_TIPO") & vbCrLf & _
        "PRESENTACIÓN: " & res1.Fields("PROD_PRESENTACION") '& " " & RES1.Fields("PROD_UNIMED_PRESENT")
        
        If IsNull(res1.Fields("PROD_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write res1.Fields("PROD_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
        Else
            imgFoto(0).Picture = LoadPicture("")
        End If
        
    Else
        imgFoto(0).Picture = LoadPicture("")
        lInfo(0).Caption = "Datos del Producto:"
        
        MsgBox "No se ha encontrado información con el valor proporcionado.", vbInformation
    End If
    
    
End Sub

Private Sub Check1_Click(Index As Integer)

    ListaProd(Index).Redraw = False
    If Check1(Index).value = Unchecked Then
        For b1 = 1 To ListaProd(Index).Rows - 1
            ListaProd(Index).Col = 0
            ListaProd(Index).Row = b1
            ListaProd(Index).TextMatrix(b1, 0) = Chr(168)
        Next b1
    Else
        For b1 = 1 To ListaProd(Index).Rows - 1
            ListaProd(Index).Col = 0
            ListaProd(Index).Row = b1
            ListaProd(Index).TextMatrix(b1, 0) = Chr(254)
        Next b1
    End If
    ListaProd(Index).Redraw = True
    
End Sub

Private Sub cmBoton_Click(Index As Integer)
    If Index = 0 Then
        If tipoPedido = "LIBRE" Or tipoPedido = "EDITAR1" Then
            generarPedido (tipoPedido)
        Else
            If tipoPedido = "GUARDADO" Then
                generaPedidoDetalle
            Else
                If tipoPedido = "EDITAR2" Then
                    ''''falta para guardar la edición del detalle
                    '''Checar si aqui o en al momento de cambiar el numero en detalle
                End If
            End If
        End If
    Else
        If Index = 3 Then
            ques = MsgBox("'¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        Else
            If Index = 1 Then
                If Val(txtProd(2).Text) > 0 Then
                    If txtProd(1).Text <> "" Then
                        Call anexarProducto("CODIGO_PROV", "1")
                    Else
                        If txtProd(0).Text <> "" Then
                            Call anexarProducto("CODIGO", "0")
                        Else
                            MsgBox "Debe seleccionar un producto. ", vbInformation
                        End If
                    End If
                Else
                    MsgBox "No se puede agregar a la lista sin un cantidad mayor a cero.", vbInformation
                End If
            Else
                If Index = 7 Then
                    If tipoPedido = "GUARDADO" Then
                        pasarLista ("SEL")
                    Else
                        MsgBox "Debe guardar primero la información del pedido.", vbInformation
                    End If
                Else
                    If Index = 9 Then
                        If tipoPedido = "GUARDADO" Then
                            pasarLista ("TODO")
                        End If
                    Else
                        If Index = 10 Then
                            regresarLista ("TODO")
                        Else
                            If Index = 8 Then
                                regresarLista ("SEL")
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub regresarLista(tipo As String)
'On Error Resume Next
Dim texto As String
Dim ques As String
Dim num As Long

    
    texto = ""
    
    If tipo = "SEL" Then
        For b1 = 1 To ListaProd(0).Rows - 1
            If ListaProd(0).TextMatrix(b1, 0) = Chr(254) Then
                texto = texto & vbCrLf & "Código: " & ListaProd(0).TextMatrix(b1, 1) & "  Producto: " & ListaProd(0).TextMatrix(b1, 2)
            End If
        Next b1
        
        If texto <> "" Then
            ques = MsgBox("Los siguientes productos serán eliminados de la lista ¿Continuar? " & vbCrLf & vbclrf & texto, vbYesNo + vbInformation)
            If ques = vbYes Then
                num = 0
                ListaProd(0).Redraw = False
                For b1 = 1 To ListaProd(0).Rows - 1
                    num = num + 1
                    If ListaProd(0).TextMatrix(num, 0) = Chr(254) Then
                        If ListaProd(0).Rows > 2 Then
                            ListaProd(0).RemoveItem (num)
                            num = num - 1
                        Else
                            ListaProd(0).Rows = 1
                            'b1 = 1
                        End If
                    End If
                Next b1
            End If
            ListaProd(0).Redraw = True
        
        Else
            MsgBox "No se encontraron elementos. Verifique.", vbInformation
        End If
    Else
        If tipo = "TODO" Then
            ques = MsgBox("Se eliminarán " & ListaProd(0).Rows - 1 & " de la lista ¿Continuar? ", vbYesNo + vbInformation)
            If ques = vbYes Then
                ListaProd(0).Rows = 1
            End If
        End If
    End If
End Sub
Private Sub pasarLista(tipo As String)

Dim texto As String
Dim encontro As Boolean
Dim tipoValor As String


    ListaProd(0).Redraw = False
    texto = ""
    For b1 = 1 To ListaProd(1).Rows - 1
        
        If tipo = "SEL" Then
            tipoValor = Chr(254)
        Else
            If tipo = "TODO" Then
                tipoValor = ListaProd(1).TextMatrix(b1, 0)
            End If
        End If
        
        If ListaProd(1).TextMatrix(b1, 0) = tipoValor Then
            encontro = False
        
            For c1 = 1 To ListaProd(0).Rows - 1
                If ListaProd(0).TextMatrix(c1, 1) = ListaProd(1).TextMatrix(b1, 1) Then
                    texto = texto & vbCrLf & "Código: " & ListaProd(0).TextMatrix(c1, 1) & "  Producto: " & ListaProd(0).TextMatrix(c1, 2)
                    encontro = True
                    Exit For
                End If
            Next c1
        
            If encontro = False Then
                ListaProd(0).AddItem ""
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 0) = ListaProd(1).TextMatrix(b1, 0)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 1) = ListaProd(1).TextMatrix(b1, 1)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 2) = ListaProd(1).TextMatrix(b1, 2)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 3) = ListaProd(1).TextMatrix(b1, 3)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 4) = ListaProd(1).TextMatrix(b1, 4)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 5) = ListaProd(1).TextMatrix(b1, 5)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 6) = ListaProd(1).TextMatrix(b1, 6)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 7) = ListaProd(1).TextMatrix(b1, 7)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 8) = ListaProd(1).TextMatrix(b1, 8)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 9) = ListaProd(1).TextMatrix(b1, 9)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 10) = ListaProd(1).TextMatrix(b1, 10)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 11) = ListaProd(1).TextMatrix(b1, 11)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 12) = ListaProd(1).TextMatrix(b1, 12)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 13) = ListaProd(1).TextMatrix(b1, 13)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 14) = ListaProd(1).TextMatrix(b1, 14)
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 15) = ListaProd(1).TextMatrix(b1, 15)
                        
                ListaProd(0).Row = ListaProd(0).Rows - 1
                ListaProd(0).Col = 7
                ListaProd(0).CellForeColor = &H80&
                ListaProd(0).CellFontSize = 12
                ListaProd(0).CellFontBold = True
                ListaProd(0).ColAlignment(7) = (4)
                
                ListaProd(0).Col = 6
                ListaProd(0).CellForeColor = &H4000&
                ListaProd(0).CellFontSize = 12
                ListaProd(0).CellFontBold = True
                ListaProd(0).ColAlignment(6) = (4)
                                    
                If Val(ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 6)) = 0 Then
                    cantAgotados = cantAgotados + 1
                    ListaProd(0).Col = 6
                    ListaProd(0).CellForeColor = &HC0&
                    ListaProd(0).CellFontSize = 12
                    ListaProd(0).CellFontBold = True
                    ListaProd(0).ColAlignment(6) = (4)
                Else
                    If Val(ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 6)) <= Val(ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 10)) Then
                        ListaProd(0).Col = 6
                        ListaProd(0).CellForeColor = &H800080
                        ListaProd(0).CellFontSize = 12
                        ListaProd(0).CellFontBold = True
                        ListaProd(0).ColAlignment(6) = (4)
                    Else
                        ListaProd(0).Col = 6
                        ListaProd(0).CellForeColor = &H0&
                        ListaProd(0).CellFontSize = 12
                        ListaProd(0).CellFontBold = True
                        ListaProd(0).ColAlignment(6) = (4)
                    End If
                End If
            
            
                ListaProd(0).Row = ListaProd(0).Rows - 1
                ListaProd(0).Col = 0
                ListaProd(0).CellFontName = "Wingdings"
                ListaProd(0).CellFontBold = True
                ListaProd(0).CellFontSize = 16
                ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 0) = Chr(254)
                
                
                
            End If
        End If
    Next b1
    
ListaProd(0).Redraw = True
lInfo(3).Caption = "Productos en pedido: " & ListaProd(0).Rows - 1
    If texto <> "" Then
        MsgBox "Los siguientes productos ya se encuentran en la lista y no se consideraron: " & vbCrLf & vbclrf & texto, vbInformation
    End If


End Sub

Private Sub generaPedidoDetalle()
    With ListaProd(0)
        If .Rows > 1 Then
            For b1 = 1 To .Rows - 1
                sql1 = "INSERT INTO PEDIDOS_DETALLE (ped_CtPedId, ped_ProdId, ped_ProdServ, ped_Cantidad, ped_CantidadCierre) VALUES " & _
                "('" & pedId & "', '" & .TextMatrix(b1, 1) & "', 'P', '" & .TextMatrix(b1, 7) & "', '" & .TextMatrix(b1, 7) & "' )"
                con.Execute (sql1)
            Next b1
        End If
    End With
    
    cancelar
    MsgBox "Informacón de pedido guardada.", vbInformation
    cargaPedidos
    SSTab1.Tab = 0

End Sub
Private Sub cancelar()
    pedId = 0
    tipoPedido = "LIBRE"
    ListaProd(0).Rows = 1
    
    For b1 = 0 To 4
        txtProd(b1).Text = ""
    Next b1
    txtProd(2).Text = "0"
    
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    
    cmBoton(3).Visible = False
    
End Sub
Private Sub anexarProducto(prodCodigo As String, Numero As Integer)

    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE  SUBTIPO = 'PRODUCTO' AND " & prodCodigo & " = '" & txtProd(Numero).Text & "' "
    'MsgBox SQL1
    Set res1 = con.Execute(sql1)
        
    ListaProd(0).Redraw = False
    'ListaProd(0).Rows = 1
    Do While Not res1.EOF
        ListaProd(0).AddItem ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 1) = res1.Fields("PROD_ID")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 2) = res1.Fields("CODIGO")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 3) = res1.Fields("NOMBRE")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 4) = res1.Fields("TIPO")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 5) = res1.Fields("MARCA")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 6) = res1.Fields("CANTIDAD")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 7) = txtProd(2).Text
        
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 8) = FormatCurrency(res1.Fields("PRECIO_VENTA"))
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 9) = res1.Fields("STATUS")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 10) = res1.Fields("STOCK_MIN")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 11) = res1.Fields("STOCK_MAX")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 12) = res1.Fields("PRESENTACION") & ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 13) = res1.Fields("UNIDAD_mEDIDA")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 14) = res1.Fields("PROVEEDOR") & ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 15) = FormatCurrency(res1.Fields("PRECIO_COSTO"))
                    
        ListaProd(0).Row = ListaProd(0).Rows - 1
        ListaProd(0).Col = 7
        'ListaProd(0).CellForeColor = &HC000&
        ListaProd(0).CellFontSize = 10
        'ListaProd(0).CellAlignment = 0
        
                    
        If res1.Fields("ID_STATUS") = "I" Or res1.Fields("CANTIDAD") = 0 Then
            cantAgotados = cantAgotados + 1
            ListaProd(0).Row = ListaProd(0).Rows - 1
            For b1 = 0 To ListaProd(0).Cols - 1
                ListaProd(0).Col = b1
                ListaProd(0).CellForeColor = vbRed
            Next b1
        Else
            If res1.Fields("CANTIDAD") <= res1.Fields("STOCK_MIN") Then
                ListaProd(0).Row = ListaProd(0).Rows - 1
                For b1 = 0 To ListaProd(0).Cols - 1
                    ListaProd(0).Col = b1
                    ListaProd(0).CellForeColor = &HC0C0&
                Next b1
            End If
        End If
        
        res1.MoveNext
    Loop
    lInfo(3).Caption = "Productos en pedido: " & ListaProd(0).Rows - 1
    ListaProd(0).Redraw = True
    
    txtProd(0).Text = ""
    txtProd(1).Text = ""
    txtProd(2).Text = "0"
    

End Sub
Private Sub generarPedido(tipo As String)


    valDatos = True
    checkValores
    If valDatos = False Then
        MsgBox "Verifique la información en el campo: " & campo, vbInformation
        'tipoPedido = "LIBRE"
        Exit Sub
    Else
        
    Dim prodPerId As String
    Dim prodTipoId As String
    Dim prodTipo As String
    
    prodPerId = "NULL"
    prodTipoId = "NULL"
    prodTipo = "NULL"
        
    If cmbProd(3).Text <> "" Then
        sql1 = "SELECT PERTP_TIPO_ID, PERTP_PER_ID, PERTP_PER_TIPO FROM PER_TIPO WHERE PERTP_PER_ID = '" & cmbProd(3).ItemData(cmbProd(3).ListIndex) & "' "
        Set res1 = con.Execute(sql1)
        
        If Not res1.EOF Then
            prodPerId = res1.Fields("PERTP_PER_ID")
            prodTipoId = res1.Fields("PERTP_TIPO_ID")
            prodTipo = res1.Fields("PERTP_PER_TIPO")
        End If
    End If
        
    If tipo = "LIBRE" Then
        sql1 = "INSERT INTO CAT_PEDIDOS (ctPed_ProvPerId, ctPed_ProvTipoId, ctPed_ProvTipo, ctPed_ClaveFactura, ctPed_Descripcion, " & _
        "ctPed_UserPerId, ctPed_UserTipoId, ctPed_UserTipo, ctPed_FechaCreacion, ctPed_FechaPedido, ctPed_Status) VALUES " & _
        "('" & prodPerId & "', '" & prodTipoId & "', '" & prodTipo & "', '" & txtProd(3).Text & "', '" & txtProd(4).Text & "',   " & _
        " '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', now(), '" & Format(dtFecha1(0), "yyyy-MM-dd") & "', 'G')"
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() pedId"
        Set resMaxId = con.Execute(sql1)
        If Not resMaxId.EOF Then
            pedId = resMaxId.Fields("pedId")
        End If
           
        MsgBox "Datos principales del pedido guardados." & vbCrLf & vbCrLf & "Puede anexar los productos al pedido.", vbInformation
    
        desabilitado ("True")
        
        tipoPedido = "GUARDADO"
    Else
        If tipo = "EDITAR1" Then
            sql1 = "UPDATE CAT_PEDIDOS SET ctPed_ProvPerId = '" & prodPerId & "', ctPed_ProvTipoId = '" & prodTipoId & "', ctPed_ProvTipo = '" & prodTipo & "',  " & _
            "ctPed_ClaveFactura = '" & txtProd(3).Text & "', ctPed_Descripcion = '" & txtProd(4).Text & "', ctPed_FechaPedido = '" & Format(dtFecha1(0), "yyyy-MM-dd") & "'"
            con.Execute (sql1)
            
            cancelar
            cargaPedidos
            SSTab1.Tab = 0
            MsgBox "Edición realizada. Verifique.", vbInformation
            
        End If
    End If
    
    End If
    
End Sub
Private Sub cargaProveedor()

    sql1 = "SELECT PER_ID, CONCAT(PER_ALIAS, ' - ', PER_NOMBRE, ' ', PER_PATERNO, ' ', PER_MATERNO) PROVEEDOR " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'V'  "
    Set res1 = con.Execute(sql1)
    
    cmbProd(3).Clear
    Do While Not res1.EOF
        cmbProd(3).AddItem res1.Fields("PROVEEDOR")
        cmbProd(3).ItemData(cmbProd(3).ListCount - 1) = res1.Fields("PER_ID")
        res1.MoveNext
    Loop
    If cmbProd(3).ListCount > 0 Then
        cmbProd(3).ListIndex = 0
    End If

End Sub

Private Sub checkValores()
    If txtProd(3).Text = "" Then
        valDatos = False
    Else
        If dtFecha1(0).value < Date Then
            If tipoPedido = "EDITAR1" Then
            Else
                valDatos = False
            End If
        Else
            If txtProd(4).Text = "" Then
                valDatos = False
            End If
        End If
    End If
    
End Sub

Private Sub cmbProd_GotFocus(Index As Integer)
    cmBoton(6).Visible = True
End Sub

Private Sub cmbProd_LostFocus(Index As Integer)
    cmBoton(6).Visible = False
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    cargaInicial
    cargaPedidos
    desabilitado ("False")
    
End Sub
Private Sub cargaPedidos()
    sql1 = "SELECT * FROM VIEW_CAT_PEDIDOS ORDER BY CREADO DESC"
    Set res1 = con.Execute(sql1)
    ListaPed(1).Rows = 1
    ListaPed(0).Rows = 1
    ListaPed(0).Redraw = False
    Do While Not res1.EOF
        ListaPed(0).AddItem ""
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 0) = res1.Fields("CLAVE")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 1) = res1.Fields("PROVEEDOR")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 2) = res1.Fields("FOLIO_NOTA")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 3) = res1.Fields("FECHA_PEDIDO")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 4) = res1.Fields("STATUS")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 5) = res1.Fields("USUARIO")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 6) = res1.Fields("CREADO")
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 7) = res1.Fields("USUARIO_CIERRE") & ""
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 8) = res1.Fields("FECHA_CIERRE") & ""
        ListaPed(0).TextMatrix(ListaPed(0).Rows - 1, 9) = res1.Fields("DESCRIPCION") & ""
        
        res1.MoveNext
    Loop
    ListaPed(0).Redraw = True
    
    lInfo(4).Caption = "Productos en pedido: " & ListaPed(1).Rows - 1
End Sub
Private Sub desabilitado(valor As String)
'On Error Resume Next
'
'    For b1 = 0 To 2
'        txtProd(b1).Enabled = valor
'        If b1 > 0 Then
'            cmBoton(b1).Enabled = valor
'        End If
'    Next b1


End Sub

Private Sub cargaInicial()
    cmBoton(6).Visible = False
    ListaProd(0).Rows = 1
    valDatos = False
    dtFecha1(0).value = Date
    cargaProveedor
    tipoPedido = "LIBRE"
    cargaProductos
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    cmBoton(3).Visible = False
    
End Sub
Private Sub cargaProductos()

    Dim claveProd, prodNom, claveProv As String
    
    Check1(1).value = Checked
    
    If txtProd(0).Text = "CLAVE CODIGO" Then
        claveProd = ""
    Else
        claveProd = txtProd(0).Text
    End If
    
    If txtProd(5).Text = "PRODUCTO" Then
        prodNom = ""
    Else
        prodNom = txtProd(5).Text
    End If
    
    If txtProd(1).Text = "CLAVE PROVEEDOR" Then
        claveProv = ""
    Else
        claveProv = txtProd(1).Text
    End If
    
    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE  SUBTIPO = 'PRODUCTO' " & _
    "AND CODIGO LIKE '%" & claveProd & "%' " & _
    "AND upper(NOMBRE) LIKE upper('%" & prodNom & "%') " & _
    "AND CODIGO_PROV LIKE '%" & claveProv & "%' ORDER BY NOMBRE ASC"
    'MsgBox SQL1
    Set res1 = con.Execute(sql1)
        
    ListaProd(1).Redraw = False
    ListaProd(1).Rows = 1
    ListaProd(1).ColWidth(7) = 0
    Do While Not res1.EOF
        ListaProd(1).AddItem ""
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 1) = res1.Fields("PROD_ID")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 2) = res1.Fields("CODIGO")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 3) = res1.Fields("NOMBRE")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 4) = res1.Fields("TIPO")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 5) = res1.Fields("MARCA")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 6) = res1.Fields("CANTIDAD")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 7) = "0"
        
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 8) = FormatCurrency(res1.Fields("PRECIO_VENTA"))
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 9) = res1.Fields("STATUS")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 10) = res1.Fields("STOCK_MIN")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 11) = res1.Fields("STOCK_MAX")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 12) = res1.Fields("PRESENTACION") & ""
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 13) = res1.Fields("UNIDAD_mEDIDA")
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 14) = res1.Fields("PROVEEDOR") & ""
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 15) = FormatCurrency(res1.Fields("PRECIO_COSTO"))
                    
        ListaProd(1).Row = ListaProd(1).Rows - 1
'        ListaProd(1).Col = 8
'        ListaProd(1).CellForeColor = &H80&
'        ListaProd(1).CellFontSize = 12
'        ListaProd(1).CellFontBold = True
        
        ListaProd(1).Col = 6
        ListaProd(1).CellForeColor = &H4000&
        ListaProd(1).CellFontSize = 12
        ListaProd(1).CellFontBold = True
        ListaProd(1).ColAlignment(6) = (4)
        
                    
        ListaProd(1).Row = ListaProd(1).Rows - 1
        ListaProd(1).Col = 0
        ListaProd(1).CellFontName = "Wingdings"
        ListaProd(1).CellFontBold = True
        ListaProd(1).CellFontSize = 16
        ListaProd(1).TextMatrix(ListaProd(1).Rows - 1, 0) = Chr(254)
                    
        If res1.Fields("ID_STATUS") = "I" Or res1.Fields("CANTIDAD") = 0 Then
            cantAgotados = cantAgotados + 1
            ListaProd(1).Row = ListaProd(1).Rows - 1
            For b1 = 0 To ListaProd(1).Cols - 1
                ListaProd(1).Col = b1
                ListaProd(1).CellForeColor = &HC0&
            Next b1
        Else
            If res1.Fields("CANTIDAD") <= res1.Fields("STOCK_MIN") Then
                ListaProd(1).Row = ListaProd(1).Rows - 1
                For b1 = 0 To ListaProd(1).Cols - 1
                    ListaProd(1).Col = b1
                    ListaProd(1).CellForeColor = &H800080
                Next b1
            End If
        End If
        
        res1.MoveNext
    Loop
    lInfo(2).Caption = "Productos en inventario: " & ListaProd(1).Rows - 1
    ListaProd(1).Redraw = True

    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Dim a As String
        
        a = MsgBox("¿Salir?", vbQuestion + vbYesNo)
        If a = vbYes Then
            If tipoPedido = "GUARDADO" Then
                a = MsgBox("¿Salir sin guardar?. " & vbCrLf & vbCrLf & "Se perderá la información anexada a la lista.", vbQuestion + vbYesNo)
                If a = vbYes Then
                    Cancel = 0
                Else
                    Cancel = 1
                End If
            End If
        Else
            Cancel = 1
        End If

End Sub

Private Sub ListaPed_Click(Index As Integer)
    
    If Index = 0 Then
        cargaDetalle (ListaPed(0).TextMatrix(ListaPed(0).Row, 0))
    End If

End Sub
Private Sub cargaDetalle(pedido As String)

    ListaPed(1).Rows = 1

    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO T1, PEDIDOS_DETALLE T2 " & _
    "WHERE  T1.SUBTIPO = 'PRODUCTO' AND T2.PED_CTPEDID = '" & pedido & "' " & _
    "AND PED_PRODID = T1.PROD_ID ORDER BY NOMBRE ASC "
    
    Set res1 = con.Execute(sql1)

    
    ListaPed(1).Redraw = False
    ListaPed(1).Rows = 1
    'ListaPed(1).ColWidth(7) = 0
    Do While Not res1.EOF
        ListaPed(1).AddItem ""
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 0) = res1.Fields("PROD_ID")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 1) = res1.Fields("CODIGO")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 2) = res1.Fields("NOMBRE")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 3) = res1.Fields("TIPO")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 4) = res1.Fields("MARCA")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 5) = res1.Fields("CANTIDAD")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 6) = res1.Fields("PED_CANTIDAD")
        
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 7) = res1.Fields("PED_CANTIDAD")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 8) = FormatCurrency(res1.Fields("PRECIO_VENTA"))
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 9) = res1.Fields("STATUS")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 10) = res1.Fields("STOCK_MIN")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 11) = res1.Fields("STOCK_MAX")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 12) = res1.Fields("PRESENTACION") & ""
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 13) = res1.Fields("UNIDAD_mEDIDA")
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 14) = res1.Fields("PROVEEDOR") & ""
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 15) = FormatCurrency(res1.Fields("PRECIO_COSTO"))
        ListaPed(1).TextMatrix(ListaPed(1).Rows - 1, 16) = res1.Fields("PED_CTPEDID") & ""
                    
        ListaPed(1).Row = ListaPed(1).Rows - 1
        ListaPed(1).Col = 7
        ListaPed(1).CellForeColor = &H80&
        ListaPed(1).CellFontSize = 12
        ListaPed(1).CellFontBold = True
        ListaPed(1).ColAlignment(7) = (4)

        
        ListaPed(1).Col = 6
        ListaPed(1).CellForeColor = &H4000&
        ListaPed(1).CellFontSize = 12
        ListaPed(1).CellFontBold = True
        ListaPed(1).ColAlignment(6) = (4)
                            
        If res1.Fields("ID_STATUS") = "I" Or res1.Fields("CANTIDAD") = 0 Then
            cantAgotados = cantAgotados + 1
            ListaPed(1).Col = 5
            ListaPed(1).CellForeColor = &HC0&
            ListaPed(1).CellFontSize = 12
            ListaPed(1).CellFontBold = True
            ListaPed(1).ColAlignment(5) = (4)
        
        Else
            If res1.Fields("CANTIDAD") <= res1.Fields("STOCK_MIN") Then
                ListaPed(1).Col = 5
                ListaPed(1).CellForeColor = &H800080
                ListaPed(1).CellFontSize = 12
                ListaPed(1).CellFontBold = True
                ListaPed(1).ColAlignment(5) = (4)
            Else
                ListaPed(1).Col = 5
                ListaPed(1).CellForeColor = &H0&
                ListaPed(1).CellFontSize = 12
                ListaPed(1).CellFontBold = True
                ListaPed(1).ColAlignment(5) = (4)
            End If
        End If
        
        res1.MoveNext
    Loop

    lInfo(4).Caption = "Productos en pedido: " & ListaPed(1).Rows - 1
    ListaPed(1).Redraw = True

    


End Sub

Private Sub cargaDetalleProdEdit(pedido As String)

    ListaProd(0).Rows = 1

    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO T1, PEDIDOS_DETALLE T2 " & _
    "WHERE  T1.SUBTIPO = 'PRODUCTO' AND T2.PED_CTPEDID = '" & pedido & "' " & _
    "AND PED_PRODID = T1.PROD_ID ORDER BY NOMBRE ASC "
    
    Set res1 = con.Execute(sql1)

    
    ListaProd(0).Redraw = False
    ListaProd(0).Rows = 1
    'listaprod(0).ColWidth(7) = 0
    Do While Not res1.EOF
        ListaProd(0).AddItem ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 1) = res1.Fields("PROD_ID")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 2) = res1.Fields("CODIGO")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 3) = res1.Fields("NOMBRE")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 4) = res1.Fields("TIPO")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 5) = res1.Fields("MARCA")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 6) = res1.Fields("CANTIDAD")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 7) = res1.Fields("PED_CANTIDAD")
    
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 8) = FormatCurrency(res1.Fields("PRECIO_VENTA"))
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 9) = res1.Fields("STATUS")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 10) = res1.Fields("STOCK_MIN")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 11) = res1.Fields("STOCK_MAX")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 12) = res1.Fields("PRESENTACION") & ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 13) = res1.Fields("UNIDAD_mEDIDA")
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 14) = res1.Fields("PROVEEDOR") & ""
        ListaProd(0).TextMatrix(ListaProd(0).Rows - 1, 15) = FormatCurrency(res1.Fields("PRECIO_COSTO"))
                    
        ListaProd(0).Row = ListaProd(0).Rows - 1
        ListaProd(0).Col = 6
        ListaProd(0).CellForeColor = &H80&
        ListaProd(0).CellFontSize = 12
        ListaProd(0).CellFontBold = True
        ListaProd(0).ColAlignment(7) = (4)

                                    
        If res1.Fields("ID_STATUS") = "I" Or res1.Fields("CANTIDAD") = 0 Then
            cantAgotados = cantAgotados + 1
            ListaProd(0).Col = 5
            ListaProd(0).CellForeColor = &HC0&
            ListaProd(0).CellFontSize = 12
            ListaProd(0).CellFontBold = True
            ListaProd(0).ColAlignment(5) = (4)
        
        Else
            If res1.Fields("CANTIDAD") <= res1.Fields("STOCK_MIN") Then
                ListaProd(0).Col = 5
                ListaProd(0).CellForeColor = &H800080
                ListaProd(0).CellFontSize = 12
                ListaProd(0).CellFontBold = True
                ListaProd(0).ColAlignment(5) = (4)
            Else
                ListaProd(0).Col = 5
                ListaProd(0).CellForeColor = &H0&
                ListaProd(0).CellFontSize = 12
                ListaProd(0).CellFontBold = True
                ListaProd(0).ColAlignment(5) = (4)
            End If
        End If
        
        res1.MoveNext
    Loop

    lInfo(3).Caption = "Productos en pedido: " & ListaProd(0).Rows - 1
    ListaProd(0).Redraw = True

    
End Sub


Private Sub ListaPed_GotFocus(Index As Integer)
    ConScroll ListaPed(Index)
End Sub

Private Sub ListaPed_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim valor As Long
    If Index = 1 Then
        If ListaPed(Index).Col = 7 Then
            valor = ListaPed(Index).TextMatrix(ListaPed(Index).Row, 7)
            If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Then
                ListaPed(Index).Text = ListaPed(Index).Text & Chr(KeyAscii)
                ListaPed(Index).Text = Val(ListaPed(Index).Text)
                
                sql1 = "SELECT CANTIDAD, STOCK_MAX FROM VIEW_PRODUCTOS_INVENTARIO WHERE CODIGO = '" & ListaPed(Index).TextMatrix(ListaPed(Index).Row, 1) & "'"
                'MsgBox SQL1
                Set res1 = con.Execute(sql1)
                
                If Not res1.EOF Then
                    If Val(ListaPed(Index).TextMatrix(ListaPed(Index).Row, 7)) > Val(res1.Fields("STOCK_MAX")) And Val(res1.Fields("STOCK_MAX")) > 0 Then
                        MsgBox "No se puede agregar una cantidad mayor al establecido como máxim. Verifique", vbInformation
                        ListaPed(Index).TextMatrix(ListaPed(Index).Row, 7) = valor
                    Else
                        sql1 = "UPDATE PEDIDOS_DETALLE SET PED_CANTIDADCIERRE = '" & ListaPed(Index).TextMatrix(ListaPed(Index).Row, 7) & "' " & _
                        "WHERE PED_CTPEDID = '" & ListaPed(Index).TextMatrix(ListaPed(Index).Row, 16) & "' AND PED_CTPEDID = PED_PRODID = '" & ListaPed(Index).TextMatrix(ListaPed(Index).Row, 0) & "' "
                        con.Execute (sql1)
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub ListaPed_LostFocus(Index As Integer)
    SinScroll ListaPed(Index)
End Sub

Private Sub ListaPed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
        If ListaPed(0).Rows > 1 Then
            If Button = vbRightButton Then
                ListaPed_Click (0)
                If ListaPed(0).TextMatrix(ListaPed(0).Row, 4) = "GENERADO" Then
                    mn_Cerrar.Enabled = True
                    mn_Editar1.Enabled = True
                    mn_Editar2.Enabled = True
                    
                Else
                    If ListaPed(0).TextMatrix(ListaPed(0).Row, 4) = "CERRADO" Then
                        mn_Cerrar.Enabled = False
                        mn_Editar1.Enabled = False
                        mn_Editar2.Enabled = False
                    
                    Else
                        mn_Editar1.Enabled = True
                        mn_Editar2.Enabled = True
                        mn_Cerrar.Enabled = True
                    End If
                End If
                'mn_RePag.Caption = "Realizar pago apartado folio: " & Lista1.TextMatrix(Lista1.MouseRow, 0) & " cliente: " & Lista1.TextMatrix(Lista1.MouseRow, 6) & ". Faltante: " & Lista1.TextMatrix(Lista1.Row, 4)
    
                ListaPed(0).Row = ListaPed(0).MouseRow
                PopupMenu mn_Opciones, vbPopupMenuLeftAlign
            End If
        End If
    End If
    
End Sub

Private Sub ListaPed_SelChange(Index As Integer)
    If Index = 1 Then
        ListaPed_Click (1)
    End If
End Sub

Private Sub ListaProd_Click(Index As Integer)
    If Index = 1 Then
        checkProducto
    End If
End Sub

Private Sub ListaProd_DblClick(Index As Integer)
    If ListaProd(Index).MouseRow = 0 Then
        Call ordenarLista(ListaProd(Index))
    Else
        If ListaProd(Index).Col = 0 Then
            Dim b1 As Long
            b1 = ListaProd(Index).Row
            
            ListaProd(Index).Row = b1
            ListaProd(Index).Col = 0
            If ListaProd(Index).TextMatrix(b1, 0) = Chr(168) Then
                ListaProd(Index).TextMatrix(b1, 0) = Chr(254)
            Else
                ListaProd(Index).TextMatrix(b1, 0) = Chr(168)
            End If
        End If
    End If

End Sub

Private Sub ListaProd_GotFocus(Index As Integer)
    ConScroll ListaProd(Index)
End Sub

Private Sub ListaProd_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim valor As Long
    If Index = 0 Then
        If ListaProd(Index).Col = 7 Then
            valor = ListaProd(Index).TextMatrix(ListaProd(Index).Row, 7)
            If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Then
                ListaProd(Index).Text = ListaProd(Index).Text & Chr(KeyAscii)
                ListaProd(Index).Text = Val(ListaProd(Index).Text)
                
                sql1 = "SELECT CANTIDAD FROM VIEW_PRODUCTOS_INVENTARIO WHERE CODIGO = '" & ListaProd(Index).TextMatrix(ListaProd(Index).Row, 1) & "'"
                'MsgBox SQL1
                Set res1 = con.Execute(sql1)
                
                If Not res1.EOF Then
                    If Val(ListaProd(Index).TextMatrix(ListaProd(Index).Row, 7)) > Val(res1.Fields("STOCK_MAX")) And Val(res1.Fields("STOCK_MAX")) > 0 Then
                        MsgBox "No se puede agregar una cantidad mayor al establecido como máxim. Verifique", vbInformation
                        ListaProd(Index).TextMatrix(ListaProd(Index).Row, 7) = valor
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub ListaProd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If ListaProd(Index).Col = 7 Then
            Select Case KeyCode
                Case vbKeyDelete
                    ListaProd(Index).Text = "0"
                Case vbKeyBack
                    If Len(ListaProd(Index).Text) > 0 Then
                        ListaProd(Index).Text = Val(Left(ListaProd(Index).Text, Len(ListaProd(Index).Text) - 1))
                        If ListaProd(Index).Text = "" Then
                            ListaProd(Index).Text = "0"
                        End If
                    End If
            End Select
        End If
    End If

End Sub

Private Sub ListaProd_LostFocus(Index As Integer)
    SinScroll ListaProd(Index)
    
End Sub

Private Sub ListaProd_SelChange(Index As Integer)
    ListaProd_Click (Index)
End Sub

Private Sub mn_Cerrar_Click()
Dim ques As String
    If ListaPed(0).TextMatrix(ListaPed(0).Row, 4) = "GENERADO" Then
        ques = MsgBox("Cerrar el pedido: " & vbCrLf & vbCrLf & "Folio: " & ListaPed(0).TextMatrix(ListaPed(0).Row, 2) & vbCrLf & _
        "Clave: " & ListaPed(0).TextMatrix(ListaPed(0).Row, 0) & vbCrLf & "Fecha de pedido: " & _
        ListaPed(0).TextMatrix(ListaPed(0).Row, 3) & vbCrLf & _
        "Proveedor: " & ListaPed(0).TextMatrix(ListaPed(0).Row, 1), vbYesNo + vbQuestion)
        
        If ques = vbYes Then
            ques = MsgBox("Al cerrar el pedido se actualizará la cantidad de los productos seleccionados." & vbCrLf & vbCrLf & _
            "Esta acción no podrá deshacerse." & vbCrLf & vbCrLf & "¿Continuar con el cierre?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                sql1 = "UPDATE CAT_PEDIDOS SET CTPED_STATUS = 'C', ctPed_fechaCierre = NOW(),  " & _
                "CTPED_USERPERID2 = '" & FRM_Menu.menuBarra2.Panels(7).Text & "', " & _
                "CTPED_USERTIPOID2 =  '" & FRM_Menu.menuBarra2.Panels(8).Text & "', CTPED_USERTIPO2 =  'U' " & _
                "WHERE CTPED_ID = '" & ListaPed(0).TextMatrix(ListaPed(0).Row, 0) & "'"
                con.Execute (sql1)
                
                cargaPedidos
                
                MsgBox "Pedido cerrar. Inventario actualizado. Verfique.", vbInformation
            
            End If
        End If
    End If
End Sub

Private Sub mn_Editar1_Click()
    
    
    Dim ques As String
    
    
    With ListaPed(0)
        ques = MsgBox("Editar datos generales del pedido: " & vbCrLf & vbCrLf & "Folio: " & .TextMatrix(ListaPed(0).Row, 2) & vbCrLf & _
        "Clave: " & .TextMatrix(ListaPed(0).Row, 0) & vbCrLf & "Fecha de pedido: " & _
        .TextMatrix(ListaPed(0).Row, 3) & vbCrLf & _
        "Proveedor: " & .TextMatrix(ListaPed(0).Row, 1), vbYesNo + vbQuestion)
        
        If ques = vbYes Then
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = False
            SSTab1.Tab = 1
            'cargaProveedor
            tipoPedido = "EDITAR1"
            pedId = .TextMatrix(.Row, 0)
            txtProd(3).Text = .TextMatrix(.Row, 2)
            dtFecha1(0) = .TextMatrix(.Row, 3)
            cmbProd(3).Text = .TextMatrix(.Row, 1)
            txtProd(4).Text = .TextMatrix(.Row, 9)
            
        Else
            tipoPedido = ""
        End If
        
    End With
    cmBoton(3).Visible = True
    
    
End Sub

Private Sub mn_Editar2_Click()

    Dim ques As String
    cmBoton(3).Visible = True
    
    With ListaPed(0)
        ques = MsgBox("Editar información de detalle del pedido: " & vbCrLf & vbCrLf & "Folio: " & .TextMatrix(ListaPed(0).Row, 2) & vbCrLf & _
        "Clave: " & .TextMatrix(ListaPed(0).Row, 0) & vbCrLf & "Fecha de pedido: " & _
        .TextMatrix(ListaPed(0).Row, 3) & vbCrLf & _
        "Proveedor: " & .TextMatrix(ListaPed(0).Row, 1), vbYesNo + vbQuestion)
        
        If ques = vbYes Then
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(0) = False
            SSTab1.Tab = 1
            'cargaProveedor
            tipoPedido = "EDITAR2"
            pedId = .TextMatrix(.Row, 0)
            txtProd(3).Text = .TextMatrix(.Row, 2)
            dtFecha1(0) = .TextMatrix(.Row, 3)
            cmbProd(3).Text = .TextMatrix(.Row, 1)
            txtProd(4).Text = .TextMatrix(.Row, 9)
            cargaDetalleProdEdit (ListaPed(0).TextMatrix(ListaPed(0).Row, 0))
        Else
            tipoPedido = ""
        End If
        
    End With

End Sub

Private Sub mn_Nuevo_Click()
    
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    cmBoton(3).Visible = True
    SSTab2.Tab = 0
    SSTab2.TabEnabled(1) = False
     
    
End Sub

Private Sub mn_Opciones_Click()
    If ListaPed(0).TextMatrix(ListaPed(0).Row, 4) = "GENERADO" Then
        mn_Cerrar.Enabled = True
    Else
        If ListaPed(0).TextMatrix(ListaPed(0).Row, 4) = "CERRADO" Then
            mn_Cerrar.Enabled = False
        Else
            mn_Cerrar.Enabled = True
        End If
    End If

End Sub

Private Sub mn_Salir_Click()
    Unload Me
End Sub

Private Sub timeCarga_Timer()
    timeCarga.Enabled = False
    SSTab1.width = Me.width - 200
    SSTab1.height = Me.height - 800
    SSTab2.width = Me.width - 200
    SSTab2.height = Me.height - 800
    
    
    ListaPed(0).width = Me.width - 450
    ListaPed(1).width = Me.width - 450
    ListaPed(1).height = ListaPed(1).height + 550
    ListaProd(1).width = Me.width - 450
    ListaProd(0).width = Me.width - 450
    ListaProd(0).height = ListaProd(0).height + 550
    
    
'    .width = Me.width - 500
'    Lista.width = Me.width - 500
    
End Sub

Private Sub txtProd_GotFocus(Index As Integer)
    If Index = 0 Then
        If txtProd(Index).Text = "CLAVE CODIGO" Then
            txtProd(Index).Text = ""
        End If
    Else
        If Index = 1 Then
            If txtProd(Index).Text = "CLAVE PROVEEDOR" Then
                txtProd(Index).Text = ""
            End If
        Else
            If Index = 5 Then
                If txtProd(Index).Text = "PRODUCTO" Then
                    txtProd(Index).Text = ""
                End If
            End If
        End If
    End If
End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 2 Then
        Numeros (KeyAscii)
    End If
    If KeyAscii = 13 Then
        If Index = 0 Then
            txtProd(1).Text = "CLAVE PROVEEDOR"
            txtProd(5).Text = "PRODUCTO"
            'Call checkProducto("PROD_CODIGO", "0")
            cargaProductos
         Else
            If Index = 1 Then
                txtProd(0).Text = "CLAVE CODIGO"
                txtProd(5).Text = "PRODUCTO"
                'Call checkProducto("PROD_CODIGO_PROV", "1")
                cargaProductos
            Else
                If Index = 5 Then
                    txtProd(0).Text = "CLAVE CODIGO"
                    txtProd(1).Text = "CLAVE PROVEEDOR"
                    'Call checkProducto("PROD_CODIGO_PROV", "1")
                    cargaProductos
                End If
            End If
         End If
        txtProd(Index).SelStart = 0
        txtProd(Index).SelLength = Len(txtProd(Index).Text)
    End If

End Sub

Private Sub txtProd_LostFocus(Index As Integer)
    If Index = 0 Then
        If txtProd(Index).Text = "" Then
            txtProd(Index).Text = "CLAVE CODIGO"
        End If
    Else
        If Index = 1 Then
            If txtProd(Index).Text = "" Then
                txtProd(Index).Text = "CLAVE PROVEEDOR"
            End If
        Else
            If Index = 5 Then
                If txtProd(Index).Text = "" Then
                    txtProd(Index).Text = "PRODUCTO"
                End If
            End If
        End If
    End If
End Sub
