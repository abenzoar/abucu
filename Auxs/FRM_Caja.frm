VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Caja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15855
   Icon            =   "FRM_Caja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRM_Caja.frx":08CA
   ScaleHeight     =   10575
   ScaleWidth      =   15855
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstab1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   18653
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   12
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Datos generales"
      TabPicture(0)   =   "FRM_Caja.frx":1194
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(13)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(14)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lista"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtFecha1(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtFecha1(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAccion(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAccion(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdMes"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAccion(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtInfo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFondo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdAccion(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "timeSize"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtFondoObser"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdAccion(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdAccion(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Check1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   " Cortes de caja"
      TabPicture(1)   =   "FRM_Caja.frx":172E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(1)"
      Tab(1).Control(1)=   "Label1(3)"
      Tab(1).Control(2)=   "Line1(2)"
      Tab(1).Control(3)=   "Label1(4)"
      Tab(1).Control(4)=   "Lista6"
      Tab(1).Control(5)=   "lista4"
      Tab(1).Control(6)=   "cmdAccion(13)"
      Tab(1).Control(7)=   "cmdAccion(15)"
      Tab(1).Control(8)=   "cmdAccion(16)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   " Detalle general"
      TabPicture(2)   =   "FRM_Caja.frx":1CC8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAccion(8)"
      Tab(2).Control(1)=   "cmdAccion(7)"
      Tab(2).Control(2)=   "Lista3"
      Tab(2).Control(3)=   "lista5"
      Tab(2).Control(4)=   "Line1(4)"
      Tab(2).Control(5)=   "Label1(6)"
      Tab(2).Control(6)=   "Line1(3)"
      Tab(2).Control(7)=   "Label1(5)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   " Detalle por usuario"
      TabPicture(3)   =   "FRM_Caja.frx":2262
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(7)"
      Tab(3).Control(1)=   "Line1(5)"
      Tab(3).Control(2)=   "Label1(8)"
      Tab(3).Control(3)=   "Line1(6)"
      Tab(3).Control(4)=   "listaPagos"
      Tab(3).Control(5)=   "lista2"
      Tab(3).Control(6)=   "cmdAccion(3)"
      Tab(3).Control(7)=   "cmdAccion(12)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   " Consumo interno"
      TabPicture(4)   =   "FRM_Caja.frx":27FC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "listaCI"
      Tab(4).Control(1)=   "Line1(7)"
      Tab(4).Control(2)=   "Label1(9)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   " Gastos"
      TabPicture(5)   =   "FRM_Caja.frx":2D96
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdAccion(17)"
      Tab(5).Control(1)=   "cmdAccion(9)"
      Tab(5).Control(2)=   "listaGST2"
      Tab(5).Control(3)=   "listaGST"
      Tab(5).Control(4)=   "Label1(17)"
      Tab(5).Control(5)=   "Line1(14)"
      Tab(5).Control(6)=   "Line1(8)"
      Tab(5).Control(7)=   "Label1(10)"
      Tab(5).ControlCount=   8
      TabCaption(6)   =   " Membresias"
      TabPicture(6)   =   "FRM_Caja.frx":3330
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListaMbr"
      Tab(6).Control(1)=   "Line1(9)"
      Tab(6).Control(2)=   "Label1(11)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   " Apartados"
      TabPicture(7)   =   "FRM_Caja.frx":38CA
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdAccion(10)"
      Tab(7).Control(1)=   "listaApt"
      Tab(7).Control(2)=   "Line1(10)"
      Tab(7).Control(3)=   "Label1(12)"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   " Asistencias"
      TabPicture(8)   =   "FRM_Caja.frx":3E64
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Line1(12)"
      Tab(8).Control(1)=   "Label1(15)"
      Tab(8).Control(2)=   "Label1(19)"
      Tab(8).Control(3)=   "Line1(16)"
      Tab(8).Control(4)=   "Label4"
      Tab(8).Control(5)=   "ListaAsts2"
      Tab(8).Control(6)=   "ListaAsts"
      Tab(8).Control(7)=   "cmdAccion(6)"
      Tab(8).Control(8)=   "ListaAsts3"
      Tab(8).Control(9)=   "cmdAccion(14)"
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "  Monedero"
      TabPicture(9)   =   "FRM_Caja.frx":43FE
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "listMonederos"
      Tab(9).Control(1)=   "Line1(13)"
      Tab(9).Control(2)=   "Label1(16)"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "Pagos usuarios"
      TabPicture(10)  =   "FRM_Caja.frx":4998
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lista_Pagos"
      Tab(10).Control(1)=   "Label1(18)"
      Tab(10).Control(2)=   "Line1(15)"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "Cancelaciones"
      TabPicture(11)  =   "FRM_Caja.frx":49B4
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Label1(20)"
      Tab(11).Control(1)=   "Line1(17)"
      Tab(11).Control(2)=   "Line1(18)"
      Tab(11).Control(3)=   "Label1(21)"
      Tab(11).Control(4)=   "ListaReimpresiones"
      Tab(11).Control(5)=   "ListaCancel"
      Tab(11).ControlCount=   6
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar gastos agrupados"
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
         Index           =   17
         Left            =   -68520
         Picture         =   "FRM_Caja.frx":49D0
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar todos los cortes"
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
         Index           =   16
         Left            =   -61080
         Picture         =   "FRM_Caja.frx":4F5A
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar indiviudal "
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
         Index           =   15
         Left            =   -61080
         Picture         =   "FRM_Caja.frx":54E4
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5520
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar lista 2 (excel)"
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
         Index           =   14
         Left            =   -70560
         Picture         =   "FRM_Caja.frx":5A6E
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   9480
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid ListaAsts3 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   60
         Top             =   5160
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":6338
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
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir corte"
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
         Index           =   13
         Left            =   -61080
         Picture         =   "FRM_Caja.frx":63F7
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dejar fondo para el día siguiente"
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
         Left            =   11880
         TabIndex        =   58
         Top             =   7200
         Width           =   3735
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realizar pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   12
         Left            =   -60840
         Picture         =   "FRM_Caja.frx":6CC1
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cajon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   11
         Left            =   14520
         Picture         =   "FRM_Caja.frx":724B
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir apartados"
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
         Index           =   10
         Left            =   -74760
         Picture         =   "FRM_Caja.frx":7B15
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   8400
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar gastos"
         Default         =   -1  'True
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
         Index           =   9
         Left            =   -74880
         Picture         =   "FRM_Caja.frx":83DF
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   9360
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportarr agrupación"
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
         Index           =   8
         Left            =   -61920
         Picture         =   "FRM_Caja.frx":8969
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   9480
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar detalle"
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
         Index           =   7
         Left            =   -61920
         Picture         =   "FRM_Caja.frx":8EF3
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   8040
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar lista 1 (excel)"
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
         Index           =   6
         Left            =   -74760
         Picture         =   "FRM_Caja.frx":947D
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   9480
         Width           =   3855
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir resumen y detalle"
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
         Height          =   1335
         Index           =   5
         Left            =   13800
         Picture         =   "FRM_Caja.frx":9D47
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox txtFondoObser 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   11880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   7920
         Width           =   3735
      End
      Begin VB.Timer timeSize 
         Interval        =   500
         Left            =   10920
         Top             =   10080
      End
      Begin MSFlexGridLib.MSFlexGrid ListaMbr 
         Height          =   6015
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         FormatString    =   $"FRM_Caja.frx":A611
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
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar fondo de caja"
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
         Index           =   4
         Left            =   13920
         Picture         =   "FRM_Caja.frx":A6F7
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox txtFondo 
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
         Left            =   11880
         TabIndex        =   20
         Top             =   6360
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   12000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Text            =   "FRM_Caja.frx":AFC1
         Top             =   10320
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar selección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   3
         Left            =   -60840
         Picture         =   "FRM_Caja.frx":AFC7
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Imprimir resumen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   2
         Left            =   11880
         Picture         =   "FRM_Caja.frx":B551
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4320
         Width           =   1815
      End
      Begin VB.ComboBox cmdMes 
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
         Left            =   12000
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realizar corte de Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   11880
         Picture         =   "FRM_Caja.frx":BE1B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar fechas y obtener datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   14400
         Picture         =   "FRM_Caja.frx":C6E5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Index           =   0
         Left            =   12000
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   310575105
         CurrentDate     =   40829
      End
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Index           =   1
         Left            =   12000
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   310575105
         CurrentDate     =   40829
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   9375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   16536
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Tipo                                        | Total                         | Cantidad   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Lista3 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   11668
         _Version        =   393216
         Cols            =   19
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":CFAF
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
      Begin MSFlexGridLib.MSFlexGrid lista2 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   9
         Top             =   960
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D11B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid lista4 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   10
         Top             =   960
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid Lista6 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   11
         Top             =   4080
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D1CA
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
      Begin MSFlexGridLib.MSFlexGrid listaPagos 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   15
         Top             =   4320
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D27D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid listaCI 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   17
         Top             =   960
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   11668
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D31E
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
      Begin MSFlexGridLib.MSFlexGrid lista5 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   23
         Top             =   8040
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D3F2
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
      Begin MSFlexGridLib.MSFlexGrid listaApt 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   34
         Top             =   960
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   12938
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         FormatString    =   $"FRM_Caja.frx":D498
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
      Begin MSFlexGridLib.MSFlexGrid ListaAsts 
         Height          =   975
         Left            =   -74520
         TabIndex        =   40
         Top             =   7920
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1720
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D562
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
      Begin MSFlexGridLib.MSFlexGrid listMonederos 
         Height          =   8415
         Left            =   -74760
         TabIndex        =   47
         Top             =   1080
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14843
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D63D
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
      Begin MSFlexGridLib.MSFlexGrid listaGST2 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   49
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   " Concepto                         |   Total                      "
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
      Begin MSFlexGridLib.MSFlexGrid listaGST 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   51
         Top             =   5280
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D718
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
      Begin MSFlexGridLib.MSFlexGrid lista_Pagos 
         Height          =   8415
         Left            =   -74760
         TabIndex        =   54
         Top             =   1080
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   14843
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D83C
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
      Begin MSFlexGridLib.MSFlexGrid ListaAsts2 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   56
         Top             =   960
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid ListaCancel 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   66
         Top             =   1080
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":D8EC
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
      Begin MSFlexGridLib.MSFlexGrid ListaReimpresiones 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   69
         Top             =   5880
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Caja.frx":DA1B
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Impresiones"
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
         Left            =   -74760
         TabIndex        =   68
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   18
         X1              =   -74760
         X2              =   -59160
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   17
         X1              =   -74760
         X2              =   -59160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelaciones"
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
         Left            =   -74760
         TabIndex        =   67
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Horario: "
         Height          =   495
         Left            =   -66000
         TabIndex        =   65
         Top             =   9720
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   16
         X1              =   -74760
         X2              =   -59160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencias"
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
         TabIndex        =   57
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de pagos a usuarios"
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
         Left            =   -74760
         TabIndex        =   55
         Top             =   720
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   15
         X1              =   -74760
         X2              =   -59160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos agrupados por tipo"
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
         Left            =   -74880
         TabIndex        =   50
         Top             =   600
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   14
         X1              =   -74880
         X2              =   -59280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   13
         X1              =   -74760
         X2              =   -59160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de monedero"
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
         Left            =   -74760
         TabIndex        =   48
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de asistencias"
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   4800
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   12
         X1              =   -74760
         X2              =   -59160
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Index           =   14
         Left            =   11880
         TabIndex        =   38
         Top             =   7560
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   11
         X1              =   11880
         X2              =   15600
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fondo de Caja"
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
         Left            =   11880
         TabIndex        =   36
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   10
         X1              =   -74880
         X2              =   -59280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de apartados"
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
         TabIndex        =   35
         Top             =   600
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   9
         X1              =   -74760
         X2              =   -59160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de membresias"
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
         TabIndex        =   33
         Top             =   600
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   8
         X1              =   -74880
         X2              =   -59280
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de gastos"
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
         Left            =   -74880
         TabIndex        =   32
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   7
         X1              =   -74880
         X2              =   -59280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de consumo interno"
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
         Left            =   -74880
         TabIndex        =   31
         Top             =   600
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   6
         X1              =   -74760
         X2              =   -59160
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de pagos/comisiones por usuario"
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   3960
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   5
         X1              =   -74760
         X2              =   -59160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de operaciones por usuario"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   4
         X1              =   -74880
         X2              =   -59280
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agrupación de operaciones por insumo"
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   7680
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   3
         X1              =   -74880
         X2              =   -59280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle específico de cada operación con los datos correspondientes"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle del corte de caja seleccionado agrupado por insumo"
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
         Left            =   -74880
         TabIndex        =   26
         Top             =   3720
         Width           =   6135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   -74880
         X2              =   -59280
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Resumen del corte de caja por usuario"
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   600
         Width           =   3855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   1
         X1              =   -74880
         X2              =   -59280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   0
         X1              =   120
         X2              =   11640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Resumen general de las operaciones principales"
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
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad $"
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
         Left            =   11880
         TabIndex        =   19
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Left            =   12000
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "a"
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
         Left            =   12000
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   12000
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Menu mn_Corte 
      Caption         =   "Corte"
      Visible         =   0   'False
      Begin VB.Menu mn_CorteCaja2 
         Caption         =   "Corte Caja"
      End
   End
End
Attribute VB_Name = "FRM_Caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim totProductos As Double
Dim totServicios As Double
Dim totDescuentos As Double
Dim totGastos As Double
Dim totMebresias As Double
Dim totApartados As Double
Dim totCambios As Double
Dim totMonederos As Double
Dim totPagosUsuarios As Double
Dim fondoIni As Double
Dim tipo As String
Dim tipoBusqueda As Boolean
Dim personas As Long
Dim tipoHorario As String

Private Sub format_Listas()
    
    lista4.Rows = 2
    
    lista4.TextMatrix(0, 0) = "Sesión - Corte de Caja"
    lista4.TextMatrix(0, 1) = "Sesión - Corte de Caja"
    
    lista4.TextMatrix(1, 0) = "Usuario"
    lista4.TextMatrix(1, 1) = "Fecha/Hora"
    
    lista4.TextMatrix(0, 2) = "Información caja sesión"
    lista4.TextMatrix(0, 3) = "Información caja sesión"
    lista4.TextMatrix(0, 4) = "Información caja sesión"
    lista4.TextMatrix(0, 5) = "Información caja sesión"
    lista4.TextMatrix(0, 6) = "Información caja sesión"
        
    lista4.TextMatrix(1, 2) = "Servicios"
    lista4.TextMatrix(1, 3) = "Serv Cant"
    lista4.TextMatrix(1, 4) = "Productos"
    lista4.TextMatrix(1, 5) = "Prod Cant"
    lista4.TextMatrix(1, 6) = "Total"
    
    lista4.TextMatrix(0, 7) = "Información caja general"
    lista4.TextMatrix(0, 8) = "Información caja general"
    lista4.TextMatrix(0, 9) = "Información caja general"
    lista4.TextMatrix(0, 10) = "Información caja general"
         
    lista4.TextMatrix(1, 7) = "Servicios"
    lista4.TextMatrix(1, 8) = "Serv Cant"
    lista4.TextMatrix(1, 9) = "Productos"
    lista4.TextMatrix(1, 10) = "Prod Cant"
    lista4.TextMatrix(1, 11) = "Total"
    
    lista4.TextMatrix(0, 12) = " "
    lista4.TextMatrix(1, 12) = "Clave corte"
    
    
    lista4.MergeCells = flexMergeRestrictRows
    lista4.MergeRow(0) = True
    lista4.MergeRow(1) = True
    
    lista4.ColWidth(0) = 3700
    lista4.ColWidth(1) = 2700


End Sub
Private Sub cmdAccion_Click(Index As Integer)
    Dim ques As String
    
'    sstab1.Tab = 8
    
    Select Case Index
        Case 0:
            If dtFecha1(0) <= dtFecha1(1) Then
                tipoBusqueda = True
                cargaDatos
            Else
                MsgBox "No se puede realizar la operación. Verifique.", vbInformation
            End If
        Case 1:
            ques = MsgBox("Va a realizar el corte de caja y quedará registrado. " & vbCrLf & vbCrLf & _
            "¿Continuar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                If dtFecha1(0) <> dtFecha1(1) Then
                    MsgBox "No se puede ralizar un corte de caja con más de una fecha. Verifique.", vbInformation
                Else
                    corteCaja
                    envioMailInfo
                End If
            End If
        Case 2:
            ques = MsgBox("Imprimir corte de caja." & vbCrLf & "Verifique la conexión con la impresora.", vbYesNo + vbQuestion)
            If ques = vbYes Then
                'impresionCaja
                resumenCaja
            End If
        Case 3:
            ques = MsgBox("Imprimir corte de usuario." & vbCrLf & "Verifique la conexión con la impresora.", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Dim fila  As Long
                fila = lista2.Row
                For b1 = 1 To lista2.Rows - 1
                    If lista2.TextMatrix(b1, 8) = Chr(254) Then
                        lista2.Row = b1
                        lista2_Click
                        cargaDetallePagos (lista2.TextMatrix(b1, 1))
                        notaUsuario (b1)
                    End If
                Next b1
                cargaDetallePagos (lista2.TextMatrix(fila, 1))
            End If
        Case 4:
            If Check1.value = Checked Then
                ques = MsgBox("Actualizar el valor para el día siguiente: " & Date + 1, vbYesNo + vbQuestion)
                If ques = vbYes Then
                        sql1 = "INSERT INTO CAT_CORTECAJA  (CRTCAJA_FECHA, CRTCAJA_MONTO, CRTCAJA_USERID, CRTCAJA_USERPERID, CRTCAJA_USERTIPO, CRTCAJA_OBSERVACIONES) VALUES (" & _
                        " '" & Format(Date + 1, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', '" & Val(Format(txtFondo.Text, "General Number")) & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U', '" & txtFondoObser.Text & "' )"
                        'MsgBox SQL1
                        con.Execute (sql1)
                        cmdAccion_Click (0)
                        MsgBox "Informaciòn actualizada."
                End If
            Else
                ques = MsgBox("Actualizar el valor de fondo de caja." & vbcrf & "Esto afectarà en el total del dìa " & dtFecha1(0), vbYesNo + vbQuestion)
                If ques = vbYes Then
                    If dtFecha1(0) = dtFecha1(1) Then
                        sql1 = "INSERT INTO CAT_CORTECAJA  (CRTCAJA_FECHA, CRTCAJA_MONTO, CRTCAJA_USERID, CRTCAJA_USERPERID, CRTCAJA_USERTIPO, CRTCAJA_OBSERVACIONES) VALUES (" & _
                        " '" & Format(dtFecha1(0), "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', '" & Val(Format(txtFondo.Text, "General Number")) & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U', '" & txtFondoObser.Text & "' )"
        '                MsgBox SQL1
                        con.Execute (sql1)
                        cmdAccion_Click (0)
                        MsgBox "Informaciòn actualizada."
                    Else
                        MsgBox "No se puede editar fondo de caja. Verifique que las fechas coincidan.", vbInformation
                    End If
                End If
            End If
        Case 5:
            ques = MsgBox("Imprimir resumen con detalle de la caja." & vbCrLf & "Verifique la conexión con la impresora.", vbYesNo + vbQuestion)
            If ques = vbYes Then
                'impresionCaja
                'resumenCaja2
            End If
        
        Case 6:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(ListaAsts2)
            End If
        Case 7:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(Lista3)
            End If
        Case 8:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(lista5)
            End If
        Case 9:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(listaGST)
            End If
        Case 10:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(listaApt)
            End If
        Case 11:
            Call abrirCajon
        Case 12:
'            Call abrirCajon
            pagoUsuarios
        Case 13
            resumenUsuario
        Case 14:
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(ListaAsts3)
            End If
        Case 16:
            ques = MsgBox("¿Exportar cortes? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(lista4)
            End If
        Case 15:
            ques = MsgBox("¿Exportar cortes? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(Lista6)
            End If
        Case 17:
            ques = MsgBox("¿Exportar gastos? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(listaGST2)
            End If
            
    End Select
End Sub


Private Sub pagoUsuarios()
    ques = MsgBox("Realizar pago a usuarios seleccionados en la lista" & vbCrLf & vbCrLf & "Periodo: " & dtFecha1(0) & " al " & dtFecha1(1) & vbCrLf & vbCrLf & "¿Continuar? ", vbYesNo + vbQuestion)
    
    
    If ques = vbYes Then
    
        For b1 = 1 To lista2.Rows - 1
            If lista2.TextMatrix(b1, 8) = Chr(254) Then
                lista2.Row = b1
                lista2_Click
                cargaDetallePagos (lista2.TextMatrix(b1, 1))
                For c1 = 1 To listaPagos.Rows - 1
                    With lista2
                        sql1 = "INSERT INTO PAGOS (PG_PERTP_TIPO_ID, PG_PERTP_PER_ID, PG_PERTP_PER_TIPO, PG_CTPG_ID, PG_FECHA_HORA, PG_PERIODO_INI, PG_PERIODO_FIN, PG_MONTO, PG_USUARIO_TIPO_ID, PG_USUARIO_PER_ID, PG_USUARIO_PER_TIPO) " & _
                        "VALUES ( '" & .TextMatrix(.Row, 9) & "',  '" & .TextMatrix(.Row, 1) & "', '" & .TextMatrix(.Row, 10) & "', '" & listaPagos.TextMatrix(c1, 5) & "', Now(), '" & Format(dtFecha1(0), "yyyy-MM-dd") & "', '" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '" & Format(listaPagos.TextMatrix(c1, 4), "General Number") & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U'  )"
                        con.Execute (sql1)
                        
                    End With
                Next c1
                'notaUsuario (b1)
            End If
        Next b1
        
        MsgBox "Pago realizado", vbInformation
        
'    cargaDetallePagos (lista2.TextMatrix(fila, 1))
        
        
        'METER EN TABLA PAGOS LOS DATOS Y LUEGO REFLEJARLOS EN RESUMEN Y PONER UN STATUS DE PAGADO EN GASTOS
'        SQL1 = "INSERT INTO PAGOS (PG_PERTP_TIPO_ID, PG_PERTP_PER_ID, PG_PERTP_PER_TIPO, PG_CTPG_ID, PG_FECHA_HORA, PG_PERIODO_INI, PG_PERIODO_FIN, PG_MONTO, PG_USUARIO_TIPO_ID, PG_USUARIO_PER_ID, PG_USUARIO_PER_TIPO) " & _
'        "VALUES ()"
'        With lista2
'            MsgBox .TextMatrix(.Row, 9) & " " & .TextMatrix(.Row, 1) & " " & .TextMatrix(.Row, 10), Now(), listaPagos.TextMatrix(n1, 5)
'        End With
    End If


End Sub
Private Sub envioMailInfo()
    Call enviar_Mail("CAJA", "Corte de caja " & Format(Date, "Short Date") & "", "", "")
End Sub

Private Sub impresionCaja()
    Dim rec_General As Recordset
    Dim b1 As Long
    Dim Fecha As String
    Dim Sucursal As String
    Dim usuarioCorte As String
    
    Set rec_General = New Recordset
    With rec_General.Fields
        .Append "Tipo", adVarChar, 50
        .Append "Total", adVarChar, 50
        .Append "Cantidad", adVarChar, 50
    End With
    rec_General.Open
    
    Fecha = "Caja del " & dtFecha1(0) & " al " & dtFecha1(1)
    Sucursal = FRM_Menu.menuBarra2.Panels(9).Text & vbCrLf & "Sucursal: " & FRM_Menu.menuBarra2.Panels(3).Text
    usuarioCorte = "Usuario que realiza el corte de caja: " & vbCrLf & FRM_Menu.menuBarra2.Panels(5).Text
    Report_CajaGral.Sections.Item("Section4").Controls("lFecha").Caption = Fecha
    Report_CajaGral.Sections.Item("Section4").Controls("lSucursal").Caption = Sucursal
    Report_CajaGral.Sections.Item("Section4").Controls("lPerCorte").Caption = usuarioCorte
    
    With lista
        For b1 = 1 To .Rows - 1
            rec_General.AddNew Array("Tipo", "Total", "Cantidad"), _
            Array(.TextMatrix(b1, 0), .TextMatrix(b1, 1), .TextMatrix(b1, 2))
        Next b1
    End With
    
    Set Report_CajaGral.DataSource = rec_General
    Report_CajaGral.Show vbModal
    
    If Not rec_General.State = adStateOpen Then
        rec_General.Close
    End If

    If Not rec_General Is Nothing Then
        Set rec_General = Nothing
    End If
    
    
End Sub
Private Sub corteCaja()
    ''''----
    Dim corteId
    
    If dtFecha1(0) <> Date Or dtFecha1(1) <> Date Then
        MsgBox "No se puede realizar un corte de caja si no es la fecha actual. ", vbInformation
    Else
        sql1 = "INSERT INTO CORTE_CAJA (FECHA, USUARIO1_ID, USUARIO1_TIPOID, USUARIO1_TIPO, PRODUCTOS_GRAL, SERVICIOS_GRAL, TOTAL_GRAL, PROD_CANT_GRAL, SERV_CANT_GRAL, fondo_caja) " & _
        "VALUES (NOW(), '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', " & _
        "'U', '" & Val(Format(lista.TextMatrix(1, 1), "General Number")) & "', '" & Val(Format(lista.TextMatrix(2, 1), "General Number")) & "', " & _
        "'" & Val(Format(lista.TextMatrix(12, 1), "General Number")) & "', '" & lista.TextMatrix(1, 2) & "',  '" & lista.TextMatrix(2, 2) & "', '" & Val(Format(txtFondo.Text, "General number")) & "' )"
        con.Execute (sql1)
        
        sql1 = "select last_insert_id() corteId"
        Set RES1 = con.Execute(sql1)
        If Not RES1.EOF Then
            corteId = RES1.Fields("corteId")
        End If
        
        With lista5
            For b1 = 1 To .Rows - 1
                sql1 = "INSERT INTO corteCaja_Detalle (IdCorte, prodCodigo, prodNombre, prodTipo, prodPrecio, prodCantidad, prodTotal, prodInventario) " & _
                "values ('" & corteId & "', '" & .TextMatrix(b1, 2) & "', '" & .TextMatrix(b1, 1) & "', '" & .TextMatrix(b1, 0) & "', " & _
                "'" & Val(Format(.TextMatrix(b1, 3), "General Number")) & "', '" & .TextMatrix(b1, 4) & "', '" & Val(Format(.TextMatrix(b1, 5), "General Number")) & "', '" & Val(.TextMatrix(b1, 6)) & "' )"
                con.Execute (sql1)
            Next b1
        End With
        
        cargaCorte
    
    '''''pARA GUARDAR EN FONDO DE CAJA EL MONTO QUE QUEDE...
    sql1 = "SELECT SUC_CORTE_FONDO FROM SUCURSAL"
    Set RES1 = con.Execute(sql1)
    
        If RES1.Fields("SUC_CORTE_FONDO") = "A" Then
            sql1 = "INSERT INTO CAT_CORTECAJA  (CRTCAJA_FECHA, CRTCAJA_MONTO, CRTCAJA_USERID, CRTCAJA_USERPERID, CRTCAJA_USERTIPO, CRTCAJA_OBSERVACIONES) VALUES (" & _
            " (DATE_ADD(NOW(), INTERVAL 1 DAY)), '" & Val(Format(lista.TextMatrix(13, 1), "General Number")) & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U', 'FONDO DE CAJA: TOTAL DEL DIA ANTERIOR' ) "
                       ' MsgBox SQL1
            con.Execute (sql1)
            cmdAccion_Click (0)
        End If
    
    End If
    
    
    


End Sub
Private Sub cargaCorte()
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "AND date_format(FECHA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " AND date_format(FECHA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
            Else
                tipo = tipo & " AND date_format(FECHA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
'            tipo = tipo & " AND FECHA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
        End If
    End If
    
    lista4.Rows = 2
    sql1 = "SELECT IDCORTE, FECHA, PRODUCTOS_GRAL, PROD_CANT_GRAL, SERV_CANT_GRAL, SERVICIOS_GRAL, TOTAL_GRAL, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) USUARIO, " & _
    "PRODUCTOS_SESION, SERVICIOS_sESION, TOTAL_SESION, PROD_CANT_SESION, SER_CANT_sESION  " & _
    "FROM CORTE_CAJA T1, PERSONA T2, SUCURSAL T5 " & _
    "WHERE T1.USUARIO1_ID = T2.PER_ID " & _
    tipo & " ORDER BY FECHA asc"
    '"AND date_format(FECHA, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    lista4.Redraw = False
    Do While Not RES1.EOF
        lista4.AddItem ""
        lista4.TextMatrix(lista4.Rows - 1, 0) = RES1.Fields("USUARIO")
        lista4.TextMatrix(lista4.Rows - 1, 1) = RES1.Fields("FECHA")
        
        lista4.TextMatrix(lista4.Rows - 1, 2) = FormatCurrency(RES1.Fields("SERVICIOS_SESION"))
        lista4.TextMatrix(lista4.Rows - 1, 3) = RES1.Fields("SER_CANT_SESION") & ""
        lista4.TextMatrix(lista4.Rows - 1, 4) = FormatCurrency(RES1.Fields("PRODUCTOS_SESION"))
        lista4.TextMatrix(lista4.Rows - 1, 5) = RES1.Fields("PROD_CANT_SESION") & ""
        lista4.TextMatrix(lista4.Rows - 1, 6) = FormatCurrency(RES1.Fields("TOTAL_SESION"))
        
        lista4.TextMatrix(lista4.Rows - 1, 7) = FormatCurrency(RES1.Fields("SERVICIOS_GRAL"))
        lista4.TextMatrix(lista4.Rows - 1, 8) = RES1.Fields("SERV_CANT_GRAL") & ""
        lista4.TextMatrix(lista4.Rows - 1, 9) = FormatCurrency(RES1.Fields("PRODUCTOS_GRAL"))
        lista4.TextMatrix(lista4.Rows - 1, 10) = RES1.Fields("PROD_CANT_GRAL") & ""
        lista4.TextMatrix(lista4.Rows - 1, 11) = FormatCurrency(RES1.Fields("TOTAL_GRAL"))
        lista4.TextMatrix(lista4.Rows - 1, 12) = RES1.Fields("IDCORTE")
        
        lista4.Row = lista4.Rows - 1
        For b1 = 7 To 11
            lista4.Col = b1
            lista4.CellBackColor = &HC0C0FF
        Next b1
        
        For b1 = 2 To 6
            lista4.Col = b1
            lista4.CellBackColor = &HFFC0C0
        Next b1
        
        RES1.MoveNext
    Loop
    lista4.Redraw = True

End Sub

Private Sub cmdMes_Click()
    inicio_Mes = DateSerial(dtFecha1(0).Year, cmdMes.ListIndex + 1, 1)
    fin_mes = DateSerial(dtFecha1(0).Year, cmdMes.ListIndex + 2, 1)
    fin_mes = DateAdd("d", -1, fin_mes)

    'MsgBox inicio_mes & "  " & fin_mes
    dtFecha1(0) = inicio_Mes
    dtFecha1(1) = fin_mes
    
End Sub



Private Sub Command1_Click()
    ListaAsts.Rows = 1
    ListaAsts2.Rows = 1
    ListaAsts3.Rows = 1
End Sub

Private Sub Form_Load()

    SSTab1.Tab = 0
    dtFecha1(0) = Date
    dtFecha1(1) = Date
    cargaMes
    
    format_Listas
    tipoBusqueda = False
    cargaDatos
    totDescuentos = 0
    totProductos = 0
    totServicios = 0
    

    If permAdd = "SI" Then
        dtFecha1(0).Enabled = True
        dtFecha1(1).Enabled = True
        cmdAccion(0).Enabled = True
        cmdMes.Enabled = True
        cmdAccion(2).Enabled = True
        cmdAccion(5).Enabled = True
        cmdAccion(4).Enabled = True
        txtFondo.Enabled = True
        txtFondoObser.Enabled = True
    Else
        dtFecha1(0).Enabled = False
        dtFecha1(1).Enabled = False
        cmdAccion(0).Enabled = False
        cmdMes.Enabled = False
        cmdAccion(2).Enabled = False
        cmdAccion(5).Enabled = False
        cmdAccion(4).Enabled = False
        txtFondo.Enabled = False
        txtFondoObser.Enabled = False
    End If

    If permEdit = "SI" Then
    
        SSTab1.TabEnabled(0) = True
        For b1 = 1 To 9
            SSTab1.TabEnabled(b1) = True
        Next b1
    Else
        For b1 = 1 To 9
            SSTab1.TabEnabled(b1) = False
        Next b1
    
    End If


End Sub
Private Sub cargaMes()
    
    cmdMes.Clear
    For b1 = 1 To 12
        cmdMes.AddItem MonthName(b1)
    Next b1
End Sub
Private Sub cargaDatos()
        
    checkCajaFondo
    pagosProdServ
    pagosApartados
    pagosCambios
    pagosDescuento
    consumoInterno
    cargaFondo
    gastos
    monederos
    fondoInicial
    pagosTotal
    pagosTipo
    cancelaciones
    reimpresiones
    
    asistenciaResumen
'    asistenciaPorDia
    
    pagosUsuarios
    

        
    ventaGroup
    detalleVenta
    cargaCorte
    cargaApartados
    
    membresias
    pagosGeneradosUsuarios

    If mesas = True Then
        comensales
    End If
    
'    monederos
'    pagosTotal
    
End Sub
Private Sub reimpresiones()

    Dim MONTO_CANCEL As Double
    Dim reimpresiones As Double
    Dim tipo As String
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE date_format(FECHAHORA_IMPRESION, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " WHERE FECHAHORA_IMPRESION BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " WHERE FECHAHORA_IMPRESION BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " WHERE FECHAHORA_IMPRESION BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
        End If
    End If
    
    
    
    ListaReimpresiones.Rows = 1
    sql1 = "SELECT * FROM VIEW_REIMPRESIONES, SUCURSAL T5  " & tipo & "   ORDER BY FOLIO DESC "
    Set RES1 = con.Execute(sql1)
    
    MONTO_CANCEL = 0
    reimpresiones = 0
    Do While Not RES1.EOF
        ListaReimpresiones.AddItem ""
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 0) = RES1.Fields("FOLIO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 1) = RES1.Fields("TIPO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 2) = RES1.Fields("PRODUCTO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 3) = RES1.Fields("CANTIDAD")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 4) = FormatCurrency(RES1.Fields("PRECIO"))
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 5) = RES1.Fields("CODIGO_PROD")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 6) = RES1.Fields("IMPRESIONES") - 1
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 7) = RES1.Fields("MOSTRADOR")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 8) = RES1.Fields("ATENDIO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 9) = RES1.Fields("AUTORIZO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 10) = RES1.Fields("FECHAHORA_VENTA")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 11) = RES1.Fields("PRODUCTO_FECHAHORA")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 12) = RES1.Fields("FECHAHORA_IMPRESION")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 13) = RES1.Fields("CLAVE_REGISTRO")
        ListaReimpresiones.TextMatrix(ListaReimpresiones.Rows - 1, 14) = RES1.Fields("MOTIVO")
        reimpresiones = reimpresiones + (RES1.Fields("IMPRESIONES") - 1)
        MONTO_CANCEL = MONTO_CANCEL + (RES1.Fields("PRECIO") * RES1.Fields("CANTIDAD"))
        RES1.MoveNext
    Loop

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "RE-IMPRESIONES"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(MONTO_CANCEL * reimpresiones)
    lista.TextMatrix(lista.Rows - 1, 2) = reimpresiones
    lista.Row = lista.Rows - 1
    lista.Col = 0
    lista.CellForeColor = vbRed
    lista.Col = 1
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellForeColor = vbRed
    

End Sub

Private Sub cancelaciones()
    Dim MONTO_CANCEL As Double
    Dim tipo As String
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE date_format(VENDET_FECHAHORACANCEL, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " WHERE VENDET_FECHAHORACANCEL BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " WHERE VENDET_FECHAHORACANCEL BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " WHERE VENDET_FECHAHORACANCEL BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
        End If
    End If
    
    
    
    ListaCancel.Rows = 1
    sql1 = "SELECT * FROM VIEW_CANCELACIONES, SUCURSAL T5 " & tipo & " ORDER BY FOLIO DESC "
    Set RES1 = con.Execute(sql1)
    
    MONTO_CANCEL = 0
    Do While Not RES1.EOF
        ListaCancel.AddItem ""
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 0) = RES1.Fields("FOLIO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 1) = RES1.Fields("TIPO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 2) = RES1.Fields("PRODUCTO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 3) = RES1.Fields("CANTIDAD")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 4) = FormatCurrency(RES1.Fields("PRECIO"))
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 5) = RES1.Fields("CODIGO_PROD")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 6) = RES1.Fields("MOSTRADOR")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 7) = RES1.Fields("ATENDIO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 8) = RES1.Fields("AUTORIZO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 9) = RES1.Fields("FECHAHORA_VENTA")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 10) = RES1.Fields("PRODUCTO_FECHAHORA")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 11) = RES1.Fields("VENDET_FECHAHORACANCEL")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 12) = RES1.Fields("CLAVE_REGISTRO")
        ListaCancel.TextMatrix(ListaCancel.Rows - 1, 13) = RES1.Fields("MOTIVO_CANCEL")
        MONTO_CANCEL = MONTO_CANCEL + (RES1.Fields("PRECIO") * RES1.Fields("CANTIDAD"))
        RES1.MoveNext
    Loop

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "CANCELACIONES"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(MONTO_CANCEL)
    lista.TextMatrix(lista.Rows - 1, 2) = ListaCancel.Rows - 1
    lista.Row = lista.Rows - 1
    lista.Col = 0
    lista.CellForeColor = vbRed
    lista.Col = 1
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellForeColor = vbRed
    

End Sub
Private Sub comensales()
    On Error Resume Next
    Dim tipo As String
    tipo = ""
    
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " date_format(fechaHora, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            'tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " FECHAHORa BETWEEN CONCAT((DATE_FORMAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " FECHAHORA BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If

'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
            
        End If
    End If
    
    
    
    sql1 = "SELECT sum(PERSONAS) personas, sum(TOTAL1) total FROM VIEW_VENTAS t1, SUCURSAL T3 WHERE " & tipo
    Set RES1 = con.Execute(sql1)
    
    personas = 0
    total = 0
    If Not RES1.EOF Then
        If IsNull(RES1.Fields("PERSONAS")) = True Then
            personas = 0
        Else
            personas = Val(RES1.Fields("PERSONAS"))
        End If
        If IsNull(RES1.Fields("total")) = True Then
            total = 0
        Else
            total = Val(RES1.Fields("total"))
        End If
        
    End If
    
            
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Cheque promedio"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Val(total) / Val(personas))
    lista.TextMatrix(lista.Rows - 1, 2) = ""
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Comensales"
    lista.TextMatrix(lista.Rows - 1, 1) = ""
    lista.TextMatrix(lista.Rows - 1, 2) = Val(personas)
    
    
    
End Sub
Private Sub asistenciaPorDia()
    Dim dias As Long, tipo As String
    ListaAsts2.Cols = 1
    ListaAsts2.Rows = 1
    dias = dtFecha1(1) - dtFecha1(0)
    ListaAsts2.Cols = dias + 6
    ListaAsts2.ColWidth(0) = 2500
    
    For b1 = 0 To ListaAsts2.Cols - 1
        If b1 = 0 Then
            ListaAsts2.TextMatrix(0, b1) = "PERSONA"
        Else
            If b1 = 1 Then
                ListaAsts2.TextMatrix(0, b1) = "JORNADA"
            Else
                If b1 = 2 Then
                    ListaAsts2.TextMatrix(0, b1) = "DIAS"
                Else
                    If b1 = 3 Then
                        ListaAsts2.TextMatrix(0, b1) = "ASISTENCIAS"
                    Else
                        If b1 = 4 Then
                            ListaAsts2.TextMatrix(0, b1) = "HORAS"
                        Else
                            ListaAsts2.TextMatrix(0, b1) = dtFecha1(0) + (b1 - 5)
                        End If
                    End If
                End If
            End If
        End If
    Next b1
    
    
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE DATE_FORMAT(FECHA_HORA, '%Y-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            tipo = tipo & " WHERE FECHA_HORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If
    
    
    sql1 = "SELECT DISTINCT(CLIENTE)  PERSONA, JORNADA FROM VIEW_ASISTENCIAS, SUCURSAL T3" & tipo
    Set RES1 = con.Execute(sql1)
    
    Dim HORAS As Long
    
    Do While Not RES1.EOF
        ListaAsts2.AddItem ""
        ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 0) = RES1.Fields("PERSONA")
        ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 1) = RES1.Fields("JORNADA") & "HRS"
        ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 2) = dtFecha1(1) - dtFecha1(0)
        ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 3) = dtFecha1(1) - dtFecha1(0)
        ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 4) = (dtFecha1(1) - dtFecha1(0)) * RES1.Fields("JORNADA")
        
        'SQL1 = "SELECT     "
        For b1 = 5 To ListaAsts2.Cols - 1
            ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, b1) = RES1.Fields("JORNADA")
        Next b1
        
        RES1.MoveNext
    Loop
        




End Sub
Private Sub pagosGeneradosUsuarios()
    Dim monto As Double
    
    sql1 = "SELECT * FROM VIEW_PAGOS WHERE INICIO >= '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND FIN <='" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    totPagosUsuarios = 0
    monto = 0
    lista_Pagos.Rows = 1
    Do While Not RES1.EOF
        lista_Pagos.AddItem ""
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 0) = RES1.Fields("RECIBIO")
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 1) = FormatCurrency(RES1.Fields("MONTO"))
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 2) = RES1.Fields("INICIO")
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 3) = RES1.Fields("FIN")
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 4) = RES1.Fields("USUARIO")
        lista_Pagos.TextMatrix(lista_Pagos.Rows - 1, 5) = RES1.Fields("REGISTRO")
        monto = monto + RES1.Fields("monto")
        RES1.MoveNext
    Loop
            
    If monto > 0 Then
    totPagosUsuarios = monto
    End If
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "PAGOS USUARIOS"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(monto)
    lista.TextMatrix(lista.Rows - 1, 2) = lista_Pagos.Rows - 1
    
    lista.Row = lista.Rows - 1
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 0
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    


End Sub


Private Sub checkCajaFondo()
    sql1 = "SELECT SUC_CORTE_FONDO FROM SUCURSAL"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        If RES1.Fields("SUC_CORTE_FONDO") = "A" Then
            cmdAccion(4).Enabled = False
            txtFondo.Enabled = False
            txtFondoObser.Enabled = False
        Else
            cmdAccion(4).Enabled = True
            txtFondo.Enabled = True
            txtFondoObser.Enabled = True
        End If
    End If

End Sub
Private Sub monederos()
Dim MONEDERO_RECIBE, MONEDERO_ENTREGA As Double
Dim cantidad_recibe, cantidad_entrega As Long


    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "where DATE_FORMAT(fechahora, '%Y-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ORDER BY FECHAHORA DESC"
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then

            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " where DATE_FORMAT(fechahora, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
            Else
                tipo = tipo & " where DATE_FORMAT(fechahora, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If

'            tipo = tipo & " where FECHAHORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) ORDER BY FECHAHORA DESC "
        End If
    End If
    
    sql1 = "SELECT * fROM VIEW_PUNTOS_ADMIN, SUCURSAL T5 " & _
    tipo
    '"where DATE_FORMAT(fechahora, '%x-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ORDER BY FECHAHORA DESC"
    Set RES1 = con.Execute(sql1)
    
    listMonederos.Rows = 1
    MONEDERO_RECIBE = 0
    MONEDERO_ENTREGA = 0
    cantidad_recibe = 0
    cantidad_entrega = 0
    totMonederos = 0
    
    listMonederos.Redraw = False
    Do While Not RES1.EOF
        If RES1.Fields("tipo") = "RECIBE" Then
            MONEDERO_RECIBE = MONEDERO_RECIBE + RES1.Fields("MONEDERO")
            cantidad_recibe = cantidad_recibe + 1
        Else
            If RES1.Fields("tipo") = "ENTREGA" Then
                MONEDERO_ENTREGA = MONEDERO_ENTREGA - RES1.Fields("MONEDERO")
                cantidad_entrega = cantidad_entrega + 1
                totMonederos = totMonederos - RES1.Fields("MONEDERO")
            End If
        End If
        listMonederos.AddItem ""
        listMonederos.TextMatrix(listMonederos.Rows - 1, 0) = RES1.Fields("TIPO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 1) = RES1.Fields("CLIENTE")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 2) = RES1.Fields("ORIGEN")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 3) = FormatCurrency(RES1.Fields("MONEDERO"))
        listMonederos.TextMatrix(listMonederos.Rows - 1, 4) = RES1.Fields("FECHAHORA")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 5) = RES1.Fields("USUARIO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 6) = RES1.Fields("FOLIO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 7) = RES1.Fields("CLAVE")
        'cantidad = cantidad + 1
        RES1.MoveNext
    Loop

    listMonederos.Redraw = True

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "MONEDERO/APLICADOS"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(MONEDERO_ENTREGA)
    lista.TextMatrix(lista.Rows - 1, 2) = cantidad_entrega
    
    lista.Row = lista.Rows - 1
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 0
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "MONEDERO/GENERADOS"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(MONEDERO_RECIBE)
    lista.TextMatrix(lista.Rows - 1, 2) = cantidad_recibe
    
End Sub
Private Sub cargaFondo()
    Dim tipo As String
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE  date_format(CRTCAJA_FECHA, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " WHERE CRTCAJA_FECHA BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " WHERE CRTCAJA_FECHA BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " WHERE CRTCAJA_FECHA BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
        End If
    End If
    
'    tipo = ""
'    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
'        tipo = tipo & " where DATE_FORMAT(CRTCAJA_FECHA , '%Y-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ) "
'    Else
'        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " WHERE CRTCAJA_FECHA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA)) "
'        End If
'    End If
    
    sql1 = "SELECT crtcaja_monto, crtcaja_observaciones fROM CAT_CORTECAJA, SUCURSAL T5 where crtcaja_fecha = (select max(crtcaja_fecha) from cat_cortecaja  " & _
    tipo & ")"
    'MsgBox sql1
    '" where DATE_FORMAT(CRTCAJA_FECHA , '%x-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' )"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        txtFondo.Text = FormatCurrency(RES1.Fields("CRTCAJA_MONTO"))
        txtFondoObser.Text = RES1.Fields("CRTCAJA_OBSERVACIONES")
        fondoIni = Val(RES1.Fields("CRTCAJA_MONTO"))
    Else
        txtFondo.Text = "$0.00"
        txtFondoObser.Text = ""
        fondoIni = Val(0)
    End If
    
End Sub
Private Sub membresias()
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "where date_format(adquirio, '%x-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ORDER BY adquirio "
    Else
'        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                tipo = tipo & " where DATE_FORMAT(adquirio, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
'            Else
'                tipo = tipo & " where DATE_FORMAT(adquirio, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
'            End If
'
''            tipo = tipo & " where ADQUIRIO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) ORDER BY ADQUIRIO"
'        End If

        If tipoBusqueda = False Then
            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                tipo = tipo & " where adquirio BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
            Else
'                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
                tipo = tipo & " where adquirio BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
        Else
'            tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            tipo = tipo & " where ADQUIRIO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
        End If

    
    End If
    
    ListaMbr.Rows = 1
    sql1 = "SELECT clave_membresia, inicio, fin, adquirio, cliente, membresia, dias_mem, status, PRECIO " & _
    "FROM view_membresias_asignadas, SUCURSAL T5 " & _
    tipo
    '"where DATE_FORMAT(adquirio, '%x-%m-%d') BETWEEN  '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ORDER BY adquirio"
    Set RES1 = con.Execute(sql1)
    'Text1.Text = sql1
    ListaMbr.Redraw = False
    Do While Not RES1.EOF
        ListaMbr.AddItem ""
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 0) = RES1.Fields("CLAVE_MEMBRESIA")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 1) = RES1.Fields("INICIO")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 2) = RES1.Fields("FIN")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 3) = RES1.Fields("ADQUIRIO")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 4) = RES1.Fields("CLIENTE")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 5) = RES1.Fields("membresia")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 6) = RES1.Fields("DIAS_MEM")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 7) = RES1.Fields("STATUS")
        ListaMbr.TextMatrix(ListaMbr.Rows - 1, 8) = RES1.Fields("PRECIO")
    RES1.MoveNext
    Loop
    ListaMbr.Redraw = True
    
End Sub
Private Sub fondoInicial()
'    SQL1 = "SELECT SUC_CAJA_FONDO FROM SUCURSAL"
'    Set RES1 = con.Execute(SQL1)
'
'    fondoIni = 0
'    If Not RES1.EOF Then
'        If IsNull(RES1.Fields("SUC_CAJA_FONDO")) Then
'            txtFondo.Text = FormatCurrency(0)
'        Else
'            txtFondo.Text = FormatCurrency(RES1.Fields("SUC_CAJA_FONDO"))
'            fondoIni = RES1.Fields("SUC_CAJA_FONDO")
'        End If
'    End If
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "FONDO INICIAL"
    lista.TextMatrix(lista.Rows - 1, 1) = txtFondo.Text

End Sub
Private Sub asistenciaResumen()
Dim resSuc As Recordset
Dim numAst As Long
Dim difHoras As Long
'sstab1.Tab = 8
'MsgBox "Ok"
    
    sql1 = "select SUC_HORAENTRADA, SUC_HORASALIDA FROM SUCURSAL"
    Set resSuc = con.Execute(sql1)
    
       ' MsgBox (Date) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time") & " - " & (Date) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")
    
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "WHERE date_format(fecha_Hora, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
        Label4.Caption = "Horario: " & resSuc.Fields("SUC_HORAENTRADA") & " - " & resSuc.Fields("SUC_HORASALIDA")
        difHoras = DateDiff("h", ((Date) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")), ((Date) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")))
        
        tipoHorario = "A"
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " WHERE fecha_Hora BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
                    Label4.Caption = "Horario: " & dtFecha1(0) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time") & " - " & (dtFecha1(1) + 1) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")
                    difHoras = DateDiff("h", ((Date) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")), ((Date + 1) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")))
                    tipoHorario = "B"
                Else
                    tipo = tipo & " WHERE fecha_Hora BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                    Label4.Caption = "Horario: " & (dtFecha1(0) - 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time") & " - " & (dtFecha1(1)) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")
                    difHoras = DateDiff("h", ((Date - 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")), ((Date) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")))
                    tipoHorario = "C"
                End If
            Else
                tipo = tipo & " WHERE fecha_Hora BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                Label4.Caption = "Horario: " & (dtFecha1(0) - 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time") & " - " & (dtFecha1(1)) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")
                'difHoras = ((dtFecha1(1)) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")) - ((dtFecha1(0) - 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time"))
                difHoras = DateDiff("h", ((Date - 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")), ((Date) & " " & Format(resSuc.Fields("SUC_HORASALIDA"), "Short time")))
                tipoHorario = "C"
            End If
        End If
    End If
    sql1 = "SELECT * FROM VIEW_ASISTENCIAS, SUCURSAL T3 " & _
    tipo & " ORDER BY cliente ASC"
    Set RES1 = con.Execute(sql1)
    
    numAst = 0
    ListaAsts.Rows = 1
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "ASISTENCIAS"
    lista.TextMatrix(lista.Rows - 1, 2) = "0"
    
    Dim tipo_ast As Integer
    Dim DIA1 As Date
    Dim DIA2 As Date
    Dim Entrada As Date
    Dim salida As Date
    Dim SALIDA_DOS As Date
    Dim contador As Integer
    Dim FechaA As Date
    Dim FechaB As Date
    tipo_ast = 0
    tipo = "Entrada"
    'ListaAsts.Redraw = False
    Do While Not RES1.EOF
        
    
        contador = 0
        ListaAsts.AddItem ""
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 0) = RES1.Fields("Cliente")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1) = Format(RES1.Fields("Fecha_hora"), "Short Date")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 2) = Format(RES1.Fields("fecha_Hora"), "Short Time")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) = RES1.Fields("fecha_Hora")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 4) = RES1.Fields("SubTipo")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 5) = RES1.Fields("CLAVE_CLTE")
        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 7) = RES1.Fields("CLAVE_ASTS")
                        
        ListaAsts.Refresh
        
        ''''Checamos que sea el usuario
        If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 0) <> ListaAsts.TextMatrix(ListaAsts.Rows - 2, 0) Then
            ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "ENTRADA"
        Else
            If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1) = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) Then
                If tipoHorario = "A" Then
                    Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                    salida = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                    SALIDA_DOS = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                Else
                    Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                    DIA1 = Format(ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1), "dd/mm/yyyy")
                    salida = (DIA1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                    SALIDA_DOS = (DIA1 + 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                End If
                If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) >= Entrada Then
                    If ListaAsts.TextMatrix(ListaAsts.Rows - 2, 6) = "ENTRADA" Then
                        contador = 0
                        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "SALIDA"
                    Else
                            If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) <= SALIDA_DOS And ListaAsts.TextMatrix(ListaAsts.Rows - 2, 3) <= SALIDA_DOS Then
                                If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) >= Entrada And ListaAsts.TextMatrix(ListaAsts.Rows - 2, 3) >= Entrada Then
                                    ListaAsts.TextMatrix(ListaAsts.Rows - 2, 6) = "DESCARTADO"
                                    ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "SALIDA"
                                Else
                                    ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "ENTRADA"
                                End If
                            Else
                                ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "ENTRADA"
                            End If

                    End If
                Else
                    contador = contador + 1
                    ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "SALIDA"
                End If
            Else
                FechaA = ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1)
                FechaB = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1)
                'MsgBox FechaA & "   " & FechaB
                'If Format(ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1), "MM") > Format(ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1), "MM") Then
                If FechaA > FechaB Then
                    If tipoHorario = "A" Then
                        Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                        salida = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                    Else
                        Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                        DIA1 = Format(ListaAsts.TextMatrix(ListaAsts.Rows - 2, 1), "dd/mm/yyyy")
                        salida = (DIA1 + 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
                    End If
                    If ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) >= Entrada And ListaAsts.TextMatrix(ListaAsts.Rows - 1, 3) <= salida Then
                        contador = contador + 1
                        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "SALIDA"
                    Else
                        contador = 0
                        ListaAsts.TextMatrix(ListaAsts.Rows - 1, 6) = "ENTRADA"
                    End If
                Else
                    ''''''
                End If
            End If
                         
        End If
                         
        If RES1.Fields("TIPO_AS") = "ENTRADA" Then
            numAst = numAst + 1
        End If
        
        
        RES1.MoveNext
    Loop
    ListaAsts.Redraw = True
    
      
    'If Not RES1.EOF Then
    lista.TextMatrix(lista.Rows - 1, 2) = numAst
    'End If
    
    asistenciaResumen_Fila

    asistenciaResumen_Columna
    
End Sub
Private Sub asistenciaResumen_Columna()
    Dim filas As Integer
    Dim totHoras As Double
   ' SSTab1.Tab = 8
    ListaAsts2.Rows = 1
    
    ListaAsts2.Cols = 2
    ListaAsts2.TextMatrix(0, 0) = "Usuario"
    ListaAsts2.ColWidth(0) = 2500
    ListaAsts2.TextMatrix(0, 1) = "Total Hrs"
    ListaAsts2.ColWidth(0) = 1500
    filas = 0
    totHoras = 0
    With ListaAsts3
        For b1 = 1 To .Rows - 1
            If .TextMatrix(b1, 0) <> ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 0) Then
                ListaAsts2.AddItem ""
                ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 0) = .TextMatrix(b1, 0)
                filas = 0
                If totHoras <> 0 Then
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 2, 1) = totHoras
                    totHoras = 0
                End If
                If ListaAsts2.Cols = 2 Then
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Entrada"
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 1)
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 2500
                    
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Salida"
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 2500
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 2)
                    
                    
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Horas"
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 700
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 3)
                    totHoras = totHoras + Val(.TextMatrix(b1, 3))
                Else
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 2) = .TextMatrix(b1, 1)
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 3) = .TextMatrix(b1, 2)
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 4) = .TextMatrix(b1, 3)
                    totHoras = totHoras + Val(.TextMatrix(b1, 3))
                
                End If
            Else
                If filas >= 5 Then
                    filas = filas + 3
                Else
                    filas = filas + 5
                End If
                
                If ListaAsts2.Cols <= filas Then
                'If ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, filas) <> "" Then
                
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Entrada"
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 1)
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 2500
                    
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Salida"
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 2500
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 2)
                    
                    ListaAsts2.Cols = ListaAsts2.Cols + 1
                    ListaAsts2.TextMatrix(0, ListaAsts2.Cols - 1) = "Horas"
                    ListaAsts2.ColWidth(ListaAsts2.Cols - 1) = 700
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, ListaAsts2.Cols - 1) = .TextMatrix(b1, 3)
                    totHoras = totHoras + Val(.TextMatrix(b1, 3))
                Else
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, filas) = .TextMatrix(b1, (filas - filas + 1))
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, filas + 1) = .TextMatrix(b1, (filas - filas + 2))
                    ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, filas + 2) = .TextMatrix(b1, (filas - filas + 3))
                    totHoras = totHoras + Val(.TextMatrix(b1, 3))
                
                End If
            
            
            End If
            
            
        Next b1
        If totHoras <> 0 Then
            ListaAsts2.TextMatrix(ListaAsts2.Rows - 1, 1) = totHoras
            totHoras = 0
        End If
    End With


End Sub

Private Sub asistenciaResumen_Fila()
On Error Resume Next
    Dim Entrada As Date
    Dim salida As Date
    Dim DIA1 As Date
    Dim HORAS As String
    Dim HORAS2 As Double
    Dim Minutos As String
    
    
    sql1 = "select SUC_HORAENTRADA, SUC_HORASALIDA FROM SUCURSAL"
    Set resSuc = con.Execute(sql1)
    ListaAsts3.Redraw = False
    ListaAsts3.Rows = 1
    Minutos = "0"
    With ListaAsts
        For b1 = 1 To .Rows - 1
            If .TextMatrix(b1, 6) = "ENTRADA" Then
                ListaAsts3.AddItem ""
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 0) = .TextMatrix(b1, 0)
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 1) = .TextMatrix(b1, 3)
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 4) = Format(.TextMatrix(b1, 3), "Short Date")
            Else
                If .TextMatrix(b1, 6) = "SALIDA" Then
                    ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 2) = .TextMatrix(b1, 3)
                    ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3) = .TextMatrix(b1, 8)
                End If
            End If
            
            ''''pARA ACOMPLETAR LA FECHA
            DIA1 = Format(ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 1), "dd/mm/yyyy")
            If tipoHorario = "A" Then
                Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                salida = ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
            Else
                Entrada = ListaAsts.TextMatrix(ListaAsts.Rows - 1, 1) & " " & Format(resSuc.Fields("SUC_HORAENTRADA"), "Short time")
                DIA1 = DIA1 + 1
                salida = (DIA1) & " " & Format(resSuc.Fields("SUC_HORAsalida"), "Short time")
            End If
            
            If ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 2) = "" Then
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 2) = salida
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 5) = "INCOMPLETO"
            Else
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 5) = "COMPLETO"
            End If
            'MsgBox DateDiff("n", Format(ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 1), "dd/mm/yy hh:mm"), Format(ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 2), "dd/mm/yy hh:mm"))
            HORAS = DateDiff("n", Format(ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 1), "dd/mm/yy hh:mm"), Format(ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 2), "dd/mm/yy hh:mm"))
            HORAS = (HORAS / 60)
            HORAS = Format(HORAS, "##.##")
            If Right(HORAS, 1) = "." Then
                HORAS = Replace(HORAS, ".", "")
            End If
            If Left(Right(HORAS, 3), 1) = "." Then
                Minutos = Right(HORAS, 2)
                Minutos = Format(((Minutos * 60) / 100), "##")
                If Val(Minutos) < 10 Then
                    Minutos = Format(Minutos, "00")
                End If
                HORAS = Left(HORAS, Len(HORAS) - 3)
            Else
                If Left(Right(HORAS, 2), 1) = "." Then
                    Minutos = Right(HORAS, 1)
                    Minutos = Format(Minutos, "00")
                    HORAS = Left(HORAS, Len(HORAS) - 2)
                End If
            End If
            If Minutos > 0 Then
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3) = HORAS & "." & Minutos
            Else
                ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3) = HORAS
            End If
            If ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 5) = "INCOMPLETO" Then
                'Horas2 = ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3)
                If ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3) > 8 Then
                'If Horas2 > 8 Then
                    ListaAsts3.TextMatrix(ListaAsts3.Rows - 1, 3) = 8
                End If
                ListaAsts3.Row = ListaAsts3.Rows - 1
                ListaAsts3.Col = 5
                ListaAsts3.CellForeColor = vbRed
                ListaAsts3.Col = 2
                ListaAsts3.CellForeColor = vbRed
                ListaAsts3.Col = 3
                ListaAsts3.CellForeColor = vbRed
            Else
                ListaAsts3.Row = ListaAsts3.Rows - 1
                ListaAsts3.Col = 5
                ListaAsts3.CellForeColor = vbBlack
                ListaAsts3.Col = 2
                ListaAsts3.CellForeColor = vbBlack
                ListaAsts3.Col = 3
                ListaAsts3.CellForeColor = vbBlack

            End If
                    
        
        Next b1
    End With
    ListaAsts3.Redraw = True

End Sub
Private Sub gastos()
    Dim totGST As Double
    Dim cantGST As Long
    
    totGST = 0
    cantGST = 0
    totGastos = 0
    
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE date_format(FECHA_HORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " where fecha_Hora BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
'                    tipo = tipo & " where vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " where fecha_Hora BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " where fecha_hora BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If
        End If
        
'        MsgBox tipo
        
'        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                tipo = tipo & " WHERE date_format(FECHA_HORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
'            Else
'                tipo = tipo & " WHERE date_format(FECHA_HORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
'            End If
'
'            'tipo = tipo & " WHERE FECHA_HORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
'        End If
    End If
    
    
    sql1 = "SELECT SUM(GASTO) TOTAL, COUNT(ID) CANT_TOTAL " & _
    "FROM VIEW_GASTOS, sucursal T3 " & _
    tipo & " AND CAJA = 'SI'"
    'MsgBox SQL1
'    "WHERE date_format(FECHA_HORA, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    'txtFondoObser.Text = SQL1
    'MsgBox SQL1
    Set RES1 = con.Execute(sql1)
    
    If IsNull(RES1.Fields("TOTAL")) Then
        totGST = 0
    Else
        totGST = RES1.Fields("TOTAL")
    End If
                
    totGastos = totGST
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "GASTOS"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totGST)
    lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("cant_total")
    
    lista.Row = lista.Rows - 1
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 0
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellForeColor = vbRed


    sql1 = "select TIPO_GASTO,  SUM(GASTO) GASTO " & _
    "FROM VIEW_GASTOS, sucursal T3 " & _
     tipo & " GROUP BY TIPO_GASTO " & _
     "Union select 'TOTAL' TIPO_GASTO,  SUM(GASTO) GASTO " & _
    "From VIEW_GASTOS, sucursal T3  " & _
    tipo

    
    Set RES1 = con.Execute(sql1)
    
    listaGST2.Redraw = False
    listaGST2.Rows = 1
    Do While Not RES1.EOF
        listaGST2.AddItem ""
        listaGST2.TextMatrix(listaGST2.Rows - 1, 0) = RES1.Fields("TIPO_GASTO")
        listaGST2.TextMatrix(listaGST2.Rows - 1, 1) = FormatCurrency(RES1.Fields("GASTO")) & ""
        RES1.MoveNext
    Loop
    listaGST2.Redraw = True
       
    sql1 = "SELECT * " & _
    "FROM VIEW_GASTOS, SUCURSAL T3 " & _
    tipo
    '"date_format(FECHA_HORA, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    listaGST.Rows = 1
    listaGST.Redraw = False
    Do While Not RES1.EOF
        listaGST.AddItem ""
        listaGST.TextMatrix(listaGST.Rows - 1, 0) = RES1.Fields("ID")
        listaGST.TextMatrix(listaGST.Rows - 1, 1) = RES1.Fields("FECHA_HORA")
        listaGST.TextMatrix(listaGST.Rows - 1, 2) = RES1.Fields("FECHA_FIN")
        listaGST.TextMatrix(listaGST.Rows - 1, 3) = RES1.Fields("USUARIO")
        listaGST.TextMatrix(listaGST.Rows - 1, 4) = RES1.Fields("TIPO_GRAL")
        listaGST.TextMatrix(listaGST.Rows - 1, 5) = RES1.Fields("TIPO_GASTO")
        listaGST.TextMatrix(listaGST.Rows - 1, 6) = FormatCurrency(RES1.Fields("GASTO"))
        listaGST.TextMatrix(listaGST.Rows - 1, 7) = RES1.Fields("COMPROBANTE")
        listaGST.TextMatrix(listaGST.Rows - 1, 8) = RES1.Fields("CODIGO") & ""
        listaGST.TextMatrix(listaGST.Rows - 1, 9) = RES1.Fields("CAJA")
        listaGST.TextMatrix(listaGST.Rows - 1, 10) = RES1.Fields("PROVEEDOR") & ""
        listaGST.TextMatrix(listaGST.Rows - 1, 11) = RES1.Fields("GST_DESCRIPCION")
        
        listaGST.Row = listaGST.Rows - 1
        If RES1.Fields("CAJA") = "SI" Then
            listaGST.Col = 5
            listaGST.CellBackColor = vbRed
            listaGST.CellForeColor = vbWhite
            listaGST.Col = 8
            listaGST.CellBackColor = vbRed
            listaGST.CellForeColor = vbWhite
        End If
        
        RES1.MoveNext
    Loop
    listaGST.Redraw = True


End Sub

Private Sub consumoInterno()
    Dim totCI As Double
    Dim cantCI As Long
    
    totCI = 0
    cantCI = 0
        
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE date_format(CSI_FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
             If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " WHERE date_format(CSI_FECHAHORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
            Else
                tipo = tipo & " WHERE date_format(CSI_FECHAHORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If
               
'            tipo = tipo & " WHERE CSI_FECHAHORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If
        
    sql1 = "SELECT SUM(CSI_PRECIO * CSI_CANTIDAD) CI, COUNT(CSI_ID) CANTIDAD " & _
    "FROM CONSUMO_INTERNO T1, SUCURSAL T3 " & _
    tipo
'    "WHERE date_format(CSI_FECHAHORA, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    
    If IsNull(RES1.Fields("CI")) Then
        totCI = 0
    Else
        totCI = RES1.Fields("CI")
    End If
                
                
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "CONSUMO INT"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totCI)
    lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("CANTIDAD")
    
    lista.Row = lista.Rows - 1
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 0
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellForeColor = vbRed

    
    sql1 = "SELECT CSI_ID,  CSI_FECHAHORA, CSI_CANTIDAD, CSI_PRECIO, T2.PROD_NOMBRE, T2.PROD_CODIGO, " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) ATENDIO," & _
    "CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) USUARIO " & _
    "FROM CONSUMO_iNTERNO T1, PRODUCTOS T2, PERSONA T3, PERSONA T4 " & _
    "Where CSI_PRODUCTO_ID = PROD_ID And CSI_PRODUCTO_SERV = PROD_sERV " & _
    "AND CSI_VEND_PERID = T3.PER_ID AND CSI_USER_PERID = T4.PER_ID " & _
    "AND date_format(CSI_FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    
    listaCI.Redraw = False
    listaCI.Rows = 1
    Do While Not RES1.EOF
        listaCI.AddItem ""
        listaCI.TextMatrix(listaCI.Rows - 1, 0) = RES1.Fields("CSI_ID")
        listaCI.TextMatrix(listaCI.Rows - 1, 1) = RES1.Fields("CSI_FECHAHORA")
        listaCI.TextMatrix(listaCI.Rows - 1, 2) = RES1.Fields("USUARIO")
        listaCI.TextMatrix(listaCI.Rows - 1, 3) = RES1.Fields("PROD_NOMBRE")
        listaCI.TextMatrix(listaCI.Rows - 1, 4) = RES1.Fields("PROD_CODIGO")
        listaCI.TextMatrix(listaCI.Rows - 1, 5) = RES1.Fields("CSI_CANTIDAD")
        listaCI.TextMatrix(listaCI.Rows - 1, 6) = FormatCurrency(RES1.Fields("CSI_PRECIO"))
        listaCI.TextMatrix(listaCI.Rows - 1, 7) = RES1.Fields("ATENDIO")
        RES1.MoveNext
    Loop
    listaCI.Redraw = True


End Sub
Private Sub detalleVenta()
    Dim filaFolio As String
    Dim tipofila As String
    Dim tipo As String
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
            End If
        End If
    End If
        
    sql1 = "SELECT T1.VENT_IDFOLIO, (VENDET_PRECIO * VENDET_CANTIDAD) TOTAL, IF(VENDET_PRODSERV = 'P', 'PRODUCTOS', 'SERVICIO') TIPO, " & _
    "VENDET_PRODSERV, CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CLIENTE, " & _
    "VENT_FECHAHORA_COBRO, VENT_SUBTOTAL, VENT_DESCUENTO, VENT_TOTAL, VENT_PAGADO, VENT_CAMBIO, VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE,  " & _
    "VENDET_PRODCODIGO, VENDET_PRODUCTONOMBRE, VENDET_CANTIDAD, VENDET_PRECIO, VENDET_DESCUENTO, (VENDET_CANTIDAD * VENDET_PRECIO - VENDET_DESCUENTO) PROD_TOT " & _
    "FROM VENTAS T1, VENTA_DETALLE T2, PERSONA T3, PERSONA T4, SUCURSAL T5 " & _
    "WHERE T1.VENT_IDFOLIO = T2.VENDET_fOLIO AND VENDET_PRODSERV IN ('P', 'S') AND VENT_STATUS = 'P' " & _
    "AND VENT_VENDPERID = T3.PER_ID AND VENT_CLIEPERID = T4.PER_ID AND T2.VENDET_STATUS = 'A' " & _
    tipo
    '"AND date_format(vent_fechaHora_cobro, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    'MsgBox tipo
    
    Set RES1 = con.Execute(sql1)
    
    
    Lista3.MergeCells = flexMergeRestrictColumns
    Lista3.Rows = 1
    Lista3.Redraw = False
    tipofila = "1"
    
    Do While Not RES1.EOF
        filaFolio = RES1.Fields("VENT_IDFOLIO")
        If Lista3.Rows - 1 > 0 Then
            If filaFolio <> Lista3.TextMatrix(Lista3.Rows - 1, 0) Then
                Lista3.AddItem ""
                Lista3.RowHeight(Lista3.Rows - 1) = 0
                If tipofila = "1" Then
                    tipofila = "2"
                Else
                    tipofila = "1"
                End If
            End If
        End If
        
        Lista3.AddItem ""
        Lista3.TextMatrix(Lista3.Rows - 1, 0) = RES1.Fields("VENT_IDFOLIO")
        Lista3.TextMatrix(Lista3.Rows - 1, 1) = RES1.Fields("USUARIO")
        Lista3.TextMatrix(Lista3.Rows - 1, 2) = RES1.Fields("CLIENTE")
        Lista3.TextMatrix(Lista3.Rows - 1, 3) = RES1.Fields("VENT_FECHAHORA_COBRO")
        Lista3.TextMatrix(Lista3.Rows - 1, 4) = FormatCurrency(RES1.Fields("VENT_SUBTOTAL"))
        Lista3.TextMatrix(Lista3.Rows - 1, 5) = FormatCurrency(RES1.Fields("VENT_DESCUENTO"))
        Lista3.TextMatrix(Lista3.Rows - 1, 6) = FormatCurrency(RES1.Fields("VENT_TOTAL"))
        Lista3.TextMatrix(Lista3.Rows - 1, 7) = FormatCurrency(RES1.Fields("VENT_PAGADO"))
        Lista3.TextMatrix(Lista3.Rows - 1, 8) = FormatCurrency(RES1.Fields("VENT_CAMBIO"))
        Lista3.TextMatrix(Lista3.Rows - 1, 9) = FormatCurrency(Val(RES1.Fields("VENT_PAGOEFECTIVO")) - Val(RES1.Fields("VENT_CAMBIO")))
        Lista3.TextMatrix(Lista3.Rows - 1, 10) = FormatCurrency(RES1.Fields("VENT_PAGOTARJETA"))
        Lista3.TextMatrix(Lista3.Rows - 1, 11) = FormatCurrency(RES1.Fields("VENT_PAGOCHEQUE"))
        Lista3.TextMatrix(Lista3.Rows - 1, 12) = RES1.Fields("TIPO")
        Lista3.TextMatrix(Lista3.Rows - 1, 13) = RES1.Fields("VENDET_PRODCODIGO")
        Lista3.TextMatrix(Lista3.Rows - 1, 14) = RES1.Fields("VENDET_PRODUCTONOMBRE")
        Lista3.TextMatrix(Lista3.Rows - 1, 15) = RES1.Fields("VENDET_CANTIDAD")
        Lista3.TextMatrix(Lista3.Rows - 1, 16) = FormatCurrency(RES1.Fields("VENDET_PRECIO"))
        Lista3.TextMatrix(Lista3.Rows - 1, 17) = FormatCurrency(RES1.Fields("VENDET_DESCUENTO"))
        Lista3.TextMatrix(Lista3.Rows - 1, 18) = FormatCurrency(RES1.Fields("PROD_TOT"))
        
        Lista3.Row = Lista3.Rows - 1
        
        If tipofila = "2" Then
            For b1 = 0 To 18
                Lista3.Col = b1
                Lista3.CellBackColor = &HFFFFC0
            Next b1
        End If
        
        If RES1.Fields("VENDET_dESCUENTO") > 0 Then
            For b1 = 13 To 18
                Lista3.Col = b1
                Lista3.CellBackColor = &H80C0FF
            Next b1
        End If
        'Lista3.MergeCol(0) = True
        For b1 = 0 To 10
            Lista3.MergeCol(b1) = True
        Next b1
    
        RES1.MoveNext
    Loop
    Lista3.Redraw = True
    

End Sub
Private Sub ventaGroup()
    lista5.Rows = 1
    
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' " & _
    "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
    Else
'        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA) " & _
'                "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
'            Else
'                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA)" & _
'                "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
'            End If
'        End If
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA) " & _
                    "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
                Else
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA) " & _
                    "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
                End If
            Else
                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA) " & _
                "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
            End If
        End If
    
    End If
    
    sql1 = "SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) TOTAL, IF(VENDET_PRODSERV = 'P', 'PRODUCTO', (IF(VENDET_PRODSERV = 'S', 'SERVICIO', (IF (VENDET_PRODSERV= 'R', 'MONEDERO', 'MEMBRESIA'))))  ) TIPO, " & _
    "VENDET_PRODSERV,  SUM(VENDET_CANTIDAD) CANT,   VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO, " & _
    "(SELECT IF(T3.PROD_CANT <= 0, 0, T3.PROD_CANT) FROM PRODUCTOS T3 WHERE T3.PROD_ID = T2.venDet_ProductoId and T3.prod_serv = T2.venDet_ProdServ) PROD_CANT " & _
    "FROM VENTAS T1, VENTA_DETALLE T2, SUCURSAL T4 " & _
    "WHERE T1.VENT_IDFOLIO = T2.VENDET_fOLIO AND VENDET_PRODSERV IN ('P', 'S', 'M', 'R') AND VENT_STATUS = 'P' and T2.VENDET_STATUS = 'A' " & _
    tipo
'    "AND date_format(vent_fechaHora_cobro, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' " & _
'    "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) , VENDET_PRODUCTONOMBRE, VENDET_PRECIO, VENDET_PRODCODIGO "
    Set RES1 = con.Execute(sql1)
    
    lista5.Redraw = False
    Do While Not RES1.EOF
        lista5.AddItem ""
        lista5.TextMatrix(lista5.Rows - 1, 0) = RES1.Fields("TIPO")
        lista5.TextMatrix(lista5.Rows - 1, 1) = RES1.Fields("VENDET_PRODUCTONOMBRE")
        lista5.TextMatrix(lista5.Rows - 1, 2) = RES1.Fields("VENDET_PRODCODIGO")
        lista5.TextMatrix(lista5.Rows - 1, 3) = FormatCurrency(RES1.Fields("VENDET_PRECIO"))
        lista5.TextMatrix(lista5.Rows - 1, 4) = RES1.Fields("CANT")
        lista5.TextMatrix(lista5.Rows - 1, 5) = FormatCurrency(RES1.Fields("TOTAL"))
        If IsNull(RES1.Fields("PROD_CANT")) Then
            lista5.TextMatrix(lista5.Rows - 1, 6) = "0"
        Else
            lista5.TextMatrix(lista5.Rows - 1, 6) = RES1.Fields("PROD_CANT")
        End If
    RES1.MoveNext
    Loop
    lista5.Redraw = True

End Sub
Private Sub pagosUsuarios()
    Dim tipo2 As String
    tipo = ""
    tipo2 = ""
    
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " AND date_format(T1.vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' " & _
    "GROUP BY CONCAT(t3.PER_NOMBRE, ' ', t3.PER_PATERNO, ' ', t3.PER_MATERNO), T2.VENDET_VENDPERID, T2.VENDET_VENDtipoid, T2.VENDET_VENDtipo  "
    
        tipo2 = tipo2 & " AND date_format(TV1.APPG_FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If tipoBusqueda = False Then
        
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA) " & _
                "GROUP BY CONCAT(t3.PER_NOMBRE, ' ', t3.PER_PATERNO, ' ', t3.PER_MATERNO), T2.VENDET_VENDPERID, T2.VENDET_VENDtipoid, T2.VENDET_VENDtipo  "
                    
                Else
                    tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA)" & _
                "GROUP BY CONCAT(t3.PER_NOMBRE, ' ', t3.PER_PATERNO, ' ', t3.PER_MATERNO), T2.VENDET_VENDPERID, T2.VENDET_VENDtipoid, T2.VENDET_VENDtipo  "
                    
                End If
            Else
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T4.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T4.SUC_HORASALIDA)" & _
                "GROUP BY CONCAT(t3.PER_NOMBRE, ' ', t3.PER_PATERNO, ' ', t3.PER_MATERNO), T2.VENDET_VENDPERID, T2.VENDET_VENDtipoid, T2.VENDET_VENDtipo  "
            End If
        
'            tipo = tipo & " AND T1.VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T4.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T4.SUC_HORASALIDA) " & _
'            "GROUP BY CONCAT(t3.PER_NOMBRE, ' ', t3.PER_PATERNO, ' ', t3.PER_MATERNO), T2.VENDET_VENDPERID  "
        
            tipo2 = tipo2 & " AND TV1.APPG_FECHAHORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T4.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T4.SUC_HORASALIDA) "
        End If
    End If

    sql1 = "SELECT SUM(T2.VENDET_PRECIO * T2.VENDET_CANTIDAD) TOTAL, " & _
    "SUM(IF(T2.VENDET_PRODSERV = 'P', (T2.VENDET_PRECIO * T2.VENDET_CANTIDAD), 0)) PRODUCTOS," & _
    "SUM(IF(T2.VENDET_PRODSERV = 'S', (T2.VENDET_PRECIO * T2.VENDET_CANTIDAD), 0)) SERVICIOS,     SUM(T2.VENDET_CANTIDAD) CANT, " & _
    "SUM(IF(T2.VENDET_PRODSERV = 'P', (T2.VENDET_CANTIDAD), 0)) PROD_CANT, " & _
    "SUM(IF(T2.VENDET_PRODSERV = 'S', (T2.VENDET_CANTIDAD), 0)) SERV_CANT, " & _
    "T2.VENDET_VENDPERID, T2.VENDET_VENDtipoid, T2.VENDET_VENDtipo, CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO,   " & _
    "(SELECT SUM(TV1.APPG_PAGO) FROM PAGOS_APARTADOS TV1 WHERE TV1.PAPG_MOSTPERID = T2.VENDET_VENDPERID " & tipo2 & ") APARTADOS " & _
    "FROM VENTA_DETALLE T2, PERSONA T3, VENTAS T1, SUCURSAL T4 " & _
    "WHERE T2.VENDET_PRODSERV IN ('P', 'S') AND T3.PER_ID = T2.VENDET_VENDPERID AND T1.VENT_IDFOLIO = T2.VENDET_fOLIO AND T1.VENT_STATUS = 'P' and T2.VENDET_STATUS = 'A'" & _
    tipo
    'MsgBox SQL1
    Set RES1 = con.Execute(sql1)
    
'      Text1.Text = SQL1
    
    lista2.ColWidth(9) = 0
    lista2.ColWidth(10) = 0
    
    lista2.Rows = 1
    
    totServicios = 0
    totProductos = 0
    lista2.Redraw = False
    Do While Not RES1.EOF
        lista2.AddItem ""
        lista2.TextMatrix(lista2.Rows - 1, 0) = RES1.Fields("USUARIO")
        lista2.TextMatrix(lista2.Rows - 1, 1) = RES1.Fields("VENDET_VENDPERID")
        lista2.TextMatrix(lista2.Rows - 1, 2) = FormatCurrency(RES1.Fields("PRODUCTOS"))
        lista2.TextMatrix(lista2.Rows - 1, 3) = RES1.Fields("PROD_CANT")
        lista2.TextMatrix(lista2.Rows - 1, 4) = FormatCurrency(RES1.Fields("SERVICIOS"))
        lista2.TextMatrix(lista2.Rows - 1, 5) = RES1.Fields("SERV_CANT")
        lista2.TextMatrix(lista2.Rows - 1, 6) = FormatCurrency(RES1.Fields("APARTADOS"))
        If IsNull(RES1.Fields("APARTADOS")) Then
                lista2.TextMatrix(lista2.Rows - 1, 7) = FormatCurrency(Val(RES1.Fields("PRODUCTOS")))
        Else
            lista2.TextMatrix(lista2.Rows - 1, 7) = FormatCurrency(Val(RES1.Fields("PRODUCTOS")) + Val(RES1.Fields("APARTADOS")))
        End If
        lista2.Row = lista2.Rows - 1
        lista2.Col = 8
        lista2.CellFontName = "Wingdings"
        lista2.CellFontBold = True
        lista2.CellFontSize = 16
        lista2.TextMatrix(lista2.Rows - 1, 8) = Chr(254)
        
        lista2.TextMatrix(lista2.Rows - 1, 9) = RES1.Fields("VENDET_VENDTIPOID")
        lista2.TextMatrix(lista2.Rows - 1, 10) = RES1.Fields("VENDET_VENDTIPO")
        
        RES1.MoveNext
    Loop
    lista2.Redraw = True

End Sub
Private Sub pagosTotal()
    
        lista.AddItem ""
        lista.Row = lista.Rows - 1
        lista.Col = 0
        lista.CellFontBold = True
        lista.TextMatrix(lista.Rows - 1, 0) = "SUB TOTAL"
        lista.Col = 1
        lista.CellFontBold = True
        lista.CellFontSize = 18
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totProductos + totServicios + totMebresias + fondoIni + totApartados + totCambios)
                
        lista.AddItem ""
        lista.Row = lista.Rows - 1
        lista.Col = 0
        lista.CellFontBold = True
        lista.CellFontSize = 24
        lista.TextMatrix(lista.Rows - 1, 0) = "TOTAL"
        lista.Col = 1
        lista.CellFontBold = True
        lista.CellFontSize = 24
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totProductos + totServicios + totMebresias - totDescuentos - totGastos + fondoIni + totApartados + totCambios - totMonederos)  '- totPagosUsuarios)
    
        'MsgBox "prod " & totProductos & " Serv " & totServicios & " mem " & totMebresias & " desc " & totDescuentos & " gast " & totGastos & " fondo " & fondoIni & " apart " & totApartados & " camb " & totCambios & " mone " & totMonederos

End Sub
Private Sub pagosDescuento()
    Dim cantDescuentos As Double
    
    totDescuentos = 0
    cantDescuentos = 0
        
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
            Else
                tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If

'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If
        
        
    sql1 = "SELECT SUM(VENT_DESCUENTO) DESCUENTO, COUNT(VENT_DESCUENTO) CANTIDAD " & _
    "FROM VENTAS T1, SUCURSAL T3 WHERE T1.VENT_STATUS = 'P' AND VENT_DESCUENTO > 0 " & _
    tipo
    '"AND date_format(vent_fechaHora_cobro, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    
    If IsNull(RES1.Fields("DESCUENTO")) Then
        totDescuentos = 0
    Else
        totDescuentos = RES1.Fields("DESCUENTO")
    End If
                
                
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "DESCUENTO"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totDescuentos)
    lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("CANTIDAD")
    
    lista.Row = lista.Rows - 1
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 0
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellForeColor = vbRed
End Sub
Private Sub pagosTipo()
    
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
'        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
'            Else
'                tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
'            End If
'
''            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
'        End If

        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T5.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T5.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If
        End If

    End If
        
    sql1 = "SELECT SUM(VENT_PAGOEFECTIVO - VENT_CAMBIO) EFECTIVO, SUM(VENT_PAGOTARJETA) TARJETA, SUM(VENT_PAGOCHEQUE) CHEQUE " & _
    "FROM VENTAS T1, SUCURSAL T3 WHERE T1.VENT_STATUS IN ('P', 'A', 'B') " & _
    tipo
    '"AND date_format(vent_fechaHora_cobro, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
        
    lista.Redraw = False
    If Not RES1.EOF Then
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = "EFECTIVO"
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Val(Format(RES1.Fields("EFECTIVO"), "General Number")))
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = "TARJETA"
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Val(Format(RES1.Fields("TARJETA"), "General Number")))
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = "CHEQUE"
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Val(Format(RES1.Fields("CHEQUE"), "General Number")))
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = "EFECTIVO TOTAL"
        lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Val(Format(RES1.Fields("EFECTIVO"), "General Number")) + Val(Format(txtFondo.Text, "General Number")) - Val(Format(totGastos, "General Number")))
    End If
    lista.Redraw = True

End Sub
Private Sub pagosCambios()
Dim cantCmbs As Long
totCambios = 0
catCmbs = 0

    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " WHERE date_format(FECHA_DEVO, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            
            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " WHERE date_format(FECHA_DEVO, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
            Else
                tipo = tipo & " WHERE date_format(FECHA_DEVO, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If
            
'            tipo = tipo & "WHERE FECHA_DEVO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If

    sql1 = "SELECT   T1.TOT_DIF TOTAL " & _
    "FROM VIEW_CAMBIOS T1, SUCURSAL T3 " & _
    tipo
'    "WHERE date_format(FECHA_DEVO, '%x-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        If Val(RES1.Fields("total")) > 0 Then
            totCambios = totCambios + Val(RES1.Fields("TOTAL"))
        End If
        cantCmbs = cantCmbs + 1
        RES1.MoveNext
    Loop
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Cambios"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totCambios)
    lista.TextMatrix(lista.Rows - 1, 2) = cantCmbs

End Sub


Private Sub pagosApartados()
Dim cantApar As Long
totApartados = 0
catApar = 0


    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "WHERE date_format(FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                tipo = tipo & " WHERE date_format(FECHAHORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
            Else
                tipo = tipo & " WHERE date_format(FECHAHORA, '%Y-%m-%d') BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If

'            tipo = tipo & "WHERE FECHAHORA BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If

    sql1 = "SELECT T1.PAGO " & _
    "FROM VIEW_pagos_APARTADO T1, SUCURSAL T3 " & _
    tipo
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        totApartados = totApartados + Val(RES1.Fields("PAGO"))
        cantApar = cantApar + 1
        RES1.MoveNext
    Loop
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Apartados"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totApartados)
    lista.TextMatrix(lista.Rows - 1, 2) = cantApar

End Sub

Private Sub cargaApartados()
    'On Error Resume Next
    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & " date_format(fecha_HORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
        'tipo = tipo & " fecha BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            tipo = tipo & " date_format(fecha_HORA, '%Y-%m-%d') BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
        End If
    End If
    
    tipo = tipo & " ORDER by FECHA DESC "
    
    listaApt.Rows = 1
    
    sql1 = "SELECT * fROM VIEW_PAGOS_APARTADO T1, SUCURSAL T3 WHERE " & tipo
    Set RES1 = con.Execute(sql1)
    
    listaApt.Redraw = False
    Do While Not RES1.EOF
        listaApt.AddItem ""
        listaApt.TextMatrix(listaApt.Rows - 1, 0) = RES1.Fields("FOLIO_APRT")
        listaApt.TextMatrix(listaApt.Rows - 1, 1) = RES1.Fields("FOLIO_VENTA")
        listaApt.TextMatrix(listaApt.Rows - 1, 2) = RES1.Fields("PRODUCTO")
        listaApt.TextMatrix(listaApt.Rows - 1, 3) = RES1.Fields("CODIGO")
        listaApt.TextMatrix(listaApt.Rows - 1, 4) = FormatCurrency(RES1.Fields("PRECIO_pROD"))
        listaApt.TextMatrix(listaApt.Rows - 1, 5) = FormatCurrency(RES1.Fields("DESCUENTO_PROD"))
        listaApt.TextMatrix(listaApt.Rows - 1, 6) = FormatCurrency(RES1.Fields("TOTAL_PROD"))
        listaApt.TextMatrix(listaApt.Rows - 1, 7) = RES1.Fields("CLIENTE")
        listaApt.TextMatrix(listaApt.Rows - 1, 8) = RES1.Fields("FECHA")
        listaApt.TextMatrix(listaApt.Rows - 1, 9) = FormatCurrency(RES1.Fields("PAGO"))
        RES1.MoveNext
    Loop
    listaApt.Redraw = True
    
End Sub
Private Sub pagosProdServ()
    
Dim cantSer As Long, cantProd As Long, cantMem As Long

    tipo = ""
    If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
        tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' "
    Else
        If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
            'tipo = tipo & "AND date_format(vent_fechaHora_cobro, '%Y-%m-%d') BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
            If tipoBusqueda = False Then
                If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
                    tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA) "
                Else
                    tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
                End If
            Else
                tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T3.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T3.SUC_HORASALIDA)"
            End If

'            tipo = tipo & " AND VENT_FECHAHORA_COBRO BETWEEN CONCAT('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', ' ', T3.SUC_HORAENTRADA) AND CONCAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), ' ', T3.SUC_HORASALIDA) "
            
        End If
    End If

    tipo = tipo & "GROUP BY IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) "

    sql1 = "SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) TOTAL, IF(VENDET_PRODSERV = 'P', 'PRODUCTO', IF(VENDET_PRODSERV = 'S', 'SERVICIO', 'MEMBRESIA')) TIPO, " & _
    "VENDET_PRODSERV, VENT_STATUS, SUM(VENDET_CANTIDAD) CANT, SUM(VENT_PERSONAS) PERSONAS " & _
    "FROM VENTAS T1, VENTA_DETALLE T2, SUCURSAL T3 " & _
    "WHERE T1.VENT_IDFOLIO = T2.VENDET_fOLIO AND VENDET_PRODSERV IN ('P', 'S', 'M') AND VENT_STATUS IN ('P', 'A') and T2.VENDET_STATUS = 'A'" & _
    tipo
    
    Set RES1 = con.Execute(sql1)
    
    
    
    
    lista.Rows = 1
    
    totServicios = 0
    totProductos = 0
    totMebresias = 0
    cantSer = 0
    cantProd = 0
    cantMem = 0
    
    Do While Not RES1.EOF
        If RES1.Fields("VENDET_PRODSERV") = "S" And RES1.Fields("VENT_STATUS") = "P" Then
            totServicios = totServicios + Val(RES1.Fields("TOTAL"))
            cantSer = cantSer + Val(RES1.Fields("CANT"))
        Else
            If RES1.Fields("VENDET_PRODSERV") = "P" And RES1.Fields("VENT_STATUS") = "P" Then
                totProductos = totProductos + Val(RES1.Fields("TOTAL"))
                cantProd = cantProd + Val(RES1.Fields("CANT"))
            Else
                If RES1.Fields("VENDET_PRODSERV") = "M" And RES1.Fields("VENT_STATUS") = "P" Then
                    totMebresias = totMebresias + Val(RES1.Fields("TOTAL"))
                    cantMem = cantMem + Val(RES1.Fields("CANT"))
                End If
            End If
        End If
        RES1.MoveNext
    Loop

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Productos"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totProductos)
    lista.TextMatrix(lista.Rows - 1, 2) = cantProd

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Servicio"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totServicios)
    lista.TextMatrix(lista.Rows - 1, 2) = cantSer

    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = "Membresias"
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(totMebresias)
    lista.TextMatrix(lista.Rows - 1, 2) = cantMem

End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
'        mn_Eliminar.Enabled = True
        PopupMenu mn_Corte, vbPopupMenuLeftAlign
    End If

End Sub

Private Sub lista2_Click()
    If lista2.Rows > 1 Then
        cargaDetallePagos (lista2.TextMatrix(lista2.Row, 1))
    End If
End Sub
Private Sub cargaDetallePagos(claveUser As Long)

    '''''----
    listaPagos.Rows = 1
    sql1 = "SELECT  T2.CTPG_NOMBRE, T2.CTPG_VALOR, IF(T2.CTPG_TIPOVALOR='E', 'EFECTIVO', 'PORCENTAJE') TIPO_VALOR, " & _
    "CTPG_APLICAVALORES, CTPG_TIPOVALOR, CTPG_APLICATIPO, IF(CTPG_TIPOPAGO='C', 'COMISION', 'HONORARIOS') TIPO_PAGO, T2.CTPG_REGLA, T1.PG_CTPG_ID " & _
    "FROM COMISIONES T1, CAT_PAGOS T2 WHERE T1.PG_PERTP_PER_ID = '" & claveUser & "' AND T2.CTPG_ID = T1.PG_CTPG_ID AND PG_STATUS = 'A'" & _
    "UNION ALL " & _
    "SELECT 'Consumo interno' CTPG_NOMBRE,  SUM(CSI_CANTIDAD * CSI_PRECIO) CTPG_VALOR, 'EFECTIVO' TIPO_VALOR,  " & _
    "'S' CTPG_APLICAVALORES, 'E' CTPG_TIPOVALOR, 'S' APLICA_TIPO, 'DEDUCCION' TIPO_PAGO, 'NO APLICA' CTPG_REGLA, '0' PG_CTPG_ID " & _
    "FROM CONSUMO_INTERNO WHERE " & _
    "date_format(CSI_FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' AND CSI_USER_PERID = '" & claveUser & "' "

    listaPagos.Redraw = False
    Set RES1 = con.Execute(sql1)
    Dim regla As String
    Do While Not RES1.EOF
        If IsNull(RES1.Fields("CTPG_VALOR")) = False Then
            listaPagos.AddItem ""
            listaPagos.TextMatrix(listaPagos.Rows - 1, 0) = RES1.Fields("CTPG_NOMBRE")
            listaPagos.TextMatrix(listaPagos.Rows - 1, 1) = RES1.Fields("TIPO_PAGO")
            listaPagos.TextMatrix(listaPagos.Rows - 1, 2) = RES1.Fields("CTPG_VALOR")
            listaPagos.TextMatrix(listaPagos.Rows - 1, 3) = RES1.Fields("TIPO_VALOR")
            listaPagos.TextMatrix(listaPagos.Rows - 1, 5) = RES1.Fields("PG_CTPG_ID")
            
            If RES1.Fields("CTPG_TIPOVALOR") = "E" Then
                listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency(RES1.Fields("CTPG_VALOR"))
            Else
                If RES1.Fields("CTPG_TIPOVALOR") = "P" Then
                    Select Case RES1.Fields("CTPG_APLICAVALORES")
                        Case 1:
                            If RES1.Fields("CTPG_APLICATIPO") = "S" Then
                                listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency((RES1.Fields("CTPG_VALOR") / 100) * Val(Format(lista2.TextMatrix(lista2.Row, 4), "General number")))
                            Else
                                If RES1.Fields("CTPG_APLICATIPO") = "P" Then
                                    listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency((RES1.Fields("CTPG_VALOR") / 100) * (Val(Format(lista2.TextMatrix(lista2.Row, 2), "General number")) + Val(Format(lista2.TextMatrix(lista2.Row, 6), "General number"))))
                                End If
                            End If
                        Case 0:
                            listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency((RES1.Fields("CTPG_VALOR") / 100) * Val(Format(lista2.TextMatrix(lista2.Row, 6), "General number")))
                    End Select
                Else
                    If RES1.Fields("CTPG_TIPOVALOR") = "C" Then
                        regla = RES1.Fields("CTPG_REGLA")
                        'regla = Replace(regla, "TOTAL", Val(Format(lista2.TextMatrix(lista2.Row, 4))))
                        regla = Replace(regla, "USUARIO_LISTA", "'" & UCase(lista2.TextMatrix(lista2.Row, 0)) & "'")
                        regla = Replace(regla, "FECHA1", "'" & Format(dtFecha1(0), "yyyy-MM-dd") & "'")
                        regla = Replace(regla, "FECHA2", "'" & Format(dtFecha1(1), "yyyy-MM-dd") & "'")
                        'Text1.Text = regla
                        'MsgBox regla
                        Dim RES2 As Recordset
                        
                        Set RES2 = con.Execute(regla)
                        If Not RES2.EOF Then
                            If IsNull(RES2.Fields("COMISION")) Then
                                listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency(0)
                            Else
                                listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = FormatCurrency(RES2.Fields("COMISION"))
                            End If
                        Else
                            listaPagos.TextMatrix(listaPagos.Rows - 1, 4) = "ERROR"
                        End If
                    
                    End If
                
                End If
            End If
        End If
        RES1.MoveNext
    Loop
    listaPagos.Redraw = True

End Sub

Private Sub lista2_DblClick()
    Dim b1 As Long
    b1 = lista2.Row
    lista2.Row = b1
    lista2.Col = 8
    If lista2.TextMatrix(b1, 8) = Chr(168) Then
        lista2.TextMatrix(b1, 8) = Chr(254)
    Else
        lista2.TextMatrix(b1, 8) = Chr(168)
    End If

End Sub

Private Sub Lista3_DblClick()
'    If Lista3.MouseRow = 0 Then
'        Call ordenarLista(Lista3)
'    End If

End Sub

Private Sub Lista3_GotFocus()
    ConScroll Lista3

End Sub

Private Sub Lista3_LostFocus()
    SinScroll Lista3

End Sub

Private Sub lista4_Click()
    ''''---------
    If lista4.Rows > 2 Then
        carcagDetalle (lista4.TextMatrix(lista4.Row, 12))
    End If
End Sub
Private Sub carcagDetalle(corteId As Long)
    'On Error Resume Next
    Dim texto As String
    Dim filas As Long
    Dim fila As Integer
    '''''----
    Lista6.Rows = 1
    Lista6.Redraw = False
    texto = ""
    
    fila = lista4.Row
    filas = lista4.Rows - 1
    
    'MsgBox fila & "  " & filas
    
    
    If fila = filas Then
        If fila > 2 Then
            If Val(lista4.TextMatrix(lista4.Row - 1, 12)) > 0 Then
                texto = " AND FECHA_HORA > '" & Format(lista4.TextMatrix(lista4.Row - 1, 1), "yyyy-MM-dd") & " " & Format(lista4.TextMatrix(lista4.Row - 1, 1), "HH:MM:SS") & "' "
            End If
        Else
            'TEXTO = " AND FECHA_HORA > '" & lista4.TextMatrix(lista4.Row - 1, 1) & "' "
        End If
    Else
        If fila = 2 Then
            texto = " AND FECHA_HORA < '" & Format(lista4.TextMatrix(lista4.Row, 1), "yyyy-MM-dd") & " " & Format(lista4.TextMatrix(lista4.Row, 1), "HH:MM:SS") & "' "
        Else
            If fila < filas Then
                texto = " AND FECHA_HORA > '" & Format(lista4.TextMatrix(lista4.Row - 1, 1), "yyyy-MM-dd") & " " & Format(lista4.TextMatrix(lista4.Row - 1, 1), "HH:MM:SS") & "' AND FECHA_HORA <= '" & Format(lista4.TextMatrix(lista4.Row, 1), "yyyy-MM-dd") & " " & Format(lista4.TextMatrix(lista4.Row, 1), "HH:MM:SS") & "' "
            End If
        End If
    End If
    
    
    sql1 = "SELECT CODIGO, PRODUCTO, PRECIO, CANTIDAD, ((PRECIO * CANTIDAD) - DESCUENTO) TOTAL, '' PRODINVENTARIO,  TIPO_PROD TIPO" & _
    " FROM VIEW_VENTASDETALLE WHERE date_format(FECHA_HORA, '%Y-%m-%d') = '" & Format(lista4.TextMatrix(lista4.Row, 1), "yyyy-MM-dd") & "' " & texto
    
'    MsgBox sql1
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        Lista6.AddItem ""
        Lista6.TextMatrix(Lista6.Rows - 1, 0) = RES1.Fields("TIPO")
        Lista6.TextMatrix(Lista6.Rows - 1, 1) = RES1.Fields("PRODUCTO")
        Lista6.TextMatrix(Lista6.Rows - 1, 2) = RES1.Fields("CODIGO")
        Lista6.TextMatrix(Lista6.Rows - 1, 3) = FormatCurrency(RES1.Fields("PRECIO"))
        Lista6.TextMatrix(Lista6.Rows - 1, 4) = RES1.Fields("CANTIDAD")
        Lista6.TextMatrix(Lista6.Rows - 1, 5) = FormatCurrency(RES1.Fields("TOTAL"))
        Lista6.TextMatrix(Lista6.Rows - 1, 6) = RES1.Fields("PRODINVENTARIO")
    RES1.MoveNext
    Loop
    
'    sql1 = "SELECT PRODCODIGO, PRODNOMBRE, PRODTIPO, PRODPRECIO, PRODCANTIDAD, PRODTOTAL, PRODINVENTARIO " & _
'    "FROM CORTECAJA_dETALLE WHERE IDCORTE = '" & corteId & "' "
'    Set res1 = con.Execute(sql1)
'    Do While Not res1.EOF
'        Lista6.AddItem ""
'        Lista6.TextMatrix(Lista6.Rows - 1, 0) = res1.Fields("PRODTIPO")
'        Lista6.TextMatrix(Lista6.Rows - 1, 1) = res1.Fields("PRODNOMBRE")
'        Lista6.TextMatrix(Lista6.Rows - 1, 2) = res1.Fields("PRODCODIGO")
'        Lista6.TextMatrix(Lista6.Rows - 1, 3) = FormatCurrency(res1.Fields("PRODPRECIO"))
'        Lista6.TextMatrix(Lista6.Rows - 1, 4) = res1.Fields("PRODCANTIDAD")
'        Lista6.TextMatrix(Lista6.Rows - 1, 5) = FormatCurrency(res1.Fields("PRODTOTAL"))
'        Lista6.TextMatrix(Lista6.Rows - 1, 6) = res1.Fields("PRODINVENTARIO")
'        res1.MoveNext
'    Loop
    Lista6.Redraw = True

End Sub

Private Sub lista5_DblClick()
    If lista5.MouseRow = 0 Then
        Call ordenarLista(lista5)
    End If
End Sub

Private Sub mn_CorteCaja2_Click()
'    MsgBox "Corte"
    resumenCaja2
End Sub

Private Sub TimeSize_Timer()

    TimeSize.Enabled = False
    SSTab1.width = Me.width - 200
    lista4.width = Me.width - 450
    Lista3.width = Me.width - 450
    listaCI.width = Me.width - 450
    listaGST.width = Me.width - 450
    ListaMbr.width = Me.width - 450
    listaApt.width = Me.width - 450
    'ListaAsts.Visible = False
    listMonederos.width = Me.width - 450
    ListaAsts2.width = Me.width - 450
    ListaAsts3.width = Me.width - 450
    ListaCancel.width = Me.width - 450
    ListaReimpresiones.width = Me.width - 450
    
    
End Sub

Private Sub txtFondo_KeyPress(KeyAscii As Integer)
    NumerosPunto (KeyAscii)
    
    If KeyAscii = 13 Then
        cmdAccion_Click (4)
    End If
End Sub
