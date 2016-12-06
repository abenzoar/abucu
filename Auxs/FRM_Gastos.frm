VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Gastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos"
   ClientHeight    =   9825
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   16455
   Icon            =   "FRM_Gastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   16455
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   697
      TabCaption(0)   =   "  Gastos"
      TabPicture(0)   =   "FRM_Gastos.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAccion(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmbTipo(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbTipo(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmbTipo(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "listaGST"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "  Datos generales"
      TabPicture(1)   =   "FRM_Gastos.frx":0B24
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblUserId(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblUserId(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblUserId(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblDatos(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblInfo(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "imgFoto(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "label0(9)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(9)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line1(9)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Line1(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Line1(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Line1(2)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(7)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(8)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(10)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label1(14)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(15)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "dtFecha1(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "dtFecha1(0)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lista"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Check2"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Check1(0)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmBoton(2)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtGst(1)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtGst(0)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmbTipo(0)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtGst(3)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "cmBoton(1)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "cmBoton(0)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtPrintCopias"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "cmbTipo(1)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Check1(1)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtGst(2)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "cmbTipo(2)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "cmBoton(4)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "cmBoton(3)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).ControlCount=   43
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar lista"
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
         Index           =   3
         Left            =   -63960
         Picture         =   "FRM_Gastos.frx":10BE
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   480
         Width           =   2535
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
         Height          =   375
         Index           =   3
         Left            =   6840
         Picture         =   "FRM_Gastos.frx":1648
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Height          =   375
         Index           =   4
         Left            =   6840
         Picture         =   "FRM_Gastos.frx":1BD2
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtGst 
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
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Afecta valor de caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtPrintCopias 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   8040
         TabIndex        =   24
         Text            =   "2"
         Top             =   9120
         Width           =   375
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
         Height          =   1215
         Index           =   0
         Left            =   360
         Picture         =   "FRM_Gastos.frx":215C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8280
         Width           =   2415
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
         Height          =   1215
         Index           =   1
         Left            =   12480
         Picture         =   "FRM_Gastos.frx":2A26
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8280
         Width           =   1935
      End
      Begin VB.TextBox txtGst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Index           =   3
         Left            =   7440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1440
         Width           =   4815
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtGst 
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
         Index           =   0
         Left            =   15120
         TabIndex        =   6
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtGst 
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar y agregar a la lista"
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
         Left            =   9840
         Picture         =   "FRM_Gastos.frx":32F0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Existe comprobante del gasto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir ticket del gasto generado"
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
         Left            =   8040
         TabIndex        =   12
         Top             =   8640
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   3255
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         WordWrap        =   -1  'True
         FormatString    =   $"FRM_Gastos.frx":3BBA
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
         Height          =   8175
         Left            =   -74880
         TabIndex        =   4
         Top             =   1320
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   14420
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"FRM_Gastos.frx":3D0B
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
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   43
         Top             =   3480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120324097
         CurrentDate     =   40829
      End
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Format          =   120324097
         CurrentDate     =   40829
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "a"
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
         Index           =   15
         Left            =   3960
         TabIndex        =   46
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del gasto"
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
         Index           =   14
         Left            =   2160
         TabIndex        =   44
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Index           =   13
         Left            =   -72000
         TabIndex        =   39
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Left            =   -68160
         TabIndex        =   37
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Left            =   2880
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de comprobante"
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
         Left            =   240
         TabIndex        =   34
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   240
         X2              =   4200
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de gastos"
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
         Left            =   240
         TabIndex        =   29
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   1
         X1              =   12360
         X2              =   16200
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario en mostrador"
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
         Left            =   12360
         TabIndex        =   28
         Top             =   600
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   0
         X1              =   240
         X2              =   12240
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos del gasto"
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
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   9
         X1              =   8040
         X2              =   12000
         Y1              =   8520
         Y2              =   8520
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Impresión de ticket de gasto"
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
         Left            =   8040
         TabIndex        =   26
         Top             =   8280
         Width           =   3495
      End
      Begin VB.Label label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Número de copías"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   8520
         TabIndex        =   25
         Top             =   9120
         Width           =   2535
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Index           =   2
         Left            =   12360
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrador"
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
         Index           =   1
         Left            =   13800
         TabIndex        =   23
         Top             =   1080
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
         Left            =   13800
         TabIndex        =   22
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   21
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   20
         Top             =   8160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   19
         Top             =   8520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   7440
         TabIndex        =   18
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
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
         Left            =   15120
         TabIndex        =   16
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_Export 
         Caption         =   "Exportar información"
      End
      Begin VB.Menu mn_Editar 
         Caption         =   "Editar gasto"
      End
      Begin VB.Menu mn_Print 
         Caption         =   "Imprimir ticket "
      End
      Begin VB.Menu mn_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mn_Cat 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_TipGasto 
         Caption         =   "Tipo de gasto"
      End
      Begin VB.Menu mn_Proveedor 
         Caption         =   "Proveedor"
      End
   End
End
Attribute VB_Name = "FRM_Gastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RES1 As Recordset
Dim sql1 As String
Dim valorEditar As Boolean
Dim idGst1 As Long

Private Sub cmBoton_Click(Index As Integer)
    
    Select Case Index
        Case 0:
            If valorEditar = False Then
                aceptar_add
            Else
                aceptar_edit
            End If
        Case 2:
            addLista
        Case 1:
            Dim ques As String
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        Case 4:
            mn_TipGasto_Click
        Case 3:
            tipoCatTipo = "G"
            tipoPersona = "PROVEEDOR_G"
            ADD_Cliente.Show vbModal
    End Select
    
End Sub
Private Sub aceptar_edit()
Dim gastoId As Long
Dim tipoProv As String
Dim prodPerId As String
Dim prodTipoId As String
Dim prodTipo As String
Dim tipoGST As String
Dim comprobante As String
Dim afecta As String

    If cmbTipo(2).Text <> "" Then
        sql1 = "SELECT PERTP_TIPO_ID, PERTP_PER_ID, PERTP_PER_TIPO FROM PER_TIPO WHERE PERTP_PER_ID = '" & cmbTipo(2).ItemData(cmbTipo(2).ListIndex) & "' "
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
            prodPerId = RES1.Fields("PERTP_PER_ID")
            prodTipoId = RES1.Fields("PERTP_TIPO_ID")
            prodTipo = RES1.Fields("PERTP_PER_TIPO")
        End If
    End If
                
    If cmbTipo(1).Text = "Gasto general" Then
        tipoGST = "G"
    Else
        If cmbTipo(1).Text = "Gasto inversión" Then
            tipoGST = "I"
        End If
    End If
                
    If Check1(0).value = 1 Then
        comprobante = "SI"
    Else
        comprobante = "NO"
    End If
    
    If Check1(1).value = 1 Then
        afecta = "SI"
    Else
        afecta = "NO"
    End If
                
                
    sql1 = "UPDATE GASTOS SET GST_FECHAHORA = '" & Format(dtFecha1(0), "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', GST_FECHAHORA_FIN = '" & Format(dtFecha1(1), "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', " & _
    "GST_USER_PERID = '" & lblUserId(0).Caption & "', GST_USER_PERTIPOID = '" & lblUserId(1).Caption & "', GST_USER_PERTIPO = '" & lblUserId(2).Caption & "', " & _
    "GST_TIPO_ID = '" & cmbTipo(0).ItemData(cmbTipo(0).ListIndex) & "', GST_DESCRIPCION = '" & txtGst(3).Text & "', GST_TOTAL = '" & Format(txtGst(1).Text, "General Number") & "',  " & _
    "GST_COMPROBANTE = '" & Left(comprobante, 1) & "', GST_CAJA = '" & Left(afecta, 1) & "', GST_CODIGO = '" & txtGst(2).Text & "', " & _
    "GST_TIPO_GRAL = '" & tipoGST & "', GST_PROV_PERID = '" & prodPerId & "', GST_PROV_TIPOID = '" & prodTipoId & "', GST_PROV_TIPO = '" & prodTipo & "' " & _
    " WHERE GST_ID = '" & idGst1 & "' "
    Set RES1 = con.Execute(sql1)
                                                        
    If Check2.value = Checked Then
        For c1 = 1 To Val(txtPrintCopias.Text)
            notaGasto (idGst1)
            MsgBox "Operación realizada. ", vbInformation
        Next c1
    End If
        
    cargaLista
    MsgBox "Información guardada.", vbInformation
    cancelar
        

End Sub
Private Sub aceptar_add()
Dim gastoId As Long
Dim tipoProv As String
    If lista.Rows > 1 Then
        With lista
            For b1 = 1 To .Rows - 1
                
                If .TextMatrix(b1, 6) = "" Or .TextMatrix(b1, 6) = "NULL" Then
                    tipoProv = "NULL"
                Else
                    tipoProv = "'" & .TextMatrix(b1, 6) & "'"
                End If
                
                sql1 = "INSERT INTO GASTOS (GST_FECHAHORA, GST_REGISTRO, GST_FECHAHORA_FIN, GST_USER_PERID, GST_USER_PERTIPOID, GST_USER_PERTIPO, GST_TIPO_ID, GST_TIPO, " & _
                "GST_DESCRIPCION, GST_TOTAL, GST_COMPROBANTE, GST_CAJA, GST_CODIGO, GST_TIPO_GRAL, GST_PROV_PERID, GST_PROV_TIPOID, GST_PROV_TIPO) VALUES ('" & Format(dtFecha1(0), "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', NOW(), '" & Format(dtFecha1(1), "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', '" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', " & _
                "'" & lblUserId(2).Caption & "', '" & .TextMatrix(b1, 2) & "', 'G', '" & .TextMatrix(b1, 11) & "', '" & Val(Format(.TextMatrix(b1, 7), "General Number")) & "', " & _
                "'" & Left(.TextMatrix(b1, 8), 1) & "', '" & Left(.TextMatrix(b1, 10), 1) & "',  '" & .TextMatrix(b1, 9) & "', '" & .TextMatrix(b1, 12) & "', " & _
                "" & .TextMatrix(b1, 4) & ", " & .TextMatrix(b1, 5) & ", " & tipoProv & ") "
                'MsgBox SQL1
                Set RES1 = con.Execute(sql1)
                            
                sql1 = "select last_insert_id() folioId"
                Set RES1 = con.Execute(sql1)
                If Not RES1.EOF Then
                    gastoId = RES1.Fields("folioId")
                End If
                            
                If Check2.value = Checked Then
                    For c1 = 1 To Val(txtPrintCopias.Text)
                        notaGasto (gastoId)
                        MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
                    Next c1
                End If
            Next b1
        
        End With
        cargaLista
        MsgBox "Información guardada.", vbInformation
        cancelar
        
    Else
        MsgBox "No se puede realizar la operación. Verifique.", vbInformation
    End If
End Sub
Private Sub cancelar()
    For b1 = 1 To 3
        txtGst(b1).Text = ""
    Next b1
    lista.Rows = 1
    cmBoton(2).Enabled = True
    valorEditar = False
    
End Sub
Private Sub addLista()
    Dim prodPerId As String
    Dim prodTipoId As String
    Dim prodTipo As String
    Dim tipoGST As String
    
    prodPerId = "NULL"
    prodTipoId = "NULL"
    prodTipo = "NULL"
    
    For b1 = 0 To 1
        If txtGst(b1).Text = "" Then
            MsgBox "Falta un valor. Verifique.", vbInformation
            Exit Sub
        End If
    Next b1
    
    If cmbTipo(1).Text = "Gasto general" Then
        tipoGST = "G"
    Else
        If cmbTipo(1).Text = "Gasto inversión" Then
            tipoGST = "I"
        End If
    End If
    
    If cmbTipo(2).Text <> "" Then
        sql1 = "SELECT PERTP_TIPO_ID, PERTP_PER_ID, PERTP_PER_TIPO FROM PER_TIPO WHERE PERTP_PER_ID = '" & cmbTipo(2).ItemData(cmbTipo(2).ListIndex) & "' "
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
            prodPerId = RES1.Fields("PERTP_PER_ID")
            prodTipoId = RES1.Fields("PERTP_TIPO_ID")
            prodTipo = RES1.Fields("PERTP_PER_TIPO")
        End If
    End If
    
    lista.AddItem ""
    lista.TextMatrix(lista.Rows - 1, 0) = cmbTipo(1).Text
    lista.TextMatrix(lista.Rows - 1, 1) = cmbTipo(0).Text
    lista.TextMatrix(lista.Rows - 1, 2) = cmbTipo(0).ItemData(cmbTipo(0).ListIndex)
    lista.TextMatrix(lista.Rows - 1, 3) = cmbTipo(2).Text
    
    lista.TextMatrix(lista.Rows - 1, 4) = prodPerId
    lista.TextMatrix(lista.Rows - 1, 5) = prodTipoId
    lista.TextMatrix(lista.Rows - 1, 6) = prodTipo

    lista.TextMatrix(lista.Rows - 1, 7) = FormatCurrency(txtGst(1).Text)
    
    If Check1(0).value = 1 Then
        lista.TextMatrix(lista.Rows - 1, 8) = "SI"
    Else
        lista.TextMatrix(lista.Rows - 1, 8) = "NO"
    End If
    
    lista.TextMatrix(lista.Rows - 1, 9) = txtGst(2).Text
        
    If Check1(1).value = 1 Then
        lista.TextMatrix(lista.Rows - 1, 10) = "SI"
    Else
        lista.TextMatrix(lista.Rows - 1, 10) = "NO"
    End If
    lista.TextMatrix(lista.Rows - 1, 11) = txtGst(3).Text
    lista.TextMatrix(lista.Rows - 1, 12) = tipoGST

    For b1 = 1 To 3
        txtGst(b1).Text = ""
    Next b1


End Sub

Public Sub cargaProveedor()

    sql1 = "SELECT PER_ID, CONCAT(PER_ALIAS, ' - ', PER_NOMBRE, ' ', PER_PATERNO, ' ', PER_MATERNO) PROVEEDOR " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'V'  "
    Set RES1 = con.Execute(sql1)
    
    cmbTipo(2).Clear
    Do While Not RES1.EOF
        cmbTipo(2).AddItem RES1.Fields("PROVEEDOR")
        cmbTipo(2).ItemData(cmbTipo(2).ListCount - 1) = RES1.Fields("PER_ID")
        
        cmbTipo(3).AddItem RES1.Fields("PROVEEDOR")
        cmbTipo(3).ItemData(cmbTipo(3).ListCount - 1) = RES1.Fields("PER_ID")
        
        RES1.MoveNext
    Loop
    
    cmbTipo(2).AddItem ""
    cmbTipo(3).AddItem "TODOS"
    
'    If cmbProd(3).ListCount > 0 Then
'        cmbProd(3).ListIndex = 0
'    End If

End Sub

Private Sub cmbTipo_Click(Index As Integer)


If Index = 3 Then
    cargaLista
Else
    If Index = 4 Then
        cargaLista
    Else
        If Index = 5 Then
            cargaLista
        End If
    End If
End If



End Sub

Private Sub cmdAccion_Click(Index As Integer)
    mn_Export_Click
    
End Sub

Private Sub dtFecha1_Change(Index As Integer)
    If Index = 0 Then
        dtFecha1(1) = dtFecha1(0)
    End If
End Sub

Private Sub dtFecha1_Click(Index As Integer)
'    If Index = 0 Then
'        dtFecha1(1) = dtFecha1(0)
'    End If
End Sub

Private Sub Form_Load()
    cargaDatos
    SSTab1.Tab = 0
    dtFecha1(0) = Date
    dtFecha1(1) = Date
    listaGST.ColWidth(2) = 0
    
End Sub
Private Sub cargaDatos()
    lblDatos(1).Caption = FRM_Menu.menuBarra2.Panels(5).Text
    lista.Rows = 1
    lista.ColWidth(2) = 0
    lista.ColWidth(4) = 0
    lista.ColWidth(5) = 0
    lista.ColWidth(6) = 0
    
    Call cargaFotoMostrador("M", 2)
    cmbTipo(1).AddItem "Gasto general"
    cmbTipo(1).AddItem "Gasto inversión"
    
    cmbTipo(4).AddItem "Gasto general"
    cmbTipo(4).AddItem "Gasto inversión"
    
    cmbTipo(4).AddItem "TODOS"
    
    cargaTipoGasto
    cargaLista
    cargaProveedor
End Sub
Private Sub cargaLista()
    Dim texto1 As String
    Dim num As Integer
    
    num = 0
    texto1 = ""
    If cmbTipo(3).Text <> "TODOS" And cmbTipo(3).Text <> "" Then
        num = num + 1
        texto1 = texto1 & "AND upper(PROVEEDOR) LIKE upper('%" & cmbTipo(3).Text & "%') "
    Else
        If cmbTipo(4).Text <> "TODOS" And cmbTipo(4).Text <> "" Then
            num = num + 1
            texto1 = texto1 & "AND upper(TIPO_GRAL) LIKE upper('%" & cmbTipo(4).Text & "%') "
        Else
            If cmbTipo(5).Text <> "TODOS" And cmbTipo(5).Text <> "" Then
                num = num + 1
                texto1 = texto1 & "AND upper(TIPO_GASTO) LIKE upper('%" & cmbTipo(5).Text & "%') "
            End If
       End If
    End If

    If num > 0 Then
        sql1 = "SELECT * " & _
        "FROM VIEW_GASTOS WHERE ID > 0 " & texto1 & " order BY FECHA_HORA DESC "
    Else
        sql1 = "SELECT * " & _
        "FROM VIEW_GASTOS order BY FECHA_HORA DESC "
    End If
    
    Set RES1 = con.Execute(sql1)
    listaGST.Redraw = False
    listaGST.Rows = 1
    Do While Not RES1.EOF
        listaGST.AddItem ""
        listaGST.TextMatrix(listaGST.Rows - 1, 0) = RES1.Fields("ID")
        listaGST.TextMatrix(listaGST.Rows - 1, 1) = Format(RES1.Fields("FECHA_HORA"), "dddd") & " " & Format(RES1.Fields("FECHA_HORA"), "Short Date")
        listaGST.TextMatrix(listaGST.Rows - 1, 2) = Format(RES1.Fields("FECHA_FIN"), "dddd") & " " & Format(RES1.Fields("FECHA_FIN"), "Short Date")
        listaGST.TextMatrix(listaGST.Rows - 1, 3) = RES1.Fields("USUARIO")
        listaGST.TextMatrix(listaGST.Rows - 1, 4) = RES1.Fields("TIPO_GRAL")
        listaGST.TextMatrix(listaGST.Rows - 1, 5) = RES1.Fields("TIPO_GASTO")
        listaGST.TextMatrix(listaGST.Rows - 1, 6) = FormatCurrency(RES1.Fields("GASTO"))
        listaGST.TextMatrix(listaGST.Rows - 1, 7) = RES1.Fields("COMPROBANTE")
        listaGST.TextMatrix(listaGST.Rows - 1, 8) = RES1.Fields("CODIGO") & ""
        listaGST.TextMatrix(listaGST.Rows - 1, 9) = RES1.Fields("CAJA")
        listaGST.TextMatrix(listaGST.Rows - 1, 10) = RES1.Fields("PROVEEDOR") & ""
        listaGST.TextMatrix(listaGST.Rows - 1, 11) = RES1.Fields("GST_DESCRIPCION")
        listaGST.TextMatrix(listaGST.Rows - 1, 12) = RES1.Fields("REGISTRO")
        
        
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
Public Sub cargaTipoGasto()
    sql1 = ("SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'G' ORDER BY CTPT_TIPO")
    Set RES1 = con.Execute(sql1)
    
    cmbTipo(0).Clear
    Do While Not RES1.EOF
        cmbTipo(0).AddItem RES1.Fields("CTPT_TIPO")
        cmbTipo(0).ItemData(cmbTipo(0).ListCount - 1) = RES1.Fields("CTPT_ID")
        
        cmbTipo(5).AddItem RES1.Fields("CTPT_TIPO")
        cmbTipo(5).ItemData(cmbTipo(5).ListCount - 1) = RES1.Fields("CTPT_ID")
        
        RES1.MoveNext
    Loop

cmbTipo(5).AddItem "TODOS"

End Sub
Public Sub cargaFotoMostrador(tipo As String, num As Integer)
    Dim idPer As String
    
    If tipo = "M" Then
        idPer = FRM_Menu.menuBarra2.Panels(7).Text
    Else
        If tipo = "U" Then
            idPer = lblUserId(3).Caption
        End If
    End If
    
    sql1 = "SELECT PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_TIPO_ID, PERTP_PER_TIPO, PER_ID, PER_FOTO " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' " & _
    "AND T1.PER_ID = '" & idPer & "'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
    If tipo = "M" Then
        lblDatos(1).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
        lblUserId(0).Caption = RES1.Fields("PER_ID")
        lblUserId(1).Caption = RES1.Fields("PERTP_TIPO_ID")
        lblUserId(2).Caption = RES1.Fields("PERTP_PER_TIPO")
    Else
        If tipo = "U" Then
            lblDatos(0).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
            lblUserId(3).Caption = RES1.Fields("PER_ID")
            lblUserId(4).Caption = RES1.Fields("PERTP_TIPO_ID")
            lblUserId(5).Caption = RES1.Fields("PERTP_PER_TIPO")
        End If
    End If
        If IsNull(RES1.Fields("PER_fOTO")) = False Then
            Dim Imagen1 As Stream
            Set Imagen1 = New Stream
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES1.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            imgFoto(num).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            imgFoto(num).Picture = LoadPicture("")
        End If
    End If
        'txtClave(1).SetFocus
End Sub


Private Sub listaGST_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If listaGST.Rows > 1 Then
        If Button = vbRightButton Then
            PopupMenu mn_menu, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Editar_Click()
   
   editar (listaGST.TextMatrix(listaGST.Row, 0))
   valorEditar = True
    
End Sub
Private Sub editar(idGst As Long)
    On Error Resume Next
    sql1 = "SELECT * FROM VIEW_GASTOS WHERE ID = " & idGst & ""
    Set RES1 = con.Execute(sql1)
        
    idGst1 = idGst
    
    If Not RES1.EOF Then
        'listaGST.AddItem ""
        'listaGST.TextMatrix(listaGST.Rows - 1, 0) = RES1.Fields("ID")
        dtFecha1(0) = RES1.Fields("FECHA_HORA")
        dtFecha1(1) = RES1.Fields("FECHA_FIN")
        'listaGST.TextMatrix(listaGST.Rows - 1, 2) = RES1.Fields("USUARIO")
        cmbTipo(1).Text = RES1.Fields("TIPO_GRAL")
        cmbTipo(0).Text = RES1.Fields("TIPO_GASTO")
        txtGst(1).Text = FormatCurrency(RES1.Fields("GASTO"))
        If RES1.Fields("COMPROBANTE") = "SI" Then
            Check1(0).value = Checked
        Else
            Check1(0).value = Unchecked
        End If
        If RES1.Fields("CAJA") = "SI" Then
            Check1(1).value = Checked
        Else
            Check1(1).value = Unchecked
        End If
        txtGst(2).Text = RES1.Fields("CODIGO") & ""
        listaGST.TextMatrix(listaGST.Rows - 1, 8) = RES1.Fields("CAJA")
        cmbTipo(2).Text = RES1.Fields("PROVEEDOR") & ""
        txtGst(3).Text = RES1.Fields("GST_DESCRIPCION")
        SSTab1.Tab = 1
        cmBoton(2).Enabled = False
    End If

End Sub
Private Sub mn_Export_Click()
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(listaGST)
            End If

End Sub

Private Sub mn_Print_Click()
    notaGasto (listaGST.TextMatrix(listaGST.Row, 0))
End Sub

Private Sub mn_TipGasto_Click()
    tipoCatTipo = "G"
    CAT_Tipo.Show vbModal

End Sub

Private Sub txtGst_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
    Call NumerosPunto(KeyAscii)
End If
End Sub
