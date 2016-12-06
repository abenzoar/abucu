VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Mensajes 
   Caption         =   "Mensajes"
   ClientHeight    =   9210
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15225
   Icon            =   "FRM_Mensajes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   15225
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16325
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Mensajes de publicidad por email"
      TabPicture(0)   =   "FRM_Mensajes.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lista"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos generales de los mensajes"
      TabPicture(1)   =   "FRM_Mensajes.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Borde(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lProd(12)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lProd(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Borde(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lProd(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Borde(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lProd(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Borde(12)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "iFoto"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Borde(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lProd(3)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Shape1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lProd(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lProd(7)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Shape1(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lProd(8)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Borde(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cMd1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtMsj(0)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtMsj(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtMsj(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Command1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtMsj(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmBoton(0)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmBoton(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Option1(0)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Option1(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Option1(2)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Option1(3)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtMsj(5)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Envio de mensajes a clientes"
      TabPicture(2)   =   "FRM_Mensajes.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lProd(5)"
      Tab(2).Control(1)=   "Borde(4)"
      Tab(2).Control(2)=   "lProd(6)"
      Tab(2).Control(3)=   "Borde(5)"
      Tab(2).Control(4)=   "ListUsuarios"
      Tab(2).Control(5)=   "txtMsj(4)"
      Tab(2).Control(6)=   "cmBoton(2)"
      Tab(2).Control(7)=   "cmBoton(3)"
      Tab(2).Control(8)=   "Check1"
      Tab(2).Control(9)=   "cmbMsj"
      Tab(2).Control(10)=   "Barra"
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtMsj 
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
         Left            =   360
         MaxLength       =   350
         TabIndex        =   2
         Top             =   6480
         Width           =   7095
      End
      Begin ComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   -74640
         TabIndex        =   29
         Top             =   7800
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   450
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Notificación citas/agenda"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   9
         Top             =   7800
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Notificación apartado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   8
         Top             =   7440
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Notificación venta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   7440
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Envios general"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   6
         Top             =   7440
         Width           =   1695
      End
      Begin VB.ComboBox cmbMsj 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   8895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Seleccionar/Deseleccionar"
         Height          =   195
         Left            =   -63720
         TabIndex        =   13
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviar emails"
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
         Left            =   -67560
         Picture         =   "FRM_Mensajes.frx":05DE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8160
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
         Index           =   2
         Left            =   -62520
         Picture         =   "FRM_Mensajes.frx":0EA8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   8160
         Width           =   2055
      End
      Begin VB.TextBox txtMsj 
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
         Left            =   -74640
         MaxLength       =   65
         TabIndex        =   15
         Top             =   8520
         Width           =   6135
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
         Left            =   5400
         Picture         =   "FRM_Mensajes.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   360
         Picture         =   "FRM_Mensajes.frx":203C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8160
         Width           =   3375
      End
      Begin VB.TextBox txtMsj 
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
         Left            =   7920
         MaxLength       =   350
         TabIndex        =   5
         Top             =   2640
         Width           =   6015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   14160
         TabIndex        =   4
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtMsj 
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
         Left            =   7920
         MaxLength       =   65
         TabIndex        =   3
         Top             =   1680
         Width           =   6015
      End
      Begin VB.TextBox txtMsj 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   1
         Left            =   360
         MaxLength       =   4900
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   2640
         Width           =   7095
      End
      Begin VB.TextBox txtMsj 
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
         MaxLength       =   350
         TabIndex        =   0
         Top             =   1680
         Width           =   7095
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   7455
         Left            =   -74760
         TabIndex        =   19
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Mensajes.frx":2906
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
      Begin MSComDlg.CommonDialog cMd1 
         Left            =   14280
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid ListUsuarios 
         Height          =   6375
         Left            =   -74640
         TabIndex        =   14
         Top             =   1440
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Mensajes.frx":2A36
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
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   6
         Left            =   360
         Top             =   6480
         Width           =   7125
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Con copia:"
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
         TabIndex        =   30
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   1
         Left            =   360
         Top             =   7320
         Width           =   7215
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Asignación del mensaje"
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
         Left            =   360
         TabIndex        =   28
         Top             =   7080
         Width           =   2895
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   405
         Index           =   5
         Left            =   -74640
         Top             =   960
         Width           =   8925
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
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
         Left            =   -74640
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   4
         Left            =   -74640
         Top             =   8505
         Width           =   6165
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Enviar correos con copia a : "
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
         Left            =   -74640
         TabIndex        =   26
         Top             =   8160
         Width           =   3495
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo adjunto"
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
         Index           =   4
         Left            =   7920
         TabIndex        =   25
         Top             =   840
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   7920
         Top             =   1080
         Width           =   7215
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del archivo"
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
         Left            =   7920
         TabIndex        =   24
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   7920
         Top             =   2640
         Width           =   6045
      End
      Begin VB.Image iFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   5295
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   5055
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   5355
         Index           =   12
         Left            =   7920
         Top             =   3360
         Width           =   5085
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta del archivo"
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
         Left            =   7920
         TabIndex        =   23
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   7920
         Top             =   1680
         Width           =   6045
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje *"
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
         TabIndex        =   22
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   3315
         Index           =   1
         Left            =   360
         Top             =   2640
         Width           =   7125
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto *"
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
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos generales del mensaje"
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
         TabIndex        =   20
         Top             =   840
         Width           =   2895
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
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   360
         Top             =   1665
         Width           =   7125
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mn_Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FRM_Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim tipo As String
Dim SQL1 As String
Dim RES1 As Recordset
Dim msjId As Long


Private Sub Check1_Click()
    ListUsuarios.Redraw = False
    If Check1.value = Checked Then
        For b1 = 1 To ListUsuarios.Rows - 1
            ListUsuarios.Col = 3
            ListUsuarios.Row = b1
            ListUsuarios.TextMatrix(b1, 3) = Chr(254)
        Next b1
    Else
        If Check1.value = Unchecked Then
            For b1 = 1 To ListUsuarios.Rows - 1
                ListUsuarios.Col = 3
                ListUsuarios.Row = b1
                ListUsuarios.TextMatrix(b1, 3) = Chr(168)
            Next b1
        End If
    End If
    ListUsuarios.Redraw = True

End Sub

Private Sub cmBoton_Click(index As Integer)
    If index = 0 Then
        agregarMsj
    Else
        If index = 3 Then
            If cmbMsj.Text <> "" Then
                envioMailInfo
            Else
                MsgBox "Seleccione un mensaje.", vbInformation
            End If
        Else
            If index = 1 Then
                cancelar
            End If
        End If
    End If
End Sub

Private Sub envioMailInfo()
    Dim numUsuarios As Long
    Dim mensaje As String
    Dim aprueba As Boolean
    
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    aprueba = False
    
    Barra.Min = 0
    Barra.Max = Lista.Rows
    Barra.value = 0
    Barra.Visible = True
    For b1 = 1 To Lista.Rows - 1
        If Lista.TextMatrix(b1, 0) = cmbMsj.ItemData(cmbMsj.ListIndex) Then
            mensaje = Lista.TextMatrix(b1, 2)
            If Lista.TextMatrix(b1, 3) <> "" Then
                SQL1 = "SELECT msj_anexo, msj_anexo_nombre FROM MENSAJES_EMAIL WHERE MSJ_ID = '" & Lista.TextMatrix(b1, 0) & "'"
                Set RES1 = con.Execute(SQL1)
                checarCarpetaTemp
                Imagen1.Open
                Imagen1.Write RES1.Fields("msj_anexo")
                Imagen1.SaveToFile direccionSistema & "\Temp\" & RES1.Fields("msj_anexo_nombre"), adSaveCreateOverWrite
                Imagen1.Close
                'iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
                adjuntoDir = direccionSistema & "\Temp\" & RES1.Fields("msj_anexo_nombre")
            Else
                adjuntoDir = ""
            End If
            aprueba = True
            Exit For
        End If
        Barra.value = Barra.value + 1
    Next b1
    
    If aprueba = True Then
        Open App.Path & "\LogErrMail2.txt" For Append As #3
        Print #3, Date & "  " & Time & " ----------------- Envio de Mensajes por correo----------------- "
        
        For b1 = 1 To ListUsuarios.Rows - 1
            If ListUsuarios.TextMatrix(b1, 3) = Chr(254) And ListUsuarios.TextMatrix(b1, 1) <> "" Then
                Call enviar_Mail("MENSAJES", cmbMsj.Text, ListUsuarios.TextMatrix(b1, 1), mensaje)
                Print #3, Date & "  " & Time & " Envio de Mensajes por correo: " & vbCrLf & vbCrLf & ListUsuarios.TextMatrix(b1, 0) & "  " & ListUsuarios.TextMatrix(b1, 1)
            
            End If
        Next b1
        Close #3
    Else
        MsgBox "No se encontró un mensaje para el mensaje seleccionado. Verifique.", vbInformation
    End If
    adjuntoDir = ""
    Barra.Visible = False

End Sub

Private Sub cargaEdit()
    On Error Resume Next
    
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    
    txtMsj(0).Text = Lista.TextMatrix(Lista.Row, 1)
    txtMsj(1).Text = Lista.TextMatrix(Lista.Row, 2)
    txtMsj(3).Text = Lista.TextMatrix(Lista.Row, 3)
    txtMsj(5).Text = Lista.TextMatrix(Lista.Row, 6)
        
    If Lista.TextMatrix(Lista.Row, 4) = "G" Then
        Option1(0).value = True
    Else
        If Lista.TextMatrix(Lista.Row, 4) = "V" Then
            Option1(1).value = True
        Else
            If Lista.TextMatrix(Lista.Row, 4) = "A" Then
                Option1(2).value = True
            Else
                If Lista.TextMatrix(Lista.Row, 4) = "C" Then
                    Option1(3).value = True
                End If
            End If
        End If
    End If
        
        
    SQL1 = "SELECT msj_anexo, msj_anexo_nombre FROM MENSAJES_EMAIL WHERE MSJ_ID = '" & Lista.TextMatrix(Lista.Row, 0) & "'"
    Set RES1 = con.Execute(SQL1)
    checarCarpetaTemp
    Imagen1.Open
    Imagen1.Write RES1.Fields("msj_anexo")
    Imagen1.SaveToFile direccionSistema & "\Temp\" & RES1.Fields("msj_anexo_nombre"), adSaveCreateOverWrite
    Imagen1.Close
    iFoto.Picture = LoadPicture(direccionSistema & "\Temp\" & RES1.Fields("msj_anexo_nombre"))
        

End Sub

Private Sub agregarMsj()
    Dim tipoMsj As String

    If Option1(0).value = True Then
        tipoMsj = "G"
    Else
        If Option1(1).value = True Then
            tipoMsj = "V"
        Else
            If Option1(2).value = True Then
                tipoMsj = "A"
            Else
                If Option1(3).value = True Then
                    tipoMsj = "C"
                End If
            End If
        End If
    End If
    


    If tipo = "Add" Then
    
        SQL1 = "INSERT INTO MENSAJES_EMAIL (MSJ_NOMBRE, MSJ_DESCRIPCION, MSJ_ANEXO_NOMBRE, MSJ_FECHA, MSJ_TIPO, MSJ_COPIA) VALUES " & _
        "('" & txtMsj(0).Text & "', '" & txtMsj(1).Text & "', '" & txtMsj(3).Text & "', NOW(), '" & tipoMsj & "', '" & txtMsj(5).Text & "'  )"
        con.Execute (SQL1)
        
        SQL1 = "select last_insert_id() msjId"
        Set RES1 = con.Execute(SQL1)
        If Not RES1.EOF Then
            msjId = RES1.Fields("msjId")
        End If
        
    Else
        If tipo = "Edit" Then
            SQL1 = "UPDATE MENSAJES_EMAIL SET MSJ_NOMBRE = '" & txtMsj(0).Text & "', MSJ_TIPO = '" & tipoMsj & "', " & _
            "MSJ_DESCRIPCION = '" & txtMsj(1).Text & "', MSJ_ANEXO_NOMBRE = '" & txtMsj(3).Text & "' WHERE MSJ_ID = '" & Lista.TextMatrix(Lista.Row, 0) & "' "
            con.Execute (SQL1)
            msjId = Lista.TextMatrix(Lista.Row, 0)
        End If
    End If

    If iFoto.Picture <> 0 Then
        Dim res As ADODB.Recordset
        Set res = New ADODB.Recordset
        Dim Imagen1 As ADODB.Stream
        Set Imagen1 = New ADODB.Stream
        
        res.Open "SELECT * FROM MENSAJES_EMAIL WHERE msj_id = '" & msjId & "'", con, adOpenStatic, adLockOptimistic
        If res.EOF Then
        Else
            Imagen1.Type = adTypeBinary
            Imagen1.Open
            Imagen1.LoadFromFile txtMsj(2).Text
            res.Fields("msj_anexo") = Imagen1.Read
            res.Update
        End If
    End If
        
        MsgBox "Información guardada.", vbInformation
        
        cancelar
        cargaInicial



End Sub
Private Sub cancelar()
                txtMsj(0).Text = ""
                txtMsj(1).Text = ""
                txtMsj(2).Text = ""
                txtMsj(3).Text = ""
                tipo = ""
                iFoto.Picture = LoadPicture("")
                SSTab1.Tab = 0
                SSTab1.TabEnabled(1) = False
                SSTab1.TabEnabled(0) = True
                SSTab1.TabEnabled(2) = True
    
End Sub
Private Sub Command1_Click()
    buscarImagen
End Sub
Private Sub cargaClientes()
ListUsuarios.Rows = 1
 ListUsuarios.Redraw = False
 
SQL1 = "select * fROM VIEW_CLTS_MSJS ORDER BY EMAIL DESC"
Set RES1 = con.Execute(SQL1)


Do While Not RES1.EOF
    ListUsuarios.AddItem ""
    ListUsuarios.TextMatrix(ListUsuarios.Rows - 1, 0) = RES1.Fields("CLIENTE")
    ListUsuarios.TextMatrix(ListUsuarios.Rows - 1, 1) = RES1.Fields("EMAIL") & ""
    ListUsuarios.TextMatrix(ListUsuarios.Rows - 1, 2) = RES1.Fields("TELEFONOS") & ""
    ListUsuarios.TextMatrix(ListUsuarios.Rows - 1, 4) = RES1.Fields("CLAVE")
    
    ListUsuarios.Row = ListUsuarios.Rows - 1
    ListUsuarios.Col = 3
    ListUsuarios.CellFontName = "Wingdings"
    ListUsuarios.CellFontBold = True
    ListUsuarios.CellFontSize = 16
    ListUsuarios.TextMatrix(ListUsuarios.Rows - 1, 3) = Chr(254)
    
    
    RES1.MoveNext
Loop
 ListUsuarios.Redraw = True
    
End Sub

Private Sub cargaLista()
Lista.Redraw = False

    Lista.Rows = 1
    cmbMsj.Clear
    
    SQL1 = "SELECT MSJ_ID, MSJ_NOMBRE, MSJ_DESCRIPCION, MSJ_ANEXO_NOMBRE, MSJ_TIPO, MSJ_FECHA, MSJ_COPIA from MENSAJES_EMAIL ORDER BY MSJ_FECHA DESC"
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("MSJ_ID")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("MSJ_NOMBRE")
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("MSJ_DESCRIPCION")
        Lista.TextMatrix(Lista.Rows - 1, 3) = RES1.Fields("MSJ_ANEXO_NOMBRE")
        Lista.TextMatrix(Lista.Rows - 1, 4) = RES1.Fields("MSJ_TIPO")
        Lista.TextMatrix(Lista.Rows - 1, 5) = RES1.Fields("MSJ_FECHA")
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("MSJ_COPIA") & ""
        
        cmbMsj.AddItem RES1.Fields("MSJ_NOMBRE")
        cmbMsj.ItemData(cmbMsj.ListCount - 1) = RES1.Fields("MSJ_ID")
                
        RES1.MoveNext
    Loop
Lista.Redraw = True
    
End Sub
Private Sub Form_Load()
    cargaInicial
    cargaClientes
End Sub
Private Sub cargaInicial()
    
    
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    cargaLista
    Check1.value = Checked
    Barra.Visible = False
    
End Sub

Private Sub buscarImagen()
    cMd1.DialogTitle = "Buscando imagen..."
    'cMd1.Filter = "Archivos de Imagenes|*.jpg*||*.bmp*||*.gif*||*.wmf*||*.emf*||*.png*|"
    cMd1.FileName = ""
    cMd1.ShowOpen
    If cMd1.FileName <> "" Then
        guardarImagen
    End If
End Sub
Private Sub guardarImagen()
On Error Resume Next
    With cMd1
        iFoto.Visible = True
        iFoto.Picture = LoadPicture(.FileName)
        txtMsj(2).Text = .FileName
        txtMsj(3).Text = .FileTitle
        
    End With
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Lista.Rows > 1 Then
        If Button = vbRightButton Then
            PopupMenu mn_Menu, vbPopupMenuLeftAlign
        End If
    End If


End Sub

Private Sub ListUsuarios_Click()
'''''
End Sub

Private Sub ListUsuarios_DblClick()
        
        
        If ListUsuarios.Col = 3 Then
            Dim b1 As Long
            b1 = ListUsuarios.Row
            
            ListUsuarios.Row = b1
            ListUsuarios.Col = 3
            If ListUsuarios.TextMatrix(b1, 3) = Chr(168) Then
                ListUsuarios.TextMatrix(b1, 3) = Chr(254)
            Else
                ListUsuarios.TextMatrix(b1, 3) = Chr(168)
            End If
        Else
           Call ordenarLista(ListUsuarios)
        
        End If




End Sub

Private Sub mn_Add_Click()
    tipo = "Add"
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    Option1(1).value = True
    
End Sub

Private Sub mn_Editar_Click()
Dim ques As String

    ques = MsgBox("Editar mensaje " & Lista.TextMatrix(Lista.Row, 1) & "?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        tipo = "Edit"
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
        cargaEdit
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = False
        
    End If
End Sub

Private Sub mn_Salir_Click()
    Unload Me
End Sub
