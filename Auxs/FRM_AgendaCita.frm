VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_AgendaCita 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cita o reservación"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15495
   Icon            =   "FRM_AgendaCita.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15495
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUsuario 
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
      Left            =   8520
      MaxLength       =   50
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   6840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   5160
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtObservaciones 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   8760
      Width           =   12975
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guardar cita"
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
      Left            =   11400
      Picture         =   "FRM_AgendaCita.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar y salir"
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
      Left            =   13320
      Picture         =   "FRM_AgendaCita.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   14760
      Picture         =   "FRM_AgendaCita.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5640
      Width           =   495
   End
   Begin VB.ComboBox cmbServ 
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
      Index           =   2
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5640
      Width           =   4095
   End
   Begin VB.ComboBox cmbServ 
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
      Index           =   1
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5640
      Width           =   5295
   End
   Begin VB.ComboBox cmbServ 
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
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5640
      Width           =   4455
   End
   Begin VB.ComboBox cmbHora 
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
      Index           =   3
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbHora 
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
      Index           =   2
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbHora 
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
      Index           =   1
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbHora 
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
      Index           =   0
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid listaClte 
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   9520683
      ForeColorFixed  =   16777215
      BackColorBkg    =   15329769
      GridColor       =   16711680
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"FRM_AgendaCita.frx":1CA8
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
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   10560
      Picture         =   "FRM_AgendaCita.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   9520683
      ForeColorFixed  =   16777215
      BackColorBkg    =   15329769
      GridColor       =   16711680
      FormatString    =   $"FRM_AgendaCita.frx":22F2
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
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
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
      Format          =   100401153
      CurrentDate     =   40956
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente seleccionado para cita"
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
      Left            =   11280
      TabIndex        =   46
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   11280
      TabIndex        =   45
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos:"
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
      Left            =   11280
      TabIndex        =   44
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   11280
      TabIndex        =   43
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfonos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   11280
      TabIndex        =   42
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   4
      X1              =   120
      X2              =   15120
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información fecha y hora para cita"
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
      Left            =   120
      TabIndex        =   41
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label lInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0    Clientes en lista"
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
      Left            =   8640
      TabIndex        =   40
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información del servicio para cita"
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
      TabIndex        =   39
      Top             =   5040
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   3
      X1              =   120
      X2              =   15120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Left            =   8520
      TabIndex        =   38
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones de la cita"
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
      Left            =   11280
      TabIndex        =   37
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2"
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
      Left            =   6840
      TabIndex        =   36
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1"
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
      Left            =   5160
      TabIndex        =   35
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Ap Materno"
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
      Left            =   3480
      TabIndex        =   34
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Ap Paterno"
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
      Left            =   1800
      TabIndex        =   33
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   2
      X1              =   120
      X2              =   11040
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda de cliente"
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
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   11280
      X2              =   15360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   11280
      X2              =   15360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de la cita: "
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
      Left            =   9120
      TabIndex        =   31
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Label lUsuario 
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
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Index           =   8
      Left            =   10320
      TabIndex        =   25
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
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
      Index           =   7
      Left            =   4800
      TabIndex        =   23
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de servicio"
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
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   18
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Fin"
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
      Left            =   5400
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   14
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mn_CancelarCita 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "FRM_AgendaCita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim checkError As Boolean
Dim RES1 As Recordset
Dim SQL1 As String
Dim RES2 As Recordset
Dim SQL2 As String
Dim datos As Boolean
Dim userPertpId As Long
Dim cliePertpId As Long
Dim cliePerId As Long
Dim salida As Boolean
Dim email As String
Dim mensaje As String

Private Sub cmbHora_Click(Index As Integer)
    If Index = 0 Or Index = 1 Then
        If cmbHora(0).ListIndex < cmbHora(0).ListCount - 1 Then
            If cmbHora(1).Text = "00" Then
                cmbHora(2).ListIndex = cmbHora(0).ListIndex
                cmbHora(3).Text = "30"
            Else
                If cmbHora(1).Text = "30" Then
                    cmbHora(2).ListIndex = cmbHora(0).ListIndex + 1
                    cmbHora(3).Text = "00"
                End If
            End If
            
        Else
            MsgBox "La hora especificado es la última hora del horario del negocio. Verifique.", vbInformation
            cmbHora(0).ListIndex = cmbHora(0).ListIndex - 1
        End If
    Else
'        If Index = 2 Then
'            If Format(cmbHora(0).Text & ":" & cmbHora(1).Text, "Short Time") > Format(cmbHora(2).Text & ":" & cmbHora(3).Text, "Short Time") Then
'                MsgBox "La hora de término no puede ser menor que la hora de inicio. Verifique.", vbInformation
'                cmbHora_Click (0)
'            End If
'        End If
    End If
End Sub

Private Sub cmBoton_Click(Index As Integer)
    If Index = 1 Then
         Unload Me
    Else
        If Index = 0 Then
            If lInfo(2).Caption = "Nombre: " Then
                MsgBox "Debe de asociar un cliente para la cita. Verifique.", vbInformation
            Else
                If lista.Rows = 1 Then
                    MsgBox "Debe asignar un servicio a la cita. Verifique.", vbInformation
                Else
                    If tipoCita = "Creacion" Then
                        guardarCita
                    Else
                        If tipoCita = "Edicion" Then
                            editarCita
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub editarCita()
    Dim textoCita As String
    
    SQL1 = "UPDATE AGENDA SET AGD_OBSERVACIONES = '" & txtObservaciones.Text & "', " & _
    "AGD_CLIE_TIPOID =  '" & cliePertpId & "', " & _
    "AGD_CLIE_PERID = '" & cliePerId & "', " & _
    "AGD_CLIE_PERTIPO = 'C' " & _
    "WHERE AGD_ID = '" & clavesCitas(FRM_Agenda.listaDia.Row, FRM_Agenda.listaDia.Col) & "' "
    con.Execute (SQL1)
    
    salida = True
    FRM_Agenda.cmdCitas_Click
        
    MsgBox "Información guardada", vbInformation
    
    
    textoCita = textoCita & vbCrLf & vbCrLf & "Edición de cita: " & vbCrLf & vbCrLf & txtObservaciones.Text & vbCrLf & vbclrf & "Te recomendamos estar 15 minutos antes de tu cita. " & vbCrLf & vbclrf & "Que tengas un excelente día."
    If email <> "" Then
        Call enviar_Mail("CITA", "Cambio de cita " & FRM_Menu.menuBarra2.Panels(9).Text & " " & dtFecha1, email, textoCita)
    Else
        MsgBox "No se puede enviar confirmación de cita por correo por falta de información. Verifique.", vbInformation
    End If
    
    
    
    Unload Me
        
End Sub
Private Sub guardarCita()
    Dim agendaId As Long
    Dim textoCita As String
    
    SQL1 = "INSERT INTO AGENDA (AGD_FECHAHORA, AGD_OBSERVACIONES, AGD_STATUS, AGD_USER_TIPOID, AGD_USER_PERID, AGD_USER_PERTIPO, " & _
    "AGD_CLIE_TIPOID, AGD_CLIE_PERID, AGD_CLIE_PERTIPO) VALUES (" & _
    "now(), '" & txtObservaciones.Text & "', 'A', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U',   " & _
    "'" & cliePertpId & "', '" & cliePerId & "', 'C' )"
    con.Execute (SQL1)
    
    SQL1 = "select last_insert_id() agendaId"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        agendaId = RES1.Fields("agendaId")
    End If
    
    textoCita = "Gracias por su preferencia. " & vbCrLf & vbCrLf & "Se ha generado su cita con la siguiente información: " & vbCrLf & vbclrf & vbCrLf & _
    "Cliente: " & vbCrLf & lInfo(2).Caption & vbCrLf & lInfo(3).Caption & vbCrLf & vbCrLf
 
    With lista
        For b1 = 1 To .Rows - 1
            SQL1 = "INSERT INTO AGENDA_SERVICIOS (agds_agdId, agds_ServId, agds_SerTipo, agds_Inicio, agds_Fin, agds_Status, " & _
            "agds_Usuario_Id, agds_Usuario_PerId, agds_Usuario_PerTipo, agds_ServPrecio, agds_FechaHora, agds_Tipo) VALUES ( " & _
            "'" & agendaId & "', '" & .TextMatrix(b1, 6) & "', 'S', '" & Format(.TextMatrix(b1, 2), "yyyy-MM-dd") & " " & Format(.TextMatrix(b1, 3), "hh:mm:ss") & "',  " & _
            "'" & Format(.TextMatrix(b1, 2), "yyyy-MM-dd") & " " & Format(.TextMatrix(b1, 4), "hh:mm:ss") & "', 'A', '" & .TextMatrix(b1, 7) & "', '" & .TextMatrix(b1, 8) & "', '" & .TextMatrix(b1, 9) & "', '0.0', now(), 'G')"
            con.Execute (SQL1)
            
            textoCita = textoCita & vbCrLf & "Servicio: " & .TextMatrix(b1, 0) & vbCrLf & "Tipo de Servicio: " & .TextMatrix(b1, 1) & vbCrLf & "Fecha: " & .TextMatrix(b1, 2) & vbCrLf & "Hora: " & .TextMatrix(b1, 3) & _
            vbCrLf & "Atendera: " & .TextMatrix(b1, 5) & vbCrLf
        Next b1
    End With
    
    textoCita = textoCita & vbCrLf & vbCrLf & txtObservaciones.Text & vbCrLf & vbclrf & "Te recomendamos estar 15 minutos antes de tu cita. " & vbCrLf & vbclrf & "Que tengas un excelente día."

    
    salida = True
    FRM_Agenda.cmdCitas_Click
    
    MsgBox "Información guardada", vbInformation
    
    If email <> "" Then
        Call enviar_Mail("CITA", "Cita " & FRM_Menu.menuBarra2.Panels(9).Text & " " & dtFecha1, email, textoCita)
    Else
        MsgBox "No se puede enviar confirmación de cita por correo por falta de información. Verifique.", vbInformation
    End If
    
    Unload Me
End Sub
Private Sub cmbServ_Click(Index As Integer)
    Select Case Index
        Case 0:
            cargaServicio
        Case 2:
            valorUsuario
    End Select
End Sub
Private Sub valorUsuario()
    SQL1 = "SELECT PERTP_TIPO_ID FROM PER_TIPO WHERE PERTP_PER_ID = '" & cmbServ(2).ItemData(cmbServ(2).ListIndex) & "' AND PERTP_per_TIPO = 'U'"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        userPertpId = RES1.Fields("PERTP_TIPO_ID")
    End If
    
End Sub
Private Sub cargaServicio()
    cmbServ(1).Clear
    SQL1 = "SELECT PROD_ID, PROD_NOMBRE FROM PRODUCTOS WHERE PROD_SERV = 'S' AND PROD_TIPO = '" & cmbServ(0).ItemData(cmbServ(0).ListIndex) & "' AND PROD_SUBTIPO = 'S'"
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        cmbServ(1).AddItem RES1.Fields("PROD_NOMBRE")
        cmbServ(1).ItemData(cmbServ(1).ListCount - 1) = RES1.Fields("PROD_ID")
        RES1.MoveNext
    Loop

End Sub

Private Sub cmdAdd_Click(Index As Integer)
    
    If Index = 0 Then
        agregarCliente
    Else
        If Index = 1 Then
            addServicioCita
        End If
    End If
End Sub


Private Sub agregarCliente()
'    ADD_Cliente.txtUsuario(0).Text = UCase(txtUsuario(0).Text)
'    ADD_Cliente.Show vbModal
    
    checarCampos
    If checkError = False Then
        crearCliente
    Else
        MsgBox "Se detecto un error. Por favor verifique. ", vbExclamation
    End If
    
    
End Sub
Private Sub checarCampos()
    checkError = False
        
    If txtUsuario(0).Text = "" Then
        checkError = True
        lUsuario(0).ForeColor = vbRed
        Exit Sub
    Else
        If txtUsuario(1).Text = "" Then
            checkError = True
            lUsuario(0).ForeColor = vbRed
            Exit Sub
        End If
    End If
    
    For b1 = 2 To 5
        If txtUsuario(b1).Text = "" Then
            txtUsuario(b1).Text = "-"
        End If
    Next b1
    
End Sub


Private Sub crearCliente()

    Dim status As String
    Dim idEstado As String
    Dim idMunicipio As String
    Dim idEstadoNac As String
    Dim genero As String
    Dim cp As String
    Dim tel1 As String
    Dim tel2 As String
    Dim telAccdte As String
    Dim membresia As String
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    Dim membresiaCodigo As String
    Dim tipoPerId As Long
    
    'genero = Left(cmbUser(2).Text, 1)
    If txtUsuario(3).Text = "-" Or txtUsuario(3).Text = "" Then
        tel1 = "null"
    Else
        tel1 = txtUsuario(3).Text
    End If
    If txtUsuario(4).Text = "-" Or txtUsuario(4).Text = "" Then
        tel2 = "null"
    Else
        tel2 = txtUsuario(4).Text
    End If
                    
    SQL1 = "INSERT INTO PERSONA (PER_NOMBRE, PER_PATERNO, PER_MATERNO, PER_FEC_NAC, " & _
    "PER_EMAIL, PER_FECHA_SISTEMA, PER_GENERO, PER_TEL1, PER_TEL2) VALUES " & _
    "('" & txtUsuario(0).Text & "', '" & txtUsuario(1).Text & "', '" & txtUsuario(2).Text & "', NOW(), " & _
    "'" & txtUsuario(5).Text & "', now(), 'M', " & tel1 & ", " & tel2 & " )"
    'MsgBox SQL1
    con.Execute (SQL1)
    
    SQL1 = "select last_insert_id() perId"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        perId = RES1.Fields("perId")
    End If
    
    membresiaCodigo = perId
    
    SQL1 = "select CTPT_ID  CTID From cat_tipo where ctpt_subtipo = 'C' LIMIT 0, 1"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        tipoPerId = RES1.Fields("CTID")
    Else
        tipoPerId = ""
    End If
    
    
    SQL1 = "INSERT INTO PER_TIPO (PERTP_TIPO_ID, PERTP_PER_ID, PERTP_FECHA, PERTP_PER_TIPO, PERTP_STATUS, PERTP_ALTA, PERTP_CODIGO_MEMBRESIA, " & _
    "PERTP_PERALTA_ID, PERTP_PERALTA_TIPO_ID, PERTP_PERALTA_TIPO, PERTP_PERALTA_FECHA) " & _
    "VALUES " & _
    "(" & tipoPerId & ", " & perId & ", now(), 'C', 'A', now(), " & _
    "'" & membresiaCodigo & "', " & _
    "'" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', NOW())"
    con.Execute (SQL1)
    
    Call buscarCliente(txtUsuario(0).Text, txtUsuario(1).Text, txtUsuario(2).Text, txtUsuario(3).Text, txtUsuario(4).Text, txtUsuario(5).Text)
    
    MsgBox "Cliente guardado.", vbInformation

    'Unload Me
    
End Sub

Private Sub addServicioCita()
        checkDatos
        
        If datos = True Then
            MsgBox mensaje, vbInformation
        Else
            lista.AddItem ""
            lista.TextMatrix(lista.Rows - 1, 0) = cmbServ(1).Text
            lista.TextMatrix(lista.Rows - 1, 1) = cmbServ(0).Text
            lista.TextMatrix(lista.Rows - 1, 2) = dtFecha1
            lista.TextMatrix(lista.Rows - 1, 3) = Format((cmbHora(0).Text & ":" & cmbHora(1).Text), "Short Time")
            lista.TextMatrix(lista.Rows - 1, 4) = Format((cmbHora(2).Text & ":" & cmbHora(3).Text), "Short Time")
            lista.TextMatrix(lista.Rows - 1, 5) = cmbServ(2).Text
            lista.TextMatrix(lista.Rows - 1, 6) = cmbServ(1).ItemData(cmbServ(1).ListIndex)
            lista.TextMatrix(lista.Rows - 1, 7) = userPertpId
            lista.TextMatrix(lista.Rows - 1, 8) = cmbServ(2).ItemData(cmbServ(2).ListIndex)
            lista.TextMatrix(lista.Rows - 1, 9) = "U"
            
            If tipoCita = "Edicion" Then
                SQL1 = "INSERT INTO AGENDA_SERVICIOS (agds_agdId, agds_ServId, agds_SerTipo, agds_Inicio, agds_Fin, agds_Status, " & _
                "agds_Usuario_Id, agds_Usuario_PerId, agds_Usuario_PerTipo, agds_ServPrecio, agds_FechaHora, agds_Tipo) VALUES ( " & _
                "'" & idAgenda & "', '" & lista.TextMatrix(lista.Rows - 1, 6) & "', 'S', '" & Format(lista.TextMatrix(lista.Rows - 1, 2), "yyyy-MM-dd") & " " & Format(lista.TextMatrix(lista.Rows - 1, 3), "hh:mm:ss") & "',  " & _
                "'" & Format(lista.TextMatrix(lista.Rows - 1, 2), "yyyy-MM-dd") & " " & Format(lista.TextMatrix(lista.Rows - 1, 4), "hh:mm:ss") & "', 'A', '" & lista.TextMatrix(lista.Rows - 1, 7) & "', '" & lista.TextMatrix(lista.Rows - 1, 8) & "', '" & lista.TextMatrix(lista.Rows - 1, 9) & "', '0.0', now(), 'G')"
                'MsgBox SQL1
                con.Execute (SQL1)
            End If
    
        End If

End Sub
Private Sub checkDatos()
    datos = False
    mensaje = "Sin motivo"
    If listaClte.Row < 1 Then
        If tipoCita = "Creacion" Then
            mensaje = "Debe haber seleccionado un cliente para generar la cita" & vbCrLf & vbCrLf & _
                    "Verifique"
            datos = True
            Exit Sub
        End If
    Else
        For b1 = 0 To 2
            If cmbServ(b1).Text = "" Then
                mensaje = "Debe haber seleccionado un servicio" & vbCrLf & vbCrLf & _
                    "Verifique"
                datos = True
                Exit Sub
            End If
        Next b1
        
        For b1 = 1 To lista.Rows - 1
            'MsgBox Format((cmbHora(0).Text & ":" & cmbHora(1).Text), "Short Time") & " >= " & Format(lista.TextMatrix(b1, 3), "Short Time") & " And " & Format((cmbHora(0).Text & ":" & cmbHora(1).Text), "Short Time") & " < " & Format(lista.TextMatrix(b1, 4), "Short Time")
            If Format((cmbHora(0).Text & ":" & cmbHora(1).Text), "Short Time") >= Format(lista.TextMatrix(b1, 3), "Short Time") And Format((cmbHora(0).Text & ":" & cmbHora(1).Text), "Short Time") < Format(lista.TextMatrix(b1, 4), "Short Time") Then
                If cmbServ(2).Text = lista.TextMatrix(b1, 5) Then
                    mensaje = "Se ha detectado un servicio en la lista con la misma hora al que quiere agregar. " & vbCrLf & vbCrLf & _
                    "Verifique"
                    datos = True
                    Exit Sub
                End If
            End If
        Next b1
    End If
    
End Sub
Private Sub Form_Load()
salida = False
cargaHoras
cargaDatos
listaClte.Rows = 1
End Sub
Private Sub cargaDatos()
    lista.Rows = 1
    lista.ColWidth(6) = 0
    lista.ColWidth(7) = 0
    lista.ColWidth(8) = 0
    lista.ColWidth(9) = 0
    
    dtFecha1 = Date
    cargaServTipo
    cargaUsuarios
    
    If tipoCita = "Creacion" Then
        citaNueva
        lInfo(0).Visible = False
    Else
        citaEdit
        lInfo(0).Visible = True
        lInfo(0).Caption = "Clave cita: " & clavesCitas(FRM_Agenda.listaDia.Row, FRM_Agenda.listaDia.Col) & ", Ubicación: " & FRM_Agenda.listaDia.Row & "-" & FRM_Agenda.listaDia.Col
        idAgenda = clavesCitas(FRM_Agenda.listaDia.Row, FRM_Agenda.listaDia.Col)
    End If
    
End Sub
Private Sub citaEdit()
    lista.Rows = 1
    With FRM_Agenda.listaDia
        SQL2 = "SELECT * FROM VIEW_CITAS WHERE CLAVE = '" & clavesCitas(.Row, .Col) & "'"
        Set RES2 = con.Execute(SQL2)
        If Not RES2.EOF Then
            'txtUsuario(0).Text = RES2.Fields("clie_Nombre")
            lInfo(2).Caption = "Nombre: " & RES2.Fields("CLIE_NOMBRE")
            lInfo(3).Caption = "Apellidos: " & RES2.Fields("CLIE_PATERNO") & " " & RES2.Fields("CLIE_MATERNO")
            lInfo(7).Caption = "Email: " & RES2.Fields("CLIE_EMAIL")
            lInfo(8).Caption = "Teléfonos: " & RES2.Fields("CLIE_TEL1") & " " & RES2.Fields("CLIE_TEL2")
            'lInfo(2).Caption = "Nombre: " & RES2.Fields("CLIE_NOMBRE")
            
            For b1 = 1 To listaClte.Rows - 1
                If listaClte.TextMatrix(b1, 1) = RES2.Fields("CLIE_ID") Then
                    listaClte_Click
                End If
            Next b1
        End If
        Do While Not RES2.EOF
            lista.AddItem ""
            lista.TextMatrix(lista.Rows - 1, 0) = RES2.Fields("SERVICIO")
            lista.TextMatrix(lista.Rows - 1, 1) = RES2.Fields("TIPO_SERVICIO")
            lista.TextMatrix(lista.Rows - 1, 2) = RES2.Fields("FECHA")
            lista.TextMatrix(lista.Rows - 1, 3) = RES2.Fields("HORA_INI")
            lista.TextMatrix(lista.Rows - 1, 4) = RES2.Fields("HORA_FIN")
            lista.TextMatrix(lista.Rows - 1, 5) = RES2.Fields("USUARIO")
            lista.TextMatrix(lista.Rows - 1, 6) = RES2.Fields("SERV_ID")
            lista.TextMatrix(lista.Rows - 1, 7) = RES2.Fields("USUARIO_PERTPID")
            lista.TextMatrix(lista.Rows - 1, 8) = RES2.Fields("USUARIO_ID")
            lista.TextMatrix(lista.Rows - 1, 9) = "C"
            
            RES2.MoveNext
        Loop
        
        
    End With
End Sub
Private Sub citaNueva()
    With FRM_Agenda.listaDia
        cmbServ(2).Text = .TextMatrix(1, .Col)
        cmbHora(0).Text = Left(.TextMatrix(.Row, 0), 2)
        cmbHora(1).Text = Right(.TextMatrix(.Row, 0), 2)
        dtFecha1 = FRM_Agenda.dtFecha1
        If dtFecha1 < Date Then
            MsgBox "La fecha de la cita es posterior a la fecha actual.", vbInformation
        End If
    End With
End Sub
Private Sub cargaServTipo()
    SQL1 = "SELECT CTPT_ID, CTPT_TIPO FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'S'"
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        cmbServ(0).AddItem RES1.Fields("CTPT_TIPO")
        cmbServ(0).ItemData(cmbServ(0).ListCount - 1) = RES1.Fields("CTPT_ID")
        RES1.MoveNext
    Loop
End Sub
Private Sub cargaUsuarios()
    SQL1 = "SELECT PER_ID, CONCAT(PER_NOMBRE, ' ', PER_PATERNO, ' ', PER_MATERNO) USUARIO " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND PERTP_PER_TIPO = 'U' AND PERTP_STATUS = 'A' and T2.PERTP_AGENDA = '1' "
    Set RES1 = con.Execute(SQL1)

    Do While Not RES1.EOF
        cmbServ(2).AddItem RES1.Fields("USUARIO")
        cmbServ(2).ItemData(cmbServ(2).ListCount - 1) = RES1.Fields("PER_ID")
        RES1.MoveNext
    Loop

End Sub
Private Sub cargaHoras()
    Dim media As String
    Dim tiempo As Long
    Dim tiempo2 As String
    Dim hora As String
    
    SQL1 = "SELECT SUC_HORAENTRADA ENTRADA, SUC_HORASALIDA SALIDA FROM SUCURSAL WHERE SUC_LOCAL = 'S'"
    Set RES1 = con.Execute(SQL1)
    
    
    For b1 = 0 To 3
        cmbHora(b1).Clear
    Next b1
    
    If Not RES1.EOF Then
        media = "00"
        cmbHora(1).AddItem "00"
        cmbHora(1).AddItem "30"
        cmbHora(3).AddItem "00"
        cmbHora(3).AddItem "30"
        
        'listaDia.RowHeight(listaDia.Rows - 1) = 730
        tiempo = DateDiff("n", Format(RES1.Fields("Entrada"), "Short Time"), Format(RES1.Fields("Salida"), "Short Time"))
        tiempo = Val(tiempo) / 60
        For b1 = 0 To tiempo
            hora = Hour(Format(RES1.Fields("Entrada"), "Short Time")) + b1
            tiempo2 = Format(hora, "00") '& ":00"
            cmbHora(0).AddItem tiempo2
            cmbHora(2).AddItem tiempo2
        Next b1
            
    Else
        MsgBox "Debe de establecer un horario para la sucursal en la cual está laborando. ", vbInformation
    End If

    cmbHora(0).ListIndex = 0
    cmbHora(1).ListIndex = 0
    cmbHora(2).ListIndex = 1
    cmbHora(3).ListIndex = 0
    
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If salida = False Then
        Dim ques As String
        ques = MsgBox("¿Salir?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            Cancel = 0
        Else
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lista.Rows >= 1 Then
        If Button = vbRightButton Then
            mn_CancelarCita.Enabled = True
            PopupMenu menu, vbPopupMenuLeftAlign
        End If
    Else
        If Button = vbRightButton Then
            mn_CancelarCita.Enabled = False
            PopupMenu menu, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub listaClte_Click()
    If listaClte.Rows > 1 Then
        muestraInfo (listaClte.TextMatrix(listaClte.Row, 4))
    End If
End Sub
Private Sub muestraInfo(valor As Long)
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    SQL1 = "SELECT T2.PERTP_TIPO_ID , T1.PER_NOMBRE, concat(T1.PER_PATERNO, ' ', T1.PER_MATERNO) APELLIDOS, T1.PER_ID, T1.PER_EMAIL, CONCAT(PER_TEL1, ' ', PER_TEL2) TELEFONOS, PER_FOTO " & _
    "FROM PERSONA T1, PER_TIPO T2 WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'C' AND  T1.PER_ID =  '" & valor & "' "
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        cliePertpId = RES1.Fields("PERTP_TIPO_ID")
        cliePerId = RES1.Fields("PER_ID")
        lInfo(2).Caption = "Nombre: " & RES1.Fields("PER_NOMBRE")
        lInfo(3).Caption = "Apellidos: " & RES1.Fields("APELLIDOS")
        lInfo(7).Caption = "Email: " & RES1.Fields("PER_EMAIL")
        lInfo(8).Caption = "Teléfonos: " & RES1.Fields("TELEFONOS")
        
        If RES1.Fields("per_EMAIL") <> "" Then
            email = RES1.Fields("per_EMAIL")
        Else
            email = ""
        End If

'        If IsNull(RES1.Fields("PER_fOTO")) = False Then
'            checarCarpetaTemp
'            Imagen1.Open
'            Imagen1.Write RES1.Fields("PER_FOTO")
'            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
'            Imagen1.Close
'            fotoUser.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
'        Else
'            fotoUser.Picture = LoadPicture("")
'        End If
    Else
        borraInfo
    End If
End Sub
Private Sub borraInfo()
    lInfo(2).Caption = "Nombre: "
    lInfo(3).Caption = "Apellidos: "
    lInfo(7).Caption = "Email: "
    lInfo(8).Caption = "Teléfonos: "
End Sub

Private Sub listaClte_GotFocus()
    ConScroll listaClte
End Sub

Private Sub listaClte_LostFocus()
    SinScroll listaClte
End Sub

Private Sub mn_CancelarCita_Click()
    Dim ques As String
    
    
    If tipoCita = "Creacion" Then
        ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            If lista.Rows > 2 Then
                lista.RemoveItem (lista.Row)
            Else
                lista.Rows = 1
            End If
        End If
    Else
        If tipoCita = "Edicion" Then
            ques = MsgBox("Al cancelar será eliminado definitivamente de la cita." & vbCrLf & "¿Continuar cancelación?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                SQL1 = "DELETE FROM AGENDA_SERVICIOS  " & _
                "WHERE agds_agdId = '" & clavesCitas(FRM_Agenda.listaDia.Row, FRM_Agenda.listaDia.Col) & "' AND " & _
                "agds_ServId = '" & lista.TextMatrix(lista.Row, 6) & "' AND agds_Usuario_PerId = '" & lista.TextMatrix(lista.Row, 8) & "'"
                con.Execute (SQL1)
                If lista.Rows > 2 Then
                    lista.RemoveItem (lista.Row)
                Else
                    lista.Rows = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub txtUsuario_Change(Index As Integer)
    Call buscarCliente(txtUsuario(0).Text, txtUsuario(1).Text, txtUsuario(2).Text, txtUsuario(3).Text, txtUsuario(4).Text, txtUsuario(5).Text)
End Sub
Private Sub buscarCliente(Nombre As String, Paterno As String, Materno As String, tel1 As String, tel2 As String, email As String)
    borraInfo
    listaClte.Rows = 1

    SQL1 = "SELECT CONCAT(T1.PER_NOMBRE, ' ', T1.PER_PATERNO, ' ', T1.PER_MATERNO) CLIENTE, T1.PER_ID, T1.PER_EMAIL, T1.PER_TEL1, T1.PER_TEL2 " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'C' AND  UPPER(T1.PER_NOMBRE) LIKE  UPPER('%" & Nombre & "%') AND UPPER(T1.PER_PATERNO) LIKE  UPPER('%" & Paterno & "%') AND UPPER(T1.PER_MATERNO) LIKE UPPER('%" & Materno & "%') " & _
    "AND UPPER(T1.PER_EMAIL) LIKE  UPPER('%" & email & "%') " & _
    "AND (T1.PER_TEL1 LIKE  '%" & tel1 & "%' OR T1.PER_TEL1 IS NULL) AND ( T1.PER_TEL2 LIKE  '%" & tel2 & "%' OR T1.PER_TEL2 IS NULL)  " & _
    "Limit 0, 100"
    Set RES1 = con.Execute(SQL1)
    
    Do While Not RES1.EOF
        listaClte.AddItem ""
        listaClte.TextMatrix(listaClte.Rows - 1, 0) = RES1.Fields("CLIENTE")
        listaClte.TextMatrix(listaClte.Rows - 1, 1) = RES1.Fields("PER_TEL1") & ""
        listaClte.TextMatrix(listaClte.Rows - 1, 2) = RES1.Fields("PER_TEL2") & ""
        listaClte.TextMatrix(listaClte.Rows - 1, 3) = RES1.Fields("PER_EMAIL")
        listaClte.TextMatrix(listaClte.Rows - 1, 4) = RES1.Fields("PER_ID")
    RES1.MoveNext
    Loop
    
    lInfo(1).Caption = listaClte.Rows - 1 & "   Clientes en lista"
End Sub

Private Sub txtUsuario_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call buscarCliente(txtUsuario(0).Text, txtUsuario(1).Text, txtUsuario(2).Text, txtUsuario(3).Text, txtUsuario(4).Text, txtUsuario(5).Text)
    End If
End Sub
