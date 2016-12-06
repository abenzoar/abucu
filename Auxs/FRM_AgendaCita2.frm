VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_AgendaCita2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la cita"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   Icon            =   "FRM_AgendaCita2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   873
      TabCaption(0)   =   "Datos generales"
      TabPicture(0)   =   "FRM_AgendaCita2.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDatos(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgFoto(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatos(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgFoto(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatos(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgFoto(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line1(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line1(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lUsuario(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lUsuario(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lUsuario(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line1(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(23)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblDatos(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(22)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblDatos(6)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(20)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblDatos(5)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(19)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblDatos(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(18)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblDatos(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lista"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "dtFecha1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtClave(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtClave(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtClave(2)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmbHora(3)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmbHora(2)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmbHora(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmbHora(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtObservaciones"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdAdd(1)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Time_listaRapida"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "Modo Touch"
      TabPicture(1)   =   "FRM_AgendaCita2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Timer Time_listaRapida 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   14040
         Top             =   1080
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
         Left            =   6360
         TabIndex        =   25
         Top             =   -5000
         Width           =   8175
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
         Left            =   8280
         Picture         =   "FRM_AgendaCita2.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3480
         Width           =   495
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
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   6840
         Width           =   12975
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3480
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3480
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3480
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
         Index           =   3
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3480
         Width           =   1215
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
         Left            =   8640
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
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
         Left            =   5040
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
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
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3480
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
         Format          =   113115137
         CurrentDate     =   40956
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   2295
         Left            =   240
         TabIndex        =   24
         Top             =   4200
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
         FormatString    =   $"FRM_AgendaCita2.frx":0E8C
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
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   13800
         TabIndex        =   42
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   13800
         TabIndex        =   41
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblClieId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   13800
         TabIndex        =   40
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   2
         Left            =   12840
         TabIndex        =   39
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   1
         Left            =   12840
         TabIndex        =   38
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         Caption         =   "Label10"
         Height          =   255
         Index           =   0
         Left            =   12840
         TabIndex        =   37
         Top             =   -5000
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   10560
         TabIndex        =   36
         Top             =   3360
         Width           =   5415
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
         Left            =   10560
         TabIndex        =   35
         Top             =   720
         Width           =   1335
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
         Left            =   10560
         TabIndex        =   34
         Top             =   960
         Width           =   6375
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
         Left            =   10560
         TabIndex        =   33
         Top             =   1920
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
         Left            =   10560
         TabIndex        =   32
         Top             =   2160
         Width           =   975
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
         Left            =   11760
         TabIndex        =   31
         Top             =   1920
         Width           =   1095
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
         Left            =   11760
         TabIndex        =   30
         Top             =   2520
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
         Left            =   11760
         TabIndex        =   29
         Top             =   2280
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
         Left            =   13440
         TabIndex        =   28
         Top             =   2520
         Width           =   1095
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
         Left            =   13440
         TabIndex        =   27
         Top             =   2280
         Width           =   855
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
         Left            =   10560
         TabIndex        =   26
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   1
         X1              =   240
         X2              =   3360
         Y1              =   3120
         Y2              =   3120
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
         Left            =   240
         TabIndex        =   21
         Top             =   3240
         Width           =   2415
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
         Left            =   2760
         TabIndex        =   20
         Top             =   3240
         Width           =   975
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
         Left            =   5520
         TabIndex        =   19
         Top             =   3240
         Width           =   975
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   3
         X1              =   7200
         X2              =   10200
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   5
         X1              =   5040
         X2              =   6720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   6
         X1              =   8640
         X2              =   10320
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   2
         Left            =   7200
         Stretch         =   -1  'True
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
         Index           =   2
         Left            =   8640
         TabIndex        =   12
         Top             =   1080
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
         Left            =   7200
         TabIndex        =   11
         Top             =   720
         Width           =   2175
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
         Left            =   8640
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   1
         Left            =   3720
         Stretch         =   -1  'True
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
         Left            =   5040
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   3720
         X2              =   6720
         Y1              =   960
         Y2              =   960
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
         Left            =   5040
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
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
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   0
         X1              =   240
         X2              =   3240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
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
         Height          =   855
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   2
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   4
         X1              =   1560
         X2              =   3240
         Y1              =   2040
         Y2              =   2040
      End
   End
End
Attribute VB_Name = "FRM_AgendaCita2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset

Private Sub lista_rapida_Click()
'''
End Sub

Private Sub lista_rapida_DblClick()
    'cargaFrom_ListaRapida
End Sub

Private Sub lista_rapida_GotFocus()
    Time_listaRapida.Enabled = False
End Sub

Private Sub lista_rapida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cargaFrom_ListaRapida
    End If
End Sub

Private Sub lista_rapida_LostFocus()
    lista_rapida.Visible = False
End Sub

Private Sub lista_SelChange()
    'Lista_Click
End Sub

Private Sub Time_listaRapida_Timer()
Time_listaRapida.Enabled = False
lista_rapida.Visible = False

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
   
    
    'Set FrmFocus = Me
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


Private Sub checkProducto()

'    On Error Resume Next
        
    lista_rapida.Visible = False
     monedero = False
'    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
'        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
'        Exit Sub
'    End If
    
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
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
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
        
        'addLista
        
'        If monedero = True Then
'            Call addMonedero(0, 0)
'        End If
    Else
        
       ' checkServicio
    End If
    
    
End Sub

Private Sub checkUsuario()
On Error Resume Next

'    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
'        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
'        Exit Sub
'    End If

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
'        updateVenta (Val(lInfo(1).Caption))
'        txtClave(0).SetFocus
    Else
        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation
    End If
    
End Sub
Private Sub checkCliente()
'    On Error Resume Next
'
'    lista_rapida.Visible = False
'
''    If UCase(lInfo(2).Caption) <> "ABIERTO" Then
''        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
''        Exit Sub
''    End If
'
'    sql1 = "SELECT PERTP_USUARIO, IF(PERTP_MEMBRESIA ='S', 'SI', 'NO') MEMBRESIA, PERTP_CODIGO_MEMBRESIA, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_PER_TIPO, PERTP_TIPO_ID, CTPT_TIPO, T1.PER_ID, PER_FOTO, PER_EMAIL, t2.TEMP_MONEDERO, (SELECT T4.TOTAL FROM VIEW_MONEDERO_CLIENTES T4 WHERE T1.PER_ID = T4.PER_ID) TOTAL " & _
'    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
'    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'C' " & _
'    "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'C' AND " & _
'    "T2.PERTP_CODIGO_MEMBRESIA = '" & txtClave(2).Text & "'"
'    'MsgBox SQL1
'    Set RES1 = con.Execute(sql1)
'
'    If Not RES1.EOF Then
'
''       PARA CHECAR QUE SOLO UNA VENTA POR CLIENTA EXISTA
'        sql1 = "SELECT COUNT(*) NUM  fROM VENTAS " & _
'        "WHERE VENT_CLIEPERID = '" & RES1.Fields("PER_ID") & "' AND VENT_STATUS = 'G' AND VENT_IDFOLIO <> '" & Val(lInfo(1).Caption) & "'"
'
'        Set RES2 = con.Execute(sql1)
'
'        If Not RES2.EOF Then
'            If Val(RES2.Fields("num")) > 0 And RES1.Fields("PERTP_CODIGO_MEMBRESIA") <> "CLTE" Then
''                    lInfo(0).Caption = "0"
''                    lblDatos(3).Caption = ""
'                    MsgBox "El cliente que quiere agregar tiene " & RES2.Fields("NUM") & " operaciones generadas.  " & vbCrLf & vbCrLf & _
'                    "Verifique para cerrar o cancelar la venta del cliente. Verfique.", vbInformation
'                    'Exit Sub
'
'            End If
'        End If
'
'
'        lblDatos(2).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
'        lblDatos(4).Caption = RES1.Fields("PER_NOMBRE") & " " & RES1.Fields("PER_PATERNO") & " " & RES1.Fields("PER_MATERNO")
'        lblDatos(5).Caption = RES1.Fields("MEMBRESIA")
'
'        If IsNull(RES1.Fields("total")) Then
'        lblDatos(6).Caption = FormatCurrency(0)
'        Else
'        lblDatos(6).Caption = FormatCurrency(Val(RES1.Fields("TOTAL")))
'        End If
'        lblClieId(0).Caption = RES1.Fields("PER_ID")
'        lblClieId(1).Caption = RES1.Fields("PERTP_TIPO_ID")
'        lblClieId(2).Caption = RES1.Fields("PERTP_PER_TIPO")
'        lblDatos(3).Caption = RES1.Fields("PER_EMAIL")
'        'lInfo(0).Caption = FormatCurrency(RES1.Fields("TEMP_MONEDERO"))
'        Me.Caption = "Operación Ticket " & lInfo(1).Caption & " Clte: " & lblDatos(2).Caption
'
'        If IsNull(RES1.Fields("PER_fOTO")) = False Then
'            Dim Imagen1 As Stream
'            Set Imagen1 = New Stream
'            Imagen1.Type = adTypeBinary
'            checarCarpetaTemp
'            Imagen1.Open
'            Imagen1.Write RES1.Fields("PER_FOTO")
'            Imagen1.SaveToFile direccionSistema & "\Temp\TempClie.dat", adSaveCreateOverWrite
'            Imagen1.Close
'            imgFoto(2).Picture = LoadPicture(direccionSistema & "\Temp\TempClie.dat")
'        Else
'            imgFoto(2).Picture = LoadPicture("")
'        End If
'
'        updateVenta (Val(lInfo(1).Caption))
'
'    Else
'        lInfo(0).Caption = "0"
'        lblDatos(3).Caption = ""
'        MsgBox "Información incorrecta. Por favor verifique. ", vbInformation
'    End If
'
End Sub


