VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PRINT_Impresora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresora"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3600
      Width           =   4455
   End
   Begin VB.CommandButton cSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   855
      Index           =   1
      Left            =   3960
      Picture         =   "PRINT_Impresora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cImprimir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cambiar"
      Height          =   855
      Index           =   1
      Left            =   2040
      Picture         =   "PRINT_Impresora.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cGuardar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guardar"
      Height          =   855
      Left            =   120
      Picture         =   "PRINT_Impresora.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog dialogo1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus de la Impresora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "PRINT_Impresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cGuardar_Click()
Dim a As String
a = MsgBox("¿Guardar estatus " & Combo1.Text & "?", vbYesNo + vbQuestion)
If a = vbYes Then
    SQL1 = "UPDATE sucursal SET SUC_ESTATUSTICKET = '" & Combo1.ListIndex & "'"
    con.Execute (SQL1)
    MsgBox "Datos guardados.", vbInformation
End If

If impresoraTicket = "cobro" Then
    If Combo1.ListIndex = 1 Then
        FRM_Cobro.Image1(0).Visible = True
        FRM_Cobro.Image1(1).Visible = False
    Else
        FRM_Cobro.Image1(1).Visible = True
        FRM_Cobro.Image1(0).Visible = False
    End If
    FRM_Cobro.txtPrint.Text = Printer.DeviceName
End If
End Sub

Private Sub cImprimir_Click(Index As Integer)
a = MsgBox("¿Cambiar la impresora actual?", vbYesNo + vbQuestion)

If a = vbYes Then
    MsgBox "Haga doble click sobre la impresora que desea seleccionar como predetermiada.", vbInformation
    dialogo1.ShowPrinter
    txtPrint.Text = Printer.DeviceName
End If

End Sub

Private Sub cSalir_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
    'txtPrint.Text = Printer.DeviceName
    cargaCombo
    checkIMpresora
End Sub
Private Sub cargaCombo()
    Combo1.Clear
    Combo1.AddItem "Inactivo"
    Combo1.AddItem "Activo"
    
End Sub
Private Sub checkIMpresora()
    txtPrint.Text = Printer.DeviceName
    
    SQL1 = "SELECT SUC_ESTATUSTICKET FROM SUCURSAL"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        If RES1.Fields("SUC_ESTATUSTICKET") = 1 Then
            Combo1.Text = "Activo"
        Else
            Combo1.Text = "Inactivo"
        End If

    Else
            Combo1.Text = "Inactivo"
    End If
    
End Sub
