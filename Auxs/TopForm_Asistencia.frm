VERSION 5.00
Begin VB.Form TopForm_Asistencia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Datos de asistencia"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar Hs1 
      Height          =   255
      Left            =   3600
      Max             =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer TimHs2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimHs1 
      Interval        =   50
      Left            =   0
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   0
      Top             =   840
   End
   Begin VB.Label lDatos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Image iFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   480
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "TopForm_Asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dato As Boolean

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Call Aplicar_Transparencia(Me.hwnd, CByte(0))
    Ubicacion
    Dato = True

End Sub
Private Sub Ubicacion()
    Me.Left = Screen.width - Me.width
    Me.Top = Screen.height - Me.height
'    Me.Top = 0
End Sub

Private Sub Hs1_Change()
    Call Aplicar_Transparencia(Me.hwnd, CByte(Hs1.Value))

End Sub

Private Sub Timer1_Timer()
Unload Me

End Sub

Private Sub TimHs1_Timer()
    Hs1.Value = Hs1.Value + 12
    If Hs1 > 230 Then
        Hs1.Value = 255
        TimHs1.Enabled = False
        Timer1.Enabled = True
    End If

End Sub

Private Sub TimHs2_Timer()
    Hs1.Value = Hs1.Value - 12
    If Hs1 <= 50 Then
        Hs1.Value = 0
        'FRM_HuellasCheck.Imagen(3).Picture = LoadPicture("")
        'FRM_HuellasCheck.Image1.Visible = False
        'FRM_HuellasCheck.Image2.Visible = False
        'FRM_HuellasCheck.AreaVerificar.Caption = ""
        'FRM_HuellasCheck.NombreVerificar.Caption = ""
        'FRM_HuellasCheck.Mensajes.Caption = ""
        'FRM_HuellasCheck.Detector.Caption = ""
        
        TimHs2.Enabled = False
        Dato = False
        Unload Me
    End If

End Sub
