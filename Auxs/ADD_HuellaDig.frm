VERSION 5.00
Begin VB.Form ADD_HuellaDig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de huella digital"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
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
      Height          =   855
      Index           =   0
      Left            =   1440
      Picture         =   "ADD_HuellaDig.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
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
      Left            =   3360
      Picture         =   "ADD_HuellaDig.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox grFinger 
      Height          =   480
      Left            =   3960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Mensajes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Huella 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Huella 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Index           =   1
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Image Imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Index           =   2
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   3255
      Left            =   1080
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "ADD_HuellaDig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImagenNumero As Integer
Dim SQL1 As String
Dim RES1 As Recordset
Dim SALIDA As Boolean

Private Sub cmBoton_Click(Index As Integer)
Select Case Index
    Case 0: botonGuardar
    Case 1: cancelar
    
End Select
End Sub
Private Sub cancelar()
    Unload Me
End Sub
Private Sub botonGuardar()
    If Imagen(1).Picture <> 0 And Imagen(2).Picture <> 0 Then
        Dim rs As New ADODB.Recordset
'        rs.Open "Select * from usuarioshuella where 1=0", con, adOpenStatic, adLockOptimistic
'        With rs
'          .AddNew
'          '.Fields("nombre") = Text1.Text
'          .Fields("huella1") = template(1).tpt
'          .Fields("huella2") = template(2).tpt
'          .Update
'        End With
        If tipoHuellas = "Clientes" Then
            rs.Open "Select * from per_tipo where PERTP_PER_TIPO = 'C' and PERTP_PER_ID = '" & idUserHuella & "'", con, adOpenStatic, adLockOptimistic
            With rs
              '.Fields("nombre") = Text1.Text
              .Fields("PERTP_huella1") = template(1).tpt
              .Fields("PERTP_huella2") = template(2).tpt
              .Update
            End With
            rs.Close
        Else
            If tipoHuellas = "Usuarios" Then
                rs.Open "Select * from per_tipo where PERTP_PER_TIPO = 'U' and PERTP_PER_ID = '" & idUserHuella & "'", con, adOpenStatic, adLockOptimistic
                With rs
                  '.Fields("nombre") = Text1.Text
                  .Fields("PERTP_huella1") = template(1).tpt
                  .Fields("PERTP_huella2") = template(2).tpt
                  .Update
                End With
                rs.Close
            End If
        End If
        'FRM_Usuarios.cHuellas.Caption = "Modificar huella"
        MsgBox "Las huellas digitales han sido guardadas.", vbInformation
        SALIDA = True
        Unload Me
    Else
        SALIDA = False
        MsgBox "No se puede realizar la operación. Verifique.", vbInformation
    End If
End Sub
Private Sub Form_Load()
    ImagenNumero = 1
    Imagen_Click (1)
    SALIDA = False
    Shape1.Left = Imagen(1).Left
    Dim Error As Integer
    Error = Inicializar(Me)
End Sub


Private Sub Imagen_Click(Index As Integer)
 ImagenNumero = Index
    If Index = 1 Then
      Shape1.Left = Imagen(1).Left
      lInfo(2).Caption = "Coloque su dedo sobre el lector y espere a que la captura 1 haya concluido."
    Else
      Shape1.Left = Imagen(2).Left
      lInfo(2).Caption = "Coloque su dedo (el mismo que el anterior) sobre el lector y espere a que la captura 2 haya concluido."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ques As String
    If SALIDA = False Then
        ques = MsgBox("¿Salir?", vbYesNo + vbQuestion)
        If ques = vbYes Then
'            grFinger_SensorUnplug (idSensor)
'            grFinger.CapStopCapture (idSensor)
'            grFinger.CapFinalize
'            grFinger.Finalize
            
            Cancel = 0
        Else
            Cancel = 1
        End If
    Else
        Cancel = 0
'            grFinger_SensorUnplug (idSensor)
'            grFinger.CapStopCapture (idSensor)
'            grFinger.CapFinalize
'            grFinger.Finalize
        'grFinger.Finalize
    End If
End Sub



''''''DE AQUI AL FINAL PARA EL DLL DE LAS HUELLAS
''''''DE AQUI AL FINAL PARA EL DLL DE LAS HUELLAS
''''''DE AQUI AL FINAL PARA EL DLL DE LAS HUELLAS
''''''DE AQUI AL FINAL PARA EL DLL DE LAS HUELLAS
''''''DE AQUI AL FINAL PARA EL DLL DE LAS HUELLAS
Private Sub grFinger_FingerDown(ByVal idSensor As String)
 ' Aqui detecta cuando pones el dedo (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Detector = "Huella detectada"
End Sub

Private Sub grFinger_FingerUp(ByVal idSensor As String)
 ' Aqui detecta cuando pones el dedo (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Detector = "Huella removida"
End Sub

Private Sub grFinger_ImageAcquired(ByVal idSensor As String, ByVal width As Long, ByVal height As Long, rawImage As Variant, ByVal res As Long)
 ' Capturar Imagen (Este mensaje es muy raro que se vea. Si se muestra pero muy rapido)
 Mensajes = "Capturando imagen..."
 
 With raw
   .img = rawImage
   .height = height
   .width = width
   .res = res
 End With
 
' If OptionGuardar.Value = True Then
   CapturaHuella False, GR_DEFAULT_CONTEXT, Me, Me.Imagen(ImagenNumero), ImagenNumero
   If EncuentraPuntos(Me, Mensajes, Imagen(ImagenNumero), ImagenNumero) = True Then
     ' Aqui entra si la Imagen se detecta bien
     If ImagenNumero = 1 Then
       Imagen_Click 2
     Else
       Imagen_Click 1
     End If
   End If
' End If

' If OptionVerificar.Value = True Then
'   CapturaHuella False, GR_DEFAULT_CONTEXT, Form1, Form1.Imagen(3), 3
'   If EncuentraPuntos(Form1, Mensajes, Imagen(3), 3) = True Then
'     ' El numero 3 es por el Template que es el numero 3
'     CambiaFoco Identificar(Form1, 3, Form1.NombreVerificar, Form1.AreaVerificar)
'   End If
' End If

End Sub
Private Sub grFinger_SensorPlug(ByVal idSensor As String)
 ' Inicializar la Captura del dispositivo
 'grFinger.CapStartCapture (idSensor)
End Sub
Private Sub grFinger_SensorUnplug(ByVal idSensor As String)
 ' Finalizar la Captura del dispositivo
 'grFinger.CapStopCapture (idSensor)
End Sub


