VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CAT_Etiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   14850
   StartUpPosition =   1  'CenterOwner
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
      Left            =   12120
      Picture         =   "CAT_Etiquetas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   240
      Width           =   1695
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
      Height          =   855
      Index           =   0
      Left            =   10320
      Picture         =   "CAT_Etiquetas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   240
      Width           =   1695
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Nombre sucursal"
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
      Index           =   5
      Left            =   10560
      TabIndex        =   26
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Razón social"
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
      Index           =   4
      Left            =   10560
      TabIndex        =   25
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Logo"
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
      Index           =   3
      Left            =   10560
      TabIndex        =   24
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Precio"
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
      Index           =   2
      Left            =   7080
      TabIndex        =   23
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Código"
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
      Index           =   1
      Left            =   7080
      TabIndex        =   22
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Nombre"
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
      Index           =   0
      Left            =   7080
      TabIndex        =   19
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   7
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   6
      Left            =   5040
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   5
      Left            =   3480
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   4
      Left            =   5040
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Index           =   3
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
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
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   $"CAT_Etiquetas.frx":1194
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
   Begin VB.Label lbStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus:"
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
      Left            =   9720
      TabIndex        =   29
      Top             =   7320
      Width           =   4695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   7080
      X2              =   13800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Campos a imprimir"
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
      Index           =   10
      Left            =   7200
      TabIndex        =   21
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   6960
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Métricas"
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
      Index           =   9
      Left            =   360
      TabIndex        =   20
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Y 2"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor X 2"
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
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "No etiquetas horizontales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5040
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "No etiquetas verticales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Alto"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Largo"
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
      Index           =   3
      Left            =   3480
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Y Inicial"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor X Inicial"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "CAT_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Dim idEtiqueta


Private Sub cmBoton_Click(Index As Integer)
    If Index = 0 Then
        If lbStatus.Caption = "Estatus: Agregando" Then
            agregar
        Else
            If lbStatus.Caption = "Estatus: Editando" Then
                editar
            End If
        End If
    Else
        If Index = 1 Then
        Dim ques As String
        ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                limpiaCampos
                lbStatus.Caption = "Estatus: Agregando"
            End If
        End If
    End If
End Sub
Private Sub editar()
    For b1 = 0 To 8
        If txtInfo(b1).Text = "" Then
            MsgBox "Información incompleta. Verifique.", vbInformation
            Exit Sub
        End If
    Next b1
        
    Dim NUM As Long
    NUM = 0
    
    For b1 = 0 To 5
        If chkInfo(b1).Value = Checked Then
            NUM = NUM + 1
        End If
    Next b1
    If NUM = 0 Then
        MsgBox "Por lo menos un campo a imprimir debe de estar seleccionado. Verifique.", vbInformation
        Exit Sub
    End If
    
    SQL1 = "UPDATE CAT_ETIQUETAS SET ETQ_NOMBRE = '" & txtInfo(0).Text & "', " & _
    "ETQ_VALORX = '" & txtInfo(1).Text & "',  " & _
    "ETQ_VALORY = '" & txtInfo(2).Text & "',  " & _
    "ETQ_LARGO = '" & txtInfo(3).Text & "',  " & _
    "ETQ_ALTO = '" & txtInfo(4).Text & "',  " & _
    "ETQ_VERTCL = '" & txtInfo(5).Text & "',  " & _
    "ETQ_HORZT = '" & txtInfo(6).Text & "',  " & _
    "ETQ_VALORX2 = '" & txtInfo(7).Text & "',  " & _
    "ETQ_VALORY2 = '" & txtInfo(8).Text & "',  " & _
    "ETQ_NOMPROD = '" & chkInfo(0).Value & "',  " & _
    "ETQ_CODIGO = '" & chkInfo(1).Value & "',  " & _
    "ETQ_PRECIO = '" & chkInfo(2).Value & "',  " & _
    "ETQ_LOGO = '" & chkInfo(3).Value & "',  " & _
    "ETQ_RAZONSOCIAL = '" & chkInfo(4).Value & "',  " & _
    "ETQ_SUCURSAL = '" & chkInfo(5).Value & "'  " & _
    "WHERE IDETIQUETA = '" & idEtiqueta & "'"
    con.Execute (SQL1)
    
    MsgBox "Información guardada.", vbInformation
    limpiaCampos
    cargaLista
    lbStatus.Caption = "Estatus: Agregando"

End Sub
Private Sub agregar()
    For b1 = 0 To 8
        If txtInfo(b1).Text = "" Then
            MsgBox "Información incompleta. Verifique.", vbInformation
            Exit Sub
        End If
    Next b1
        
    Dim NUM As Long
    NUM = 0
    
    For b1 = 0 To 5
        If chkInfo(b1).Value = Checked Then
            NUM = NUM + 1
        End If
    Next b1
    If NUM = 0 Then
        MsgBox "Por lo menos un campo a imprimir debe de estar seleccionado. Verifique.", vbInformation
        Exit Sub
    End If
    
    SQL1 = "INSERT INTO CAT_ETIQUETAS (ETQ_NOMBRE, ETQ_VALORX, ETQ_VALORY, ETQ_LARGO, ETQ_ALTO, ETQ_VERTCL, " & _
    "ETQ_HORZT, ETQ_VALORX2, ETQ_VALORY2, ETQ_NOMPROD, ETQ_CODIGO, ETQ_PRECIO, ETQ_LOGO, ETQ_RAZONSOCIAL, " & _
    "ETQ_SUCURSAL) VALUES ('" & txtInfo(0).Text & "', '" & txtInfo(1).Text & "', '" & txtInfo(2).Text & "', " & _
    "'" & txtInfo(3).Text & "', '" & txtInfo(4).Text & "', '" & txtInfo(5).Text & "', '" & txtInfo(6).Text & "', " & _
    "'" & txtInfo(7).Text & "', '" & txtInfo(8).Text & "', '" & chkInfo(0).Value & "', '" & chkInfo(1).Value & "', " & _
    "'" & chkInfo(2).Value & "', '" & chkInfo(3).Value & "', '" & chkInfo(4).Value & "', '" & chkInfo(5).Value & "')"
    con.Execute (SQL1)
        
    MsgBox "Información guardada.", vbInformation
    limpiaCampos
    cargaLista
    lbStatus.Caption = "Estatus: Agregando"
    

End Sub
Private Sub Form_Load()
    cargaLista
    limpiaCampos
    lbStatus.Caption = "Estatus: Agregando"
    
End Sub
Private Sub limpiaCampos()
    For b1 = 0 To 8
        txtInfo(b1).Text = ""
    Next b1
    For b1 = 0 To 5
        chkInfo(b1).Value = Unchecked
    Next b1
End Sub
Private Sub cargaLista()
    SQL1 = "SELECT ETQ_NOMBRE, IDETIQUETA, ETQ_VALORX, ETQ_VALORY, ETQ_LARGO, ETQ_ALTO, ETQ_VERTCL, " & _
    "ETQ_HORZT, ETQ_VALORX2, ETQ_VALORY2, if(ETQ_NOMPROD=1, 'SI', 'NO') ETQ_NOMPROD, " & _
    "IF(ETQ_CODIGO=1, 'SI', 'NO') ETQ_CODIGO, IF(ETQ_PRECIO=1, 'SI', 'NO') ETQ_PRECIO, " & _
    "IF(ETQ_LOGO=1, 'SI', 'NO') ETQ_LOGO, IF(ETQ_RAZONSOCIAL=1, 'SI', 'NO') ETQ_RAZONSOCIAL, " & _
    "IF(ETQ_SUCURSAL=1, 'SI', 'NO') ETQ_SUCURSAL FROM CAT_ETIQUETAS"
    Set RES1 = con.Execute(SQL1)
    lista.Rows = 1
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("IdEtiqueta")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("Etq_Nombre")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("Etq_Codigo")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("Etq_NomProd")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("Etq_Precio")
        lista.TextMatrix(lista.Rows - 1, 5) = RES1.Fields("Etq_Logo")
        lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("Etq_RazonSocial")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("Etq_Sucursal")
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("Etq_valorX")
        lista.TextMatrix(lista.Rows - 1, 9) = RES1.Fields("Etq_valory")
        lista.TextMatrix(lista.Rows - 1, 10) = RES1.Fields("Etq_largo")
        lista.TextMatrix(lista.Rows - 1, 11) = RES1.Fields("Etq_alto")
        lista.TextMatrix(lista.Rows - 1, 12) = RES1.Fields("Etq_vertcl")
        lista.TextMatrix(lista.Rows - 1, 13) = RES1.Fields("Etq_horzt")
        lista.TextMatrix(lista.Rows - 1, 14) = RES1.Fields("Etq_valorX2")
        lista.TextMatrix(lista.Rows - 1, 15) = RES1.Fields("Etq_valorY2")
        
        RES1.MoveNext
    Loop

End Sub

Private Sub lista_Click()
'ASDSAD
End Sub

Private Sub lista_DblClick()
    idEtiqueta = lista.TextMatrix(lista.Row, 0)
    SQL1 = "SELECT ETQ_NOMBRE, IDETIQUETA, ETQ_VALORX, ETQ_VALORY, ETQ_LARGO, ETQ_ALTO, ETQ_VERTCL, " & _
    "ETQ_HORZT, ETQ_VALORX2, ETQ_VALORY2, ETQ_NOMPROD, " & _
    "ETQ_CODIGO, ETQ_PRECIO, " & _
    "ETQ_LOGO, ETQ_RAZONSOCIAL, " & _
    "ETQ_SUCURSAL FROM CAT_ETIQUETAS WHERE IDETIQUETA = '" & idEtiqueta & "'"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        txtInfo(0).Text = RES1.Fields("ETQ_NOMBRE")
        txtInfo(1).Text = RES1.Fields("ETQ_VALORX")
        txtInfo(2).Text = RES1.Fields("ETQ_VALORY")
        txtInfo(3).Text = RES1.Fields("ETQ_LARGO")
        txtInfo(4).Text = RES1.Fields("ETQ_ALTO")
        txtInfo(5).Text = RES1.Fields("ETQ_VERTCL")
        txtInfo(6).Text = RES1.Fields("ETQ_HORZT")
        txtInfo(7).Text = RES1.Fields("ETQ_VALORX2")
        txtInfo(8).Text = RES1.Fields("ETQ_VALORY2")
        chkInfo(0).Value = RES1.Fields("ETQ_NOMPROD")
        chkInfo(1).Value = RES1.Fields("ETQ_CODIGO")
        chkInfo(2).Value = RES1.Fields("ETQ_PRECIO")
        chkInfo(3).Value = RES1.Fields("ETQ_LOGO")
        chkInfo(4).Value = RES1.Fields("ETQ_RAZONSOCIAL")
        chkInfo(5).Value = RES1.Fields("ETQ_SUCURSAL")

        
        lbStatus.Caption = "Estatus: Editando"
        MsgBox "Editando", vbInformation
    End If
End Sub
