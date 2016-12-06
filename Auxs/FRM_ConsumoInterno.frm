VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_ConsumoInterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo interno"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14820
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCant 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   17
      Top             =   960
      Width           =   615
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
      Left            =   120
      Picture         =   "FRM_ConsumoInterno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
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
      Left            =   1920
      Picture         =   "FRM_ConsumoInterno.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtClave 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Left            =   5040
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtClave 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   16777215
      AllowUserResizing=   1
      FormatString    =   $"FRM_ConsumoInterno.frx":1194
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
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   5
      Left            =   11520
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   4
      Left            =   11520
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   3
      Left            =   11520
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto/Servicio"
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
      Index           =   2
      Left            =   5040
      TabIndex        =   11
      Top             =   240
      Width           =   1695
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
      Left            =   5040
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario consumidor"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   2055
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
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   2175
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
      Left            =   8280
      TabIndex        =   7
      Top             =   240
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
      Left            =   8280
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   10200
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mn_Busqueda 
      Caption         =   "Busqueda"
      Begin VB.Menu mn_BusUsuarios 
         Caption         =   "Búsqueda de usuarios"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mn_BusProductos 
         Caption         =   "Búsqueda de productos"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "FRM_ConsumoInterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset

Public Sub checkProdCI()
    'On Error Resume Next
    SQL1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "PROD_CODIGO = '" & txtClave(0).Text & "' AND PROD_STATUS = 'A'"
    Set RES1 = con.Execute(SQL1)
    Dim b1 As Long
    If Not RES1.EOF Then
        lblDatos(2).Caption = RES1.Fields("PROD_NOMBRE")
      
        
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
        
        addLista
    Else
        MsgBox "Clave no encontrada. Verifique. ", vbInformation
    End If
    
    
End Sub

Private Sub addLista()
    If RES1.Fields("PROD_SERV") = "P" Then
        If RES1.Fields("PROD_CANT") = 0 Then
            MsgBox "No hay productos en existencia. Verifique.", vbInformation
            Exit Sub
        End If
        For b1 = 1 To Lista.Rows - 1
            If Lista.TextMatrix(b1, 1) = RES1.Fields("PROD_CODIGO") Then
                If RES1.Fields("PROD_CANT") > Lista.TextMatrix(b1, 3) Then
                    Lista.TextMatrix(b1, 3) = Lista.TextMatrix(b1, 3) + 1
                    'updateVentDet (b1)
                    checkPrecio (b1)
                    Exit Sub
                Else
                    MsgBox "No hay productos en existencia. Verifique.", vbInformation
                    Exit Sub
                End If
            End If
        Next b1
    Else
        MsgBox "Se quiere asignar un servicio. Verifique. ", vbInformation
    End If
        
    Lista.AddItem ""
    Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("TIPO_PROD")
    Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("PROD_CODIGO")
    Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("PROD_NOMBRE")
    Lista.TextMatrix(Lista.Rows - 1, 3) = "1"
    Lista.TextMatrix(Lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
    Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("PROD_SERV")
    Lista.TextMatrix(Lista.Rows - 1, 7) = RES1.Fields("PROD_ID")
    Lista.TextMatrix(Lista.Rows - 1, 8) = lblDatos(0).Caption
    Lista.TextMatrix(Lista.Rows - 1, 9) = lblUserId(3).Caption
    Lista.TextMatrix(Lista.Rows - 1, 10) = lblUserId(4).Caption
    checkPrecio (Lista.Rows - 1)
    'addVentDet
End Sub
Public Sub checkPrecio(fila As Long)
    Lista.TextMatrix(fila, 5) = Lista.TextMatrix(fila, 3) * Lista.TextMatrix(fila, 4)
    Lista.TextMatrix(fila, 5) = FormatCurrency(Lista.TextMatrix(fila, 5))
    'checkPrecioFinal
End Sub

Private Sub cmBoton_Click(Index As Integer)
    
    Dim ques As String
    If Index = 0 Then
        ques = MsgBox("Va registrar " & Lista.Rows - 1 & " operaciones como consumo interno. ¿Continuar?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            cobrar
        End If
    Else
        ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
        If ques = vbYes Then
           Unload Me
        End If
    End If
End Sub
Private Sub cobrar()
    If Lista.Rows > 1 Then
        For b1 = 1 To Lista.Rows - 1
             With Lista
                SQL1 = "INSERT INTO CONSUMO_iNTERNO (CSI_FECHAHORA, CSI_VEND_PERID, CSI_VEND_PERTIPOID, CSI_VEND_PERTIPO, " & _
                "CSI_USER_PERID, CSI_USER_PERTIPOID, CSI_USER_PERTIPO, CSI_PRODUCTO_ID, CSI_PRODUCTO_SERV, CSI_CANTIDAD, CSI_PRECIO) " & _
                "VALUES (NOW(), '" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', '" & lblUserId(2).Caption & "', " & _
                " '" & .TextMatrix(b1, 9) & "', '" & .TextMatrix(b1, 10) & "', 'U', '" & .TextMatrix(b1, 7) & "', " & _
                "'" & .TextMatrix(b1, 6) & "', '" & .TextMatrix(b1, 3) & "', '" & Val(Format(.TextMatrix(b1, 4), "General Number")) & "')"
                con.Execute (SQL1)
                'MsgBox SQL1
                                
                SQL1 = "UPDATE PRODUCTOS SET PROD_CANT = PROD_CANT - (" & Val(.TextMatrix(b1, 3)) & ") " & _
                "WHERE PROD_ID = '" & .TextMatrix(b1, 7) & "' AND PROD_SERV = '" & .TextMatrix(b1, 6) & "'"
                'MsgBox SQL1
                con.Execute (SQL1)
            End With
        Next b1

        MsgBox "Consumo interno efectivo. " & vbCrLf & Lista.Rows - 1 & " registro(s) realizado(s).", vbInformation
        Lista.Rows = 1
        Unload Me
    Else
        MsgBox "No se puede realizar la operación. Verifique.", vbInformation
    End If
End Sub
Private Sub Form_Load()
    cargaDatos
End Sub
Private Sub cargaDatos()
    txtCant.Visible = False
    lblDatos(1).Caption = FRM_Menu.menuBarra2.Panels(5).Text
    Lista.ColWidth(6) = 0
    Lista.ColWidth(7) = 0
    Lista.ColWidth(9) = 0
    Lista.ColWidth(10) = 0
    Lista.Rows = 1
    For b1 = 0 To 5
        lblUserId(b1).Caption = ""
    Next b1
    
    Call cargaFotoMostrador("M", 2)
    
End Sub

Public Sub cargaFotoMostrador(tipo As String, Num As Integer)
    Dim idPer As String
    
    If tipo = "M" Then
        idPer = FRM_Menu.menuBarra2.Panels(7).Text
    Else
        If tipo = "U" Then
            idPer = lblUserId(3).Caption
        End If
    End If
    
    SQL1 = "SELECT PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_TIPO_ID, PERTP_PER_TIPO, PER_ID, PER_FOTO " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' " & _
    "AND T1.PER_ID = '" & idPer & "'"
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
        userId = txtClave(1).Text
            
        
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
            imgFoto(Num).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            imgFoto(Num).Picture = LoadPicture("")
        End If
    End If
        'txtClave(1).SetFocus
End Sub

Private Sub Lista_DblClick()
Select Case Lista.Col
    Case 3:
        ''''Para la cantidad
            'ListaOper.Row = fila
            'ListaOper.Col = columna
                txtCant.Top = Lista.CellTop + Lista.Top
                txtCant.Left = Lista.CellLeft + Lista.Left
                txtCant.height = Lista.CellHeight
                txtCant.width = Lista.CellWidth
                txtCant.Text = Lista.TextMatrix(Lista.Row, Lista.Col)
                txtCant.Visible = True
                txtCant.SelStart = 0
                txtCant.SelLength = Len(txtCant.Text)
                txtCant.SetFocus

    End Select

End Sub

Private Sub mn_BusProductos_Click()
    If lblUserId(3).Caption <> "" Then
        modBusqueda = "ConsumoInterno"
        BUSQ_ProdSer.Show vbModal
    Else
        MsgBox "Debe seleccioonar un usuario primerio. Verifique.", vbInformation
    End If
End Sub

Private Sub mn_BusUsuarios_Click()
        tipoBusqueda = "U"
        modBusqueda = "ConsumoInterno"
        BUSQ_Usuarios.Caption = "Búsqueda de usuarios."
        BUSQ_Usuarios.Show vbModal

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    NumerosPunto (txtCant.Text)
    If KeyAscii = 13 Then
        If Val(txtCant.Text) <> Val(Lista.TextMatrix(Lista.Row, 3)) And Val(txtCant.Text) > 0 Then
            If Lista.TextMatrix(Lista.Row, 6) = "P" Then
                SQL1 = "SELECT PROD_CANT FROM PRODUCTOS WHERE PROD_CODIGO = '" & Lista.TextMatrix(Lista.Row, 1) & "'"
                Set RES1 = con.Execute(SQL1)
                If Not RES1.EOF Then
                    If RES1.Fields("PROD_CANT") >= Val(txtCant.Text) Then
                        Lista.TextMatrix(Lista.Row, 3) = txtCant.Text
                        'updateVentDet (lista.Row)
                        checkPrecio (Lista.Row)
                        txtCant.Text = ""
                        txtCant.Visible = False
                        Exit Sub
                    Else
                        MsgBox "La cantidad supera los productos en existencia. Verifique.", vbInformation
                        txtCant.SelStart = 0
                        txtCant.SelLength = Len(txtCant.Text)
                        txtCant.SetFocus
                        Exit Sub
                    End If
                End If
'            Else
'                If lista.TextMatrix(lista.Row, 6) = "S" Then
'                    lista.TextMatrix(lista.Row, 3) = txtCant.Text
'                    updateVentDet (lista.Row)
'                    checkPrecio (lista.Row)
'                    txtCant.Text = ""
'                    txtCant.Visible = False
'                    Exit Sub
'                Else
'                    MsgBox "Operación no permitida. Verifique.", vbInformation
'                End If
            End If
        Else
            txtCant.Text = ""
            txtCant.Visible = False
            Exit Sub
        End If
    Else
        If KeyAscii = 27 Then
            txtCant.Text = ""
            txtCant.Visible = False
        End If
    End If

End Sub

Private Sub txtCant_LostFocus()
    txtCant.Text = ""
    txtCant.Visible = False
End Sub

Private Sub txtClave_KeyPress(Index As Integer, KeyAscii As Integer)
    checkProdCI
End Sub
