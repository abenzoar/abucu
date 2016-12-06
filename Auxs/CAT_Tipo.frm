VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CAT_Tipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de tipo"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10590
   Icon            =   "CAT_Tipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10590
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox iFoto2 
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   8040
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   -5000
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cMd1 
      Left            =   9840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editar"
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
      Index           =   4
      Left            =   9480
      Picture         =   "CAT_Tipo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
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
      Index           =   3
      Left            =   8280
      Picture         =   "CAT_Tipo.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar imagen"
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
      Left            =   8280
      Picture         =   "CAT_Tipo.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eliminar imagen"
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
      Index           =   2
      Left            =   9480
      Picture         =   "CAT_Tipo.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtProd 
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
      MaxLength       =   65
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtProd 
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
      Left            =   120
      MaxLength       =   65
      TabIndex        =   1
      Top             =   1560
      Width           =   7935
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
      Height          =   735
      Index           =   0
      Left            =   8280
      Picture         =   "CAT_Tipo.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid listCatalogo 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Clave   | Tipo                                | Imagen | Descripción                                           "
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
   Begin VB.PictureBox iFoto 
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   2355
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   2475
      Index           =   12
      Left            =   8280
      Top             =   4320
      Width           =   2205
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   120
      Top             =   1560
      Width           =   7965
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   120
      Top             =   720
      Width           =   3405
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo *"
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
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lProd 
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos en lista:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label lInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   8415
      Index           =   1
      Left            =   -240
      Picture         =   "CAT_Tipo.frx":247C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
   Begin VB.Menu mn_Ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "CAT_Tipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SQL1 As String
    Dim RES1 As Recordset
    Dim tipoId As Long

Private Sub cmBoton_Click(Index As Integer)
    Select Case Index
        Case 0:
            If txtProd(0).Text <> "" Then
                If lInfo(0).Caption = "Agregar" Then
                    guardarMarca
                Else
                    editaMarca
                End If
            Else
                MsgBox "Se ha detectado un error. Por favor verifique.", vbExclamation
            End If
        Case 4:
            txtProd(0).Text = listCatalogo.TextMatrix(listCatalogo.Row, 1)
            txtProd(1).Text = listCatalogo.TextMatrix(listCatalogo.Row, 2)
            tipoId = listCatalogo.TextMatrix(listCatalogo.Row, 0)
            lInfo(0).Caption = "Editar"
            cmBoton(0).Enabled = True
            cmBoton(1).Enabled = True
            cmBoton(2).Enabled = True
        
        Case 3:
            txtProd(0).Text = ""
            txtProd(1).Text = ""
            lInfo(0).Caption = "Agregar"
            cmBoton(0).Enabled = True
            cmBoton(1).Enabled = False
            cmBoton(2).Enabled = False
        Case 1:
            buscarImagen
        Case 2:
            eliminarImagen (listCatalogo.TextMatrix(listCatalogo.Row, 0))
        End Select
        
        
End Sub
Private Sub eliminarImagen(ctTipoId As String)
    If iFoto.Picture <> 0 Then
    
        SQL1 = "UDPATE CAT_TIPO SET CTPT_FOTO = NULL WHERE CTPT_ID = '" & ctTipoId & "'"
        con.Execute (SQL1)
        
        iFoto.Picture = LoadPicture("")
    End If
    
End Sub
Private Sub buscarImagen()
    cMd1.DialogTitle = "Buscando imagen..."
    cMd1.Filter = "Archivos de Imagenes|*.jpg*||*.bmp*||*.gif*||*.wmf*||*.emf*|"
    cMd1.FileName = ""
    cMd1.ShowOpen
    If cMd1.FileName <> "" Then
        mostrarImagen
    End If
End Sub
Private Sub mostrarImagen()
    With cMd1
        iFoto2.Picture = LoadPicture(.FileName)
        iFoto.AutoRedraw = True
        iFoto.PaintPicture iFoto2.Picture, _
            iFoto.ScaleLeft, iFoto.ScaleTop, _
                iFoto.ScaleWidth, iFoto.ScaleHeight, _
            iFoto2.ScaleLeft, iFoto2.ScaleTop, _
                iFoto2.ScaleWidth, iFoto2.ScaleHeight
        iFoto.Picture = iFoto.Image
        
    End With
End Sub

Private Sub guardarImagen(ctTipoId As String)
    Dim res As ADODB.Recordset
    Set res = New ADODB.Recordset
    Dim Imagen1 As ADODB.Stream
    Set Imagen1 = New ADODB.Stream
    'Para la fotoi
    If iFoto.Picture <> 0 Then
        checarCarpetaTemp
        SavePicture iFoto.Picture, (direccionSistema & "\Temp\TempProd.dat")
        'If Not RES1.EOF Then
            res.Open "SELECT * FROM CAT_TIPO WHERE CTPT_ID = '" & ctTipoId & "'", con, adOpenStatic, adLockOptimistic
            'MsgBox "-" & prodId & "-"
            If Not res.EOF Then
                '''NO DEBE
            'Else
                Imagen1.Type = adTypeBinary
                Imagen1.Open
                Imagen1.LoadFromFile (direccionSistema & "\Temp\TempProd.dat")
                res.Fields("ctpt_foto") = Imagen1.Read
                res.Update
            Else
                ''''
            End If
        'End If
    End If
    
End Sub
Private Sub editaMarca()
    SQL1 = "UPDATE CAT_TIPO SET CTPT_TIPO = '" & txtProd(0).Text & "', " & _
    "CTPT_DESCRIPCION = '" & txtProd(1).Text & "' " & _
    "WHERE CTPT_ID = '" & tipoId & "'"
    con.Execute (SQL1)
    
    guardarImagen (tipoId)
    
    MsgBox "Información guardada.", vbInformation
    iFoto.Picture = LoadPicture("")
    txtProd(1).Text = ""
    lInfo(0).Caption = "Agregar"
    cargaLista
    checkProducto
    cmBoton(0).Enabled = False
    cmBoton(1).Enabled = False
    cmBoton(2).Enabled = False
End Sub
Private Sub guardarMarca()
    SQL1 = "INSERT INTO CAT_TIPO (CTPT_TIPO, CTPT_DESCRIPCION, CTPT_SUBTIPO) VALUES " & _
    "('" & txtProd(0).Text & "', '" & txtProd(1).Text & "', '" & tipoCatTipo & "')"
    con.Execute (SQL1)
    
    SQL1 = "select last_insert_id() prodid"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        tipoId = RES1.Fields("prodid")
    End If
    
    guardarImagen (tipoId)
    
    MsgBox "Información guardada.", vbInformation
    iFoto.Picture = LoadPicture("")
    txtProd(1).Text = ""
    lInfo(0).Caption = "Agregar"
    cargaLista
    checkProducto
    cmBoton(0).Enabled = False
    cmBoton(1).Enabled = False
    cmBoton(2).Enabled = False

End Sub
Private Sub checkProducto()
    If tipoCatTipo = "P" Then
        FRM_Productos.cmdTipo_Click
    Else
        If tipoCatTipo = "U" Then
            FRM_Usuarios.cmdTipoUsuario_Click
        Else
            If tipoCatTipo = "S" Then
                FRM_Servicios.cmdTipo_Click
            Else
                If tipoCatTipo = "C" Then
                    FRM_Clientes.cmdTipoUsuario_Click
                Else
                    If tipoCatTipo = "G" Then
                        FRM_Gastos.cargaTipoGasto
                    End If
                End If
            End If
        End If
    End If

End Sub
Private Sub Form_Load()
    cargaLista
    lInfo(0).Caption = "-"
    cmBoton(0).Enabled = False
    cmBoton(1).Enabled = False
    cmBoton(2).Enabled = False
End Sub
Private Sub cargaLista()
    SQL1 = "SELECT CTPT_ID, CTPT_TIPO, CTPT_DESCRIPCION, if(isnull(ctpt_Foto),'NO','SI') FOTO_SN FROM CAT_TIPO " & _
    "WHERE CTPT_TIPO LIKE '%" & txtProd(0).Text & "%' AND CTPT_SUBTIPO = '" & tipoCatTipo & "'"
    Set RES1 = con.Execute(SQL1)
    
    listCatalogo.Rows = 1
    
    listCatalogo.Redraw = False
    Do While Not RES1.EOF
        listCatalogo.AddItem ""
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 0) = RES1.Fields("CTPT_ID")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 1) = RES1.Fields("CTPT_TIPO")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 2) = RES1.Fields("foto_sn")
        
        If IsNull(RES1.Fields("CTPT_DESCRIPCION")) Then
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 3) = ""
        Else
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 3) = RES1.Fields("CTPT_DESCRIPCION")
        End If
        RES1.MoveNext
    Loop
    listCatalogo.Redraw = True

    lInfo(10).Caption = "Tipos en lista: " & listCatalogo.Rows - 1
End Sub

Private Sub listCatalogo_Click()
    SQL1 = "SELECT ctpt_foto FROM cat_tipo WHERE ctpt_id = '" & listCatalogo.TextMatrix(listCatalogo.Row, 0) & "'"
    Set RES1 = con.Execute(SQL1)
    
    If IsNull(RES1.Fields("ctpt_fOTO")) = False Then
        Dim Imagen1 As Stream
        Set Imagen1 = New Stream
        Imagen1.Type = adTypeBinary
        checarCarpetaTemp
        Imagen1.Open
        Imagen1.Write RES1.Fields("ctpt_FOTO")
        Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
        Imagen1.Close
        iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
    Else
        iFoto.Picture = LoadPicture("")
    End If
End Sub

Private Sub listCatalogo_DblClick()
'    txtProd(0).Text = listCatalogo.TextMatrix(listCatalogo.Row, 1)
'    txtProd(1).Text = listCatalogo.TextMatrix(listCatalogo.Row, 2)
'    tipoId = listCatalogo.TextMatrix(listCatalogo.Row, 0)
'    lInfo(0).Caption = "Editar"
End Sub

Private Sub listCatalogo_GotFocus()
    ConScroll listCatalogo
End Sub

Private Sub listCatalogo_LostFocus()
    SinScroll listCatalogo
End Sub

Private Sub mn_Ayuda_Click()
    MsgBox "Para agregar un tipo escriba los valores en los cuadros de texto y de clic en aceptar. " & vbCrLf & vbCrLf & _
    "Para editar un tipo de doble clic sobre el tipo que desea editar en la lista, los valores se mostrarán en los cuadros de texto y podrá cambiarlos, al concluir de clic en aceptar." & vbCrLf & vbCrLf & _
    "En la parte inferior derecha se muestra una leyenda para verificar si esta agregando o editando.", vbInformation
End Sub

Private Sub txtProd_Change(Index As Integer)
    If Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    End If
End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 And Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub


