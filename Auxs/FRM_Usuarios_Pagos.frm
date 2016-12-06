VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Usuarios_Pagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de pagos a usuarios"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   14565
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista general"
      TabPicture(0)   =   "FRM_Usuarios_Pagos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lista1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos generales asignación/edición"
      TabPicture(1)   =   "FRM_Usuarios_Pagos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbStatus"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lInfo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lista2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbInfo"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmBoton(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmBoton(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmBoton(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "dtFecha1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Left            =   10800
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   104267777
         CurrentDate     =   40877
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modificar"
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
         Left            =   6240
         Picture         =   "FRM_Usuarios_Pagos.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
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
         Left            =   240
         Picture         =   "FRM_Usuarios_Pagos.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7200
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
         Left            =   2040
         Picture         =   "FRM_Usuarios_Pagos.frx":11CC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7200
         Width           =   1695
      End
      Begin VB.ComboBox cmbInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   5775
      End
      Begin MSFlexGridLib.MSFlexGrid lista2 
         Height          =   4215
         Left            =   240
         TabIndex        =   2
         Top             =   2760
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Usuarios_Pagos.frx":1A96
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
      Begin MSFlexGridLib.MSFlexGrid lista1 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   12938
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Usuarios_Pagos.frx":1B4D
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
      Begin VB.Label lInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   7335
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
         Left            =   4080
         TabIndex        =   7
         Top             =   7680
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo pago"
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   4455
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Editar 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "FRM_Usuarios_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Private Sub cmbInfo_Click()
    sql1 = "SELECT CTPG_ID, CTPG_NOMBRE, IF(CTPG_TIPOPAGO='C', 'COMISIONES', 'HONORARIOS/PAGOS FIJOS') TIPO_PAGO, CTPG_IDPERIODO, " & _
    "CTPG_VALOR, IF(CTPG_TIPOVALOR='E', 'EFECTIVO', 'PORCENTAJE') TIPO_VALOR, CTPR_PERIODO, CONCAT(CTPR_PERIODO, ' DIAS: ', CTPR_DIAS) PERIODO, " & _
    "CTPG_APLICAVALORES, IF(CTPG_APLICATIPO=NULL, 'NO', IF(CTPG_APLICATIPO='P', 'PRODUCTOS', 'SERVICIOS')) APLICA_TIPO, " & _
    "IF(CTPG_APLICASUBTIPO IS NULL, 'NO', (IF(CTPG_APLICASUBTIPO='G', 'GENERAL', 'SUBTIPOS'))) APLICA_SUBTIPO, PG_STATUS, PG_FECHA_INI, " & _
    "PER_ID, PER_NOMBRE, PER_PATERNO, PER_MATERNO, T6.CTPT_TIPO, PERTP_USUARIO, if(PERTP_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PERTP_TIPO_ID, PERTP_PER_TIPO  " & _
    "FROM CAT_PAGOS T1, CAT_PERIODO T2, COMISIONES T3, PER_TIPO T4, PERSONA T5, CAT_TIPO T6 " & _
    "WHERE CTPG_IDPERIODO = CTID_PERIODO AND CTPG_ID = '" & cmbInfo.ItemData(cmbInfo.ListIndex) & "' AND " & _
    "T1.CTPG_ID = T3.PG_CTPG_ID AND T3.PG_PERTP_TIPO_ID = T4.PERTP_TIPO_ID AND T3.PG_PERTP_PER_ID = T4.PERTP_PER_ID AND " & _
    "T3.PG_PERTP_PER_TIPO = T4.PERTP_PER_TIPO AND T4.PERTP_PER_ID = T5.PER_ID AND T4.PERTP_TIPO_ID = T6.CTPT_ID AND T4.PERTP_PER_TIPO = T6.CTPT_SUBTIPO "
    Set RES2 = con.Execute(sql1)
    lista2.Rows = 1
    If Not RES2.EOF Then
        lInfo.Caption = "Tipo de pago: " & RES2.Fields("TIPO_PAGO") & vbCrLf & "Periodo: " & RES2.Fields("CTPR_PERIODO") & _
        vbCrLf & "Valor: " & RES2.Fields("CTPG_VALOR") & " Tipo valor: " & RES2.Fields("TIPO_VALOR")
    End If
    Do While Not RES2.EOF
        lista2.AddItem ""
        lista2.TextMatrix(lista2.Rows - 1, 0) = RES2.Fields("PER_ID")
        lista2.TextMatrix(lista2.Rows - 1, 1) = RES2.Fields("PER_NOMBRE")
        lista2.TextMatrix(lista2.Rows - 1, 2) = RES2.Fields("PER_PATERNO")
        lista2.TextMatrix(lista2.Rows - 1, 3) = RES2.Fields("PER_MATERNO")
        lista2.TextMatrix(lista2.Rows - 1, 4) = RES2.Fields("CTPT_TIPO")
        lista2.TextMatrix(lista2.Rows - 1, 5) = RES2.Fields("PERTP_USUARIO")
        lista2.TextMatrix(lista2.Rows - 1, 6) = RES2.Fields("STATUS")
        lista2.TextMatrix(lista2.Rows - 1, 8) = RES2.Fields("PERTP_TIPO_ID")
        lista2.TextMatrix(lista2.Rows - 1, 9) = RES2.Fields("PERTP_PER_TIPO")
        lista2.TextMatrix(lista2.Rows - 1, 10) = RES2.Fields("PG_FECHA_INI") & ""
        
        lista2.Row = lista2.Rows - 1
        lista2.Col = 7
        lista2.CellFontName = "Wingdings"
        lista2.CellFontBold = True
        lista2.CellFontSize = 16
        'ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = Chr(254)
        If RES2.Fields("PG_STATUS") = "A" Then
            lista2.TextMatrix(lista2.Rows - 1, 7) = Chr(254)
        Else
            lista2.TextMatrix(lista2.Rows - 1, 7) = Chr(168)
        End If
        RES2.MoveNext
    Loop
End Sub

Private Sub cmBoton_Click(Index As Integer)
    Dim num As Long
    Dim ques As String
    
    Select Case Index
        Case 0:
            num = 0
            For b1 = 1 To lista2.Rows - 1
                If lista2.TextMatrix(b1, 7) = Chr(254) Then
                    num = num + 1
                End If
            Next b1
                ques = MsgBox("¿Asociar " & num & " registros al tipo de pago: " & cmbInfo.Text & "?", vbYesNo + vbQuestion)
                If ques = vbYes Then
                    asignarRegistro
                End If
        Case 2:
            If cmbInfo.Text <> "" Then
                lbStatus.Caption = "Estatus: Modificando pago"
                cmbInfo.Enabled = False
                cmBoton(2).Enabled = False
                
                
            Else
                lbStatus.Caption = "Estatus: Asignando pago"
            End If
        Case 1:
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cargaDatos
            End If
    End Select
End Sub
Private Sub asignarRegistro()
    Dim b1 As Long
    If lbStatus.Caption = "Estatus: Modificando pago" Then
        With lista2
            For b1 = 1 To .Rows - 1
                If .TextMatrix(b1, 7) = Chr(254) Then
                    'MsgBox "254"
                    If .TextMatrix(b1, 10) <> "" Then
                        sql1 = "UPDATE COMISIONES SET PG_STATUS = 'A', PG_FECHA_INI = '" & Format(.TextMatrix(b1, 10), "yyyy-MM-dd") & "' WHERE PG_PERTP_TIPO_ID = '" & .TextMatrix(b1, 8) & "' AND " & _
                        "PG_PERTP_PER_ID = '" & .TextMatrix(b1, 0) & "' AND  PG_PERTP_PER_tIPO = '" & .TextMatrix(b1, 9) & "' AND PG_cTPG_ID = '" & cmbInfo.ItemData(cmbInfo.ListIndex) & "'"
                        con.Execute (sql1)
                    Else
                        MsgBox "Debe establecer una fecha de inicio.", vbInformation
                    End If
                Else
                    If .TextMatrix(b1, 7) = Chr(168) Then
                        'MsgBox "168"
                        sql1 = "UPDATE COMISIONES SET PG_STATUS = 'I' WHERE PG_PERTP_TIPO_ID = '" & .TextMatrix(b1, 8) & "' AND " & _
                        "PG_PERTP_PER_ID = '" & .TextMatrix(b1, 0) & "' AND  PG_PERTP_PER_tIPO = '" & .TextMatrix(b1, 9) & "' AND PG_cTPG_ID = '" & cmbInfo.ItemData(cmbInfo.ListIndex) & "'"
                        'MsgBox SQL1
                        con.Execute (sql1)
                    End If
                End If
            Next b1
        End With
        MsgBox "Información guardada.", vbInformation
        cargaDatos
    Else
    End If
End Sub

Private Sub dtFecha1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        lista2.TextMatrix(lista2.Row, lista2.Col) = dtFecha1
        dtFecha1.Visible = False
    End If

End Sub

Private Sub dtFecha1_KeyPress(KeyAscii As Integer)
    If keyascci = 13 Then
        lista2.TextMatrix(lista2.Row, lista2.Col) = dtFecha1
        dtFecha1.Visible = False
    End If
End Sub

Private Sub dtFecha1_LostFocus()
    dtFecha1.Visible = False
End Sub

Private Sub Form_Load()
    lista2.ColWidth(8) = 0
    lista2.ColWidth(9) = 0
    cargaDatos
End Sub
Private Sub cargaDatos()
    
    lbStatus.Caption = "Estatus: Asignando pago"
    cmbInfo.Enabled = True
    cmBoton(2).Enabled = True
    lista2.Rows = 1
    lInfo.Caption = "Tipo de pago: " & "" & vbCrLf & "Periodo: " & "" & _
    vbCrLf & "Valor: " & "" & " Tipo valor: " & ""
    dtFecha1.Visible = False
    SSTab1.Tab = 0
    sql1 = "SELECT CTPG_ID, CTPG_NOMBRE FROM CAT_PAGOS"
    Set RES1 = con.Execute(sql1)
    cmbInfo.Clear
    Do While Not RES1.EOF
        cmbInfo.AddItem RES1.Fields("CTPG_NOMBRE")
        cmbInfo.ItemData(cmbInfo.ListCount - 1) = RES1.Fields("CTPG_ID")
        RES1.MoveNext
    Loop
    cargaLista1
End Sub
Private Sub cargaLista1()
    sql1 = "SELECT CTPG_ID, CTPG_NOMBRE, IF(CTPG_TIPOPAGO='C', 'COMISIONES', 'HONORARIOS/PAGOS FIJOS') TIPO_PAGO, CTPG_IDPERIODO, " & _
    "CTPG_VALOR, IF(CTPG_TIPOVALOR='E', 'EFECTIVO', 'PORCENTAJE') TIPO_VALOR, CTPR_PERIODO, CONCAT(CTPR_PERIODO, ' DIAS: ', CTPR_DIAS) PERIODO, " & _
    "CTPG_APLICAVALORES, IF(CTPG_APLICATIPO=NULL, 'NO', IF(CTPG_APLICATIPO='P', 'PRODUCTOS', 'SERVICIOS')) APLICA_TIPO, " & _
    "IF(CTPG_APLICASUBTIPO IS NULL, 'NO', (IF(CTPG_APLICASUBTIPO='G', 'GENERAL', 'SUBTIPOS'))) APLICA_SUBTIPO, PG_STATUS, PG_FECHA_INI, " & _
    "PER_ID, PER_NOMBRE, PER_PATERNO, PER_MATERNO, T6.CTPT_TIPO, PERTP_USUARIO, if(PERTP_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PERTP_TIPO_ID, PERTP_PER_TIPO  " & _
    "FROM CAT_PAGOS T1, CAT_PERIODO T2, COMISIONES T3, PER_TIPO T4, PERSONA T5, CAT_TIPO T6 " & _
    "WHERE CTPG_IDPERIODO = CTID_PERIODO AND " & _
    "T1.CTPG_ID = T3.PG_CTPG_ID AND T3.PG_PERTP_TIPO_ID = T4.PERTP_TIPO_ID AND T3.PG_PERTP_PER_ID = T4.PERTP_PER_ID AND PG_STATUS = 'A' AND " & _
    "T3.PG_PERTP_PER_TIPO = T4.PERTP_PER_TIPO AND T4.PERTP_PER_ID = T5.PER_ID AND T4.PERTP_TIPO_ID = T6.CTPT_ID AND T4.PERTP_PER_TIPO = T6.CTPT_SUBTIPO ORDER BY PERTP_USUARIO"
    Set RES1 = con.Execute(sql1)
    lista1.Rows = 1
    Do While Not RES1.EOF
        lista1.AddItem ""
        lista1.TextMatrix(lista1.Rows - 1, 0) = RES1.Fields("PER_ID")
        lista1.TextMatrix(lista1.Rows - 1, 1) = RES1.Fields("PERTP_USUARIO") & ""
        lista1.TextMatrix(lista1.Rows - 1, 2) = RES1.Fields("PER_NOMBRE")
        lista1.TextMatrix(lista1.Rows - 1, 3) = RES1.Fields("PER_PATERNO")
        lista1.TextMatrix(lista1.Rows - 1, 4) = RES1.Fields("PER_MATERNO")
        lista1.TextMatrix(lista1.Rows - 1, 5) = RES1.Fields("CTPT_TIPO")
        lista1.TextMatrix(lista1.Rows - 1, 6) = RES1.Fields("CTPG_NOMBRE")
        
        lista1.TextMatrix(lista1.Rows - 1, 7) = RES1.Fields("TIPO_PAGO")
        lista1.TextMatrix(lista1.Rows - 1, 8) = RES1.Fields("CTPG_VALOR")
        lista1.TextMatrix(lista1.Rows - 1, 9) = RES1.Fields("TIPO_VALOR")
        lista1.TextMatrix(lista1.Rows - 1, 10) = RES1.Fields("PERIODO")
        lista1.TextMatrix(lista1.Rows - 1, 11) = RES1.Fields("CTPG_APLICAVALORES")
        
        
        RES1.MoveNext
    Loop
End Sub
Private Sub cargaLista2()
    lista2.Rows = 1
    sql1 = "SELECT PER_ID, PER_PATERNO, PER_MATERNO, PER_NOMBRE, PER_FEC_NAC, if(PERTP_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, " & _
    "(YEAR(CURDATE()) - YEAR(PER_FEC_NAC)) EDAD, PERTP_TIPO_ID, CTPT_TIPO, PERTP_USUARIO, PERTP_PER_TIPO " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3, COMISIONES T4 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_TIPO_ID = T3.CTPT_ID  AND T2.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
    "AND PERTP_PER_TIPO = 'U'  AND T4."
    Set RES1 = con.Execute(sql1)
    
    Do While Not RES1.EOF
        lista2.AddItem ""
        
        
        
        lista2.TextMatrix(lista2.Rows - 1, 0) = RES1.Fields("PER_ID")
        lista2.TextMatrix(lista2.Rows - 1, 1) = RES1.Fields("PER_NOMBRE")
        lista2.TextMatrix(lista2.Rows - 1, 2) = RES1.Fields("PER_PATERNO")
        lista2.TextMatrix(lista2.Rows - 1, 3) = RES1.Fields("PER_MATERNO")
        lista2.TextMatrix(lista2.Rows - 1, 4) = RES1.Fields("CTPT_TIPO")
        lista2.TextMatrix(lista2.Rows - 1, 5) = RES1.Fields("PERTP_USUARIO")
        lista2.TextMatrix(lista2.Rows - 1, 6) = RES1.Fields("STATUS")
        lista2.TextMatrix(lista2.Rows - 1, 8) = RES1.Fields("PERTP_TIPO_ID")
        lista2.TextMatrix(lista2.Rows - 1, 9) = RES1.Fields("PERTP_PER_TIPO")
        
        lista2.Row = lista2.Rows - 1
        lista2.Col = 7
        lista2.CellFontName = "Wingdings"
        lista2.CellFontBold = True
        lista2.CellFontSize = 16
        'ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = Chr(254)
        lista2.TextMatrix(lista2.Rows - 1, 7) = Chr(168)
        
        
        RES1.MoveNext
    Loop

End Sub
Private Sub lista2_DblClick()
    
Select Case lista2.Col
    Case 7:
        If cmbInfo.Enabled = False And lbStatus.Caption = "Estatus: Modificando pago" Then
            Dim b1 As Long
            b1 = lista2.Row
            lista2.Row = b1
            lista2.Col = 7
            If lista2.TextMatrix(b1, 7) = Chr(168) Then
                If lista2.TextMatrix(b1, 10) <> "" Then
                    lista2.TextMatrix(b1, 7) = Chr(254)
                Else
                    MsgBox "Debe establecer una fecha de inicio. " & vbCrLf & vbCrLf & _
                    "(Haga doble clic sobre la casilla fecha para establecer alguna)", vbInformation
                End If
            Else
                lista2.TextMatrix(b1, 7) = Chr(168)
            End If
        End If
    Case 10:
        If cmbInfo.Enabled = False And lbStatus.Caption = "Estatus: Modificando pago" Then
            dtFecha1.Top = lista2.CellTop + lista2.Top
            dtFecha1.Left = lista2.CellLeft + lista2.Left
            dtFecha1.height = lista2.CellHeight
            dtFecha1.width = lista2.CellWidth
            If lista2.TextMatrix(lista2.Row, lista2.Col) <> "" Then
                dtFecha1 = lista2.TextMatrix(lista2.Row, lista2.Col)
            Else
                dtFecha1 = Date
            End If
            dtFecha1.Visible = True
            dtFecha1.SetFocus
        End If
End Select

    
End Sub
