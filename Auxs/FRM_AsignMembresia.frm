VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_AsignMembresia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de membresía"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   15315
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Asignación de membresías"
      TabPicture(0)   =   "FRM_AsignMembresia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lTitulo(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lTitulo(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "listaCL"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmBoton(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmBoton(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmBoton(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtFecha1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Lista de membresias"
      TabPicture(1)   =   "FRM_AsignMembresia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textBus(4)"
      Tab(1).Control(1)=   "textBus(1)"
      Tab(1).Control(2)=   "textBus(0)"
      Tab(1).Control(3)=   "listaMbrs"
      Tab(1).Control(4)=   "lBus(4)"
      Tab(1).Control(5)=   "lBus(1)"
      Tab(1).Control(6)=   "lBus(0)"
      Tab(1).ControlCount=   7
      Begin VB.TextBox textBus 
         Height          =   285
         Index           =   4
         Left            =   -68160
         TabIndex        =   15
         Text            =   "50"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -71880
         TabIndex        =   12
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtFecha1 
         Height          =   375
         Left            =   12000
         TabIndex        =   9
         Top             =   3240
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
         Format          =   100859905
         CurrentDate     =   40877
      End
      Begin MSFlexGridLib.MSFlexGrid listaMbrs 
         Height          =   8175
         Left            =   -74880
         TabIndex        =   10
         Top             =   1080
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   14420
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         WordWrap        =   -1  'True
         FormatString    =   $"FRM_AsignMembresia.frx":0038
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
         Left            =   12000
         Picture         =   "FRM_AsignMembresia.frx":00F4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5640
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
         Left            =   12000
         Picture         =   "FRM_AsignMembresia.frx":09BE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Asignar"
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
         Left            =   12000
         Picture         =   "FRM_AsignMembresia.frx":1288
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   5415
         Left            =   120
         TabIndex        =   2
         Top             =   3960
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9551
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_AsignMembresia.frx":1B52
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
      Begin MSFlexGridLib.MSFlexGrid listaCL 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_AsignMembresia.frx":1BE5
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
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Núm reg"
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
         Left            =   -68160
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lBus 
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
         Height          =   255
         Index           =   1
         Left            =   -71880
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lBus 
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
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de membresias del cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label lTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de membresias para asignar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   6975
      End
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mn_Cancel 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "FRM_AsignMembresia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Dim activo As Boolean
Dim fechaUltima As Date
Private Sub cmBoton_Click(Index As Integer)
    If Index = 0 Then
        Dim Num As Long
        Num = 0
        For b1 = 1 To listaCL.Rows - 1
            If listaCL.TextMatrix(b1, 7) = "ASIGNANDO" Then
                Num = Num + 1
            End If
        Next b1
        
        If Num = 0 Then
            MsgBox "No se ha asociado ninguna membresía al cliente. Verifique.", vbInformation
        Else
            cargaOper
        End If
    Else
        If Index = 2 Then
            asignarUP
        Else
            Unload Me
        End If
    End If
End Sub
Private Sub asignarUP()
    Dim Num As Long
    Dim ques As String
    Num = 0
    For b1 = 1 To Lista.Rows - 1
        If Lista.TextMatrix(b1, 7) = Chr(254) Then
            Num = Num + 1
        End If
    Next b1
    Dim fecha2 As Date
    If Num > 0 Then
        ques = MsgBox("Se van a asignar " & Num & " membresias al cliente. ¿Continuar?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            cmBoton(2).Enabled = False
            For b1 = 1 To Lista.Rows - 1
                If Lista.TextMatrix(b1, 7) = Chr(254) Then
                    listaCL.Row = 0
                    listaCL.AddItem ""
                    listaCL.Col = 7
                    listaCL.Sort = 1
                    listaCL.TextMatrix(1, 7) = "ASIGNANDO"
                    listaCL.TextMatrix(1, 0) = Lista.TextMatrix(b1, 0)
                    listaCL.TextMatrix(1, 1) = Lista.TextMatrix(b1, 1)
                    If activo = True Then
                        listaCL.TextMatrix(1, 2) = fechaUltima + 1
                    Else
                        listaCL.TextMatrix(1, 2) = Date
                    End If
                    fecha2 = listaCL.TextMatrix(1, 2)
                    fecha2 = fecha2 + Val(Lista.TextMatrix(b1, 4))
                    listaCL.TextMatrix(1, 3) = fecha2
                    listaCL.TextMatrix(1, 4) = Lista.TextMatrix(b1, 4)
                    listaCL.TextMatrix(1, 5) = Lista.TextMatrix(b1, 6)
                    listaCL.TextMatrix(1, 8) = Lista.TextMatrix(b1, 2)
                    
                    
                    listaCL.Row = 1
                    listaCL.Col = 7
                    listaCL.CellFontBold = True
                    listaCL.CellBackColor = vbCyan
                    listaCL.CellForeColor = vbBlack
                End If
            Next b1
        End If
    End If
End Sub
Private Sub cargaOper()
    For b1 = 1 To listaCL.Rows - 1
        If listaCL.TextMatrix(b1, 7) = "ASIGNANDO" Then
            SQL1 = "INSERT INTO MEMBRESIAS " & _
            "(MBR_CTMBID, MBR_PERTP_TIPO_ID, MBR_PERTP_PER_ID, MBR_PERTP_PER_TIPO, MBR_INICIO, MBR_FIN, MBR_STATUS, MBR_FECHA, MBR_VENTAFOLIO) " & _
            "VALUES ('" & listaCL.TextMatrix(b1, 0) & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(0).Caption & "', " & _
            "'" & FrmFocus.lblClieId(2).Caption & "', '" & Format(listaCL.TextMatrix(b1, 2), "yyyy-MM-dd") & "',  " & _
            "'" & Format(listaCL.TextMatrix(b1, 3), "yyyy-MM-dd") & "', 'I', NOW(), '" & FrmFocus.lInfo(1).Caption & "')"
            con.Execute (SQL1)
            
            FrmFocus.Lista.AddItem ""
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 0) = "MEMBRESIA"
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 1) = listaCL.TextMatrix(b1, 0)
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 2) = listaCL.TextMatrix(b1, 1)
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 3) = "1"
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 4) = listaCL.TextMatrix(b1, 8)
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 6) = "M"
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 7) = listaCL.TextMatrix(b1, 0)
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 8) = FrmFocus.lblDatos(1).Caption
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 9) = FrmFocus.lblUserId(1).Caption
            FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Rows - 1, 10) = FrmFocus.lblUserId(0).Caption
            FrmFocus.checkPrecio (FrmFocus.Lista.Rows - 1)
            FrmFocus.addVentDet
        End If
    Next b1

    Unload Me
End Sub
Private Sub dtFecha1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dtFecha1 <= fechaUltima Then
            MsgBox "La fecha inicial no puede ser menor o igual que la última fecha de finalización de la membresía. Verifique.", vbInformation
        Else
            listaCL.TextMatrix(listaCL.Row, listaCL.Col) = dtFecha1
            listaCL.TextMatrix(listaCL.Row, 3) = dtFecha1 + listaCL.TextMatrix(listaCL.Row, 4)
            dtFecha1.Visible = False
        End If
    Else
        If KeyCode = 27 Then
            dtFecha1.Visible = False
        End If
    End If

End Sub

Private Sub dtFecha1_KeyPress(KeyAscii As Integer)
    If KeyCode = 13 Then
        If dtFecha1 <= fechaUltima Then
            MsgBox "La fecha inicial no puede ser menor o igual que la última fecha de finalización de la membresía. Verifique.", vbInformation
        Else
            listaCL.TextMatrix(listaCL.Row, listaCL.Col) = dtFecha1
            listaCL.TextMatrix(listaCL.Row, 3) = dtFecha1 + listaCL.TextMatrix(listaCL.Row, 4)
            dtFecha1.Visible = False
        End If
    Else
        If KeyCode = 27 Then
            dtFecha1.Visible = False
        End If
    End If
End Sub

Private Sub dtFecha1_LostFocus()
    dtFecha1.Visible = False
End Sub

Private Sub Form_Load()
    listaCL.ColWidth(8) = 0
    SSTab1.Tab = 0
    cargaListaClte
    cargaLista
    cargaMbrs
End Sub

Private Sub cargaMbrs()
    listaMbrs.Rows = 1
    SQL1 = "SELECT CLAVE_MEMBRESIA, INICIO, FIN, ADQUIRIO, CLIENTE, MEMBRESIA, DIAS_MEM, STATUS, PER_ID, DIAS_PER " & _
    "FROM VIEW_MEMBRESIAS_ASIGNADAS " & _
    "WHERE upper(CLIENTE) LIKE upper('%" & textBus(1).Text & "%') " & _
    "AND upper(MEMBRESIA) LIKE upper('%" & textBus(0).Text & "%') " & _
    "ORDER BY ADQUIRIO DESC " & _
    "Limit 0, " & Val(textBus(4).Text)
    Set RES1 = con.Execute(SQL1)
    Do While Not RES1.EOF
    
        listaMbrs.AddItem ""
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 0) = RES1.Fields("CLIENTE")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 1) = RES1.Fields("PER_ID")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 2) = RES1.Fields("MEMBRESIA")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 3) = RES1.Fields("DIAS_MEM")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 4) = RES1.Fields("DIAS_PER")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 5) = RES1.Fields("STATUS")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 6) = RES1.Fields("INICIO")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 7) = RES1.Fields("FIN")
        listaMbrs.TextMatrix(listaMbrs.Rows - 1, 8) = RES1.Fields("ADQUIRIO")
        listaMbrs.Row = listaMbrs.Rows - 1
        listaMbrs.Col = 5
        If RES1.Fields("STATUS") = "VIGENTE" Then
            listaMbrs.CellFontBold = True
            listaMbrs.CellBackColor = vbBlue
            listaMbrs.CellForeColor = vbWhite
        Else
            If RES1.Fields("STATUS") = "ANTICIPADA" Then
                listaMbrs.CellFontBold = True
                listaMbrs.CellBackColor = vbGreen
                listaMbrs.CellForeColor = vbWhite
            Else
                If RES1.Fields("STATUS") = "VENCIDO" Then
                    listaMbrs.CellFontBold = True
                    listaMbrs.CellBackColor = vbRed
                    listaMbrs.CellForeColor = vbWhite
                End If
            End If
        End If
        RES1.MoveNext
    Loop
    
End Sub
Private Sub cargaLista()
    SQL1 = "SELECT ID, MEMBRESIA, PRECIO, DIAS_MEMBRESIA, DIAS_PERIODO, PERIODO, TIPO FROM VIEW_MEMBRESIAS "
    Set RES1 = con.Execute(SQL1)
    Lista.Rows = 1
    
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("ID")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("MEMBRESIA")
        Lista.TextMatrix(Lista.Rows - 1, 2) = FormatCurrency(RES1.Fields("PRECIO"))
        Lista.TextMatrix(Lista.Rows - 1, 3) = RES1.Fields("PERIODO")
        Lista.TextMatrix(Lista.Rows - 1, 4) = RES1.Fields("DIAS_PERIODO")
        Lista.TextMatrix(Lista.Rows - 1, 5) = RES1.Fields("TIPO")
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("DIAS_MEMBRESIA")
        
        Lista.Row = Lista.Rows - 1
        Lista.Col = 7
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16
        Lista.TextMatrix(Lista.Rows - 1, 7) = Chr(168)
        
        
        
        RES1.MoveNext
    Loop
End Sub
Private Sub cargaListaClte()
    lInfo.Caption = "Cliente: " & FrmFocus.lblDatos(2).Caption
    SQL1 = "SELECT CLAVE_MEMBRESIA, INICIO, FIN, ADQUIRIO, CLIENTE, MEMBRESIA, DIAS_MEM, DIAS_PER, STATUS " & _
    "FROM VIEW_MEMBRESIAS_asignadas WHERE PER_ID = '" & FrmFocus.lblClieId(0).Caption & "'"
    'MsgBox SQL1
    Set RES1 = con.Execute(SQL1)
    listaCL.Rows = 1
    activo = False
    
    If Not RES1.EOF Then
        fechaUltima = RES1.Fields("FIN")
    End If
    Do While Not RES1.EOF
        listaCL.AddItem ""
        listaCL.TextMatrix(listaCL.Rows - 1, 0) = RES1.Fields("CLAVE_MEMBRESIA")
        listaCL.TextMatrix(listaCL.Rows - 1, 1) = RES1.Fields("MEMBRESIA")
        listaCL.TextMatrix(listaCL.Rows - 1, 2) = RES1.Fields("INICIO")
        listaCL.TextMatrix(listaCL.Rows - 1, 3) = RES1.Fields("FIN")
        listaCL.TextMatrix(listaCL.Rows - 1, 4) = RES1.Fields("DIAS_MEM")
        listaCL.TextMatrix(listaCL.Rows - 1, 5) = RES1.Fields("DIAS_PER")
        listaCL.TextMatrix(listaCL.Rows - 1, 6) = RES1.Fields("ADQUIRIO")
        listaCL.TextMatrix(listaCL.Rows - 1, 7) = RES1.Fields("STATUS")
        
        If RES1.Fields("STATUS") = "VIGENTE" Then
            listaCL.Row = listaCL.Rows - 1
            listaCL.Col = 7
            listaCL.CellFontBold = True
            listaCL.CellBackColor = vbBlue
            listaCL.CellForeColor = vbWhite
        Else
            If RES1.Fields("STATUS") = "ANTICIPADA" Then
                listaCL.Row = listaCL.Rows - 1
                listaCL.Col = 7
                listaCL.CellFontBold = True
                listaCL.CellBackColor = vbGreen
                listaCL.CellForeColor = vbWhite
            Else
                If RES1.Fields("STATUS") = "VENCIDO" Then
                    listaCL.Row = listaCL.Rows - 1
                    listaCL.Col = 7
                    listaCL.CellFontBold = True
                    listaCL.CellBackColor = vbRed
                    listaCL.CellForeColor = vbWhite
                End If
            End If
        End If
        
        activo = True
        RES1.MoveNext
    Loop
End Sub
Private Sub cargaLista2()
'    On Error Resume Next
    Dim valida As Boolean
    valida = False
    
    lInfo.Caption = FrmFocus.lblDatos(2).Caption

    SQL1 = "SELECT T1.ctmb_Id, T1.ctmb_Status, IF(ctmb_Status='A', 'ACTIVO', 'INACTIVO') STATUS, " & _
    "T1.ctmb_Precio, T1.ctmb_Nombre, T1.ctmb_Descripcion, T2.ctpr_Periodo, T2.ctpr_DIAS, IF(CTMB_TIPO='P', 'PERIODICO', 'CONSECUTIVO') TIPO, CTMB_DIAS,  " & _
    "(SELECT MBR_CTMBID FROM MEMBRESIAS T3 WHERE CURDATE() BETWEEN T3.MBR_INICIO AND T3.MBR_FIN AND MBR_STATUS = 'A' " & _
    "AND T3.MBR_PERTP_PER_ID =  '" & FrmFocus.lblClieId(0).Caption & "' AND T3.MBR_PERTP_TIPO_ID = '" & FrmFocus.lblClieId(1).Caption & "') MEMBRESIA,  " & _
   "(SELECT MBR_INICIO FROM MEMBRESIAS T3 WHERE CURDATE() BETWEEN T3.MBR_INICIO AND T3.MBR_FIN AND MBR_STATUS = 'A' " & _
    "AND T3.MBR_PERTP_PER_ID =  '" & FrmFocus.lblClieId(0).Caption & "' AND T3.MBR_PERTP_TIPO_ID = '" & FrmFocus.lblClieId(1).Caption & "') INICIO,  " & _
   "(SELECT MBR_FIN FROM MEMBRESIAS T3 WHERE CURDATE() BETWEEN T3.MBR_INICIO AND T3.MBR_FIN AND MBR_STATUS = 'A' " & _
    "AND T3.MBR_PERTP_PER_ID =  '" & FrmFocus.lblClieId(0).Caption & "' AND T3.MBR_PERTP_TIPO_ID = '" & FrmFocus.lblClieId(1).Caption & "') FIN  " & _
    "FROM CAT_MEMBRESIAS T1, CAT_PERIODO T2 " & _
    "WHERE T1.CTMB_PERIODOID = T2.CTID_PERIODO AND CTMB_STATUS = 'A'"
    Set RES1 = con.Execute(SQL1)
    'MsgBox SQL1
    Lista.Rows = 1
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("CTMB_ID")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("CTMB_NOMBRE")
        Lista.TextMatrix(Lista.Rows - 1, 2) = FormatCurrency(RES1.Fields("CTMB_PRECIO"))
        Lista.TextMatrix(Lista.Rows - 1, 3) = RES1.Fields("CTPR_PERIODO")
        Lista.TextMatrix(Lista.Rows - 1, 4) = RES1.Fields("CTPR_DIAS")
        Lista.TextMatrix(Lista.Rows - 1, 5) = RES1.Fields("TIPO")
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("CTMB_DIAS")
        
        Lista.Row = Lista.Rows - 1
        Lista.Col = 9
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16
        'ListaUsers.TextMatrix(ListaUsers.Rows - 1, 7) = Chr(254)
        If RES1.Fields("MEMBRESIA") = Lista.TextMatrix(Lista.Rows - 1, 0) Then
            valida = True
            Lista.TextMatrix(Lista.Rows - 1, 7) = RES1.Fields("INICIO")
            Lista.TextMatrix(Lista.Rows - 1, 8) = RES1.Fields("FIN")
            Lista.TextMatrix(Lista.Rows - 1, 9) = Chr(254)
            Lista.TopRow = Lista.Rows - 1
            Lista.Enabled = False
            cmBoton(0).Enabled = False
            Exit Do
        End If
        RES1.MoveNext
    Loop
    
    If valida = True Then
        lInfo.Caption = lInfo.Caption & vbCrLf & vbCrLf & "El cliente tiene vigente su membresía."
        lInfo.ForeColor = vbBlack
    Else
        Lista.TextMatrix(Lista.Rows - 1, 9) = Chr(168)
        lInfo.Caption = lInfo.Caption & vbCrLf & vbCrLf & "El cliente no cuenta con membresia. Verifique."
        lInfo.ForeColor = vbRed
    End If
    
End Sub


Private Sub Lista_DblClick()

'Select Case lista.Col
'    Case 7:
            Dim b1 As Long
            b1 = Lista.Row
            Lista.Row = b1
            Lista.Col = 7
            If Lista.TextMatrix(b1, 7) = Chr(168) Then
                Lista.TextMatrix(b1, 7) = Chr(254)
            Else
                Lista.TextMatrix(b1, 7) = Chr(168)
            End If
'    Case 7:
'        'MsgBox "Ok"
'            dtFecha1.Top = lista.CellTop + lista.Top
'            dtFecha1.Left = lista.CellLeft + lista.Left
'            dtFecha1.height = lista.CellHeight
'            dtFecha1.width = lista.CellWidth
'            If lista.TextMatrix(lista.Row, lista.Col) <> "" Then
'                dtFecha1 = lista.TextMatrix(lista.Row, lista.Col)
'            Else
'                dtFecha1 = Date
'            End If
'            dtFecha1.Visible = True
'            dtFecha1.SetFocus
'End Select


End Sub

Private Sub listaCL_DblClick()
Select Case listaCL.Col
    Case 2:
'        'MsgBox "Ok"
            dtFecha1.Top = listaCL.CellTop + listaCL.Top
            dtFecha1.Left = listaCL.CellLeft + listaCL.Left
            dtFecha1.height = listaCL.CellHeight
            dtFecha1.width = listaCL.CellWidth
            If listaCL.TextMatrix(listaCL.Row, listaCL.Col) <> "" Then
                dtFecha1 = listaCL.TextMatrix(listaCL.Row, listaCL.Col)
            Else
                dtFecha1 = Date
            End If
            dtFecha1.Visible = True
            dtFecha1.SetFocus
End Select
    
End Sub

Private Sub listaCL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If listaCL.Rows > 1 Then
        If Button = vbRightButton Then
            If listaCL.TextMatrix(listaCL.Row, 7) = "ASIGNANDO" Then
                mn_Cancel.Enabled = True
                PopupMenu menu, vbPopupMenuLeftAlign
            Else
                mn_Cancel.Enabled = False
                PopupMenu menu, vbPopupMenuLeftAlign
            End If
        End If
    End If

End Sub

Private Sub listaMbrs_DblClick()
    If listaMbrs.MouseRow = 0 Then
        Call ordenarLista(listaMbrs)
    End If
End Sub

Private Sub mn_Cancel_Click()
    Dim ques As String
    ques = MsgBox("¿Cancelar " & listaCL.TextMatrix(listaCL.Row, 1) & "?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        listaCL.RemoveItem (listaCL.Row)
    End If
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaMbrs
    End If
End Sub
