VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_Monederos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monederos"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   15615
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   17595
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FRM_Monederos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lBus(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Borde(16)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "listMonederos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "time_size"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAccion(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textBus(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar detalle"
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
         Index           =   7
         Left            =   12840
         Picture         =   "FRM_Monederos.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Timer time_size 
         Interval        =   500
         Left            =   11880
         Top             =   840
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   8655
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   15266
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         AllowUserResizing=   1
         FormatString    =   "Cliente                                        | Puntos           | PerId "
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
      Begin MSFlexGridLib.MSFlexGrid listMonederos 
         Height          =   8415
         Left            =   4920
         TabIndex        =   5
         Top             =   1080
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   14843
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   9520683
         ForeColorFixed  =   16777215
         BackColorBkg    =   15329769
         GridColor       =   16711680
         AllowUserResizing=   1
         FormatString    =   $"FRM_Monederos.frx":05A6
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
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   16
         Left            =   240
         Top             =   480
         Width           =   3045
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Menu mn_global 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu reset_gral 
         Caption         =   "Resetear monedero"
      End
   End
   Begin VB.Menu mn_indiv 
      Caption         =   "Monedero individual"
      Visible         =   0   'False
      Begin VB.Menu reset_individual 
         Caption         =   "Resetar monedero individual"
      End
   End
End
Attribute VB_Name = "FRM_Monederos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RES1 As Recordset
Dim SQL1 As String

Private Sub cmbStatus_Click(Index As Integer)
    cargaLista
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(listMonederos)
    End If

End Sub

Private Sub Form_Load()
    Lista.ColWidth(2) = 0
    SSTab1.TabCaption(0) = " "
    cargaLista
End Sub

Private Sub Lista_Click()
    cargaLista_Persona
End Sub
Private Sub cargaLista_Persona()
    SQL1 = "SELECT * fROM VIEW_PUNTOS_ADMIN " & _
    "where PER_ID = '" & Lista.TextMatrix(Lista.Row, 2) & "'"
    Set RES1 = con.Execute(SQL1)
    
    listMonederos.Rows = 1
    
    listMonederos.Redraw = False
    Do While Not RES1.EOF
        listMonederos.AddItem ""
        listMonederos.TextMatrix(listMonederos.Rows - 1, 0) = RES1.Fields("TIPO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 1) = RES1.Fields("CLIENTE")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 2) = RES1.Fields("ORIGEN")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 3) = FormatCurrency(RES1.Fields("MONEDERO"))
        listMonederos.TextMatrix(listMonederos.Rows - 1, 4) = RES1.Fields("FECHAHORA")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 5) = RES1.Fields("USUARIO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 6) = RES1.Fields("FOLIO")
        listMonederos.TextMatrix(listMonederos.Rows - 1, 7) = RES1.Fields("CLAVE")
        RES1.MoveNext
    Loop

    listMonederos.Redraw = True
End Sub
Private Sub Lista_DblClick()

    Call ordenarLista(Lista)

End Sub

Private Sub Lista_GotFocus()
    ConScroll Lista

End Sub

Private Sub Lista_LostFocus()
    SinScroll Lista
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Lista.Rows > 1 Then
        If Button = vbRightButton Then
            PopupMenu mn_global, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub reset_gral_Click()
Dim a As String

a = MsgBox("Resetear los puntos a " & Lista.TextMatrix(Lista.Row, 0) & "." & vbCrLf & vbCrLf & _
"El valor de monedero quedará en cero." & vbCrLf & vbCrLf & _
"¿Continuar?", vbYesNo + vbQuestion)

If a = vbYes Then
    SQL1 = "UPDATE MONEDERO SET MND_PUNTOS = 0 WHERE MND_CLIEPERID = " & Lista.TextMatrix(Lista.Row, 2) & ""
    con.Execute (SQL1)
    
    MsgBox "Puntos reseteados en cero. Verifique."
    cargaLista_Persona
End If


End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        listMonederos.Rows = 1
        cargaLista
    End If
End Sub

Private Sub time_size_Timer()
    time_size.Enabled = False
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    'Lista.width = Me.width - 500
    listMonederos.width = Me.width - 5500
End Sub


Private Sub cargaLista()
    Dim texto1 As String
    texto1 = ""
        
    SQL1 = "SELECT * fROM VIEW_PUNTOS_TOTAL WHERE " & _
    "UPPER(CLIENTE) LIKE UPPER('%" & textBus(0).Text & "%') "
    Set RES1 = con.Execute(SQL1)
    
    Lista.Redraw = False
    Lista.Rows = 1
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("CLIENTE")
        Lista.TextMatrix(Lista.Rows - 1, 1) = Round(RES1.Fields("MONEDERO"), 2)
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("PER_ID")
        RES1.MoveNext
    Loop
    
    Lista.Redraw = True
End Sub


