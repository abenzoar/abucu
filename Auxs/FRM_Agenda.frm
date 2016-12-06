VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Agenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17865
   Icon            =   "FRM_Agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   17865
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer timerAgendaCobro 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12960
      Top             =   0
   End
   Begin VB.Timer TTime 
      Interval        =   250
      Left            =   11160
      Top             =   7920
   End
   Begin VB.CheckBox chkHora 
      Caption         =   "Movimiento de filas a la hora actual"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   8
      Top             =   240
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
   Begin VB.Timer tTiempo 
      Interval        =   60000
      Left            =   12360
      Top             =   120
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   103874561
      CurrentDate     =   40956
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10440
      Top             =   -5000
   End
   Begin VB.CommandButton cmdCitas 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   -5000
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   14160
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   -5000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   13320
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   -5000
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid listaDia 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   12303
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Atendiendo"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   15
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   5520
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagado"
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
      Index           =   2
      Left            =   4320
      TabIndex        =   14
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   3840
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Retrasado"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   1920
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Agendado"
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
      Left            =   600
      TabIndex        =   12
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   120
      X2              =   10080
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informaicón de citas"
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
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00076CF5&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lTest 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "lInfo"
      Height          =   255
      Left            =   12600
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lHora 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   12600
      Stretch         =   -1  'True
      Top             =   -5000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mn_Cita 
      Caption         =   "MenuCita"
      Visible         =   0   'False
      Begin VB.Menu mn_VerInfo 
         Caption         =   "Ver información rápida"
      End
      Begin VB.Menu mn_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Cobrar 
         Caption         =   "Cobrar"
      End
      Begin VB.Menu mn_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar cita"
      End
      Begin VB.Menu mn_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_CancelCita 
         Caption         =   "Cancelar cita"
      End
   End
End
Attribute VB_Name = "FRM_Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim SQL2 As String
Dim RES2 As Recordset
Private Const anchoImagen As Long = 64
    ' variable para la clase tooltip
'Dim tTip As clss_ToolTip
'Public clavesCitas(30, 30)
Dim textoCitas(100, 100)
Dim textoCitaCorto(100, 100)
Dim folioCobro(100, 100)
Dim horas_tiempo(100, 100)
Dim CitaServId(100, 100)
Dim CitaClieId(100, 100)
Dim CitaUserId(100, 100)
Dim fila, columna, fila1, col1 As Long
Dim Color As String
Dim filaDestino As Long
Dim colDestino As Long
Private Sub cargaCitas()
    Dim b1, c1 As Long
    listaDia.MergeCells = flexMergeRestrictColumns
    listaDia.Redraw = False
    For b1 = 4 To listaDia.Rows - 1
        For c1 = 1 To listaDia.Cols - 1
'            MsgBox Format(dtFecha1, "yyyy-MM-dd") & " USUARIO " & listaDia.TextMatrix(2, c1) & " " & listaDia.TextMatrix(b1, 0) & ">= HORA_INI AND " & listaDia.TextMatrix(b1, 0) & " < HORA_FIN "
            sql1 = "SELECT * FROM VIEW_CITAS WHERE STATUS_AGENDA <> 'C' AND date_format(FECHA_INICIO, '%Y-%m-%d') = '" & Format(dtFecha1, "yyyy-MM-dd") & "' AND USUARIO_ID = '" & listaDia.TextMatrix(2, c1) & "' AND '" & Format(listaDia.TextMatrix(b1, 0), "Short Time") & "' >= HORA_INI AND '" & Format(listaDia.TextMatrix(b1, 0), "Short Time") & "' < HORA_FIN  "
            'MsgBox SQL1
            Set RES1 = con.Execute(sql1)
            If Not RES1.EOF Then
                
                Dim HORA1, HORA2 As Long
                Dim HORAS As Double
                Dim HORAS2 As Double
                
                HORA1 = (Val(Left(RES1.Fields("HORA_FIN"), 2)) * 60) + Val(Right(RES1.Fields("HORA_FIN"), 2))
                HORA2 = (Val(Left(RES1.Fields("HORA_INI"), 2)) * 60) + Val(Right(RES1.Fields("HORA_INI"), 2))
                
                'MsgBox HORA1 & "  " & HORA2
                
                HORAS = Round(((HORA1 - HORA2) / 60), 2)
                HORAS2 = ((HORA1 - HORA2) / 30)
                    
            
                listaDia.Row = b1
                listaDia.Col = c1
                listaDia.CellFontBold = True
                listaDia.CellFontSize = 9
                If RES1.Fields("STATUS_AGENDA") = "A" Then
                    listaDia.CellBackColor = &H76CF5
                Else
                    If RES1.Fields("STATUS_AGENDA") = "P" Then
                       listaDia.CellBackColor = &HC000&
                    Else
                        If RES1.Fields("STATUS_AGENDA") = "T" Then
                            listaDia.CellBackColor = &HFF8080
                        End If
                    End If
                End If
                listaDia.CellForeColor = vbWhite
                listaDia.TextMatrix(b1, c1) = RES1.Fields("CLIENTE") & vbCrLf & vbCrLf & RES1.Fields("SERVICIO")
                listaDia.MergeCol(c1) = True
                clavesCitas(b1, c1) = RES1.Fields("CLAVE")
                folioCobro(b1, c1) = RES1.Fields("folio")
                textoCitas(b1, c1) = "Cliente: " & RES1.Fields("Cliente") & vbCrLf & "Servicio: " & RES1.Fields("Servicio") & _
                vbCrLf & "Tipo de servicio: " & RES1.Fields("Tipo_servicio") & _
                vbCrLf & "Cita: " & RES1.Fields("Fecha") & "  " & RES1.Fields("Hora_ini") & " a " & RES1.Fields("Hora_Fin") & vbCrLf & _
                vbCrLf & "Usuario: " & RES1.Fields("Usuario") & _
                vbCrLf & "Observaciones: " & RES1.Fields("Observaciones") & _
                vbCrLf & "Cita generada: " & RES1.Fields("GENERADA") & "  " & " por " & RES1.Fields("AGENDO") & vbCrLf & vbCrLf & _
                "Tiempo estimado: " & HORAS
                textoCitaCorto(b1, c1) = "Cliente: " & RES1.Fields("Cliente") & vbCrLf & "Servicio: " & RES1.Fields("Servicio") & _
                vbCrLf & "Tipo de servicio: " & RES1.Fields("Tipo_servicio") & _
                vbCrLf & "Cita: " & RES1.Fields("Fecha") & "  " & RES1.Fields("Hora_ini") & " a " & RES1.Fields("Hora_Fin")
                horas_tiempo(b1, c1) = HORAS2
                CitaServId(b1, c1) = RES1.Fields("SERV_ID")
                CitaClieId(b1, c1) = RES1.Fields("CLIE_ID")
                CitaUserId(b1, c1) = RES1.Fields("USUARIO_PERTPID")
            Else
                listaDia.TextMatrix(b1, c1) = ""
                listaDia.Row = b1
                listaDia.Col = c1
                listaDia.CellBackColor = vbWhite
            
            End If
        Next c1
    Next b1
    listaDia.Redraw = True

End Sub
Private Sub checkColumnas()
    listaDia.RowHeight(2) = 0
    listaDia.RowHeight(3) = 500
    
    For b1 = 1 To listaDia.Cols - 1
        If listaDia.TextMatrix(1, b1) = "" Then
            listaDia.ColWidth(b1) = 0
        Else
            listaDia.ColWidth(b1) = 2500
        End If
    Next b1


End Sub

Private Sub cmbTipo_Click()
    ocultarRol
    cargaImagen
End Sub
Private Sub ocultarRol()
   Dim b1 As Long
   
    If cmbTipo.Text = "(Todos)" Then
   
        For b1 = 1 To listaDia.Cols - 1
            listaDia.ColWidth(b1) = 2500
        Next b1
    Else
        For b1 = 1 To listaDia.Cols - 1
            If listaDia.TextMatrix(3, b1) <> cmbTipo.Text Then
                listaDia.ColWidth(b1) = 0
            Else
                listaDia.ColWidth(b1) = 2500
            End If
        Next b1
    End If

End Sub
Public Sub cmdCitas_Click()
    formatoAgenda
    cargaHorario
    cargaUsuarios
    checkColumnas
    cargaCitas
    checkHoraActual
    'cargaImagen

End Sub

Private Sub dtFecha1_Change()
    lFecha.Caption = Format(dtFecha1, "Long Date")
    cargaCitas
    checkHoraActual
    cargaImagen

End Sub

Private Sub dtFecha1_Click()
    lFecha.Caption = Format(dtFecha1, "Long Date")
    cargaCitas
    checkHoraActual
    cargaImagen
End Sub

Private Sub Form_Load()
'    formatoAgenda
'    cargaUsuarios
    cargaGral

End Sub
Private Sub cargaGral()
    dtFecha1 = Date
    lFecha.Caption = Format(dtFecha1, "Long Date")
    lHora.Caption = Format(Time, "Short Time")
    cargaTipo
End Sub
Private Sub cargaTipo()
    
    cmbTipo.Clear
    cmbTipo.AddItem "(Todos)"
    cmbTipo.ItemData(cmbTipo.ListCount - 1) = 0

    SQL2 = "SELECT CTPT_TIPO, CTPT_ID FROM CAT_TIPO WHERE CTPT_SUBTIPO = 'U'"
    Set RES2 = con.Execute(SQL2)
    
    Do While Not RES2.EOF
        cmbTipo.AddItem RES2.Fields("CTPT_TIPO")
        cmbTipo.ItemData(cmbTipo.ListCount - 1) = RES2.Fields("ctpt_id")
        RES2.MoveNext
    Loop
    
End Sub
Private Sub cargaHorario()
    Dim media As String
    Dim tiempo As Long
    Dim tiempo2 As String
    Dim hora As String
    
    sql1 = "SELECT SUC_HORAENTRADA ENTRADA, SUC_HORASALIDA SALIDA FROM SUCURSAL WHERE SUC_LOCAL = 'S'"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        'calendario.Value = Date
        listaDia.Rows = 4
        'listaDia.MousePointer = flexCustom
        media = "00"
        listaDia.RowHeight(listaDia.Rows - 1) = 730
        If IsNull(RES1.Fields("ENTRADA")) Then
            MsgBox "No se ha establecido un horario para la agenda, por favor verifique. ", vbInformation
            Exit Sub
        End If
        
        listaDia.Redraw = False
        tiempo = DateDiff("n", Format(RES1.Fields("Entrada"), "Short Time"), Format(RES1.Fields("Salida"), "Short Time"))
        tiempo = Val(tiempo) / 60
        For b1 = 0 To tiempo
            hora = Hour(Format(RES1.Fields("Entrada"), "Short Time")) + b1
            tiempo2 = Format(hora, "00") & ":00"
            listaDia.AddItem ""
            listaDia.TextMatrix(listaDia.Rows - 1, 0) = tiempo2
            listaDia.RowHeight(listaDia.Rows - 1) = 850
            listaDia.Col = 0
            listaDia.Row = listaDia.Rows - 1
            listaDia.CellBackColor = &HE0E0E0
            listaDia.CellForeColor = vbBlack
            listaDia.CellFontSize = 14
            listaDia.CellFontBold = True
            tiempo2 = Format(hora, "00") & ":30"
            listaDia.AddItem ""
            listaDia.TextMatrix(listaDia.Rows - 1, 0) = tiempo2
            listaDia.RowHeight(listaDia.Rows - 1) = 850
            listaDia.Col = 0
            listaDia.Row = listaDia.Rows - 1
            listaDia.CellBackColor = &HE0E0E0
            listaDia.CellForeColor = vbBlack
            listaDia.CellFontSize = 14
            listaDia.CellFontBold = True
        Next b1
            
    Else
        MsgBox "Debe de establecer un horario para la sucursal en la cual está laborando. ", vbInformation
    End If
    listaDia.Redraw = True
     
    listaDia.Col = 0
    listaDia.Row = 3
    listaDia.CellFontSize = 8
    listaDia.CellFontBold = True
    listaDia.TextMatrix(3, 0) = "Hora / Tipo"
    listaDia.Row = 0
    listaDia.CellFontSize = 8
    listaDia.CellFontBold = True
    listaDia.TextMatrix(0, 0) = "Imagen"
    listaDia.Row = 1
    listaDia.CellFontSize = 8
    listaDia.CellFontBold = True
    listaDia.TextMatrix(1, 0) = "Nombre"
    listaDia.RowHeight(0) = 0
    
End Sub
Private Sub cargaImagen()
On Error Resume Next
    Dim Imagen1 As Stream
    Dim b1 As Long

    For b1 = 1 To listaDia.Cols - 1
    
        sql1 = "SELECT CONCAT(T1.PER_NOMBRE, ' ', T1.PER_PATERNO, ' ', T1.PER_MATERNO) USUARIO, PERTP_USUARIO, PER_FOTO, CONCAT(T2.PERTP_PER_ID, T2.PERTP_PER_TIPO) IDPER  " & _
        "FROM PERSONA T1, PER_TIPO T2 " & _
        "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_ID = '" & listaDia.TextMatrix(2, b1) & "' AND T2.PERTP_PER_TIPO = 'U'"
        Set RES1 = con.Execute(sql1)
        
        If Not RES1.EOF Then
        Set Imagen1 = New Stream
        Imagen1.Type = adTypeBinary
        listaDia.Row = 0
        listaDia.Col = b1
        
            If IsNull(RES1.Fields("PER_fOTO")) = False Then
                checarCarpetaTemp
                Imagen1.Open
                Imagen1.Write RES1.Fields("PER_FOTO")
                Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
                Imagen1.Close
                Picture1.AutoSize = True
                Picture1.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
                Picture2.AutoSize = False
                Picture2.width = Picture2.width
                Picture2.height = Picture2.height
                Picture2.PaintPicture Picture1.Image, 0, 0, Picture2.width, Picture2.height
                Set listaDia.CellPicture = Picture2.Image
                'MsgBox "OK"
            Else
                Picture1.AutoSize = True
                Picture1.Picture = LoadPicture(direccionSistema & "\Temp\usuarioDefault.jpg")
                Picture2.AutoSize = False
                Picture2.width = Picture2.width
                Picture2.height = Picture2.height
                Picture2.PaintPicture Picture1.Image, 0, 0, Picture2.width, Picture2.height
                Set listaDia.CellPicture = Picture2.Image
                'MsgBox "Ok"
            End If
            
        End If
    Next b1

End Sub
Private Sub cargaUsuarios()
    Dim numUsuarios As Long
        
    sql1 = "SELECT COUNT(*) NUM " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T2.PERTP_AGENDA = '1' AND T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' AND T2.PERTP_TIPO_ID = T3.ctpt_Id AND T2.PERTP_PER_TIPO = T3.ctpt_SubTipo "
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        numUsuarios = RES1.Fields("NUM")
    Else
        numUsuarios = 0
    End If
    
    listaDia.Redraw = False
    listaDia.Cols = numUsuarios + 1
    
    sql1 = "SELECT CONCAT(T1.PER_NOMBRE, ' ', T1.PER_PATERNO, ' ', T1.PER_MATERNO) USUARIO, " & _
    "T2.PERTP_PER_ID IDPER, T3.CTPT_TIPO, T2.PERTP_TIPO_ID " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T2.PERTP_AGENDA = '1' AND T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'U' AND T2.PERTP_TIPO_ID = T3.ctpt_Id AND T2.PERTP_PER_TIPO = T3.ctpt_SubTipo "
    Set RES1 = con.Execute(sql1)
                
    Dim b1 As Long
    b1 = 0
    Do While Not RES1.EOF
        b1 = b1 + 1
                                
        '''Nombre del usuario
        listaDia.TextMatrix(0, b1) = RES1.Fields("PERTP_TIPO_ID")
        listaDia.TextMatrix(1, b1) = RES1.Fields("USUARIO")
        listaDia.TextMatrix(2, b1) = RES1.Fields("IDPER")
        listaDia.TextMatrix(3, b1) = RES1.Fields("CTPT_TIPO")
        listaDia.Row = 0
        listaDia.Col = b1
        listaDia.CellFontSize = 1
        listaDia.Row = 1
        listaDia.Col = b1
        listaDia.CellFontSize = 12
        listaDia.CellFontBold = True
        listaDia.CellBackColor = &H8000000D
        listaDia.CellForeColor = vbWhite
        listaDia.CellAlignment = 1
        listaDia.Row = 3
        listaDia.Col = b1
        listaDia.CellFontSize = 11
        listaDia.CellFontBold = True
        listaDia.CellBackColor = &H8000000D
        listaDia.CellForeColor = vbWhite
        listaDia.CellAlignment = 1
        listaDia.Row = 0
        listaDia.Col = b1
        'listaDia.CellBackColor = &H8000000D
        listaDia.CellAlignment = vbCenter
        RES1.MoveNext
    Loop
    listaDia.Redraw = True

End Sub

Private Sub formatoAgenda()
    With listaDia
        .Rows = 5
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .RowHeightMin = (anchoImagen * Screen.TwipsPerPixelX) + (Screen.TwipsPerPixelX * 2)
        .FixedCols = 1
        .FixedRows = 4
        .Row = 0
        .RowHeight(0) = 800
        .WordWrap = True
        .width = Me.width - 300
        .height = Me.height - 1500
        .RowHeight(2) = 0
        
    End With

'       ' Crear una nueva instancia de la clase
'       Set tTip = New clss_ToolTip
'       ' Establece El tipo ( balloon o normal )
'       tTip.Estilo = TTBalloon
'       ' Indica el icono a utilizar ( info, Warning , error etc..)
'       tTip.Icono = TTIconInfo
'       tTip.Delay = 50 ' Tiempo de duración

End Sub

Private Sub listaDia_Click()
'''''asdsad
    tTiempo.Enabled = False
    
End Sub

Private Sub listaDia_DblClick()
     
    If listaDia.TextMatrix(listaDia.Row, listaDia.Col) = "" Then
        tipoCita = "Creacion"
    Else
        tipoCita = "Edicion"
    End If
     'FRM_AgendaCita.Show vbModal
    FRM_AgendaCita2.Show vbModal
End Sub

Private Sub listaDia_DragDrop(Source As Control, X As Single, Y As Single)

    Dim ques As String
    
    lTest.Caption = listaDia.TextMatrix(fila, columna)
    filaDestino = listaDia.MouseRow
    colDestino = listaDia.MouseCol
    
    If listaDia.MouseRow > 2 And listaDia.MouseCol > 0 Then
        If listaDia.TextMatrix(filaDestino, colDestino) = "" Then
            For b1 = 1 To Val(horas_tiempo(fila, columna))
                listaDia.Row = filaDestino + b1 - 1
                listaDia.Col = colDestino
                listaDia.TextMatrix(filaDestino + b1 - 1, colDestino) = lTest.Caption
                listaDia.CellFontName = lInfo.Font
                listaDia.CellBackColor = Color
                listaDia.CellFontBold = True
                listaDia.CellFontSize = 9
                listaDia.CellBackColor = &H76CF5
                listaDia.CellForeColor = vbWhite
            Next b1
            listaDia.MergeCol(colDestino) = True
            
            lTest.Caption = horas_tiempo(fila, columna)
        
                ques = MsgBox(textoCitaCorto(fila, columna) & vbCrLf & vbCrLf & "¿Copiar la cita?" & vbCrLf & vbCrLf & "Al seleccionar si se crea un duplicado de la cita origen en el destino seleccionado dejando la cita original. " & vbCrLf & vbCrLf & "Al seleccionar no solo se movera la cita origen al destino seleccionado.", vbYesNoCancel + vbQuestion)
                If ques = vbYes Then
                    copiarCita
                Else
                    If ques = vbNo Then
                        moverCita
                    End If
                End If
            cargaImagen
        End If
    End If

End Sub
Private Sub copiarCita()

Dim tipoCopia As String

'    MsgBox "columna origen: " & columna & " columna destino: " & colDestino
'
    If columna = colDestino Then
        tipoCopia = "C"
    Else
        If columna <> colDestino Then
            tipoCopia = "G"
        End If
    End If
    
            sql1 = "INSERT INTO AGENDA_SERVICIOS (agds_agdId, agds_ServId, agds_SerTipo, agds_Inicio, agds_Fin, agds_Status, " & _
            "agds_Usuario_Id, agds_Usuario_PerId, agds_Usuario_PerTipo, agds_ServPrecio, agds_FechaHora, agds_Tipo) VALUES ( " & _
            "'" & clavesCitas(fila, columna) & "', '" & CitaServId(fila, columna) & "', 'S', '" & Format(dtFecha1, "yyyy-MM-dd") & " " & Format(listaDia.TextMatrix(filaDestino, 0), "hh:mm:ss") & "',  " & _
            "'" & Format(dtFecha1, "yyyy-MM-dd") & " " & Format(listaDia.TextMatrix(filaDestino + 1, 0), "hh:mm:ss") & "', 'A', '" & listaDia.TextMatrix(0, colDestino) & "', '" & listaDia.TextMatrix(2, colDestino) & "', 'U', '0.0', now(), '" & tipoCopia & "')"
           ' MsgBox SQL1
            con.Execute (sql1)
    
    cmdCitas_Click
    
    MsgBox "Cita copiada. Verifique.", vbInformation
    
End Sub
Private Sub moverCita()

    

    sql1 = "UPDATE AGENDA_SERVICIOS SET AGDS_INICIO = '" & Format(dtFecha1, "yyyy-MM-dd") & " " & Format(listaDia.TextMatrix(filaDestino, 0), "hh:mm:ss") & "',   " & _
    "AGDS_FIN = '" & Format(dtFecha1, "yyyy-MM-dd") & " " & Format(listaDia.TextMatrix(filaDestino + 1, 0), "hh:mm:ss") & "', " & _
    "AGDS_USUARIO_ID = '" & listaDia.TextMatrix(0, colDestino) & "', agds_Usuario_PerId = '" & listaDia.TextMatrix(2, colDestino) & "', agds_FechaHora = NOW() " & _
    "WHERE AGDS_AGDID = '" & clavesCitas(fila, columna) & "' AND AGDS_SERVID = '" & CitaServId(fila, columna) & "' and agds_usuario_perid = '" & listaDia.TextMatrix(2, columna) & "' "
    'MsgBox SQL1
    con.Execute (sql1)
    
    cmdCitas_Click
    
    MsgBox "Cita movida. Verifique.", vbInformation
    
    
End Sub

Private Sub listaDia_DragOver(Source As Control, X As Single, Y As Single, State As Integer)


        archivo = Dir(direccionSistema & "\Com\move.ico", vbArchive)
        If archivo <> "" Then
            lInfo.DragIcon = LoadPicture(direccionSistema & "\Com\move.ico")
        Else
            lInfo.DragIcon = LoadPicture()
        End If


End Sub

Private Sub listaDia_GotFocus()
    ConScroll listaDia
End Sub

Private Sub listaDia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        lInfo.Visible = False
        lInfo.DragIcon = LoadPicture("")
    End If
End Sub

Private Sub listaDia_LostFocus()
    tTiempo.Enabled = True
    SinScroll listaDia
End Sub

Private Sub listaDia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tTiempo.Enabled = False
    
    Dim a As Long, b As Long
    
    lInfo.DragIcon = LoadPicture()
    lInfo.Drag (0)
    lInfo.Visible = False
        
    If listaDia.Rows > 2 And listaDia.Cols > 1 Then
    
        fila = listaDia.Row
        columna = listaDia.Col
        
        
        'MsgBox "Down: " & listaDia.TextMatrix(fila, columna)
        
        If fila > 2 And columna > 0 Then
            If Button = vbRightButton Then
                If listaDia.TextMatrix(fila, columna) = "" Then
                    mn_Edit.Enabled = False
                    mn_VerInfo.Enabled = False
                Else
                    mn_Edit.Enabled = True
                    mn_VerInfo.Enabled = True
                End If
                PopupMenu mn_Cita, vbPopupMenuLeftAlign
            
            Else
                If listaDia.TextMatrix(fila, columna) <> "" Then
                    Color = listaDia.CellBackColor
                    lInfo.Visible = True
                    lInfo.Font = listaDia.CellFontName
                    lInfo.FontBold = True
                    lInfo.FontSize = 9
                    lInfo.ForeColor = vbWhite
                    lInfo.Move listaDia.CellLeft + listaDia.Left, listaDia.CellTop + listaDia.Top, listaDia.CellWidth, listaDia.CellHeight
                    lInfo.Caption = listaDia.TextMatrix(fila, columna)
                    lInfo.Drag
                End If
            End If
        End If
    End If
End Sub

Private Sub listaDia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Static f As Long
      Static c As Long
        
      With listaDia
        If .Rows > 3 And .Cols > 1 Then
            If .MouseRow > 3 And .MouseCol > 0 Then
                    
                  If f <> .MouseRow Or c <> .MouseCol Then
'                    If .TextMatrix(.MouseRow, .MouseCol) <> "" Then
'                        f = .MouseRow
'                        c = .MouseCol
'                        tTip.Titulo = "Cita"
'                        tTip.Texto = textoCitas(.MouseRow, .MouseCol)
'                        tTip.Crear .hwnd ' crea el Tips
'                    End If
                   End If
            Else
                ' Lo destruye
'                tTip.Destroy
            End If
        End If
      End With



End Sub

Private Sub mn_CancelCita_Click()
    Dim ques As String
    
    ques = MsgBox("Cancelar la cita: " & vbCrLf & vbCrLf & textoCitas(listaDia.Row, listaDia.Col), vbYesNo + vbQuestion)
    If ques = vbYes Then
        sql1 = "UPDATE AGENDA SET AGD_STATUS = 'C' WHERE AGD_ID = '" & clavesCitas(listaDia.Row, listaDia.Col) & "'"
        con.Execute (sql1)
        
        cargaCitas
        MsgBox "Cita cancelada.", vbInformation
    End If
End Sub

Private Sub mn_Cobrar_Click()
    SinScroll listaDia
    Dim ques As String
    
    ques = MsgBox("Cobrar cita: " & vbCrLf & vbCrLf & textoCitas(listaDia.Row, listaDia.Col), vbYesNo + vbQuestion)
    
    If ques = vbYes Then
        'MDIC_Operaciones.Show
        MDIC_OperTickets.Show
        MDIC_OperTickets.Opt(0).value = True
        
        For b1 = 1 To MDIC_OperTickets.lista.Rows - 1
            If MDIC_OperTickets.lista.TextMatrix(b1, 0) = folioCobro(listaDia.Row, listaDia.Col) Then
                MDIC_OperTickets.lista.Row = b1
                MDIC_OperTickets.lista.Col = 0
                 'timerAgendaCobro.Enabled = True
                'MDIC_OperTickets.Lista_DblClick
                'MDIC_OperTickets.timerCobro.Enabled = True
                MDIC_OperTickets.operacionTicket
                Exit Sub
            End If
        Next b1
    End If
    
End Sub

Private Sub mn_Edit_Click()
    listaDia_DblClick
End Sub

Private Sub mn_VerInfo_Click()
    
    MsgBox textoCitas(listaDia.Row, listaDia.Col), vbInformation
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
cmdCitas_Click

End Sub


Private Sub timerAgendaCobro_Timer()
timerAgendaCobro.Enabled = False
MDIC_OperTickets.timerCobro.Enabled = True
End Sub

Private Sub tTiempo_Timer()
    lHora.Caption = Format(Time, "Short Time")
    
    checkHoraActual
    'cargaImagen
End Sub
Private Sub checkHoraActual()
    
    On Error Resume Next
    
        For b1 = 4 To listaDia.Rows - 1
            '''Para mover la lista a la hora actual
            If Format(listaDia.TextMatrix(b1, 0), "Short Time") > Format(Time, "Short Time") Then
                listaDia.Row = b1
                listaDia.Col = 0
                listaDia.CellBackColor = &HE0E0E0
                If chkHora.value = Checked Then
                    If b1 > 1 Then
                        listaDia.TopRow = b1 - 1
                    End If
                End If
                Exit For
            Else
                listaDia.Row = b1
                listaDia.Col = 0
                listaDia.CellBackColor = &H80FF&
            End If
        Next b1

End Sub

Private Sub TTime_Timer()
    tTime.Enabled = False
    'listaDia.Visible = False
    For b1 = 0 To 3
        Shape1(b1).Top = Me.height - 800
        lUsuario(b1).Top = Me.height - 800
    Next b1
    
    Line1(0).Y1 = Me.height - 900
    Line1(0).Y2 = Me.height - 900
    Label1(0).Top = Me.height - 1100
End Sub
