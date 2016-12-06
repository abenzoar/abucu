VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form RPT_RrtMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15495
   Icon            =   "RPT_RrtMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15495
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      Picture         =   "RPT_RrtMain.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer timerTamaño 
      Interval        =   50
      Left            =   13320
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4680
      Picture         =   "RPT_RrtMain.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   706
      TabCaption(0)   =   "  Utilidades"
      TabPicture(0)   =   "RPT_RrtMain.frx":13DE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Graf1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  Gastos"
      TabPicture(1)   =   "RPT_RrtMain.frx":1978
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Graf2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "  Top ventas"
      TabPicture(2)   =   "RPT_RrtMain.frx":1F12
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "  Cuentas por pagar"
      TabPicture(3)   =   "RPT_RrtMain.frx":24AC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin MSChart20Lib.MSChart Graf2 
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "RPT_RrtMain.frx":2A46
         TabIndex        =   7
         Top             =   480
         Width           =   15135
      End
      Begin MSChart20Lib.MSChart Graf1 
         Height          =   7455
         Left            =   -74880
         OleObjectBlob   =   "RPT_RrtMain.frx":4F2D
         TabIndex        =   6
         Top             =   480
         Width           =   15135
      End
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109838337
      CurrentDate     =   40829
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109838337
      CurrentDate     =   40829
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "RPT_RrtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset



Private Sub cmdAccion_Click(Index As Integer)

End Sub

Private Sub cmdPrint_Click()
    If SSTab1.Tab = 0 Then
        Call Imprimir_grafico("Gráfico de ejemplo", Graf1, 0.5)
    Else
        If SSTab1.Tab = 1 Then
            Call Imprimir_grafico("Gráfico de ejemplo", Graf2, 0.5)
        End If
    End If
End Sub
Private Sub Imprimir_grafico(titulo As String, _
                             MsChart As MsChart, _
                             escala As Single)
      
      
    With MsChart
        ' elimina los datos del clipboard
        Clipboard.Clear
          
        ' copia la imagen del mschart al portapapeles
        .EditCopy
          
        ' sila imagen es válida
        If Clipboard.GetFormat(vbCFBitmap) Then
            'scale mode
            Printer.ScaleMode = vbTwips
            .Parent.ScaleMode = vbTwips
              
           ' titulo
            Printer.Font.Size = 10
            Printer.FontName = "Verdana"
              
            Printer.Print vbNullString
           ' Printer.Print titulo
            Printer.Print vbNullString
              
            ' dibuja la imagen
            Printer.PaintPicture Clipboard.GetData(), 100, 500, _
                                 .width * escala, .height * escala, 0, 0
          
              
            Printer.EndDoc ' envía el trabajo a la impresora
        End If
    End With
End Sub
Private Sub Command2_Click()
    
    Select Case SSTab1.Tab
        Case 0: cargaUtilidad
        Case 1: cargaGastos
    
    End Select
        
End Sub
Private Sub cargaUtilidad()

    SQL1 = "SELECT SUM(UTILIDAD) UTILIDAD, SUM(VENTA) VENTAS, SUM(GASTOS) GASTOS FROM VIEW_RPT_UTILIDADES " & _
    "WHERE fecha BETWEEN date_format('" & Format(dtFecha1(0), "yyyy-MM-dd") & "' , '%Y-%m-%d')  " & _
    "AND date_format('" & Format(dtFecha1(1), "yyyy-MM-dd") & "' , '%Y-%m-%d')  "
    Set RES1 = con.Execute(SQL1)


    Dim ventas As Double, utilidad As Double, gastos As Double
    If IsNull(RES1.Fields("VENTAS")) = True Then
        ventas = 0
    Else
        ventas = RES1.Fields("VENTAS")
    End If
    If IsNull(RES1.Fields("utilidad")) = True Then
        utilidad = 0
    Else
        utilidad = Val(RES1.Fields("UTILIDAD"))
    End If
    If IsNull(RES1.Fields("gastos")) = True Then
        gastos = 0
    Else
        gastos = Val(RES1.Fields("GASTOS"))
    End If
    
    Graf1.TitleText = "Utilidades  " & dtFecha1(0) & "  " & dtFecha1(1)
    With Graf1.DataGrid
        .RowLabel(1, 1) = FormatCurrency(ventas) & vbCrLf & vbCrLf & "  Ventas"
        .RowLabel(2, 1) = FormatCurrency(utilidad) & vbCrLf & vbCrLf & "  Utilidad"
        .RowLabel(3, 1) = FormatCurrency(gastos) & vbCrLf & vbCrLf & "  Gastos"
        .SetSize 3, 1, 3, 1
        .SetData 1, 1, ventas, 0
        .SetData 2, 1, utilidad, 0
        .SetData 3, 1, gastos, 0
    End With


End Sub
Private Sub cargaGastos()
    Dim b1 As Integer
    Dim valor(50) As String
    Dim total(50) As Double
    
    SQL1 = "SELECT SUM(TOTAL) TOTAL, GASTO FROM VIEW_RPT_GASTOS " & _
    "WHERE fecha BETWEEN date_format('" & Format(dtFecha1(0), "yyyy-MM-dd") & "' , '%Y-%m-%d')  " & _
    "AND date_format('" & Format(dtFecha1(1), "yyyy-MM-dd") & "' , '%Y-%m-%d') GROUP BY GASTO "
    Set RES1 = con.Execute(SQL1)
    
    b1 = 0

    Do While Not RES1.EOF
        b1 = b1 + 1
        total(b1) = Val(RES1.Fields("TOTAL"))
        valor(b1) = UCase(RES1.Fields("GASTO"))

        RES1.MoveNext
    Loop

    With Graf2
        
        .TitleText = "Gastos  " & dtFecha1(0) & "  " & dtFecha1(1)
        
        .chartType = VtChChartType2dPie
        
        .RowCount = 1
        .ColumnCount = b1
        .RowLabel = ""
        
        For c1 = 1 To b1
            .Row = 1
            .Column = c1
            .ColumnLabel = valor(c1)
            .Data = total(c1)
        Next c1
          
        If b1 > 0 Then
            For b1 = 1 To .Plot.SeriesCollection.Count
                .Plot.SeriesCollection(b1).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
                '.Plot.SeriesCollection(b1).DataPoints(-1).DataPointLabel.TextLayout.Orientation = VtOrientationUp
                .Plot.SeriesCollection(b1).DataPoints(-1).DataPointLabel.VtFont.Size = 10
                
            Next b1
        End If
        'chart1.Plot.SeriesCollection(1).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        'chart1.Plot.SeriesCollection(1).DataPoints(-1).DataPointLabel.TextLayout.Orientation = VtOrientationUp
        
        .ShowLegend = True
        
    End With

        
    
    
End Sub

Private Sub cargaMes()
    Dim inicio_Mes, fin_mes As Date
    
    dtFecha1(0) = Date
    dtFecha1(1) = Date
    inicio_Mes = DateSerial(dtFecha1(0).Year, Month(Date), 1)
    fin_mes = DateSerial(dtFecha1(1).Year, Month(Date) + 1, 1)
    fin_mes = DateAdd("d", -1, fin_mes)

    dtFecha1(0) = inicio_Mes
    dtFecha1(1) = fin_mes

End Sub

Private Sub Form_Load()
    cargaMes
    SSTab1.Tab = 0
    cargaUtilidad
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Command2_Click
End Sub

Private Sub timerTamaño_Timer()
    timerTamaño.Enabled = False
     SSTab1.width = Me.width - 200
     SSTab1.height = Me.height - 1500
     
     Graf1.width = Me.width - 400
     Graf1.height = Me.height - 2000
     Graf2.width = Me.width - 400
     Graf2.height = Me.height - 2000
'     Graf5.width = Me.width - 400
'     Graf5.height = Me.height - 2000
     
End Sub
