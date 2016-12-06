Attribute VB_Name = "FCN_Export"
Option Explicit
'Variables para acceder a la hoja excel
Dim objExcel As Object
Dim columna  As Integer

Public Sub exportExcel(lista As MSFlexGrid)

    Set objExcel = CreateObject("Excel.application")
    With objExcel
    .Workbooks.Add
    Dim b1 As Long
    Dim c1 As Long
    For b1 = 1 To lista.Rows
        For c1 = 1 To lista.Cols
            .cells(b1, c1).Formula = " " & lista.TextMatrix(b1 - 1, c1 - 1)
        Next c1
    Next b1
    'Formato de celdas ( fuente y color de fondo )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' autoajustar las columnas
    .Columns("A:A").EntireColumn.AutoFit
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
    .Columns("D:D").EntireColumn.AutoFit
    .Columns("E:E").EntireColumn.AutoFit
    .Columns("F:F").EntireColumn.AutoFit
    .Columns("G:G").EntireColumn.AutoFit
    .Columns("H:H").EntireColumn.AutoFit
    .Columns("I:I").EntireColumn.AutoFit
    .Columns("J:J").EntireColumn.AutoFit
    .Columns("K:K").EntireColumn.AutoFit
    .Columns("L:L").EntireColumn.AutoFit
    .Columns("M:M").EntireColumn.AutoFit
      
      
    .Range("A1:M1").Select
    With .Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With .Selection.Interior
        .colorindex = 35
    End With
'    .Range("A2:D13").Select
'    .Selection.Font.colorindex = 2
'    With .Selection.Interior
'        .colorindex = 11
'    End With
'
'    .Range("A14:D16").Select
'
'    .Selection.Interior.colorindex = 35
'
'    .Range("A14:A16").Select
'    .Selection.Font.FontStyle = "Negrita"
  
    End With
  
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ahcer visible el Excel
    objExcel.Visible = True
      
    ' eliminar la referencia
    Set objExcel = Nothing

End Sub

Public Sub exportExcel1()

    ' crear la referencia a excel
    Set objExcel = CreateObject("Excel.application")
    With objExcel
    ' Agregar un Nuevo libro
    .Workbooks.Add
    Dim i As Integer
    ' Agregar los nombres de los meses a la columna Meses
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    columna = 1
    For i = 1 To 12
        .cells(1, columna).Formula = "Meses"
        .cells(i + 1, columna).Formula = MonthName(i)
    Next
          
    ' agregar costos valores para la columna 2  ( Gastos Productos )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    columna = 2
    For i = 1 To 12
        .cells(1, columna).Formula = "Gastos Productos"
        .cells(i + 1, columna).Formula = CInt(Rnd * 255)
          
    Next
      
    ' agregar valores a la columna 3 (Gastos impuestos)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    columna = 3
    For i = 1 To 12
        .cells(1, columna).Formula = "Gastos impuestos"
        .cells(i + 1, columna).Formula = CInt(Rnd * 150)
          
    Next
      
    ' agregar valores a la columna 4 ( Otros gastos )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    columna = 4
    For i = 1 To 12
        .cells(1, columna).Formula = "Otros gastos"
        .cells(i + 1, columna).Formula = CInt(Rnd * 50)
    Next
      
      
    ' Sacar el SubTotal para cada columna
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    .cells(14, 1).Formula = "SubTotales"
    .cells(14, 2).Formula = "=SUM(B2:B13)"
      
    ' SubTotal para la columna 3
    .cells(14, 3).Formula = "=SUM(C2:C13)"
  
    ' SubTotal para la columna 4
    .cells(14, 4).Formula = "=SUM(D2:D13)"
      
      
    ' Total
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    .cells(16, 1).Formula = "Total"
    .cells(16, 4).Formula = "=SUM(B14:D14)"
  
      
    'Formato de celdas ( fuente y color de fondo )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' autoajustar las columnas
    .Columns("A:A").EntireColumn.AutoFit
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
    .Columns("D:D").EntireColumn.AutoFit
      
      
    .Range("A1:D1").Select
    With .Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With .Selection.Interior
        .colorindex = 35
  
    End With
    .Range("A2:D13").Select
      
    .Selection.Font.colorindex = 2
    With .Selection.Interior
        .colorindex = 11
    End With
      
    .Range("A14:D16").Select
      
    .Selection.Interior.colorindex = 35
      
    .Range("A14:A16").Select
    .Selection.Font.FontStyle = "Negrita"
  
    End With
  
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ahcer visible el Excel
    objExcel.Visible = True
      
    ' eliminar la referencia
    Set objExcel = Nothing

End Sub

Public Sub exportExcel2(lista As MSFlexGrid)
    Dim fila1 As Long
    Dim fila2 As Long
    Dim colu As Long
    
    Set objExcel = CreateObject("Excel.application")
    With objExcel
    .Workbooks.Add
    Dim b1 As Long
    Dim c1 As Long
    Dim d1 As Long
    fila1 = 1
    
'        If ListaUsers.TextMatrix(b1, 14) = Chr(168) Then
'            ListaUsers.TextMatrix(b1, 14) = Chr(254)
'            enviaProductoSel (ListaUsers.Row)
'        Else
'            ListaUsers.TextMatrix(b1, 14) = Chr(168)
'        End If

    
    For b1 = 1 To lista.Rows
        If Val(lista.TextMatrix(b1 - 1, 4)) > 0 And lista.TextMatrix(b1 - 1, 14) = Chr(254) Then
            For d1 = 1 To Val(lista.TextMatrix(b1 - 1, 4))
                For c1 = 1 To lista.Cols - 1
                    .cells(fila1, c1).Formula = " " & lista.TextMatrix(b1 - 1, c1 - 1)
                Next c1
    '            .cells(b1, c1).Formula = " " & lista.TextMatrix(b1 - 1, c1 - 1)
                fila1 = fila1 + 1
            Next d1
        Else
            If lista.TextMatrix(b1 - 1, 4) = "Cant" Then
                For c1 = 1 To lista.Cols - 1
                    .cells(fila1, c1).Formula = " " & lista.TextMatrix(b1 - 1, c1 - 1)
                Next c1
                fila1 = fila1 + 1
            End If
        End If
    Next b1
    'Formato de celdas ( fuente y color de fondo )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' autoajustar las columnas
    .Columns("A:A").EntireColumn.AutoFit
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
    .Columns("D:D").EntireColumn.AutoFit
    .Columns("E:E").EntireColumn.AutoFit
    .Columns("F:F").EntireColumn.AutoFit
    .Columns("G:G").EntireColumn.AutoFit
    .Columns("H:H").EntireColumn.AutoFit
      
      
    .Range("A1:H1").Select
    With .Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With .Selection.Interior
        .colorindex = 35
    End With
'    .Range("A2:D13").Select
'    .Selection.Font.colorindex = 2
'    With .Selection.Interior
'        .colorindex = 11
'    End With
'
'    .Range("A14:D16").Select
'
'    .Selection.Interior.colorindex = 35
'
'    .Range("A14:A16").Select
'    .Selection.Font.FontStyle = "Negrita"
  
    End With
  
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ahcer visible el Excel
    objExcel.Visible = True
      
    ' eliminar la referencia
    Set objExcel = Nothing

End Sub

'''''''''''''''Exportar a excel en maestro y detalle
Public Sub exportExcel_MD(Lista1 As MSFlexGrid, Lista2 As MSFlexGrid)

    Set objExcel = CreateObject("Excel.application")
    With objExcel
    .Workbooks.Add
    Dim b1 As Long
    Dim c1 As Long
    
    Dim b2 As Long
    Dim c2 As Long
    
    For b1 = 1 To Lista1.Rows
        For c1 = 1 To Lista1.Cols - 1
            .cells(b1, c1).Formula = " " & Lista1.TextMatrix(b1 - 1, c1 - 1)
        Next c1
    Next b1
    
    b1 = b1 + 1
    
    For b2 = 1 To Lista2.Rows
        For c2 = 1 To Lista2.Cols - 1
            .cells(b2 + b1, c2).Formula = " " & Lista2.TextMatrix(b2 - 1, c2 - 1)
        Next c2
    Next b2
        
    'Formato de celdas ( fuente y color de fondo )
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' autoajustar las columnas
    .Columns("A:A").EntireColumn.AutoFit
    .Columns("B:B").EntireColumn.AutoFit
    .Columns("C:C").EntireColumn.AutoFit
    .Columns("D:D").EntireColumn.AutoFit
    .Columns("E:E").EntireColumn.AutoFit
    .Columns("F:F").EntireColumn.AutoFit
    .Columns("G:G").EntireColumn.AutoFit
    .Columns("H:H").EntireColumn.AutoFit
    .Columns("I:I").EntireColumn.AutoFit
    .Columns("J:J").EntireColumn.AutoFit
    .Columns("K:K").EntireColumn.AutoFit
    .Columns("L:L").EntireColumn.AutoFit
    .Columns("M:M").EntireColumn.AutoFit
    .Columns("N:N").EntireColumn.AutoFit
      
      
    .Range("A1:N1").Select
    With .Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With .Selection.Interior
        .colorindex = 35
    End With
    
    .Range("A5:N5").Select
    With .Selection.Font
        .Name = "Arial"
        .FontStyle = "Negrita"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
    End With
    With .Selection.Interior
        .colorindex = 35
    End With
    
'    .Range("A2:D13").Select
'    .Selection.Font.colorindex = 2
'    With .Selection.Interior
'        .colorindex = 11
'    End With
'
'    .Range("A14:D16").Select
'
'    .Selection.Interior.colorindex = 35
'
'    .Range("A14:A16").Select
'    .Selection.Font.FontStyle = "Negrita"
  
    End With
  
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ahcer visible el Excel
    objExcel.Visible = True
      
    ' eliminar la referencia
    Set objExcel = Nothing

End Sub

