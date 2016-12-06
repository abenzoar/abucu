Attribute VB_Name = "BordesText"
Public Sub bordesProductos()
'On Error Resume Next
Dim b1 As Long
Dim c1 As Long
    With FRM_Productos
        c1 = 0
        For b1 = 0 To .txtProd.Count - 1
            Load .Borde(b1 + 1)
            c1 = c1 + 1
            .txtProd(b1).BorderStyle = 0
            .Borde(b1).Visible = True
            .Borde(b1).width = .txtProd(b1).width
            .Borde(b1).height = .txtProd(b1).height
            .Borde(b1).Top = .txtProd(b1).Top
            .Borde(b1).Left = .txtProd(b1).Left
            .Borde(b1).BorderWidth = 4
            .Borde(b1).BorderColor = &H80FF&
        Next b1
        For b1 = 0 To .cmbProd.Count - 1
            c1 = c1 + 1
            Load .Borde(c1)
            .cmbProd(b1).Appearance = 0
            .Borde(c1).Visible = True
            .Borde(c1).width = .cmbProd(b1).width
            .Borde(c1).height = .cmbProd(b1).height
            .Borde(c1).Top = .cmbProd(b1).Top
            .Borde(c1).Left = .cmbProd(b1).Left
            .Borde(c1).BorderWidth = 4
            .Borde(c1).BorderColor = &H80FF&
        Next b1
        

        c1 = c1 + 1
        Load .Borde(c1)
        .iFoto.BorderStyle = 0
        .Borde(c1).Visible = True
        .Borde(c1).width = .iFoto.width
        .Borde(c1).height = .iFoto.height
        .Borde(c1).Top = .iFoto.Top
        .Borde(c1).Left = .iFoto.Left
        .Borde(c1).BorderWidth = 2
        .Borde(c1).BorderColor = &H80FF&
        
    End With

End Sub

