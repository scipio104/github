Attribute VB_Name = "Module3"
Function nisoujiban(b, L, df, H, wl)

For Each shp In ActiveSheet.Shapes
    If Not (shp.Type = msoOLEControlObject Or shp.Type = msoFormControl) Then shp.Delete
Next shp
    Set rngStart = Range("b4")
    Set rngEnd = Range("f17")
        bx = rngStart.Left
        by = rngStart.Top
        ex = rngEnd.Left
        ey = rngEnd.Top

'GL
    With ActiveSheet.Shapes.AddLine(bx, by, ex + 60, by)
       .line.ForeColor.SchemeColor = 8
       .Name = "line01"
    End With
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, bx, by - 15, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textbox01"
        .TextFrame.Characters.Text = "Å§GL"
        .IncrementTop -2
        .line.Visible = mosfales
        .Fill.Visible = False
    End With

boxw = ex - bx
boxh = ey - by


    h1 = H - df
    h2 = h1 / 2

    sw = boxw / (L + h1)
    sh = boxh / H
    
    If sw < sh Then
        S = sw
    Else
        S = sh
    End If

'kabu jibann
    With ActiveSheet.Shapes.AddLine(bx - 10, by + H * S, ex + 10, by + H * S) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With

    With ActiveSheet.Shapes.AddLine(ex - (L + h1) * S, by + H * S, ex - (L + h2) * S, by + df * S) 'tate left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, by + H * S, ex - h1 / 2 * S, by + df * S) 'tate right
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
'section
    With ActiveSheet.Shapes.AddLine(ex - (L + h2) * S, by + df * S, ex - h2 * S, by + df * S) 'bottom yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S - h1 / 2 * S, by + (df - d) * S, ex - h1 / 2 * S, by + (df - d) * S) 'depth up
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - (L + h2) * S, by + df * S, ex - (L + h2) * S, by + (df - d) * S) 'depth left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - h2 * S, by + df * S, ex - h2 * S, by + (df - d) * S) 'depth right
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
    With ActiveSheet.Shapes.AddLine(ex - (h2 + L) * S, by + (df - d) * S, ex - (h2 + L) * S, by) 'column left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - h2 * S, by + (df - d) * S, ex - h2 * S, by) 'column left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
    pl = L * 0.1
    ph = pl * 2
    n = Round(L / pl, 0)
    m = Round((L + h1) / ph, 0)
    
    For i = 1 To n
        With ActiveSheet.Shapes.AddLine(ex - (h2 + L - pl * i) * S, by + df * S, ex - (h2 + L - pl * i) * S, by + df * S - 10) 'column leftt
            .line.ForeColor.SchemeColor = 8
            .Name = "line02"
        End With
    Next
    
     For i = 0 To m
        With ActiveSheet.Shapes.AddLine(ex - (h1 + L - ph * i) * S, by + H * S, ex - (h1 + L - ph * i) * S, by + H * S - 10) 'column leftt
            .line.ForeColor.SchemeColor = 8
            .Name = "line02"
        End With
    Next
    
'sunpou
    'varchi
    With ActiveSheet.Shapes.AddLine(ex + 50, by + H * S, ex + 50, by) 'tate
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 25, by + H * S, ex + 25, by) 'tate
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, by + df * S, ex + 25, by + df * S) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, by + H * S, ex + 50, by + H * S) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
    
    
'sunpou Plan bottom
    With ActiveSheet.Shapes.AddLine(ex - (L + h1) * S, by + H * S + 21, ex, by + H * S + 21)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - (L + h1) * S, by + H * S + 22, ex - (L + h1) * S, by + H * S + 3)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, by + H * S + 22, ex, by + H * S + 3)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
'text df
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by + df * S, 60, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxDf"
        .TextFrame.Characters.Text = "df=" & df
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft -10
        .IncrementTop -40
    End With
'text df
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by + H * S, 100, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxDf"
        .TextFrame.Characters.Text = "h-df=" & H - df
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft -20
        .IncrementTop -60
    End With
'text h
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, ey, 65, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxB"
        .TextFrame.Characters.Text = "H=" & H
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft 10
        .IncrementTop -150
    End With
    
'text L
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, (ex - bx) / 2, by + H * S, 150, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl"
        .TextFrame.Characters.Text = "L+H-Df=" & L + h1
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft 0
        .IncrementTop 5
    End With
'text L
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex - (L + h1) / 2 * S, by + df * S, 100, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl"
        .TextFrame.Characters.Text = "L=" & L
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft -20
        .IncrementTop -30
    End With
'text B
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, (ex - bx) / 2, by + H * S, 150, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl"
        .TextFrame.Characters.Text = "B+H-Df=" & b + h1
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft 10
        .IncrementTop 17
    End With
'text B
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex - (L + h1) / 2 * S, by + df * S, 100, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl"
        .TextFrame.Characters.Text = "B=" & b
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft -20
        .IncrementTop 0
    End With
    
'sunpou tryangle

    With ActiveSheet.Shapes.AddLine(ex, by + df * S, ex - h2 * 0.5 * S, by + df * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, by + df * S, ex, by + df * S + h2 * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, by + df * S + h2 * S, ex - h2 * 0.5 * S, by + df * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
 'text
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by + df * S, 20, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textbox2"
        .TextFrame.Characters.Text = "2"
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft -5
        .IncrementTop 5
    End With
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by + df * S, 20, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textbox2"
        .TextFrame.Characters.Text = "1"
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft -15
        .IncrementTop -15
    End With
'water level

    With ActiveSheet.Shapes.AddLine(bx, by + wl * S, bx + h2 * 0.5 * S, by + wl * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, bx, by + wl * S, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textbox01"
        .TextFrame.Characters.Text = "Å§WL"
        .IncrementTop -2
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementTop -15

    End With
    
    
    
  nisoujiban = "ó™ê}"
  
End Function


