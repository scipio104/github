Attribute VB_Name = "Module1"
'Function GLline(b, L, df, d, cB, cD, ex1, ey1)

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

GLline = boxh

    sw = boxw / L
    sh = boxh / (b + 300 + df)
    
    If sw < sh Then
        S = sw
    Else
        S = sh
    End If
'kiso cennter line v
    With ActiveSheet.Shapes.AddLine(ex - L / 2 * S, ey + 15, ex - L / 2 * S, by - 10)
       .line.ForeColor.SchemeColor = 2
       .Name = "line01"
    End With
'column cennter line vl
    With ActiveSheet.Shapes.AddLine(ex - (L / 2 - ex1) * S, ey + 15, ex - (L / 2 - ex1) * S, by - 10)
       .line.ForeColor.SchemeColor = 4
       .Name = "line01"
    End With
'kiso cennter line h
    With ActiveSheet.Shapes.AddLine(ex - L * S - 20, ey - b / 2 * S, ex + 27, ey - b / 2 * S)
       .line.ForeColor.SchemeColor = 2
       .Name = "line01"
    End With
'columun cennter line h
    With ActiveSheet.Shapes.AddLine(ex - L * S - 20, ey - (b / 2 + ey1) * S, ex + 27, ey - (b / 2 + ey1) * S)
       .line.ForeColor.SchemeColor = 4
       .Name = "line01"
    End With
    

'plan
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey, ex, ey) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey - b * S, ex, ey - b * S) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey, ex - L * S, ey - b * S) 'tate left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, ey, ex, ey - b * S) 'tate right
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
'section
    With ActiveSheet.Shapes.AddLine(ex - L * S, by + df * S, ex, by + df * S) 'bottom yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, by + (df - d) * S, ex, by + (df - d) * S) 'depth up
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, by + df * S, ex - L * S, by + (df - d) * S) 'depth left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, by + df * S, ex, by + (df - d) * S) 'depth right
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1 + cD) * S, by + (df - d) * S, ex - ((L - cD) / 2 - ex1 + cD) * S, by) 'column leftt
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With


    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1) * S, by + (df - d) * S, ex - ((L - cD) / 2 - ex1) * S, by) 'column rightt
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With

'column

    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1 + cD) * S, ey - ((b - cB) / 2 + ey1) * S, ex - ((L - cD) / 2 - ex1) * S, ey - ((b - cB) / 2 + ey1) * S) 'yoko douwn
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1 + cD) * S, ey - ((b - cB) / 2 + ey1 + cB) * S, ex - ((L - cD) / 2 - ex1) * S, ey - ((b - cB) / 2 + ey1 + cB) * S) 'yoko up
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1 + cD) * S, ey - ((b - cB) / 2 + ey1) * S, ex - ((L - cD) / 2 - ex1 + cD) * S, ey - ((b - cB) / 2 + ey1 + cB) * S) 'tate left
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - ((L - cD) / 2 - ex1) * S, ey - ((b - cB) / 2 + ey1) * S, ex - ((L - cD) / 2 - ex1) * S, ey - ((b - cB) / 2 + ey1 + cB) * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
    
'sunpou
    'varchi
    With ActiveSheet.Shapes.AddLine(ex + 50, by + df * S, ex + 50, by)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 25, by + df * S, ex + 25, by + (df - d) * S) 'tate
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, by + df * S, ex + 55, by + df * S) 'yoko
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, by + (df - d) * S, ex + 28, by + (df - d) * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
'sunpou Plan bottom
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey + 30, ex, ey + 30)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey + 15, ex, ey + 15)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex - L * S, ey + 30, ex - L * S, ey + 3)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex, ey + 30, ex, ey + 3)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
'plan right
    With ActiveSheet.Shapes.AddLine(ex + 50, ey, ex + 50, ey - b * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 25, ey, ex + 25, ey - b * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, ey, ex + 52, ey)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    With ActiveSheet.Shapes.AddLine(ex + 5, ey - b * S, ex + 52, ey - b * S)
        .line.ForeColor.SchemeColor = 8
        .Name = "line02"
    End With
    
'text df
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxDf"
        .TextFrame.Characters.Text = df
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft 18
        .IncrementTop 15
        
    End With
'text D
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, by + df * S, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxD"
        .TextFrame.Characters.Text = d
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft -5
        .IncrementTop -40
    End With
'text B
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, ey, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxB"
        .TextFrame.Characters.Text = b
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft 20
        .IncrementTop -b * S * 0.6
    End With
    
    If ey1 > 0 Then
        ey2 = ey
    Else
        ey2 = -b * S * 0.5 + ey
    End If
    
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, ey2, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl2"
        .TextFrame.Characters.Text = b / 2
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        .IncrementLeft -6
        .IncrementTop -b * S * 0.4
    End With
    
    If ey1 > 0 Then
        ey3 = ey - b * S * 0.5 - 33
    Else
        ey3 = ey - b * S * 0.5 + 10
        
    End If
    
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, ey3, 60, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxex"
        .TextFrame.Characters.Text = "ey=" & ey1
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementRotation 270
        
        .IncrementLeft -11
        .IncrementTop 0
    End With
    
    
    
'text L
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex, ey, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl"
        .TextFrame.Characters.Text = L
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft -L * S * 0.6
        .IncrementTop 15
    End With
    
    If ex1 > 0 Then
        ex2 = -L * S + ex
    Else
        ex2 = -L * S * 0.5 + ex
    End If
    
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex2, ey, 50, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxl2"
        .TextFrame.Characters.Text = L / 2
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft L / 2 * S * 0.3
        .IncrementTop 0
    End With
    
    If ex1 > 0 Then
        ex3 = -L * S * 0.5 + ex
    Else
        ex3 = -L * S * 0.5 + ex - 50
    End If
    
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, ex3, ey, 70, 20)
        .line.ForeColor.SchemeColor = 8
        .Name = "textboxex"
        .TextFrame.Characters.Text = "ex=" & ex1
        .line.Visible = mosfales
        .Fill.Visible = False
        .IncrementLeft 0
        .IncrementTop 0
    End With

  GLline = "äÓëbê°ñ@"
  
End Function
