sub 制作工资条()

  dim i as long
  for i = 2 to range("a1").currentRegion.Rows.Count - 1
  
  ActiverCell.offset(2,0).rows("1:2").EntireRow.Select
  Selection.Insert Shift:=xlDown,copyOrigin:=xlFormatFormatFromleftOrAbove
  ActiveCell,offset(-2,0).Rows("1:1").EntireRow.select
  Selection.Copy
  ActiveCell.offset(3,0).Rows("1:1").EntireRow.Select
  ActiveCell.Offset(-1,0).Range("A1:G1").Select
  Application.CutCopyMode = False
  Selection.Borders(xLDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle ＝　xlNone
  With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 5
      .TintAndshade = -0.499984740745262
      .Weight = xlThin
  End with
  With Selection.Border(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ThemeColor = 5
      .TintAndShade = -0.499984740745262
      .Weight = xlThin
  End with
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  Selection.Borders(xlInsideVertical).LineStyle = xlNone
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  ActiveCell.offset(1,0).Range("A1").Select
  
  Next

End sub
  
