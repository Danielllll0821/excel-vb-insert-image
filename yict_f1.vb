Option Explicit
Private Sub fladdy()
  Dim PicLib, varRow, CodeCol, ColorCol, PicCol, Code, Color, PicFileName, PicPath, PicLeft, PicTop, PicWidth, PicHeight
  PicLib = "D:"  '图片库路径
  varRow = 6  '商品代码起始行
  CodeCol = "B"  '商品代码所在列
  ColorCol = "D"  '商品颜色所在列
  PicCol = "A" '商品图片所在列
  Do While Len(Sheet1.Cells(varRow, CodeCol)) = 8
    Code = CStr(Sheet1.Cells(varRow, CodeCol).Value)
    Color = Sheet1.Cells(varRow, ColorCol).Value
    PicFileName = Code + Color + ".jpg"
    PicPath = PicLib + PicFileName
    Sheet1.Range(PicCol & varRow).RowHeight = 70
    Sheet1.Range(PicCol & varRow).ColumnWidth = 20
    PicLeft = Range(PicCol & varRow).Left '图片左边位置
    PicTop = Range(PicCol & varRow).Top  '图片顶部位置
    PicWidth = Range(PicCol & varRow).Width  '图片宽度
    PicHeight = Range(PicCol & varRow).Height  '图片高度
    '插入图片
    On Error Resume Next
    Sheet1.Shapes.AddPicture PicPath, msoFalse, msoTrue, PicLeft, PicTop, PicWidth, PicHeight
    varRow = varRow + 1
  Loop
  MsgBox "执行完毕"
  Exit Sub
myErr:
  MsgBox (PicFileName & " 未找到，按确定继续处理。")
  Resume Next
End Sub