Attribute VB_Name = "table"
Const com As String = ",": Const at As String = "@": Const rc As String = "2,2"
Dim sh As String, gsh As String

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX debug_macro:debug macros are for testing
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub debug_macro()
'add a new worksheet
Call add_sheet("graph")
'delete the worksheet. the default value is the graph sheet
'the option is the sheet name to delete
Call del_graph
'create data for graph
Call table_main
'create a graph using reference data
Call graph_main
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  graph_main:the main macro is for creating graphs                  XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub graph_main()
Dim cols As Variant, data As Variant, source As Variant
Dim gname As String
If sh = "" Then sh = "data"

'create a graph.source is the data sheet
'data array is option value '2'
source = add_data
Set cols = source(2)(0): Set data = source(2)(1)
Call graph_make(cols, data, 2)
'change column and row
Call graph_make(cols, data, 1)
'create a graph.source is the data sheet
'data array is option value '1'
Set cols = source(1)(0): Set data = source(1)(1)
Call graph_make(cols, data, 1)
'change row and column
Call graph_make(cols, data, 2)

'create a graph.source is the data sheet
'data array is option value '2'
source = add_data
Set cols = source(2)(0): Set data = source(2)(1)
gname = graph_make(cols, data, 2):
'add data of the 5th column to the created graph
Call graph_add(add_data(, "0,5"), gname)
'create a graph.source is the data sheet
'data array is option value '1'
Set cols = source(1)(0): Set data = source(1)(1)
gname = graph_make(cols, data, 1):
'add data of the 4th row to the created graph
Call graph_add(add_data(, "4,0"), gname)


Dim rw As Long, srw As Long, lrw As Long
Dim col As Long, scol As Long, lcol As Long
rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
srw = 10: scol = 7
With ThisWorkbook.Worksheets(sh)
  lrw = .Cells(rw, col).End(xlDown).Row
  lcol = .Cells(rw, col).End(xlToRight).Column
  
  'create a graph.data array is option value '2'
  '(1).specify data range
  source = add_data( _
    Union(.Range(.Cells(rw, col), .Cells(lrw, col)), _
    .Range(.Cells(rw, scol), .Cells(lrw, scol + 1))))
  Set cols = source(2)(0): Set data = source(2)(1)
  Call graph_make(cols, data, 2)
  'change column and row
  Call graph_make(cols, data, 1)
  
  'create a graph.data array is option value '1'
  '(1).specify data range
  source = add_data(Union(.Range(.Cells(rw, col), .Cells(rw, lcol)), _
    .Range(.Cells(srw, col), .Cells(lrw, lcol))))
  Set cols = source(1)(0): Set data = source(1)(1)
  Call graph_make(cols, data, 1)
  'change row and column
  Call graph_make(cols, data, 2)
  
  'create a graph.data array is option value '2'
  '(1).add the data of the 2th column
  '(2).specify data range
  lrw = 12: lcol = 10
  source = add_data( _
    Union(.Range(.Cells(rw, col), .Cells(rw, lcol)), _
    Union(.Range(.Cells(rw + 1, col), .Cells(rw + 1, lcol)), _
          .Range(.Cells(srw, col), .Cells(lrw, lcol)))))
  Set cols = source(2)(0): Set data = source(2)(1)
  gname = graph_make(cols, data, 2):
  Call graph_add(add_data(Union(cols, data), "0,2"), gname)
  'change column and row
  gname = graph_make(cols, data, 1):
  Call graph_add(add_data(Union(cols, data), "2,0"), gname)
  
End With
Set cols = Nothing: Set data = Nothing
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  add_data:create graph data                                        XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function add_data(Optional data As Variant, Optional src As Variant = "")
Dim r As Variant, rr As Variant
Dim c As Variant, cc As Variant
Dim cr As Variant, d As Variant
Dim i As Integer, ii As Integer, n As Integer
Dim rw As Long, col As Long

rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
With ThisWorkbook
  With .Worksheets(sh)
    d = data
    If IsObject(data) Then Set d = data
    Set cr = .Cells(rw, col).CurrentRegion
    
    If IsError(d) Or IsEmpty(d) Then
      Set d = Union(cr.Columns(1), cr.Columns(2), cr.Columns(3), cr.Columns(4))
    End If
    
    Set data = d
    If src <> "" Then n = 2 Else n = 1
    'c = "": r = "": cc = "": rr = ""
    For ii = 0 To data.Areas.count - 1
      Set d = data.Areas(ii + 1)
      For i = n To d.count
        If d.Item(i).Row = data.Row Then
          If Not IsObject(c) Then Set c = d.Item(i)
          Set c = Union(c, d.Item(i))
        Else
          If Not IsObject(cc) Then Set cc = d.Item(i)
          Set cc = Union(cc, d.Item(i))
        End If
        If d.Item(i).Column = data.Column Then
          If Not IsObject(r) Then Set r = d.Item(i)
          Set r = Union(r, d.Item(i))
        Else
          If Not IsObject(rr) Then Set rr = d.Item(i)
          Set rr = Union(rr, d.Item(i))
        End If
      Next
      n = 1
    Next
    If src <> "" Then
      src = Split(src, com): i = 0: ii = 0
      If IsNumeric(src(0)) Then i = CInt(src(0))
      If IsNumeric(src(1)) Then ii = CInt(src(1))
      If i > 0 Then _
        d = Array(c.Item(1).Offset(i - 1, -1), r, c.Offset(i - 1))
      If ii > 0 Then _
        d = Array(r.Item(1).Offset(-1, ii - 1), r, r.Offset(, ii - 1))
      If i = 0 Or ii = 0 Then add_data = d: GoTo goto1
    End If
    c = Array(c, cc): r = Array(r, rr)
    add_data = Array(, r, c)
goto1:
  End With
End With
Set d = Nothing: Set cr = Nothing: Set data = Nothing
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  graph_make:create a graph by specifying the data range            XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function graph_make(ByVal cols As Variant, data As Variant, _
    Optional plot As Integer = 2)
Dim value As String, gname As String
Dim s As Variant, w As Variant
Dim rw As Long, col As Long, srw As Long
Dim i As Integer, ii As Integer
rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
gsh = "graph"
With ThisWorkbook
  With .Worksheets(sh)
    If .Cells(rw, col) <> "" Then value = .Cells(rw, col): .Cells(rw, col) = ""
  End With
  w = 0
  For i = 0 To data.count
    If data.Item(i).Row = data.Row Then w = w + 1
    If data.Item(i).Column = data.Column Then w = w + 1
  Next
  w = w * 30
  If w > Application.UsableWidth Then
    w = Application.UsableWidth * 0.88
    'w = Application.UsableWidth * 0.75
  End If
  srw = row_graph
  With .Worksheets(gsh).ChartObjects.Add _
    (Cells(srw, col - 1).Left, Cells(srw, col - 1).Top, w, 300)
    gname = sh & .Index: .Name = gname
    With .Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, w - 50, 0, 50, 20)
      .TextFrame.Characters.Text = gname
      .TextFrame.VerticalAlignment = xlVAlignCenter
      .Name = "graph"
    End With
    With .Chart
      .ChartType = 51 'xlColumnClustered
      .SetSourceData source:=Union(cols, data), PlotBy:=plot 'xlrows=1, xlColumns=2
      .HasLegend = True
      .Legend.Position = -4107 'xlLegendPositionBottom
      .HasTitle = True
      .ChartTitle.Text = value
      .Axes(xlCategory).TickLabels.Orientation = 55
      For Each s In .SeriesCollection
        s.HasDataLabels = True
        s.DataLabels.Position = 2 'xlLabelPositionOutsideEnd
      Next
    End With
  End With
  .Worksheets(sh).Cells(rw, col) = value
End With
graph_make = gname
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  graph_add:add a graph by specifying the data range                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub graph_add(Optional data As Variant, Optional gname As String = "data1")
Dim rw As Long, col As Long
rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
With ThisWorkbook.Worksheets(gsh)
  With .ChartObjects(gname).Chart.SeriesCollection.NewSeries
    .ChartType = xlLine
    .Name = data(0)
    .XValues = data(1)
    .Values = data(2)
  End With
End With
Set g = Nothing
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  row_graph:gets the position of the graph                          XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function row_graph()
Dim rw As Integer: rw = 1
With ThisWorkbook.Worksheets(gsh)
  For Each g In .ChartObjects
    If g.BottomRightCell.Row > rw Then rw = g.BottomRightCell.Row + 1
  Next
End With
row_graph = rw
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  del_graph:delete a graph. the option is the sheet names           XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub del_graph(Optional gsh As String = "graph")
Dim g As Variant
With ThisWorkbook.Worksheets(gsh)
  For Each g In .ChartObjects: g.Delete: Next
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  table_main:the main macro is for creating sample data             XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub table_main()
If sh = "" Then sh = "data"
Dim data As String: data = sh
Dim graph As String: graph = "graph"
Call add_sheet(CVar(data))
Call table_make
Call table_line
'Call add_sheet(CVar(graph))
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  table_make:create sample data for testing                         XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub table_make()
Dim v As Variant, vv As Variant, cell As Variant
Dim i As Integer, ii As Integer
Dim rw As Long, col As Long
Dim lrw As Long, lcol As Long
rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
v = Split(get_variable, at): lrw = UBound(v) + rw
With ThisWorkbook.Worksheets(sh)
  .Cells.Clear
  For i = 0 To UBound(v)
    vv = Split(v(i), com)
    'For ii = 0 To UBound(vv)
    For ii = 0 To 0
      .Cells(i + rw, col + ii) = vv(ii)
    Next
  Next
  For i = 2020 To 2010 Step -1
    lcol = .Cells(rw, .Columns.count).End(xlToLeft).Column + 1
    .Cells(rw, lcol) = i
    For ii = rw + 1 To lrw
      .Cells(ii, lcol) = WorksheetFunction.RandBetween(10, 100)
    Next
  Next
  cell = Split(Cells(1, col + 1).Address, "$")(1)
  cell = cell & ":" & Split(Cells(1, col + 1 + 10).Address, "$")(1)
  .Columns(col).AutoFit
  .Columns(cell).ColumnWidth = 5
  .Columns(1).ColumnWidth = 3
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX table_line:add borders to the table                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub table_line()
Dim rw As Long, col As Long
Dim lrw As Long, lcol As Long
rw = CLng(Split(rc, com)(0)): col = CLng(Split(rc, com)(1))
With ThisWorkbook.Worksheets(sh)
  lrw = .Cells(Rows.count, col).End(xlUp).Row
  lcol = .Cells(rw, col).End(xlToRight).Column
  '.Cells.Borders.LineStyle = xlNone
  .Cells.Borders.LineStyle = False
  .Cells.Interior.Color = xlNone
  With .Range(.Cells(rw, col), .Cells(lrw, lcol))
    .Borders.LineStyle = xlDashDotDot
    .Borders.ColorIndex = xlAutomatic
    .Borders.Weight = xlMedium
    .Borders(xlInsideVertical).Weight = xlThick
    .BorderAround Weight:=xlThick
  End With
  With .Range(.Cells(rw, col), .Cells(rw, lcol))
    .Interior.Color = RGB(200, 240, 250)
    .BorderAround Weight:=xlThick
  End With
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  test                                                              XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub xvalu_click()
MsgBox 1
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  del_sheets:delete the sheet. option is the sheet name             XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub del_sheets(Optional sh As Variant = "data,data1")
Dim s As Variant, ss As String
Dim sheet1 As String: sheet1 = "Sheet1"
On Error Resume Next
Application.DisplayAlerts = False
With ThisWorkbook
  If sh <> "" Then
    For Each s In .Worksheets
      If InStr(com & sh & com, com & s.Name & com) Then s.Delete
    Next
  Else
    ss = get_sheet
    If InStr(com & ss & com, com & sheet1 & com) = 0 Then _
      .Worksheets.Add.Name = sheet1
    For Each s In .Worksheets
      If s.Name <> sheet1 Then s.Delete
    Next
  End If
End With
Application.DisplayAlerts = True
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  get_sheet:gata all sheet names                                    XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function get_sheet()
Dim s As Variant, sh As String
With ThisWorkbook
  For Each s In .Worksheets: sh = sh & s.Name & com: Next
End With
get_sheet = Mid(sh, 1, Len(sh) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  add_sheet:add a sheet. the option is the sheet name               XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub add_sheet(Optional sh As Variant = "data,data1,data2")
Dim s As Variant, i As Integer
Call del_sheets(sh)
sh = Split(sh, com)
With ThisWorkbook
  For i = UBound(sh) To 0 Step -1
    .Worksheets.Add.Name = sh(i)
  Next
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  get_variable:change variable value                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function get_variable()
Dim v As Variant
With ThisWorkbook.VBProject.VBComponents("abc_key").CodeModule
  'need to think
  v = get_abc
  v = Replace(Replace(Replace(v, """", ""), " ", ""), "&", "")
  v = Replace(Replace(v, vbLf, ""), vbCr, "")
  
  v = Replace(Replace(v, "com", com), "_", at)
  v = Mid(v, InStr(v, at) + 1, Len(v))
  v = "Fruit,Fruit Code" & at & v
End With
get_variable = v
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  get_abc:specify a variable to get the variable value              XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function get_abc(Optional str As String = "fruit")
Dim v As Variant, i As Integer
With ThisWorkbook.VBProject.VBComponents("abc_key").CodeModule
  v = .Lines(1, .CountOfLines - 1)
  v = Replace(Replace(v, vbLf, ""), vbCr, "")
  v = Split(v, ":")
  For i = 0 To UBound(v)
    If Split(v(i), " ")(1) = str Then
      get_abc = v(i): Exit For
    End If
  Next
End With
End Function
