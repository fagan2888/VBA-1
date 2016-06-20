'Add a prefix to all cells in column A

Private Sub GenerateURL()

 Dim prefix, addr As String
 prefix = "http://go/cr/"
 Dim ws As Worksheet
   For Each ws In ThisWorkbook.Worksheets
   Dim c As Range
        For Each c In ws.UsedRange.Columns("A").Cells
            c.Value = WorksheetFunction.Trim(c.Value)
            addr = prefix + CStr(c.Value)
            c.Hyperlinks.Add Anchor:=c, _
                Address:=addr, _
                TextToDisplay:=c.Value
        Next c
   Next ws
End Sub

' Change the row color based on the value of cell 
Private Sub Workbook_Open()
 Dim prefix, addr As String
 prefix = "http://go/cr/"
 Dim ws As Worksheet
   For Each ws In ThisWorkbook.Worksheets
        Dim c As Range
            For Each c In ws.UsedRange.Columns("A").Cells
                If c.Row > 1 Then
                    c.Value = WorksheetFunction.Trim(c.Value)
                    If Not IsEmpty(c.Value) Then
                        addr = prefix + CStr(c.Value)
                        c.Hyperlinks.Add Anchor:=c, _
                         Address:=addr, _
                         TextToDisplay:=addr
                    End If
                End If
        Next c
         
        Dim d As Range
            For Each d In ws.UsedRange.Columns("B").Cells
                d.Value = WorksheetFunction.Trim(d.Value)
                If d.Value = "Done" Then
                    ws.UsedRange.Rows(d.Row).Interior.ColorIndex = 4 '4 indicates green
                ElseIf d.Value = "Investigating" Then
                    ws.UsedRange.Rows(d.Row).Interior.ColorIndex = 7
                ElseIf InStr(d.Value, "Waiting") >= 1 Then
                    ws.UsedRange.Rows(d.Row).Interior.ColorIndex = 6
                ElseIf InStr(d.Value, "Working") >= 1 Then
                    ws.UsedRange.Rows(d.Row).Interior.ColorIndex = 28
                End If
                            
        Next d
   Next ws
End Sub