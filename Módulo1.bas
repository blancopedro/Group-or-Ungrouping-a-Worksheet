Attribute VB_Name = "Módulo1"
Sub Ungroup()
'   Sheets.Select
'   The main purpose os this code is to ungroup the whol Worksheet.
    Dim ws As Worksheet
    Dim CurSheet As Worksheet
        Set CurSheet = ActiveSheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Activate
            ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
        Next ws
        CurSheet.Activate
End Sub


Sub Group()
'   Sheets.Select
'   The main purpose os this code is to group the whol Worksheet.
    
    Dim ws As Worksheet
    Dim CurSheet As Worksheet
        Set CurSheet = ActiveSheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Activate
            ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

        Next ws
        CurSheet.Activate
End Sub

