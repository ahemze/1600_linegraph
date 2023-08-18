Attribute VB_Name = "UKG_CreateLineGraph"
Sub UKG_CreateLineGraph()

    ' Coded by Ashley Hemze, Store 1600
    ' August 2023
    
    ' Change district number and store number to your store
    ' under Private Subs "CutJobDesc_CashierSCOTCSMOffice" and "CutJobDesc_BookkeepingSvcCtr"

    Call DelExpMDHeaderWithTR
    Call DelNonFEDepts
    Call CutJobDesc_CashierSCOTCSMOffice
    Call CutJobDesc_BookkeepingSvcCtr
    Call ConvertMilitaryTime
    Call SortJobs
    Call DeleteOvernight
    Call AddHeader
    Call AddDuration
    Call AddBorder
    Call ResizeColumnsRows
    Call AlignText
    Call AddDateTitle
    Call FormatText
    Call PrintFormat

End Sub
Private Sub DelExpMDHeaderWithTR()
    
    Rows("1:10").Select
    Selection.Delete Shift:=xlUp
    
End Sub
Private Sub DelNonFEDepts()

    Dim cell As Range, cRange As Range, lastRow As Long, x As Long
    lastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    Set cRange = Range("B1:B" & lastRow)

    For x = cRange.Cells.Count To 1 Step -1
        With cRange.Cells(x)
                If .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/Front End/Front End/Cashier" And .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/General/General/Bookkeeping" And .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/Front End/Front End/CSM" And .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/Front End/Front End/CSM Office" And .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/Front End/Front End/SCOT" And .Value <> "Cub/District 1/1600-Maple Grove/Non Sales/General/General/Service Center" Then
                .EntireRow.Delete
            End If
        End With
    Next x

End Sub
Private Sub CutJobDesc_CashierSCOTCSMOffice()

    Dim rng As Range
    Set rng = Range("B1:B999")

    For Each cell In rng ' Change district and store right below
        cell.Value = Replace(cell.Value, "Cub/District 1/1600-Maple Grove/Non Sales/Front End/Front End/", "")
    Next cell

End Sub
Private Sub CutJobDesc_BookkeepingSvcCtr()

    Dim rng As Range
    Set rng = Range("B1:B999")

    For Each cell In rng ' Change district and store right below
        cell.Value = Replace(cell.Value, "Cub/District 1/1600-Maple Grove/Non Sales/General/General/", "")
    Next cell

End Sub
Private Sub ConvertMilitaryTime()

    Columns(4).NumberFormat = "h:mm AM/PM"
    Columns(5).NumberFormat = "h:mm AM/PM"

End Sub
Private Sub SortJobs()

    Columns("A:E").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B1:B105"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        CustomOrder:="Bookkeeping,Cashier,SCOT,CSM,CSM Office,Service Center", DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:E105")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
Private Sub DeleteOvernight()
    
    Dim lastRowWithData As Long
    
    lastRowWithData = ActiveSheet.Cells(ActiveSheet.Rows.Count, "D").End(xlUp).Row
    
    If ActiveSheet.Cells(2, "D").Value <> ActiveSheet.Cells(lastRowWithData, "D").Value Then
        ActiveSheet.Rows(2).Delete Shift:=xlUp
    End If
    
End Sub
Private Sub AddHeader()
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("A1").Value = "Associate"
    Range("B1").Value = "Job"
    Range("C1").Value = "Reg #"
    Range("D1").Value = "Start Time"
    Range("E1").Value = "End Time"
    Range("F1").Value = "Break 1"
    Range("G1").Value = "Break 2"
    Range("H1").Value = "Hours"
    Range("I1").Value = "Comment"
    
End Sub
Private Sub AddDuration()

    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "D").End(xlUp).Row
    
    For i = 2 To lastRow
        If ActiveSheet.Cells(i, 4).Value <> "" And ActiveSheet.Cells(i, 5) <> "" Then
            Dim durationMinutes As Double
            durationMinutes = (ActiveSheet.Cells(i, 5).Value - ActiveSheet.Cells(i, 4).Value) * 1440
            ActiveSheet.Cells(i, 8).Value = Format(durationMinutes / 60, "0.00")
            
            ActiveSheet.Cells(i, 8).NumberFormat = "0.00"
        End If
    Next i

End Sub
Private Sub AddBorder()

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim cell As Range
    
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    Set dataRange = ActiveSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(lastRow, lastCol))
    
    dataRange.BorderAround xlContinuous, xlThin
    dataRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    dataRange.Borders(xlInsideVertical).Weight = xlThin
    dataRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    dataRange.Borders(xlInsideHorizontal).Weight = xlThin

End Sub
Private Sub ResizeColumnsRows()

    Range("A1").EntireColumn.ColumnWidth = 27
    Range("B1").EntireColumn.ColumnWidth = 16.5
    Range("C1").EntireColumn.ColumnWidth = 5
    Range("D1").EntireColumn.ColumnWidth = 9
    Range("E1").EntireColumn.ColumnWidth = 9
    Range("F1").EntireColumn.ColumnWidth = 9
    Range("G1").EntireColumn.ColumnWidth = 9
    Range("H1").EntireColumn.ColumnWidth = 6
    Range("I1").EntireColumn.ColumnWidth = 33.5
    
    Rows("1:500").RowHeight = 14

End Sub
Private Sub AlignText()
    
    Columns("A:C").HorizontalAlignment = xlLeft
    Columns("D:E").HorizontalAlignment = xlRight
    Columns("H:H").HorizontalAlignment = xlCenter
    
    Rows(1).HorizontalAlignment = xlCenter

End Sub
Private Sub AddDateTitle()
    
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "D").End(xlUp).Row
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    If lastRow >= 2 Then
        ActiveSheet.Cells(1, 1).Value = "Line Schedule Report for " & Format(ActiveSheet.Cells(lastRow, "D").Value, "M/DD/YYYY")
    End If

End Sub
Private Sub FormatText()
    
    Rows(1).Font.Size = 16
    Rows("3:500").Font.Size = 10

End Sub
Private Sub PrintFormat()

    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.37)
        .RightMargin = Application.InchesToPoints(0.37)
        .TopMargin = Application.InchesToPoints(0.37)
        .BottomMargin = Application.InchesToPoints(0.37)
        .Orientation = xlLandscape
    End With

End Sub
