Attribute VB_Name = "ShowHideRowsCols"
Sub ShowHideRows(control As IRibbonControl)
    Dim row As Range
    Dim cell As Range
    Dim containsNumbers As Boolean
    Dim lastRow As Integer
    Dim currSheet As Worksheet
    Dim indexSheet As Worksheet
    Dim indexName As String

    Application.ScreenUpdating = False

    ' set current worksheet object
    Set currSheet = Sheets(ActiveSheet.Name)

    ' create sheet for index of hidden rows or use existing one
    indexName = "_" + currSheet.Name
    CreateSheetIf (indexName)
    Set indexSheet = ActiveWorkbook.Worksheets(indexName)

    ' check if the rows have already been hidden
    If Not indexSheet.Cells(2, 1) Then
        ' set hidden flag on index sheet to true
        indexSheet.Cells(1, 1) = "hide rows"
        indexSheet.Cells(2, 1) = True
        ' clear contents of index
        indexSheet.Columns(2).ClearContents

        currSheet.Activate

        ' create array with length equal to used rows
        lastRow = 1 + currSheet.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
        ReDim hiddenRows(1 To lastRow) As Boolean

        ' evaluate which rows should be hidden
        For Each row In currSheet.UsedRange.Rows
            ' by default, assume row doesn't contain numbers and shouldn't be hidden
            hiddenRows(row.row) = False
            containsNumbers = False

            ' loop through cells in selected row
            For Each cell In row.Cells
                If (IsNumeric(cell) And Not IsEmpty(cell)) Then
                    If cell.Value <> 0 Then
                        GoTo NextRow
                    Else
                        containsNumbers = True
                    End If
                End If
            Next cell

            If containsNumbers Then
                hiddenRows(row.row) = True
                row.Hidden = True
            End If

NextRow:
        Next row

        ' write the array of hidden rows to the hidden sheet
        For i = LBound(hiddenRows) To UBound(hiddenRows)
            indexSheet.Cells(i, 2) = hiddenRows(i)
        Next i

    ' rows already hidden, unhide
    Else
        ' set hidden flag on index sheet to false
        indexSheet.Cells(2, 1) = False

        ' get last used row in index column
        lastRow = currSheet.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
        For i = 1 To lastRow
            If Not indexSheet.Cells(2, i) Then
                currSheet.Rows(i).Hidden = False
            End If
        Next
    End If

    ' return focus to original sheet
    currSheet.Activate
End Sub
Sub ShowHideCols(control As IRibbonControl)
    Dim col As Range
    Dim cell As Range
    Dim containsNumbers As Boolean
    Dim lastCol As Integer
    Dim currSheet As Worksheet
    Dim indexSheet As Worksheet
    Dim indexName As String

    Application.ScreenUpdating = False

    ' set current worksheet object
    Set currSheet = Sheets(ActiveSheet.Name)

    ' create sheet for index of hidden rows or use existing one
    indexName = "_" + currSheet.Name
    CreateSheetIf (indexName)
    Set indexSheet = ActiveWorkbook.Worksheets(indexName)

    ' check if the cols have already been hidden
    If Not indexSheet.Cells(2, 3) Then
        ' set hidden flag on index sheet to true
        indexSheet.Cells(1, 3) = "hide cols"
        indexSheet.Cells(2, 3) = True
        ' clear contents of index
        indexSheet.Columns(4).ClearContents

        currSheet.Activate

        ' create array with length equal to used cols
        lastCol = 1 + currSheet.Cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
        ReDim hiddenCols(1 To lastCol) As Boolean

        ' evaluate which cols should be hidden
        For Each col In currSheet.UsedRange.Columns
            ' by default, assume col doesn't contain numbers and shouldn't be hidden
            hiddenCols(col.Column) = False
            containsNumbers = False

            ' loop through cells in selected col
            For Each cell In col.Cells
                If (IsNumeric(cell) And Not IsEmpty(cell)) Then
                    If cell.Value <> 0 Then
                        GoTo NextCol
                    Else
                        containsNumbers = True
                    End If
                End If
            Next cell

            If containsNumbers Then
                hiddenCols(col.Column) = True
                col.Hidden = True
            End If

NextCol:
        Next col

        ' write the array of hidden rows to the hidden sheet
        For i = LBound(hiddenCols) To UBound(hiddenCols)
            indexSheet.Cells(i, 4) = hiddenCols(i)
        Next i

    ' cols already hidden, unhide
    Else
        ' set hidden flag on index sheet to false
        indexSheet.Cells(2, 3) = False

        ' get last used col in index column
        lastCol = currSheet.Cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
        For i = 1 To lastCol
            If Not indexSheet.Cells(4, i) Then
                currSheet.Columns(i).Hidden = False
            End If
        Next
    End If

    ' return focus to original sheet
    currSheet.Activate
End Sub
Function CreateSheetIf(strSheetName As String) As Boolean
    Dim wsTest As Worksheet
    CreateSheetIf = False

    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(strSheetName)
    On Error GoTo 0

    If wsTest Is Nothing Then
        CreateSheetIf = True
        Worksheets.Add.Name = strSheetName
        Sheets(strSheetName).Visible = xlSheetVeryHidden
    End If

End Function
