Attribute VB_Name = "Module1"
Public Const ITEMS_LIMIT = 1000
Public Const DROPDOWN_ROWS_LIMIT = 30

Sub AddDropdowns()
Attribute AddDropdowns.VB_Description = "Add dropdown to the current column based on the information on the column to the left."
Attribute AddDropdowns.VB_ProcData.VB_Invoke_Func = "K\n14"

Dim i As Long
Dim j As Long
Dim k As Long
Dim list As String
Dim columnToRead As Integer
Dim columnToApply As Integer

columnToApply = ActiveCell.Column
columnToRead = columnToApply - 1

For i = 1 To Rows.Count
    Dim title As String
    title = Cells(i, columnToRead).Value
    If Trim(title) <> "" Then
        list = readList(title)
        If Trim(list) <> "" Then
            Call ApplyFilters(i, columnToApply, list)
        End If
    End If
If i = ITEMS_LIMIT Then Exit For
Next i
End Sub

Function readList(title As String) As String
    Dim newList As String
    For j = 1 To Worksheets("Dropdown_Data").Columns.Count
        If Trim(Worksheets("Dropdown_Data").Cells(1, j).Value) <> "" Then
            If Worksheets("Dropdown_Data").Cells(1, j).Value = title Then
                For k = 2 To Worksheets("Dropdown_Data").Rows.Count
                    newList = newList & Worksheets("Dropdown_Data").Cells(k, j).Value & ","
                If k = DROPDOWN_ROWS_LIMIT Then Exit For
                Next k
            End If
        End If
    Next j
    readList = newList
End Function

Sub ApplyFilters(rowNumber, columnToApply, listToAppend)
    Call addDropdown(Cells(rowNumber, columnToApply), listToAppend)
End Sub

Sub addDropdown(cell, listToAppend)
With cell.Validation
.Delete
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
xlBetween, Formula1:=listToAppend
.IgnoreBlank = True
.InCellDropdown = True
.InputTitle = ""
.ErrorTitle = "Invalid Input"
.InputMessage = ""
.ErrorMessage = "Please, select a valid item from the list."
.ShowInput = True
.ShowError = True
End With

cell.Interior.Color = RGB(214, 239, 237)
cell.Interior.Color = RGB(214, 239, 237)
If Trim(cell.Value) = "" Then
    cell.Value = "-select-"
End If

End Sub

