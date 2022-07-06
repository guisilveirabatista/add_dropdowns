Public Const ROWS_LIMIT = 1000
Public Const DROPDOWN_LIST_LIMIT = 30
Public Const DROPDOWN_DATA_WORKSHEET_NAME = "Dropdown_Data"

Sub AddDropdowns()

Dim i As Long
Dim j As Long
Dim k As Long
Dim list As String
Dim columnToRead As Integer
Dim columnToApply As Integer

columnToApply = ActiveCell.Column
columnToRead = columnToApply - 1

'Iterate over the rows in the selected sheet and get the titles to be matched with the titles in Dropdown_Data'
For i = 1 To Rows.Count
    Dim title As String
    title = Cells(i, columnToRead).Value
    If Trim(title) <> "" Then
        'Read data list from the sheet Dropdown_Data for the given title'
        list = ReadDropdownDataList(title)
        If Trim(list) <> "" Then
            'If there is a match of titles, apply the dropdown in the cell in the selected column'
            Call AddDropdown(Cells(i, columnToApply), list)
        End If
    End If
If i = ROWS_LIMIT Then Exit For
Next i
End Sub

Function ReadDropdownDataList(title As String) As String
    Dim newList As String
    For j = 1 To Worksheets(DROPDOWN_DATA_WORKSHEET_NAME).Columns.Count
        If Trim(Worksheets(DROPDOWN_DATA_WORKSHEET_NAME).Cells(1, j).Value) <> "" Then
            If Worksheets(DROPDOWN_DATA_WORKSHEET_NAME).Cells(1, j).Value = title Then
                For k = 2 To Worksheets(DROPDOWN_DATA_WORKSHEET_NAME).Rows.Count
                    newList = newList & Worksheets(DROPDOWN_DATA_WORKSHEET_NAME).Cells(k, j).Value & ","
                If k = DROPDOWN_LIST_LIMIT Then Exit For
                Next k
            End If
        End If
    If j = ROWS_LIMIT Then Exit For
    Next j
    ReadDropdownDataList = newList
End Function

Function AddDropdown(cell, listToAppend)
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

End Function