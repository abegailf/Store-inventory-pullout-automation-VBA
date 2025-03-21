Sub StorePullout()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, lastCol As Long, skuRow As Long, storeCol As Long
    Dim quantityNeeded As Long, pulledQuantity As Long, storeInventory As Variant
    Dim outputRow As Long, i As Long, j As Long
    Dim stores() As Variant, storeCount As Long, temp As Variant
    Dim sheetName As String

    ' Allow user to select the input sheet
    On Error Resume Next
    sheetName = Application.InputBox("Select the input sheet by clicking on any cell in the sheet:", "Select Input Sheet", Type:=8).Parent.Name
    On Error GoTo 0

    ' Check if the user canceled the input box
    If sheetName = "" Then
        MsgBox "No sheet selected. Macro canceled.", vbExclamation
        Exit Sub
    End If

    ' Set input sheet
    On Error Resume Next
    Set wsInput = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If wsInput Is Nothing Then
        MsgBox "The selected sheet was not found.", vbExclamation
        Exit Sub
    End If

    ' Prepare output sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("STORE PULLOUT OUTPUT").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.Name = "STORE PULLOUT OUTPUT"
    wsInput.Rows(1).Copy wsOutput.Rows(1)

    ' Find data bounds
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column
    outputRow = 2

    For skuRow = 2 To lastRow
        ' Copy row to output
        wsInput.Rows(skuRow).Copy wsOutput.Rows(outputRow)
        
        ' Initialize store columns to 0 in the output
        For storeCol = 7 To lastCol - 1
            wsOutput.Cells(outputRow, storeCol).Value = 0
        Next storeCol
        
        ' Get quantity needed (column AC)
        quantityNeeded = Val(wsInput.Cells(skuRow, lastCol).Value)
        If quantityNeeded <= 0 Then GoTo NextRow

        ' Collect store data (column index, inventory)
        ReDim stores(1 To lastCol - 6, 1 To 2)
        storeCount = 0
        For storeCol = 7 To lastCol - 1
            storeInventory = Val(wsInput.Cells(skuRow, storeCol).Value)
            If storeInventory > 0 Then
                storeCount = storeCount + 1
                stores(storeCount, 1) = storeCol   ' Store column index
                stores(storeCount, 2) = storeInventory ' Inventory
            End If
        Next storeCol

        ' Sort stores by inventory (descending)
        For i = 1 To storeCount - 1
            For j = i + 1 To storeCount
                If stores(i, 2) < stores(j, 2) Then
                    temp = stores(i, 1): stores(i, 1) = stores(j, 1): stores(j, 1) = temp
                    temp = stores(i, 2): stores(i, 2) = stores(j, 2): stores(j, 2) = temp
                End If
            Next j
        Next i

        ' Pull inventory in a round-robin fashion
        pulledQuantity = 0
        Do While pulledQuantity < quantityNeeded And storeCount > 0
            For i = 1 To storeCount
                If pulledQuantity >= quantityNeeded Then Exit Do
                If stores(i, 2) > 0 Then
                    ' Pull 1 unit from the store
                    wsOutput.Cells(outputRow, stores(i, 1)).Value = wsOutput.Cells(outputRow, stores(i, 1)).Value + 1
                    pulledQuantity = pulledQuantity + 1
                    stores(i, 2) = stores(i, 2) - 1
                Else
                    ' Remove exhausted stores
                    For j = i To storeCount - 1
                        stores(j, 1) = stores(j + 1, 1)
                        stores(j, 2) = stores(j + 1, 2)
                    Next j
                    storeCount = storeCount - 1
                End If
            Next i
        Loop

NextRow:
        outputRow = outputRow + 1
    Next skuRow

    ' Rotate text in columns G to AC (first row) 90 degrees upward
    With wsOutput.Range("G1:AC1")
        .Orientation = 90 ' Rotate text 90 degrees upward
        .VerticalAlignment = xlTop ' Align text to the top
    End With

    wsOutput.Columns.AutoFit
    MsgBox "Inventory pullout completed! Check the 'STORE PULLOUT OUTPUT' sheet.", vbInformation
End Sub
