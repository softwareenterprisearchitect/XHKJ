Sub ImageResize()
'
' ImageResize Macro
'
'
Dim i As Long
With ActiveDocument
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
            .ScaleHeight = 100
            .ScaleWidth = 100
        End With
    Next i
End With
End Sub
Sub AutoFitAllTableMacro()
'
' AutoFitAllTableMacro Macro
'
'

  Dim tbl As Table
  For Each tbl In ActiveDocument.Tables
    tbl.AutoFitBehavior wdAutoFitWindow
  Next tbl

End Sub
Sub AllBordersMacro()
'
' AllBordersMacro Macro
'
'
Dim Title As String
    Dim Msg As String
    Dim Style As VbMsgBoxStyle
    Dim Response As VbMsgBoxResult
    
    Dim oTable As Table
    Dim oBorderStyle As WdLineStyle
    Dim oBorderWidth As WdLineWidth
    Dim oBorderColor As WdColor
    Dim oarray As Variant
    
    Dim n As Long
    Dim i As Long
    
    '=========================
    'Change the values below to the desired style, width and color
    oBorderStyle = wdLineStyleSingle
    oBorderWidth = wdLineWidth050pt
    oBorderColor = wdColorBlack
    '=========================
    
    Title = "Apply Uniform Borders to All Tables"
    
    If ActiveDocument.Tables.Count > 0 Then
        Msg = "This command applies uniform table borders " & _
                "to all tables in the active document." & vbCr & vbCr & _
                "Do you want to continue?"
        Style = vbYesNo + vbQuestion
        Response = MsgBox(Msg, Style, Title)
        'Stop if user did not click Yes
        If Response <> vbYes Then Exit Sub
    Else
        'Stop - no tables are found
        MsgBox "The document contains no tables.", vbInformation, Title
        Exit Sub
    End If
        
    'Define array with the borders to be changed
    'Diagonal borders not included here
    oarray = Array(wdBorderTop, _
        wdBorderLeft, _
        wdBorderBottom, _
        wdBorderRight, _
        wdBorderHorizontal, _
        wdBorderVertical)
        
    For Each oTable In ActiveDocument.Tables
        'Count tables - used in message
        n = n + 1
        With oTable
            For i = LBound(oarray) To UBound(oarray)
                
                'Skip if only one row and wdBorderHorizontal
                If .Rows.Count = 1 And oarray(i) = wdBorderHorizontal Then GoTo Skip
                'Skip if only one column and wdBorderVertical
                If .Columns.Count = 1 And oarray(i) = wdBorderVertical Then GoTo Skip
                
                With .Borders(oarray(i))
                    .LineStyle = oBorderStyle
                    .LineWidth = oBorderWidth
                    .Color = oBorderColor
                End With
Skip:
            Next i
        End With
    Next oTable
    
    MsgBox "Finished applying borders to " & n & " tables.", vbOKOnly, Title
End Sub
Sub FIllPageColors()
'
' FIllPageColors Macro
'
'
ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(255, 255, 204)
    ActiveDocument.Background.Fill.Transparency = 0#
    ActiveDocument.Background.Fill.PresetTextured msoTextureParchment
End Sub
