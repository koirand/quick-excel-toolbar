Attribute VB_Name = "QXT"
Option Explicit

Sub Auto_Open()

    'disabele F1 key
    Application.OnKey "{F1}", "QXT_nothingDo"
    
    'reset cursol position and zoom
    Application.OnKey "^A", "QXT_resetCursolAndZoom"
    
    'insert row
    Application.OnKey "{INSERT}", "QXT_insert"

    'paste only value
    Application.OnKey "^V", "QXT_pasteValue"
    
    'create baloon
    Application.OnKey "^{^}", "QXT_createBaloonWithText"
    
    'reduction image
    Application.OnKey "^@", "QXT_sizeDownImage"
    
    'togle fullscreen
    Application.OnKey "{F11}", "QXT_fullscreen"
    
    'draw borders
    Application.OnKey "^k", "QXT_drawBorders"
    
End Sub

Sub QXT_nothingDo()

    'nothing do

End Sub

Sub QXT_pasteValue()

    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Application.EnableCancelKey = xlInterrupt

End Sub

Sub QXT_createBaloonWithText(Optional control As IRibbonControl = Nothing)
    
    Dim shape As shape
    Set shape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangularCallout, ActiveCell.Left, ActiveCell.Top, 160, 60)
    shape.TextFrame.Characters.Text = ActiveCell.Value
    shape.Select
    
End Sub

Sub QXT_sizeUpImage(Optional control As IRibbonControl = Nothing)

    On Error GoTo ErrorHandler
    
    Dim Pic As Picture
    
    With Selection.ShapeRange
    .LockAspectRatio = msoTrue
    .Height = .Height * 1.2
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Clear
    MsgBox "Please select a image."
    On Error GoTo 0

End Sub

Sub QXT_sizeDownImage(Optional control As IRibbonControl = Nothing)

    On Error GoTo ErrorHandler
    
    Dim Pic As Picture
    
    With Selection.ShapeRange
    .LockAspectRatio = msoTrue
    .Height = .Height * 0.8
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Clear
    MsgBox "Please select a image."
    On Error GoTo 0

End Sub


Sub QXT_insert()
    
    Application.SendKeys "+ "
    Application.SendKeys "+^{+}"
    
End Sub

Sub QXT_increment(Optional control As IRibbonControl = Nothing)

    If ActiveCell.Value = "" Then
        Exit Sub
        
    ElseIf IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = ActiveCell.Value + 1
    
    ElseIf IsDate(ActiveCell.Value) Then
        ActiveCell.Value = CDate(ActiveCell.Value) + 1
        
    End If
    
End Sub

Sub QXT_decrement(Optional control As IRibbonControl = Nothing)

    If ActiveCell.Value = "" Then
        Exit Sub
        
    ElseIf IsNumeric(ActiveCell.Value) Then
        ActiveCell.Value = ActiveCell.Value - 1
        
    ElseIf IsDate(ActiveCell.Value) Then
        ActiveCell.Value = CDate(ActiveCell.Value) - 1
            
    End If
    
End Sub

Sub QXT_fullScreen()

    If Application.DisplayFullScreen = True Then
        Application.DisplayFullScreen = False
    Else
        Application.DisplayFullScreen = True
    End If
       
End Sub

Sub QXT_changeCellFormat()

    If MsgBox( _
        "Do you want to change the cell format ?" & vbCrLf & _
        "Font Name : Meiryo UI" & vbCrLf & _
        "Font Size : 10pt" & vbCrLf & _
        "Number Format : @", _
        vbOKCancel) = vbOK Then

        ActiveSheet.Cells.Font.Name = "Meiryo UI"
        ActiveSheet.Cells.Font.Size = 10
        ActiveSheet.Cells.NumberFormat = "@"
        
    End If

End Sub

Sub QXT_drawBorders()

    Range(Selection.Address).Borders.LineStyle = xlContinuous
    
End Sub

Sub QXT_resetCursolAndZoom(Optional control As IRibbonControl = Nothing)
    Dim ws As Excel.Worksheet

    Application.ScreenUpdating = False

        On Error Resume Next
    
        For Each ws In ActiveWorkbook.Worksheets
            Application.GoTo ws.Cells(1, 1), True
            ActiveWindow.Zoom = 100
        Next

        ActiveWorkbook.Sheets(1).Select
    Application.ScreenUpdating = True

End Sub

Sub QXT_styleDefault(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(0, 0, 0), RGB(255, 255, 255), RGB(216, 216, 216))

End Sub

Sub QXT_stylePrimary(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(255, 255, 255), RGB(66, 139, 202), RGB(31, 113, 186))

End Sub

Sub QXT_styleSuccessDark(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(255, 255, 255), RGB(92, 184, 92), RGB(42, 164, 41))

End Sub

Sub QXT_styleInfoDark(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(255, 255, 255), RGB(91, 192, 222), RGB(35, 173, 214))

End Sub

Sub QXT_styleWarningDark(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(255, 255, 255), RGB(240, 173, 78), RGB(234, 144, 5))

End Sub

Sub QXT_styleDangerDark(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(255, 255, 255), RGB(217, 83, 79), RGB(210, 41, 33))

End Sub

Sub QXT_styleSuccessLight(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(61, 117, 62), RGB(223, 240, 216), RGB(61, 117, 62))

End Sub

Sub QXT_styleInfoLight(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(50, 111, 143), RGB(217, 237, 247), RGB(50, 111, 143))

End Sub

Sub QXT_styleWarningLight(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(138, 109, 59), RGB(252, 248, 227), RGB(138, 109, 59))

End Sub

Sub QXT_styleDangerLight(Optional control As IRibbonControl = Nothing)

    Call QXT_SetStyle(RGB(169, 68, 66), RGB(242, 222, 222), RGB(169, 68, 66))

End Sub

Sub QXT_SetStyle(ByVal fontColor As Long, ByVal backColor As Long, ByVal borderColor As Long)

    On Error GoTo ErrorHandler
      
    If TypeName(Selection) = "Range" Then
    
        If fontColor = RGB(0, 0, 0) Then
            Selection.Font.ColorIndex = xlAutomatic
        
        Else
            Selection.Font.Color = fontColor
        
        End If
        
        If backColor = RGB(255, 255, 255) Then
            Selection.Interior.Pattern = xlNone
        
        Else
            Selection.Interior.Color = backColor
        
        End If
        
    ElseIf TypeName(Selection) = "Picture" Then
        Selection.ShapeRange.Line.ForeColor.RGB = borderColor
        
    Else
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontColor
        Selection.ShapeRange.Fill.ForeColor.RGB = backColor
        Selection.ShapeRange.Line.ForeColor.RGB = borderColor
    
    End If
    
ErrorHandler:
    
End Sub

Sub QXT_setSameColumnWidth(Optional control As IRibbonControl = Nothing)

    If TypeName(Selection) = "Range" Then
        Selection.ColumnWidth = ActiveCell.ColumnWidth
    End If

End Sub

Sub QXT_setSameRowHeight(Optional control As IRibbonControl = Nothing)

    If TypeName(Selection) = "Range" Then
        Selection.RowHeight = ActiveCell.RowHeight
    End If

End Sub

Sub QXT_ShrinkToFit(Optional control As IRibbonControl = Nothing)
    
    If TypeName(Selection) = "Range" Then
        If Selection.ShrinkToFit = True Then
            Selection.ShrinkToFit = False
            
        Else
            Selection.ShrinkToFit = True
            
        End If
    End If

End Sub

Sub QXT_help(Optional control As IRibbonControl = Nothing)
    
    Dim Shell
    Set Shell = CreateObject("Wscript.Shell")
    Shell.Run "https://github.com/koirand/QuickExcelToolbar/blob/master/README.md", 3
    Set Shell = Nothing

End Sub


