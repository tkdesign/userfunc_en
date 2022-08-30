Attribute VB_Name = "MainMod"
Option Explicit

Public Separator As String
Public WithoutRepeat As Integer
Public ComparedDataType As Integer
Public MergeCellsSeparator As String
Public CopyFormulaSeparator As String
Public RegExpPattern As String
Public ReplacementTemplate As String
Public RegExpPattern2 As String
Public RegExpMatchNumber As Integer
Public RoundPrecision As Integer
Public CellsAddressSeparator As String

Public BackupData As Object
Public RepeatData As Object

' ----------------------------------------------------------------
' Procedure Name: FillCells
' Purpose: Filling selected cells with a selected color from the preset
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Parameter selectedId (String):
' Parameter selectedIndex (Integer):
' Author: Petr Kovalenko
' Date: 08.10.2020
' ----------------------------------------------------------------
Sub FillCells(control As IRibbonControl)
    On Error GoTo FillCells_Error
    Dim i As Range
    Dim TargetRange As Range
    Dim FillColor
    If Selection.Count = 1 Then Set TargetRange = Selection Else Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    If control.Tag = "Fill color 1" Then
        FillColor = RGB(255, 255, 0)
    ElseIf control.Tag = "Fill color 2" Then
        FillColor = RGB(255, 192, 0)
    ElseIf control.Tag = "Fill color 3" Then
        FillColor = RGB(146, 208, 80)
    ElseIf control.Tag = "Fill color 4" Then
        FillColor = RGB(0, 176, 80)
    ElseIf control.Tag = "Fill color 5" Then
        FillColor = RGB(0, 176, 240)
    ElseIf control.Tag = "Fill color 6" Then
        FillColor = RGB(255, 0, 0)
    ElseIf control.Tag = "Fill color 7" Then
        FillColor = RGB(192, 0, 0)
    ElseIf control.Tag = "Fill color 8" Then
        FillColor = RGB(112, 48, 160)
    ElseIf control.Tag = "Fill color 9" Then
        FillColor = xlNone
    End If
    Set BackupData = CreateObject("Scripting.Dictionary")
    Set RepeatData = CreateObject("Scripting.Dictionary")
    For Each i In TargetRange
        If i.Interior.ColorIndex = xlNone Then
        BackupData.Add i.Address(True, True, xlA1, False, False), xlNone
        Else
        BackupData.Add i.Address(True, True, xlA1, False, False), i.Interior.Color
        End If
        RepeatData.Add i.Address(True, True, xlA1, False, False), FillColor
        If FillColor <> xlNone Then
            i.Interior.Color = FillColor
        Else
            i.Interior.ColorIndex = xlNone
        End If
    Next
    Application.OnUndo "Cancel fill cells", "FillCells_Undo"
    Application.OnRepeat "Repeate fill cells", "FillCells_Repeat"
    On Error GoTo 0
    Exit Sub
FillCells_Error:
    Set BackupData = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FillCells, line " & Erl & "."
End Sub

Sub FillCells_Undo()
    On Error GoTo FillCells_Undo_Error
    Dim Key
    Dim ColorCode
    Dim a As Collection
    If BackupData Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    If BackupData.Count = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    For Each Key In BackupData.Keys()
        If BackupData.Item(Key) <> xlNone Then
            Range(Key).Interior.Color = BackupData.Item(Key)
        Else
            Range(Key).Interior.ColorIndex = xlNone
        End If
    Next
    On Error GoTo 0
    Exit Sub
FillCells_Undo_Error:
    Set BackupData = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FillCells_Undo, line " & Erl & "."
End Sub

Sub FillCells_Repeat()
    On Error GoTo FillCells_Repeat_Error
    Dim i As Range
    Dim TargetRange As Range
    If RepeatData Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    If RepeatData.Count = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    Dim FillColor
    FillColor = RepeatData.Items()(0)
    For Each i In TargetRange
        If FillColor <> xlNone Then
            i.Interior.Color = FillColor
        Else
            i.Interior.ColorIndex = xlNone
        End If
    Next
    On Error GoTo 0
    Exit Sub
FillCells_Repeat_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FillCells_Repeat, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: CellColor
' Purpose: The display in the pop-up window of the font color and filling the selected cell. In case of cells range, the action is repeated sequentially for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub CellColor(control As IRibbonControl)
    On Error GoTo CellColor_Error
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        MsgBox ("Background color - " & i.Interior.Color & vbCrLf & "Background color index - " & i.Interior.ColorIndex & vbCrLf & "Text color - " & i.Font.Color & vbCrLf & "Text color index - " & i.Font.ColorIndex)
    Next
    On Error GoTo 0
    Exit Sub
CellColor_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine CellColor, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ConditionalFormattingColor
' Purpose: Displays in a pop-up window text and background color from conditional formatting rule applied to the cell. In case of range, this action is repeated sequentially for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ConditionalFormattingColor(control As IRibbonControl)
    On Error GoTo ConditionalFormattingColor_Error
    Dim i As Range, n As Variant, Text As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If (i.FormatConditions.Count > 0) Then
            Text = ""
            For n = 1 To i.FormatConditions.Count Step 1
                If Text <> "" Then
                    Text = Text & vbCrLf & n & ". Background color - " & i.FormatConditions(n).Interior.Color & vbCrLf & "Background color index - " & i.FormatConditions(n).Interior.ColorIndex & vbCrLf & "Text color - " & i.FormatConditions(n).Font.Color & vbCrLf & "Text color index - " & i.FormatConditions(n).Font.ColorIndex
                Else
                    Text = "Background color - " & i.FormatConditions(n).Interior.Color & vbCrLf & "Background color index - " & i.FormatConditions(n).Interior.ColorIndex & vbCrLf & "Text color - " & i.FormatConditions(n).Font.Color & vbCrLf & "Text color index - " & i.FormatConditions(n).Font.ColorIndex
                End If
            Next
            MsgBox (Text)
        Else
            MsgBox ("The rules of conditional formatting for the cell are not specified")
        End If
    Next
    On Error GoTo 0
    Exit Sub
ConditionalFormattingColor_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ConditionalFormattingColor, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: FillColor
' Purpose: Getting background color of the specified cell. Incase of cells range, the first cell color is returned
' Procedure Kind: Function
' Procedure Access: Public
' Parameter CheckedCells (Range): The range of checked cells
' Return Type: Double
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function FillColor(CheckedCells As Range) As Double
    On Error GoTo FillColor_Error
    Application.Volatile True
    FillColor = CheckedCells.Interior.Color
    On Error GoTo 0
    Exit Function
FillColor_Error:
    FillColor = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: SumColor
' Purpose: Calculate the sum of cells with their filtering according to background color
' Procedure Kind: Function
' Procedure Access: Public
' Parameter SumRange (Range): The cell range for sum calculation
' Parameter ColorSample (): The color code for filtering cells
' Return Type: Double
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function SumColor(SumRange As Range, Optional ColorSample) As Double
    On Error GoTo FillColor_Error
    Dim Sum As Double
    Dim i As Range
    Dim SumCondition As Variant
    Application.Volatile True
    Dim TargetRange As Range
    If SumRange.Count = 1 Then Set TargetRange = SumRange Else Set TargetRange = SumRange.SpecialCells(xlCellTypeVisible)
    If IsMissing(ColorSample) Then SumCondition = 65535 Else SumCondition = ColorSample.Interior.Color
    For Each i In TargetRange
        If i.Interior.Color = SumCondition Then
            i.Value = i.Value * (-1) * (-1)
            Sum = Sum + i.Value
        End If
    Next i
    SumColor = Sum
    On Error GoTo 0
    Exit Function
FillColor_Error:
    SumColor = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: CountColor
' Purpose: Calculate the count of cells with their filtering according to background color
' Procedure Kind: Function
' Procedure Access: Public
' Parameter SumRange (Range): The cell range for count calculation
' Parameter ColorSample (Range): The color code for filtering cells
' Return Type: Double
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function CountColor(SumRange As Range, ColorSample As Range) As Double
    On Error GoTo CountColor_Error
    Dim Sum As Double
    Dim i As Range
    Application.Volatile True
    Sum = 0
    Dim TargetRange As Range
    If SumRange.Count = 1 Then
        Set TargetRange = SumRange
    Else
        Set TargetRange = SumRange.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.Interior.Color = ColorSample.Interior.Color Then
            Sum = Sum + 1
        End If
    Next i
    CountColor = Sum
    On Error GoTo 0
    Exit Function
CountColor_Error:
    CountColor = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: SumBoldFont
' Purpose: Calculate the sum of cells with their filtering according to font style
' Procedure Kind: Function
' Procedure Access: Public
' Parameter SumRange (Range): The cell range for sum calculation
' Parameter IsBold (Boolean): Summation condition: True - Summage cells with bold font, False - Summage all cells except by bold font
' Return Type: Double
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function SumBoldFont(SumRange As Range, IsBold As Boolean) As Double
    On Error GoTo SumBoldFont_Error
    Dim Sum As Double
    Dim i As Range
    Application.Volatile True
    Dim TargetRange As Range
    If SumRange.Count = 1 Then
        Set TargetRange = SumRange
    Else
        Set TargetRange = SumRange.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.Font.Bold = IsBold Then
            Sum = Sum + i.Value * (-1) * (-1)
        End If
    Next i
    SumBoldFont = Sum
    On Error GoTo 0
    Exit Function
SumBoldFont_Error:
    SumBoldFont = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: DevideValueBy10
' Purpose: Deviding cell value by 10. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DevideValueBy10(control As IRibbonControl)
    On Error GoTo DevideValueBy10_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.Value = i.Value * (-1) * (-1)
        i.Value = i.Value / 10#
        i.NumberFormatLocal = "General"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DevideValueBy10_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DevideValueBy10, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DevideValueBy100
' Purpose: Deviding cell value by 100. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DevideValueBy100(control As IRibbonControl)
    On Error GoTo DevideValueBy100_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.Value = i.Value * (-1) * (-1)
        i.Value = i.Value / 100#
        i.NumberFormatLocal = "General"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DevideValueBy100_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DevideValueBy100, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DevideValueBy1000
' Purpose: Deviding cell value by 1000. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DevideValueBy1000(control As IRibbonControl)
    On Error GoTo DevideValueBy1000_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.Value = i.Value * (-1) * (-1)
        i.Value = i.Value / 1000#
        i.NumberFormatLocal = "General"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DevideValueBy1000_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DevideValueBy1000, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DivideBy10AsFormula
' Purpose: Deviding value from selected cell by 10 with save the result as a formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DivideBy10AsFormula(control As IRibbonControl)
    On Error GoTo DivideBy10AsFormula_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.NumberFormatLocal = "General"
        i.Value = i.Value * (-1) * (-1)
        i.FormulaLocal = "=" & i.Value & "/10"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DivideBy10AsFormula_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DivideBy10AsFormula, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DivideBy100AsFormula
' Purpose: Deviding value from selected cell by 100 with save the result as a formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DivideBy100AsFormula(control As IRibbonControl)
    On Error GoTo DivideBy100AsFormula_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.NumberFormatLocal = "General"
        i.Value = i.Value * (-1) * (-1)
        i.FormulaLocal = "=" & i.Value & "/100"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DivideBy100AsFormula_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DivideBy100AsFormula, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DivideBy1000AsFormula
' Purpose: Deviding value from selected cell by 1000 with save the result as a formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DivideBy1000AsFormula(control As IRibbonControl)
    On Error GoTo DivideBy1000AsFormula_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.NumberFormatLocal = "General"
        i.Value = i.Value * (-1) * (-1)
        i.FormulaLocal = "=" & i.Value & "/1000"
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DivideBy1000AsFormula_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DivideBy1000AsFormula, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DelLastDivisor
' Purpose: Removing the last divider from the formula in the cell. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DelLastDivisor(control As IRibbonControl)
    On Error GoTo DelLastDivisor_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    myRegExp.Pattern = "(.*)(/\d+?)$"
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            Set myMatches = myRegExp.Execute(i.FormulaLocal)
            If myMatches.Count >= 1 Then
                Set m = myMatches.Item(0)
                If m.SubMatches.Count = 2 Then
                    If (m.SubMatches(0) <> "") Then
                        ResultString = myRegExp.Replace(i.FormulaLocal, "$1")
                        i.FormulaLocal = ResultString
                    End If
                End If
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DelLastDivisor_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DelLastDivisor, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DelLastMultiplier
' Purpose: Removing the last multiplier from the formula in the cell. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DelLastMultiplier(control As IRibbonControl)
    On Error GoTo DelLastMultiplier_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    myRegExp.Pattern = "(.*)(\*\d+?)$"
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            Set myMatches = myRegExp.Execute(i.FormulaLocal)
            If myMatches.Count >= 1 Then
                Set m = myMatches.Item(0)
                If m.SubMatches.Count = 2 Then
                    If (m.SubMatches(0) <> "") Then
                        ResultString = myRegExp.Replace(i.FormulaLocal, "$1")
                        i.FormulaLocal = ResultString
                    End If
                End If
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DelLastMultiplier_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DelLastMultiplier, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ReplaceWithRegExp
' Purpose: Text replace with regular expression
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ReplaceWithRegExp(control As IRibbonControl)
    On Error GoTo ReplaceWithRegExp_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    Dim vRetVal
    Dim vRetVal2
    vRetVal = InputBox("Enter regular expression:", "RegExp template", RegExpPattern)
    If StrPtr(vRetVal) = 0 Then 'The Cancel button is pressed
        Exit Sub
    End If
    RegExpPattern = vRetVal
    vRetVal2 = InputBox("Enter replacement pattern", "Replacement pattern", ReplacementTemplate)
    If StrPtr(vRetVal2) = 0 Then 'The Cancel button is pressed
        Exit Sub
    End If
    ReplacementTemplate = vRetVal2
    vRetVal2 = Replace(vRetVal2, "\n", vbCrLf)
    vRetVal2 = Replace(vRetVal2, "\t", vbTab)
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    myRegExp.Pattern = vRetVal
    For Each i In TargetRange
        If i.HasFormula <> True Then
            ResultString = myRegExp.Replace(i.FormulaLocal, vRetVal2)
            i.FormulaLocal = ResultString
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
ReplaceWithRegExp_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ReplaceWithRegExp, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ExtractWithRegExp
' Purpose: Text extract using regular expression
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ExtractWithRegExp(control As IRibbonControl)
    On Error GoTo ExtractWithRegExp_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    Dim vRetVal
    Dim vRetVal2
    vRetVal = InputBox("Enter regular expression:", "Extract template", RegExpPattern2)
    If StrPtr(vRetVal) = 0 Then 'The Cancel button is pressed
        Exit Sub
    End If
    RegExpPattern2 = vRetVal
    vRetVal2 = InputBox("Enter matching index:", "Matching index", CStr(RegExpMatchNumber))
    If StrPtr(vRetVal2) = 0 Then 'The Cancel button is pressed
        Exit Sub
    End If
    If CInt(vRetVal2) < 1 Then vRetVal2 = "1"
    RegExpMatchNumber = CInt(vRetVal2)
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    myRegExp.Pattern = vRetVal
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If myRegExp.Test(i.FormulaLocal) Then
                Set myMatches = myRegExp.Execute(i.FormulaLocal)
                If myMatches.Count >= CInt(vRetVal2) Then i.FormulaLocal = myMatches.Item(CInt(vRetVal2) - 1)
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
ExtractWithRegExp_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ExtractWithRegExp, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ValueToText
' Purpose: Cell format change to the text. In case of cells range, the action is repeated for each cell
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ValueToText(control As IRibbonControl)
    On Error GoTo ValueToText_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim tmpvar
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.NumberFormatLocal = "@"
        i.FormulaLocal = i.Text
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
ValueToText_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ValueToText, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ValueToGeneral
' Purpose: Cell format change to the general. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ValueToGeneral(control As IRibbonControl)
    On Error GoTo ValueToGeneral_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim tmpvar
    Dim TargetRange As Range
    If Selection.Count = 1 Then Set TargetRange = Selection Else Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    For Each i In TargetRange
        i.NumberFormatLocal = "General"
        i.FormulaLocal = i.Text
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
ValueToGeneral_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ValueToGeneral, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ValueToFormula
' Purpose: Convert cell to formula with setting cell format to the general. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ValueToFormula(control As IRibbonControl)
    On Error GoTo ValueToFormula_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
    Set TargetRange = Selection
    Else
    Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.NumberFormatLocal = "General"
        i.FormulaLocal = "=" & i.Text
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
ValueToFormula_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ValueToFormula, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: FormulaAsText
' Purpose: Convert cell formula as text with the representaion function name in local format
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub FormulaAsText(control As IRibbonControl)
    On Error GoTo FormulaAsText_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range, Prefix As String
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        Prefix = ""
        If i.HasFormula Then i.FormulaLocal = Chr(39) & i.FormulaLocal
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
FormulaAsText_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FormulaAsText, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: FormulaAsTextInt
' Purpose: Convert cell formula as text with the representaion function name in international format. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 25.09.2020
' ----------------------------------------------------------------
Sub FormulaAsTextInt(control As IRibbonControl)
    On Error GoTo FormulaAsTextInt_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range, Prefix As String
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        Prefix = ""
        If i.HasFormula Then i.Formula = Chr(39) & i.Formula
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
FormulaAsTextInt_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FormulaAsTextInt, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: TextToClipboard
' Purpose: Copy text value from cell into the clipboard. In case of cells range, the action is repeated for each cell
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub TextToClipboard(control As IRibbonControl)
    On Error GoTo TextToClipboard_Error
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    Select Case TypeName(Selection)
        Case Is = "TextBox"
            ResultData = Selection.Caption
        Case Is = "Range"
            If Selection.Count = 1 Then
                Set TargetRange = Selection
            Else
                Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
            End If
            For Each i In TargetRange
                If i.Text <> "" Then
                    If ResultData <> "" Then
                        ResultData = ResultData & ";" & i.Text
                    Else
                        ResultData = i.Text
                    End If
                End If
            Next
        Case Else
    End Select
    If Application.WorksheetFunction.IsText(ResultData) Then
        ClipBoard_SetData (ResultData)
    Else
        ClipBoard_SetData (Format(ResultData))
    End If
    On Error GoTo 0
    Exit Sub
TextToClipboard_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine TextToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: TextToClipboardDialog
' Purpose: Copy text value from cell into the clipboard with showing text into pop-up window. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub TextToClipboardDialog(control As IRibbonControl)
    On Error GoTo TextToClipboardDialog_Error
    Dim i As Range
    Dim a As Variant
    Dim ResultData As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.Text <> "" Then
            If ResultData <> "" Then
                ResultData = ResultData & ";" & i.Text
            Else
                ResultData = i.Text
            End If
        End If
    Next
    a = MsgBoxU(ResultData, vbOKCancel + vbInformation, "Cell text value:")
    Select Case a
        Case vbOK
            If Application.WorksheetFunction.IsText(ResultData) Then
                ClipBoard_SetData (ResultData)
            Else
                ClipBoard_SetData (Format(ResultData))
            End If
        Case Else
    End Select
    On Error GoTo 0
    Exit Sub
TextToClipboardDialog_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine TextToClipboardDialog, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: TextToClipboardSeparatorSelection
' Purpose: Coping text from selected cells into the clipboard and concatenating values with using a user-specified text delimiter. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Sub TextToClipboardSeparatorSelection(control As IRibbonControl)
    On Error GoTo TextToClipboardSeparatorSelection_Error
    InputDelimiterDialogCD.Caption = "Change text delimiter"
    InputDelimiterDialogCD.DialogDescription.Caption = "Enter new text delimiter"
    If Separator = "\t" Then
        InputDelimiterDialogCD.InputString = ""
        InputDelimiterDialogCD.Tab_Button.SetFocus
    ElseIf Separator = "\n" Then
        InputDelimiterDialogCD.InputString = ""
        InputDelimiterDialogCD.CR_Button.SetFocus
    Else
        InputDelimiterDialogCD.InputString = Separator
        InputDelimiterDialogCD.InputString.SetFocus
        InputDelimiterDialogCD.InputString.SelStart = 0
        InputDelimiterDialogCD.InputString.SelLength = Len(InputDelimiterDialog.InputString.Text)
    End If
    If WithoutRepeat = 1 Then
        InputDelimiterDialogCD.CheckDublicate.Value = True
    Else
        InputDelimiterDialogCD.CheckDublicate.Value = False
    End If
    Dim Result As Variant
    InputDelimiterDialogCD.Show 1
    Result = InputDelimiterDialogCD.DialogResult
    If Result = 0 Then
        Unload InputDelimiterDialogCD
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputDelimiterDialogCD.InputString.Text
    If Result2 = CStr(vbTab) Then
        Separator = "\t"
    ElseIf Result2 = CStr(vbCrLf) Then
        Separator = "\n"
    Else
        Separator = Result2
    End If
    If InputDelimiterDialogCD.CheckDublicate = True Then
    WithoutRepeat = 1
    Else
    WithoutRepeat = 0
    End If
    Dim i As Range, lr As Long, lc As Long, sRes As String
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If sRes <> "" Then
            sRes = sRes & Result2 & i.Text
        Else
            sRes = i.Text
        End If
    Next
    If WithoutRepeat Then
        Dim oDict As Object, sTmpStr
        Set oDict = CreateObject("Scripting.Dictionary")
        sTmpStr = Split(sRes, Result2)
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            On Error Resume Next
            oDict.Add sTmpStr(lr), sTmpStr(lr)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo TextToClipboardSeparatorSelection_Error
        Next lr
        sRes = ""
        sTmpStr = oDict.Keys
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            sRes = sRes & IIf(sRes <> "", Result2, "") & sTmpStr(lr)
        Next lr
    End If
    If Application.WorksheetFunction.IsText(sRes) Then
        ClipBoard_SetData (sRes)
    Else
        ClipBoard_SetData (Format(sRes))
    End If
    On Error GoTo 0
    Exit Sub
TextToClipboardSeparatorSelection_Error:
    Unload InputDelimiterDialogCD
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine TextToClipboardSeparatorSelection, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: Merge
' Purpose: Getting text from specified cells and concatenating with using a user-specified text delimiter
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RangeWithText (Range): Cells range for combine
' Parameter TextSeparator (String): User-defined text delimiter (space symbol by default)
' Parameter IsRepeatedText (Boolean): Exclude duplicates: 1 - yes, 0 - no (by default)
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function Merge(RangeWithText As Range, Optional TextSeparator As String = " ", Optional IsRepeatedText As Boolean = False)
    On Error GoTo Merge_Error
    Dim avData, lr As Long, lc As Long, sRes As String
    Application.Volatile True
    avData = RangeWithText.Value
    If Not IsArray(avData) Then
        Merge = avData
        Exit Function
    End If
    For lc = 1 To UBound(avData, 2)
        For lr = 1 To UBound(avData, 1)
            If Len(avData(lr, lc)) Then
                sRes = sRes & TextSeparator & avData(lr, lc)
            End If
        Next lr
    Next lc
    If Len(sRes) Then
        sRes = Mid(sRes, Len(TextSeparator) + 1)
    End If
    If IsRepeatedText Then
        Dim oDict As Object, sTmpStr
        Set oDict = CreateObject("Scripting.Dictionary")
        sTmpStr = Split(sRes, TextSeparator)
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            On Error Resume Next
            oDict.Add sTmpStr(lr), sTmpStr(lr)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo Merge_Error
        Next lr
        sRes = ""
        sTmpStr = oDict.Keys
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            sRes = sRes & IIf(sRes <> "", TextSeparator, "") & sTmpStr(lr)
        Next lr
    End If
    Merge = sRes
    On Error GoTo 0
    Exit Function
Merge_Error:
    Merge = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: MergeRegion
' Purpose: Getting text from cells areas with combined their using user-defined text delimitor
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RegionWithText (Range): Cells range for combine
' Parameter TextSeparator (String): User-defined text delimitor (space by default)
' Parameter IsRepeatedText (Boolean): Exclude duplicates: 1 - yes, 0 - no (by default)
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function MergeRegion(RegionWithText As Range, Optional TextSeparator As String = " ", Optional IsRepeatedText As Boolean = False)
    On Error GoTo MergeRegion_Error
    Dim avData, lr As Long, lc As Long, sRes As String, i
    Application.Volatile True
    i = 1
    Do
        avData = RegionWithText.Areas(i).Value
        If Not IsArray(avData) And RegionWithText.Areas.Count < 2 Then
            MergeRegion = avData
            Exit Function
        End If
        If IsArray(avData) Then
            For lc = 1 To UBound(avData, 2)
                For lr = 1 To UBound(avData, 1)
                    If Len(avData(lr, lc)) Then
                        sRes = sRes & TextSeparator & avData(lr, lc)
                    End If
                Next lr
            Next lc
        Else
            sRes = sRes & TextSeparator & avData
        End If
        i = i + 1
    Loop While i <= RegionWithText.Areas.Count
    If Len(sRes) Then
        sRes = Mid(sRes, Len(TextSeparator) + 1)
    End If
    If IsRepeatedText Then
        Dim oDict As Object, sTmpStr
        Set oDict = CreateObject("Scripting.Dictionary")
        sTmpStr = Split(sRes, TextSeparator)
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            On Error Resume Next
            oDict.Add sTmpStr(lr), sTmpStr(lr)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo MergeRegion_Error
        Next lr
        sRes = ""
        sTmpStr = oDict.Keys
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            sRes = sRes & IIf(sRes <> "", TextSeparator, "") & sTmpStr(lr)
        Next lr
    End If
    MergeRegion = sRes
    On Error GoTo 0
    Exit Function
MergeRegion_Error:
    MergeRegion = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: FormulaToClipboard
' Purpose: Copy formula from cell to clipboard in international format. In case of cells range, formulas combined using tabulation symbol as text delimiter
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub FormulaToClipboard(control As IRibbonControl)
    On Error GoTo FormulaToClipboard_Error
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    InputDelimiterDialog.Caption = "Change text delimiter"
    InputDelimiterDialog.DialogDescription.Caption = "Enter new text delimiter"
    InputDelimiterDialog.InputString = ""
    InputDelimiterDialog.InputString.Enabled = False
    InputDelimiterDialog.InputString.Locked = False
    If CopyFormulaSeparator = "\t" Then
        InputDelimiterDialog.Tab_Button.SetFocus
    ElseIf CopyFormulaSeparator = "\n" Then
        InputDelimiterDialog.CR_Button.SetFocus
    Else
        CopyFormulaSeparator = "\n"
        InputDelimiterDialog.CR_Button.SetFocus
    End If
    Dim Result As Variant
    InputDelimiterDialog.Show 1
    Result = InputDelimiterDialog.DialogResult
    If Result = 0 Then
        Unload InputDelimiterDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputDelimiterDialog.InputString.Text
    If Result2 = CStr(vbTab) Then
        CopyFormulaSeparator = "\t"
    ElseIf Result2 = CStr(vbCrLf) Then
        CopyFormulaSeparator = "\n"
    Else
        CopyFormulaSeparator = "\n"
        Result2 = vbCrLf
    End If
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            If ResultData <> "" Then
                ResultData = ResultData & Result2 & i.Formula
            Else
                ResultData = i.Formula
            End If
        End If
    Next
    If Application.WorksheetFunction.IsText(ResultData) Then
        ClipBoard_SetData (ResultData)
    Else
        ClipBoard_SetData (Format(ResultData))
    End If
    On Error GoTo 0
    Exit Sub
FormulaToClipboard_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FormulaToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: LocalFormulaToClipboard
' Purpose: Copy formula from cell to clipboard in local format. In case of cells range, formulas combined using tabulation symbol as text delimiter
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 08.10.2020
' ----------------------------------------------------------------
Sub LocalFormulaToClipboard(control As IRibbonControl)
    On Error GoTo LocalFormulaToClipboard_Error
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    InputDelimiterDialog.Caption = "Change text delimiter"
    InputDelimiterDialog.DialogDescription.Caption = "Enter new text delimiter"
    InputDelimiterDialog.InputString = ""
    InputDelimiterDialog.InputString.Enabled = False
    InputDelimiterDialog.InputString.Locked = False
    If CopyFormulaSeparator = "\t" Then
        InputDelimiterDialog.Tab_Button.SetFocus
    ElseIf CopyFormulaSeparator = "\n" Then
        InputDelimiterDialog.CR_Button.SetFocus
    Else
        CopyFormulaSeparator = "\n"
        InputDelimiterDialog.CR_Button.SetFocus
    End If
    Dim Result As Variant
    InputDelimiterDialog.Show 1
    Result = InputDelimiterDialog.DialogResult
    If Result = 0 Then
        Unload InputDelimiterDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputDelimiterDialog.InputString.Text
    If Result2 = CStr(vbTab) Then
        CopyFormulaSeparator = "\t"
    ElseIf Result2 = CStr(vbCrLf) Then
        CopyFormulaSeparator = "\n"
    Else
        CopyFormulaSeparator = "\n"
        Result2 = vbCrLf
    End If
    
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            If ResultData <> "" Then
                ResultData = ResultData & Result2 & i.FormulaLocal
            Else
                ResultData = i.FormulaLocal
            End If
        End If
    Next
    If Application.WorksheetFunction.IsText(ResultData) Then
        ClipBoard_SetData (ResultData)
    Else
        ClipBoard_SetData (Format(ResultData))
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
LocalFormulaToClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine LocalFormulaToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: NumToClipboard
' Purpose: Copy numerical value from cell to clipboard. In case of cells range, values from selected cells are summarized
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub NumToClipboard(control As IRibbonControl)
    On Error GoTo NumToClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If i.Text <> "" Then
                ResultData = ResultData + CDbl(i.Text)
            End If
        Else
            If IsNumeric(i.Value) Then
                ResultData = ResultData + i.Value
            End If
        End If
    Next
    ClipBoard_SetData (Format(ResultData))
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
NumToClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine NumToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: RangeAddressToClipboard
' Purpose: Copy range address to clipboard
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub RangeAddressToClipboard(control As IRibbonControl)
    On Error GoTo RangeAddressToClipboard_Error
    Dim x As Range
    Dim RangeAsText As String, SelectedRanges As Range
    Set x = Application.InputBox(PROMPT:="Select range", TITLE:="Select range in current workbook", Type:=8)
    If ObjPtr(x) = 0 Then
        Exit Sub
    End If
    RangeAsText = x.Address(True, True, xlA1, False, False)
    Set SelectedRanges = Range(RangeAsText)
    ClipBoard_SetData (RangeAsText)
    On Error GoTo 0
    Exit Sub
RangeAddressToClipboard_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine RangeAddressToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: CellAddressToClipboard
' Purpose: Copy range address to clipboard. In case of cells range, cells addresses combined using semicolon character
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub CellAddressToClipboard(control As IRibbonControl)
    On Error GoTo CellAddressToClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If ResultData <> "" Then
            ResultData = ResultData & ";" & i.Address(0, 0)
        Else
            ResultData = i.Address(0, 0)
        End If
    Next
    If Application.WorksheetFunction.IsText(ResultData) Then
        ClipBoard_SetData (ResultData)
    Else
        ClipBoard_SetData (Format(ResultData))
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
CellAddressToClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine CellAddressToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: FilledCellAdressToClipboard
' Purpose: Copy not empty cell address to clipboard. In case of cells range, cells addresses combined using semicolon character
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub FilledCellAdressToClipboard(control As IRibbonControl)
    On Error GoTo FilledCellAdressToClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    InputStringDialog.Caption = "Change delimiter"
    InputStringDialog.DialogDescription.Caption = "Enter new delimiter"
    InputStringDialog.InputString = CStr(CellsAddressSeparator)
    InputStringDialog.InputString.SetFocus
    InputStringDialog.InputString.SelStart = 0
    InputStringDialog.InputString.SelLength = Len(InputStringDialog.InputString.Text)
    Dim Result As Variant
    InputStringDialog.Show 1
    Result = InputStringDialog.DialogResult
    If Result = 0 Then
        Unload InputStringDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputStringDialog.InputString.Text
    If IsEmpty(Result2) Or Result2 = "" Then
        Unload InputStringDialog
        Exit Sub
    End If
    CellsAddressSeparator = CStr(Result2)
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If ResultData <> "" Then
            If i.FormulaLocal <> "" Then ResultData = ResultData & CellsAddressSeparator & i.Address(0, 0)
        Else
            If i.FormulaLocal <> "" Then ResultData = i.Address(0, 0)
        End If
    Next
    If Application.WorksheetFunction.IsText(ResultData) Then
        ClipBoard_SetData (ResultData)
    Else
        ClipBoard_SetData (Format(ResultData))
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
FilledCellAdressToClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine FilledCellAdressToClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: WrapSquareBrackets
' Purpose: Wrap cell value into square brackets. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub WrapSquareBrackets(control As IRibbonControl)
    On Error GoTo WrapSquareBrackets_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If Not i.HasFormula Then
            i.Value = "[" & i.Value & "]"
        End If
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
WrapSquareBrackets_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine WrapSquareBrackets, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: WrapSum
' Purpose: Wrap cell calcalated value into SUM() formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub WrapSum(control As IRibbonControl)
    On Error GoTo WrapSum_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If i.Text <> "" Then
                i.FormulaLocal = "=SUM(" & i.Text & ")"
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
WrapSum_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine WrapSum, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: WrapRound
' Purpose: Wrap cell calcalated value into ROUND() formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub WrapRound(control As IRibbonControl)
    On Error GoTo WrapRound_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim RoundPrecisionLocal As String
    RoundPrecisionLocal = InputBox("Enter round precision", "Round precision", CStr(RoundPrecision))
    If StrPtr(RoundPrecisionLocal) = 0 Then 'Cancel buttom pressed
        Exit Sub
    End If
    Dim i As Range
    RoundPrecision = CInt(RoundPrecisionLocal)
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If i.Text <> "" Then
                i.Value = i.Value * (-1) * (-1)
                i.FormulaLocal = "=ROUND(" & i.Value & "," & RoundPrecisionLocal & ")"
            End If
        Else
            If i.FormulaLocal <> "" Then
                Dim TempFormula
                TempFormula = Right(i.FormulaLocal, Len(i.FormulaLocal) - 1)
                i.FormulaLocal = "=ROUND((" & TempFormula & ")," & RoundPrecisionLocal & ")"
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
WrapRound_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine WrapRound, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: IncreaseRoundPrecision
' Purpose: Increase round precision by one digit into ROUND() formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub IncreaseRoundPrecision(control As IRibbonControl)
    On Error GoTo IncreaseRoundPrecision_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    Dim NewVal As Variant
    myRegExp.Global = True
    myRegExp.Pattern = "=ROUND\((.*?),(\d+)\)"
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            Set myMatches = myRegExp.Execute(i.FormulaLocal)
            If myMatches.Count >= 1 Then
                Set m = myMatches.Item(0)
                If m.SubMatches.Count = 2 Then
                    If (m.SubMatches(1) > 0) Then
                        NewVal = m.SubMatches(1)
                        NewVal = NewVal + 1
                        ResultString = myRegExp.Replace(i.FormulaLocal, "=ROUND($1," & NewVal & ")")
                        i.FormulaLocal = ResultString
                    End If
                End If
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
IncreaseRoundPrecision_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine IncreaseRoundPrecision, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DecreaseRoundPrecision
' Purpose: Decrease round precision by one digit into ROUND() formula. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DecreaseRoundPrecision(control As IRibbonControl)
    On Error GoTo DecreaseRoundPrecision_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    Dim NewVal As Variant
    myRegExp.Global = True
    myRegExp.Pattern = "=ROUND\((.*?),(\d+)\)"
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula Then
            Set myMatches = myRegExp.Execute(i.FormulaLocal)
            If myMatches.Count >= 1 Then
                Set m = myMatches.Item(0)
                If m.SubMatches.Count = 2 Then
                    If (m.SubMatches(1) > 0) Then
                        NewVal = m.SubMatches(1)
                        NewVal = NewVal - 1
                        ResultString = myRegExp.Replace(i.FormulaLocal, "=ROUND($1," & NewVal & ")")
                        i.FormulaLocal = ResultString
                    End If
                End If
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
DecreaseRoundPrecision_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DecreaseRoundPrecision, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: PasteClipboard
' Purpose: Paste value into cell from clipboard. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub PasteClipboard(control As IRibbonControl)
    On Error GoTo PasteClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myData As Object
    Set myData = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    myData.GetFromClipboard
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    TargetRange = myData.GetText()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
PasteClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine PasteClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: PastePrefixClipboard
' Purpose: Paste value into begin of cell text from clipboard. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub PastePrefixClipboard(control As IRibbonControl)
    On Error GoTo PastePrefixClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myData As Object
    Dim MyText As Variant
    Dim i As Range
    Set myData = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    myData.GetFromClipboard
    MyText = myData.GetText()
    If MyText <> "" Then
        Dim TargetRange As Range
        If Selection.Count = 1 Then
            Set TargetRange = Selection
        Else
            Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
        End If
        For Each i In TargetRange
            i = MyText & i.Text
        Next
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
PastePrefixClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine PastePrefixClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: PasteSuffixClipboard
' Purpose: Paste value into end of cell text from clipboard. In case of cells range, the action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub PasteSuffixClipboard(control As IRibbonControl)
    On Error GoTo PasteSuffixClipboard_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myData As Object
    Dim MyText As Variant
    Dim i As Range
    Set myData = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    myData.GetFromClipboard
    MyText = myData.GetText()
    If MyText <> "" Then
        Dim TargetRange As Range
        If Selection.Count = 1 Then
            Set TargetRange = Selection
        Else
            Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
        End If
        For Each i In TargetRange
            i = i.Text & MyText
        Next
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
PasteSuffixClipboard_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine PasteSuffixClipboard, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: Median
' Purpose: Selection in a range of cells with median values using setting their background to color 8 from Excel palette. You must first select a cells range
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub Median(control As IRibbonControl)
    On Error GoTo Median_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim Middle As Variant
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    Middle = Application.Median(TargetRange)
    For Each i In TargetRange
        If i.Value = Middle Then
            i.Interior.ColorIndex = 8
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
Median_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine Median, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: MedianByColumns
' Purpose: Application conditional formatting rule to the selected table columns for selection table rows contained median values
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub MedianByColumns(control As IRibbonControl)
    On Error GoTo MedianByColumns_Error
    Dim i
    Dim FirstCellsLinks() As String
    Dim Fx As String
    ReDim FirstCellsLinks(Selection.Areas.Count - 1)
    If Selection.Areas.Count < 3 Then
        Exit Sub
    End If
    For i = 1 To Selection.Areas.Count
        FirstCellsLinks(i - 1) = Selection.Areas(i).Cells(1).Address(False, False, xlA1, False, False)
    Next i
    For i = 1 To Selection.Areas.Count
    Fx = "=" & FirstCellsLinks(i - 1) & "=MEDIAN(" & Join(FirstCellsLinks, ",") & ")"
        With Selection.Areas(i)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:=Fx
            .FormatConditions(1).Interior.ColorIndex = 33
            .FormatConditions(1).Font.ColorIndex = 1
        End With
    Next i
    On Error GoTo 0
    Exit Sub
MedianByColumns_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine MedianByColumns, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: Divisors
' Purpose: Displayed dividers of numerical value from selected cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub Divisors(control As IRibbonControl)
    On Error GoTo Divisors_Error
    Dim n As Variant
    Dim i As Variant
    Dim j As Variant
    Dim s As Variant
    Dim f As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    n = TargetRange.Value
    For i = 1 To n
        If n Mod i = 0 Then s = s + i
    Next i
    Debug.Print s
    For i = 1 To n Step 1
        If n Mod i = 0 Then Debug.Print "i ="; i
    Next i
    f = 2
    For j = 1 To n Step 1
        If n / j = n \ j Then
            TargetRange.Offset(f) = j
            f = f + 1
        End If
    Next j
    On Error GoTo 0
    Exit Sub
Divisors_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine Divisors, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: TrimSpaces
' Purpose: Removing leading, trailing and double spaces into the cell text. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub TrimSpaces(control As IRibbonControl)
    On Error GoTo TrimSpaces_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Range
    Dim strSize As Long
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If i.Text <> "" Then
                If Len(i.Text) <= 255 Then
                    i = Application.WorksheetFunction.Trim(i)
                Else
                    Do
                        strSize = Len(i.Text)
                        i = Trim(Replace(i, "  ", " ", , , vbBinaryCompare))
                    Loop Until strSize = Len(i.Text)
                    strSize = 0
                End If
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
TrimSpaces_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine TrimSpaces, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: SpecialsSymbols
' Purpose: Strip whitespace (or other non-printed characters) from text in cell. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub SpecialsSymbols(control As IRibbonControl)
    On Error GoTo SpecialsSymbols_Error
    Dim i As Range
    Dim strSize As Long
    Dim TargetRange As Range
    Dim tmp As Variant
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If i.HasFormula <> True Then
            If i.Text <> "" Then
                If Len(i.Text) <= 255 Then
                    i.Value = Application.Clean(i.Value)
                Else
                    tmp = "CLEAN(""" & i.Value & """)"
                    i.Value = Application.Evaluate(tmp)
                End If
            End If
        End If
    Next
    On Error GoTo 0
    Exit Sub
SpecialsSymbols_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine SpecialsSymbols, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: MakeLocalAddress
' Purpose: Convert external links inside the formulas into local addresses. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub MakeLocalAddress(control As IRibbonControl)
    On Error GoTo MakeLocalAddress_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    Dim myMatches As Object
    myRegExp.Global = True
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    myRegExp.Pattern = "'[^']*'!"
    Dim ResultString
    For Each i In TargetRange
        If i.HasFormula = True Then
            If myRegExp.Test(i.FormulaLocal) Then
                ResultString = myRegExp.Replace(i.FormulaLocal, "")
                i.FormulaLocal = ResultString
            End If
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
MakeLocalAddress_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine MakeLocalAddress, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DecPoint
' Purpose: Converts floating point value from US to international format by replacing a decimal point with a comma
' Procedure Kind: Function
' Procedure Access: Public
' Parameter Value(#): Floating point number for converting
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function DecPoint(Value#)
    On Error GoTo DecPoint_Error
    DecPoint = Len(Split(Replace(Value#, ".", ",") & ",", ",")(1))
    On Error GoTo 0
    Exit Function
DecPoint_Error:
    DecPoint = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: UpdateCell
' Purpose: Auto reentering formula into cell. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub UpdateCell(control As IRibbonControl)
    On Error GoTo UpdateCell_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.FormulaLocal = i.FormulaLocal
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
UpdateCell_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine UpdateCell, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: Uppercase
' Purpose: Convert cell text to uppercase. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub Uppercase(control As IRibbonControl)
    On Error GoTo Uppercase_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If Not i.HasFormula Then
            i.Value = UCase(i.Value)
        End If
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
Uppercase_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine Uppercase, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: Lowercase
' Purpose: Convert cell text to lowercase. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub Lowercase(control As IRibbonControl)
    On Error GoTo Lowercase_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If Not i.HasFormula Then
            i.Value = LCase(i.Value)
        End If
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
Lowercase_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine Lowercase, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: UcFirst
' Purpose: Convert cell text to sentence case. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub UcFirst(control As IRibbonControl)
    On Error GoTo UcFirst_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If Not i.HasFormula Then
            i.Value = Application _
                .WorksheetFunction _
                .Proper(i.Value)
        End If
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
UcFirst_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine UcFirst, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ResetColumnWidth
' Purpose: Reset column width by default value. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ResetColumnWidth(control As IRibbonControl)
    On Error GoTo ResetColumnWidth_Error
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.ColumnWidth = ActiveSheet.StandardWidth
    Next
    On Error GoTo 0
    Exit Sub
ResetColumnWidth_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ResetColumnWidth, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ResetNoteSize
' Purpose: Reset the size of cell note by default values. In case of cells range, this action is repeated for each cell
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ResetNoteSize(control As IRibbonControl)
    On Error GoTo ResetNoteSize_Error
    Dim xComment As Comment, i As Variant, DPI As Integer
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Variant, objItem As Variant
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_DisplayConfiguration", , 48)
    DPI = 72
    Select Case TypeName(Selection)
        Case Is = "TextBox"
            With Selection
                .Width = (107.25 * 2.54 / DPI) * DPI / 2.54
                .Height = (59.25 * 2.54 / DPI) * DPI / 2.54
            End With
        Case Is = "Range"
            For Each i In Selection
                Set xComment = i.Comment
                With xComment.Shape
                    .Width = (107.25 * 2.54 / DPI) * DPI / 2.54
                    .Height = (59.25 * 2.54 / DPI) * DPI / 2.54
                End With
            Next i
        Case Else
    End Select
    On Error GoTo 0
    Exit Sub
ResetNoteSize_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ResetNoteSize, line " & Erl & "."
End Sub

Public Sub CellsToNotes(control As IRibbonControl)
    On Error GoTo CellsToNotes_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim i As Variant
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        i.AddComment CStr(i.FormulaLocal)
    Next i
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
CellsToNotes_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine CellsToNotes, line " & Erl & "."
End Sub


' ----------------------------------------------------------------
' Procedure Name: UngroupAndFillCells
' Purpose: Unmerge cells with inserting common group values into all cells. In case of range, this action is repeated for each group
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub UngroupAndFillCells(control As IRibbonControl)
    On Error GoTo UngroupAndFillCells_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim Address As String
    Dim Cell As Range
    If TypeName(Selection) <> "Range" Then
        Exit Sub
    End If
    If Selection.Cells.Count = 1 Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For Each Cell In Intersect(Selection, ActiveSheet.UsedRange).Cells
        If Cell.MergeCells Then
            Address = Cell.MergeArea.Address
            Cell.UnMerge
            Range(Address).Value = Cell.Value
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
UngroupAndFillCells_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine UngroupAndFillCells, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: MergeByGroups
' Purpose: Makes separate group for each select cells area. Also by group merges text values contains in cells in each area
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub MergeByGroups(control As IRibbonControl)
    On Error GoTo MergeByGroups_Error
    InputDelimiterDialog.Caption = "Change delimiter"
    InputDelimiterDialog.DialogDescription.Caption = "Enter cells values delimiter"
    If MergeCellsSeparator = "\t" Then
        InputDelimiterDialog.InputString = ""
        InputDelimiterDialog.Tab_Button.SetFocus
    ElseIf MergeCellsSeparator = "\n" Then
        InputDelimiterDialog.InputString = ""
        InputDelimiterDialog.CR_Button.SetFocus
    Else
        InputDelimiterDialog.InputString = MergeCellsSeparator
        InputDelimiterDialog.InputString.SetFocus
        InputDelimiterDialog.InputString.SelStart = 0
        InputDelimiterDialog.InputString.SelLength = Len(InputDelimiterDialog.InputString.Text)
    End If
    Dim Result As Variant
    InputDelimiterDialog.Show 1
    Result = InputDelimiterDialog.DialogResult
    If Result = 0 Then
        Unload InputDelimiterDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputDelimiterDialog.InputString.Text
    If Result2 = CStr(vbTab) Then
        MergeCellsSeparator = "\t"
    ElseIf Result2 = CStr(vbCrLf) Then
        MergeCellsSeparator = "\n"
    Else
        MergeCellsSeparator = Result2
    End If
    Dim rCell As Range
    Dim sMergeStr As String
    Dim sMergeArray() As String
    Dim cntr
    If TypeName(Selection) <> "Range" Then Exit Sub
    cntr = 1
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    Dim areaRange, areaValues, i, lr As Long, lc As Long, sRes As String
    i = 1
    Do
        Set areaRange = TargetRange.Areas(i)
        areaValues = TargetRange.Areas(i).Value
        If IsArray(areaValues) Then
            For lc = 1 To UBound(areaValues, 2)
                For lr = 1 To UBound(areaValues, 1)
                    If Len(areaValues(lr, lc)) Then
                        If (Len(sRes)) Then
                            sRes = sRes & Result2 & areaValues(lr, lc)
                        Else
                            sRes = areaValues(lr, lc)
                        End If
                    End If
                Next lr
            Next lc
        Else
            sRes = areaValues
        End If
        Application.DisplayAlerts = False
        areaRange.Merge Across:=False
        Application.DisplayAlerts = True
        areaRange.Item(1).Value = sRes
        i = i + 1
        sRes = ""
    Loop While i <= TargetRange.Areas.Count
    On Error GoTo 0
    Exit Sub
MergeByGroups_Error:
    Unload InputDelimiterDialog
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine MergeByGroups, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: MergeCellsAndText
' Purpose: Groupping selected cells with combined their values using user-defined delimiter
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub MergeCellsAndText(control As IRibbonControl)
    On Error GoTo MergeCellsAndText_Error
    InputDelimiterDialog.Caption = "Change delimiter"
    InputDelimiterDialog.DialogDescription.Caption = "Enter cells values delimiter"
    If MergeCellsSeparator = "\t" Then
        InputDelimiterDialog.InputString = ""
        InputDelimiterDialog.Tab_Button.SetFocus
    ElseIf MergeCellsSeparator = "\n" Then
        InputDelimiterDialog.InputString = ""
        InputDelimiterDialog.CR_Button.SetFocus
    Else
        InputDelimiterDialog.InputString = MergeCellsSeparator
        InputDelimiterDialog.InputString.SetFocus
        InputDelimiterDialog.InputString.SelStart = 0
        InputDelimiterDialog.InputString.SelLength = Len(InputDelimiterDialog.InputString.Text)
    End If
    Dim Result As Variant
    InputDelimiterDialog.Show 1
    Result = InputDelimiterDialog.DialogResult
    If Result = 0 Then
        Unload InputDelimiterDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputDelimiterDialog.InputString.Text
    If Result2 = CStr(vbTab) Then
        MergeCellsSeparator = "\t"
    ElseIf Result2 = CStr(vbCrLf) Then
        MergeCellsSeparator = "\n"
    Else
        MergeCellsSeparator = Result2
    End If
    Dim rCell As Range
    Dim sMergeStr As String
    Dim sMergeArray() As String
    Dim cntr
    If TypeName(Selection) <> "Range" Then Exit Sub
    cntr = 1
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    With TargetRange
        For Each rCell In .Cells
            ReDim Preserve sMergeArray(cntr - 1)
            sMergeArray(cntr - 1) = rCell.Text
            cntr = cntr + 1
        Next rCell
        sMergeStr = Join(sMergeArray, Result2)
        Application.DisplayAlerts = False
        .Merge Across:=False
        Application.DisplayAlerts = True
        .Item(1).Value = sMergeStr
    End With
    On Error GoTo 0
    Exit Sub
MergeCellsAndText_Error:
    Unload InputDelimiterDialog
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine MergeCellsAndText, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: SelectBlankRows
' Purpose: Select empty rows in a table
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub SelectBlankRows(control As IRibbonControl)
    On Error GoTo SelectBlankRows_Error
    Dim i As Long
    Dim diapaz1 As Range
    Dim diapaz2 As Range
    Set diapaz1 = Application.Range(ActiveSheet.Range("A1"), _
        ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell))
    For i = 1 To diapaz1.Rows.Count
        If WorksheetFunction.CountA(diapaz1.Rows(i).EntireRow) = 0 Then
            If diapaz2 Is Nothing Then
                Set diapaz2 = diapaz1.Rows(i).EntireRow
            Else
                Set diapaz2 = Application.Union(diapaz2, diapaz1.Rows(i).EntireRow)
            End If
        End If
    Next
    If diapaz2 Is Nothing Then
        MsgBox "No blank lines found!"
    Else
        diapaz2.Select
    End If
    On Error GoTo 0
    Exit Sub
SelectBlankRows_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine SelectBlankRows, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DeleteBlankRows
' Purpose: Remove empty rows from table
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DeleteBlankRows(control As IRibbonControl)
    On Error GoTo DeleteBlankRows_Error
    Dim i As Long
    Dim diapaz1 As Range
    Dim diapaz2 As Range
    Set diapaz1 = Application.Range(ActiveSheet.Range("A1"), _
        ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell))
    For i = 1 To diapaz1.Rows.Count
        If WorksheetFunction.CountA(diapaz1.Rows(i).EntireRow) = 0 Then
            If diapaz2 Is Nothing Then
                Set diapaz2 = diapaz1.Rows(i).EntireRow
            Else
                Set diapaz2 = Application.Union(diapaz2, diapaz1.Rows(i).EntireRow)
            End If
        End If
    Next
    If diapaz2 Is Nothing Then
        MsgBox "No blank lines found!"
    Else
        diapaz2.[Delete]
    End If
    On Error GoTo 0
    Exit Sub
DeleteBlankRows_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DeleteBlankRows, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: HideBlankRows
' Purpose: Hide empty rows in a table
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub HideBlankRows(control As IRibbonControl)
    On Error GoTo HideBlankRows_Error
    Dim i As Long
    Dim diapaz1 As Range
    Dim diapaz2 As Range
    Set diapaz1 = Application.Range(ActiveSheet.Range("A1"), _
        ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell))
    For i = 1 To diapaz1.Rows.Count
        If WorksheetFunction.CountA(diapaz1.Rows(i).EntireRow) = 0 Then
            If diapaz2 Is Nothing Then
                Set diapaz2 = diapaz1.Rows(i).EntireRow
            Else
                Set diapaz2 = Application.Union(diapaz2, diapaz1.Rows(i).EntireRow)
            End If
        End If
    Next
    If diapaz2 Is Nothing Then
        MsgBox "No blank lines found!"
    Else
        diapaz2.EntireRow.Hidden = True
    End If
    On Error GoTo 0
    Exit Sub
HideBlankRows_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine HideBlankRows, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DuplicateBlankRows
' Purpose: Duplicate empty rows in a table
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DuplicateBlankRows(control As IRibbonControl)
    On Error GoTo DuplicateBlankRows_Error
    Dim i As Long
    Dim diapaz1 As Range
    Dim diapaz2 As Range
    Set diapaz1 = Application.Range(ActiveSheet.Range("A1"), _
        ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell))
    For i = 1 To diapaz1.Rows.Count
        If WorksheetFunction.CountA(diapaz1.Rows(i).EntireRow) = 0 Then
            If diapaz2 Is Nothing Then
                Set diapaz2 = diapaz1.Rows(i).EntireRow
            Else
                Set diapaz2 = Application.Union(diapaz2, diapaz1.Rows(i).EntireRow)
            End If
        End If
    Next
    If diapaz2 Is Nothing Then
        MsgBox "No blank lines found!"
    Else
        diapaz2.[Insert]
    End If
    On Error GoTo 0
    Exit Sub
DuplicateBlankRows_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DuplicateBlankRows, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DeleteEvenRows
' Purpose: Deletes even lines in the selected range
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DeleteEvenRows(control As IRibbonControl)
    On Error GoTo DeleteEvenRows_Error
    Dim ra As Range, delra As Range, cntdel As Integer
    cntdel = 0
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    For Each ra In TargetRange.Rows
        If cntdel <> 0 Then
            ra.EntireRow.Delete
        End If
        If cntdel = 2 Then cntdel = 0
        cntdel = cntdel + 1
    Next
    On Error GoTo 0
    Exit Sub
DeleteEvenRows_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DeleteEvenRows, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: DuplicateCurrentRow
' Purpose: Duplicate the current table row
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub DuplicateCurrentRow(control As IRibbonControl)
    On Error GoTo DuplicateCurrentRow_Error
    With ActiveCell.EntireRow
        .Offset(1, 0).Insert
        .Copy Rows(.Row + 1)
    End With
    On Error GoTo 0
    Exit Sub
DuplicateCurrentRow_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DuplicateCurrentRow, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ExtractTextFirstLine
' Purpose: Gets the first line of cell text. In the case of a range, cells text is concatenated using a semicolon
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RangeWithText (Range): Range of cells with text
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function ExtractTextFirstLine(RangeWithText As Range) As String
    On Error GoTo ExtractTextFirstLine_Error
    Application.Volatile True
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    myRegExp.Global = True
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    myRegExp.Pattern = ".*"
    Dim TargetRange As Range
    If RangeWithText.Count = 1 Then
        Set TargetRange = RangeWithText
    Else
        Set TargetRange = RangeWithText.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        Set myMatches = myRegExp.Execute(i.Value)
        If myMatches.Count >= 1 Then
            Set m = myMatches.Item(0)
            If (m.Value <> "") Then
                If ResultString <> "" Then
                    ResultString = ResultString & ";" & m.Value
                Else
                    ResultString = m.Value
                End If
            End If
        End If
    Next
    ExtractTextFirstLine = ResultString
    On Error GoTo 0
    Exit Function
ExtractTextFirstLine_Error:
    ExtractTextFirstLine = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: CalculateFormula
' Purpose: Evaluation of a formula specified as text
' Procedure Kind: Function
' Procedure Access: Public
' Parameter Fx (String): Formula as text
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function CalculateFormula(Fx As String)
    On Error GoTo CalculateFormula_Error
    Application.Volatile True
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    myRegExp.Global = True
    Dim myMatches As Object
    Dim m As Variant
    Dim ResultString As Variant, FormulaString As String
    myRegExp.Global = True
    myRegExp.Pattern = "(^[^=\x20]*?)$|(^.*?)=.*?$|(^[^\x20]*?)\x20\S*?$"
    Set myMatches = myRegExp.Execute(Fx)
    If myMatches.Count >= 1 Then
        Set m = myMatches.Item(0)
        If (m.SubMatches.Item(0) <> "") Then
            FormulaString = Replace(m.SubMatches.Item(0), ",", ".")
            ResultString = Application.Evaluate(FormulaString)
            CalculateFormula = ResultString
        ElseIf (m.SubMatches.Item(1) <> "") Then
            FormulaString = Replace(m.SubMatches.Item(1), ",", ".")
            ResultString = Application.Evaluate(FormulaString)
            CalculateFormula = ResultString
        ElseIf (m.SubMatches.Item(2) <> "") Then
            FormulaString = Replace(m.SubMatches.Item(2), ",", ".")
            ResultString = Application.Evaluate(FormulaString)
            CalculateFormula = ResultString
        End If
    End If
    On Error GoTo 0
    Exit Function
CalculateFormula_Error:
    CalculateFormula = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: ExtractByRegExp
' Purpose: Returns a part of text using a regular expression
' Procedure Kind: Function
' Procedure Access: Public
' Parameter TextSrc (String): Source text
' Parameter TemplateForExtract (String): Template (regular expression)
' Parameter MatchIndex (Integer): Match index for extract
' Parameter CapturingGroupIndex (Integer): Group index for extract
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function ExtractByRegExp(TextSrc As String, TemplateForExtract As String, Optional MatchIndex As Integer = 1, Optional CapturingGroupIndex As Integer = 1) As String
    On Error GoTo ExtractByRegExp_Error
    Dim regex As Variant, myMatches As Variant, m As Variant
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = TemplateForExtract
    regex.Global = True
    If regex.Test(TextSrc) Then
        Set myMatches = regex.Execute(TextSrc)
        If myMatches.Count >= 1 Then
            If MatchIndex > 1 Then
                Set m = myMatches.Item(MatchIndex - 1)
            Else
                Set m = myMatches.Item(0)
            End If
            If (m.SubMatches.Count > 0 And CapturingGroupIndex >= 1) Then
                ExtractByRegExp = m.SubMatches.Item(CapturingGroupIndex - 1)
            Else
                ExtractByRegExp = m.Value
            End If
            Exit Function
        End If
    End If
    On Error GoTo 0
    Exit Function
ExtractByRegExp_Error:
    ExtractByRegExp = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: GetStringByNumber
' Purpose: Gets lines from cell text by line number. In the case of range, text strings are concatenated using semicolon
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RangeWithMultiLineText (Range): Cells range with source text
' Parameter ExtractTemplate (String): Template for line extract (regular expression)
' Parameter LineIndex (Integer): Line number to extract from source text
' Parameter MultiLineMode (Boolean): Multiline search: True - on, False - off (by default)
' Parameter IgnoreRegister (Boolean): Case insensitive search: True - on, False - off (by default)
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function GetStringByNumber(RangeWithMultiLineText As Range, ExtractTemplate As String, Optional LineIndex As Integer = 1, Optional MultiLineMode As Boolean = False, Optional IgnoreRegister As Boolean = False) As String
    On Error GoTo GetStringByNumber_Error
    Application.Volatile True
    Dim myRegExp As Object
    Set myRegExp = CreateObject("VBScript.RegExp")
    myRegExp.Global = True
    myRegExp.Multiline = MultiLineMode
    myRegExp.IgnoreCase = IgnoreRegister
    Dim myMatches As Object
    Dim m As Object
    Dim ResultString As String
    Dim i As Range
    myRegExp.Global = True
    myRegExp.Pattern = ExtractTemplate
    Dim TargetRange As Range
    If RangeWithMultiLineText.Count = 1 Then
        Set TargetRange = RangeWithMultiLineText
    Else
        Set TargetRange = RangeWithMultiLineText.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        Set myMatches = myRegExp.Execute(i.Value)
        If myMatches.Count >= 1 Then
            Set m = myMatches.Item(LineIndex - 1)
            If ResultString <> "" Then
                ResultString = ResultString & ";" & m.Value
            Else
                ResultString = m.Value
            End If
        End If
    Next
    GetStringByNumber = ResultString
    On Error GoTo 0
    Exit Function
GetStringByNumber_Error:
    GetStringByNumber = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: SplitString
' Purpose: Splits string using user-defined delimiter and returns value by number
' Procedure Kind: Function
' Procedure Access: Public
' Parameter StringSrc (String): Source string
' Parameter PartSeparator (String): User-defined delimiter
' Parameter SubstrIndex (Integer): Return value number from 1
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function SplitString(StringSrc As String, PartSeparator As String, Optional SubstrIndex As Integer = 1) As String
    On Error GoTo SplitString_Error
    Dim a As Variant
    If SubstrIndex < 1 Then
        SubstrIndex = 1
    End If
    a = Split(StringSrc, PartSeparator)
    If (UBound(a) + 1) > 0 And SubstrIndex <= (UBound(a) + 1) Then
        If a(SubstrIndex - 1) <> "" Then
            ActiveCell.NumberFormatLocal = "General"
            SplitString = a(SubstrIndex - 1)
        End If
    End If
    On Error GoTo 0
    Exit Function
SplitString_Error:
    SplitString = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: CountPartsSplitString
' Purpose: Splits string using user-defined delimiter and returns count of elements
' Procedure Kind: Function
' Procedure Access: Public
' Parameter StringSrc (String): Source string
' Parameter PartSeparator (String): User-defined delimiter
' Return Type: Variant
' Author: Petr Kovalenko
' Date: 19.03.2021
' ----------------------------------------------------------------
Public Function CountPartsSplitString(StringSrc As String, PartSeparator As String) As Variant
    On Error GoTo CountPartsSplitString_Error
    Dim a As Variant
    a = Split(StringSrc, PartSeparator)
    ActiveCell.NumberFormatLocal = "General"
    CountPartsSplitString = UBound(a) + 1
    On Error GoTo 0
    Exit Function
CountPartsSplitString_Error:
    CountPartsSplitString = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: RemoveHiddenNames
' Purpose: Removes hidden names from current workbook
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub RemoveHiddenNames(control As IRibbonControl)
    On Error GoTo RemoveHiddenNames_Error
    Dim n As Name
    Dim Count As Integer
    For Each n In ActiveWorkbook.Names
        If Not n.Visible Then
            n.Delete
            Count = Count + 1
        End If
    Next n
    MsgBox "Hidden names in count of " & Count & " have been removed."
    On Error GoTo 0
    Exit Sub
RemoveHiddenNames_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine RemoveHiddenNames, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: MergeWorkbooks
' Purpose: Merges several workbooks into one workbook
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub MergeWorkbooks(control As IRibbonControl)
    On Error GoTo MergeWorkbooks_Error
    Dim wbTarget As New Workbook, wbSrc As Workbook, shSrc As Worksheet, shTarget As Worksheet, arFiles, i As Integer, stbar As Boolean
    On Error GoTo 0
    With Application
        arFiles = Application.GetOpenFilename(FileFilter:="All files (*.*), *.*", MultiSelect:=True, TITLE:="Files to Merge")
        If Not IsArray(arFiles) Then End
        Set wbTarget = Workbooks.Add(template:=xlWorksheet)
        .ScreenUpdating = False
        stbar = .DisplayStatusBar
        .DisplayStatusBar = True
        .DisplayAlerts = False
        For i = 1 To UBound(arFiles)
            .StatusBar = "File processing " & i & " of " & UBound(arFiles)
            Set wbSrc = Workbooks.Open(arFiles(i), ReadOnly:=True)
            For Each shSrc In wbSrc.Worksheets
                If IsNull(shSrc.UsedRange.Text) Then
                    Set shTarget = wbTarget.Sheets.Add(after:=wbTarget.Sheets(wbTarget.Sheets.Count))
                    shTarget.Name = shSrc.Name & "-" & i
                    shSrc.Cells.Copy shTarget.Range("A1")
                End If
            Next
            wbSrc.Close False
        Next
        .ScreenUpdating = True
        .DisplayStatusBar = stbar
        .StatusBar = False
        If wbTarget.Sheets.Count = 1 Then
            MsgBox "There are no non-blank sheets in the selected books, there is nothing to save!"
            wbTarget.Close False
            End
        Else
            .DisplayAlerts = False
            wbTarget.Sheets(1).Delete
            .DisplayAlerts = True
        End If
        On Error Resume Next
        On Error GoTo 0
        arFiles = Application.GetSaveAsFilename(InitialFileName:="Result", FileFilter:="Excel 2007-365 (*.xlsx),*.xlsx", TITLE:="Save merged workbook")
        If VarType(arFiles) = vbBoolean Then
            GoTo save_err
        Else
            On Error GoTo save_err
            wbTarget.SaveAs arFiles
        End If
        End
save_err:
        MsgBox "Workbook not saved!", vbCritical
    End With
    On Error GoTo 0
    Exit Sub
MergeWorkbooks_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine MergeWorkbooks, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: RangeToSmartTable

' ----------------------------------------------------------------
' Procedure Name: ShortenRange
' Purpose: Reduces a range by the specified lines number
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RangeWithRows (Range): Source range
' Parameter RowsCount (Long): Lines number for reduce
' Return Type: Range
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Public Function ShortenRange(ByVal RangeWithRows As Range, ByVal RowsCount As Long) As Range
    On Error GoTo ShortenRange_Error
    Application.Volatile True
    If RangeWithRows Is Nothing Then Exit Function
    Dim Rows_Count As Long
    Rows_Count = RangeWithRows.Rows.Count
    If Rows_Count < 2 Or Rows_Count <= RowsCount Then
        Set ShortenRange = RangeWithRows
        Exit Function
    End If
    Set ShortenRange = RangeWithRows.Resize(Rows_Count - RowsCount, RangeWithRows.Columns.Count)
    On Error GoTo 0
    Exit Function
ShortenRange_Error:
    ShortenRange = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: CompareColumnsWithConditionalFormatting
' Purpose: Applies conditional formatting rules to two selected ranges for compare them
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub CompareColumnsWithConditionalFormatting(control As IRibbonControl)
    On Error GoTo CompareColumnsWithConditionalFormatting_Error
    If Selection.Areas.Count <> 2 Then
        Exit Sub
    End If
    Dim C1 As String, C2 As String
    InputStringDialog.Caption = "Data type"
    InputStringDialog.DialogDescription.Caption = "Enter data type (1 - number, 2 - text)"
    InputStringDialog.InputString = CStr(ComparedDataType)
    InputStringDialog.InputString.SetFocus
    InputStringDialog.InputString.SelStart = 0
    InputStringDialog.InputString.SelLength = Len(InputStringDialog.InputString.Text)
    Dim Result As Variant
    InputStringDialog.Show 1
    Result = InputStringDialog.DialogResult
    If Result = 0 Then
        Unload InputStringDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputStringDialog.InputString.Text
    If IsEmpty(Result2) Or Result2 = "" Then
        Unload InputStringDialog
        Exit Sub
    End If
    ComparedDataType = CInt(Result2)
    C1 = Selection.Areas(1).Cells(1).Address(False, False, xlA1, False, False)
    C2 = Selection.Areas(2).Cells(1).Address(False, False, xlA1, False, False)
    If ComparedDataType = 2 Then
        With Selection.Areas(1)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
            .FormatConditions(1).Interior.ColorIndex = 38
            .FormatConditions(1).Font.ColorIndex = 9
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "="""")"
            .FormatConditions(2).Interior.ColorIndex = 23
            .FormatConditions(2).Font.ColorIndex = 1
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "=""""," & C2 & "<>"""")"
            .FormatConditions(3).Interior.ColorIndex = 33
            .FormatConditions(3).Font.ColorIndex = 1
        End With
    Else
        If ComparedDataType = 1 Then
            With Selection.Areas(1)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & ">" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
                .FormatConditions(1).Interior.ColorIndex = 38
                .FormatConditions(1).Font.ColorIndex = 9
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
                .FormatConditions(2).Interior.ColorIndex = 36
                .FormatConditions(2).Font.ColorIndex = 53
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "="""")"
                .FormatConditions(3).Interior.ColorIndex = 23
                .FormatConditions(3).Font.ColorIndex = 1
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "=""""," & C2 & "<>"""")"
                .FormatConditions(4).Interior.ColorIndex = 33
                .FormatConditions(4).Font.ColorIndex = 1
            End With
        End If
    End If
    On Error GoTo 0
    Exit Sub
CompareColumnsWithConditionalFormatting_Error:
    Unload InputStringDialog
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine CompareColumnsWithConditionalFormatting, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: HighlightBlankCells
' Purpose: Sets color 8 for empty cells background
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 01.10.2020
' ----------------------------------------------------------------
Sub HighlightBlankCells(control As IRibbonControl)
    On Error GoTo HighlightBlankCells_Error
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then Set TargetRange = Selection Else Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    For Each i In TargetRange
        If IsEmpty(i) Then i.Interior.ColorIndex = 8
    Next
    On Error GoTo 0
    Exit Sub
HighlightBlankCells_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine HighlightBlankCells, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: HighlightCellsWithFormulasReturningVoid
' Purpose: Sets color 8 for cells contained formulas that return empty values
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 01.10.2020
' ----------------------------------------------------------------
Sub HighlightCellsWithFormulasReturningVoid(control As IRibbonControl)
    On Error GoTo HighlightCellsWithFormulasReturningVoid_Error
    Dim i As Range
    Dim TargetRange As Range
    If Selection.Count = 1 Then Set TargetRange = Selection Else Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    For Each i In TargetRange
        If i.HasFormula And CStr(i.Value) = "" Then i.Interior.ColorIndex = 14
    Next
    On Error GoTo 0
    Exit Sub
HighlightCellsWithFormulasReturningVoid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine HighlightCellsWithFormulasReturningVoid, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: SwapCells
' Purpose: Swaps ranges. It is necessary to select two areas (ranges) of the same size using the Ctrl key and run this procedure
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 28.10.2020
' ----------------------------------------------------------------
Sub SwapCells(control As IRibbonControl)
    On Error GoTo SwapCells_Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Dim Area1 As Range
    Dim Area2 As Range
    Dim r As Variant
    If Selection.Areas.Count <> 2 Then
        MsgBox "It is necessary to select two ranges of lines that need to be swapped." & vbCrLf & _
            "Ranges can contain more than 1 line. " & vbCrLf & _
            "Ranges currently selected: " & Selection.Areas.Count, 16, "Select two ranges"
        Exit Sub
    End If
    If Selection.Areas(1).Columns.Count <> Selection.Areas(2).Columns.Count Or _
        Selection.Areas(1).Rows.Count <> Selection.Areas(2).Rows.Count Then
        MsgBox "It is necessary to select two areas (ranges) of the same size", 16, "Select ranges of the same size."
        Exit Sub
    End If
    Set Area1 = Selection.Areas(1)
    Set Area2 = Selection.Areas(2)
    r = Area1.Value
    Area1.Value = Area2.Value
    Area2.Value = r
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub
SwapCells_Error:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine SwapCells, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ReverseOrderList
' Purpose: Rearranges list elements in reverse order
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 16.03.2021
' ----------------------------------------------------------------
Sub ReverseOrderList(control As IRibbonControl)
    On Error GoTo ReverseOrderList_Error
    Dim arrData(), n As Long
    Dim i As Range
    Dim Idx
    Dim TargetRange As Range
    If Selection.Count = 1 Then
        Set TargetRange = Selection
    Else
        Set TargetRange = Selection.SpecialCells(xlCellTypeVisible)
    End If
    If TargetRange.Count < 2 Or TargetRange.Areas.Count > 1 Then Exit Sub
    arrData = TargetRange.FormulaLocal
    Dim Result, Result2
    If TargetRange.Columns.Count > 1 And TargetRange.Rows.Count > 1 Then
        Dim k
        k = TargetRange.Columns.Count
        n = 0
        For Each i In TargetRange
            Idx = TargetRange.Rows.Count - n
            i.FormulaLocal = arrData(Idx, k)
            If k = 1 Then
                k = TargetRange.Columns.Count
                n = n + 1
            Else
                k = k - 1
            End If
        Next i
    ElseIf TargetRange.Columns.Count > 1 And TargetRange.Rows.Count = 1 Then
        For Each i In TargetRange
            Idx = UBound(arrData, 2) - n
            i.FormulaLocal = arrData(1, Idx)
            n = n + 1
        Next i
    ElseIf TargetRange.Columns.Count = 1 And TargetRange.Rows.Count > 1 Then
        For Each i In TargetRange
            Idx = UBound(arrData, 1) - n
            i.FormulaLocal = arrData(Idx, 1)
            n = n + 1
        Next i
    End If
    On Error GoTo 0
    Exit Sub
ReverseOrderList_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine ReverseOrderList, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ListRange
' Purpose: Generates string with numerical list using comma+space separator
' Procedure Kind: Function
' Procedure Access: Public
' Parameter EndVal (): Last item in list
' Parameter BeginVal (): First item in list
' Parameter ListSeparator (): List item delimiter
' Parameter ListStep (): Step between items in list
' Return Type: String
' Author: Petr Kovalenko
' Date: 06.02.2022
' ----------------------------------------------------------------
Public Function ListRange(Optional EndVal = 1, Optional BeginVal = 1, Optional ListSeparator = ", ", Optional ListStep = 1) As String
    On Error GoTo ListRange_Error
    Dim sResult As String
    sResult = ""
    Application.Volatile True
    Dim Element
    For Element = BeginVal To EndVal Step ListStep
        If sResult <> "" Then sResult = sResult & ListSeparator & Element Else sResult = "'" & Element
    Next Element
    If sResult <> "" Then ActiveCell.NumberFormatLocal = "General"
    ListRange = sResult
    On Error GoTo 0
    Exit Function
ListRange_Error:
    ListRange = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: BuildShortenRange
' Purpose: Generates string with numerical range (from-to)
' Procedure Kind: Function
' Procedure Access: Public
' Parameter EndVal (): Last item in list
' Parameter BeginVal (): First item in list
' Return Type: String
' Author: Petr Kovalenko
' Date: 06.02.2022
Public Function BuildShortenRange(Optional EndVal = 1, Optional BeginVal = 1) As String
    On Error GoTo BuildShortenRange_Error
    Dim sResult As String
    sResult = ""
    Application.Volatile True
    sResult = "'" & CStr(BeginVal) & "-" & CStr(EndVal)
    If sResult <> "" Then ActiveCell.NumberFormatLocal = "General"
    BuildShortenRange = sResult
    On Error GoTo 0
    Exit Function
BuildShortenRange_Error:
    BuildShortenRange = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: RepeatConditionalFormatting
' Purpose: Applies conditional formatting rules to two selected ranges for compare them
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter control (IRibbonControl):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub RepeatConditionalFormatting(control As IRibbonControl)
    On Error GoTo RepeatConditionalFormatting_Error
    If Selection.Areas.Count <> 2 Then
        Exit Sub
    End If
    Dim C1 As String, C2 As String
    InputStringDialog.Caption = "Data tyoe"
    InputStringDialog.DialogDescription.Caption = "Enter data type (1 - number, 2 - text)"
    InputStringDialog.InputString = CStr(ComparedDataType)
    InputStringDialog.InputString.SetFocus
    InputStringDialog.InputString.SelStart = 0
    InputStringDialog.InputString.SelLength = Len(InputStringDialog.InputString.Text)
    Dim Result As Variant
    InputStringDialog.Show 1
    Result = InputStringDialog.DialogResult
    If Result = 0 Then
        Unload InputStringDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputStringDialog.InputString.Text
    If IsEmpty(Result2) Or Result2 = "" Then
        Unload InputStringDialog
        Exit Sub
    End If
    ComparedDataType = CInt(Result2)
    C1 = Selection.Areas(1).Cells(1).Address(False, False, xlA1, False, False)
    C2 = Selection.Areas(2).Cells(1).Address(False, False, xlA1, False, False)
    If ComparedDataType = 2 Then
        With Selection.Areas(1)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
            .FormatConditions(1).Interior.ColorIndex = 38
            .FormatConditions(1).Font.ColorIndex = 9
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "="""")"
            .FormatConditions(2).Interior.ColorIndex = 23
            .FormatConditions(2).Font.ColorIndex = 1
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "=""""," & C2 & "<>"""")"
            .FormatConditions(3).Interior.ColorIndex = 33
            .FormatConditions(3).Font.ColorIndex = 1
        End With
    Else
        If ComparedDataType = 1 Then
            With Selection.Areas(1)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & ">" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
                .FormatConditions(1).Interior.ColorIndex = 38
                .FormatConditions(1).Font.ColorIndex = 9
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<" & C2 & "," & C1 & "<>""""," & C2 & "<>"""")"
                .FormatConditions(2).Interior.ColorIndex = 36
                .FormatConditions(2).Font.ColorIndex = 53
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "<>""""," & C2 & "="""")"
                .FormatConditions(3).Interior.ColorIndex = 23
                .FormatConditions(3).Font.ColorIndex = 1
                .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & C1 & "<>" & C2 & "," & C1 & "=""""," & C2 & "<>"""")"
                .FormatConditions(4).Interior.ColorIndex = 33
                .FormatConditions(4).Font.ColorIndex = 1
            End With
        End If
    End If
    On Error GoTo 0
    Exit Sub
RepeatConditionalFormatting_Error:
    Unload InputStringDialog
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine RepeatConditionalFormatting, line " & Erl & "."
End Sub

Public Sub EditSmartTableRowInDlgWnd(control As IRibbonControl)
    On Error GoTo EditSmartTableRowInDlgWnd_Error
    Dim SelectedCell As Range
    Dim TableName As String
    Dim ActiveTable As ListObject
    Set SelectedCell = ActiveCell
    TableName = SelectedCell.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    If ActiveTable Is Nothing Then
        Exit Sub
    End If
    Dim arrData()
    ReDim Preserve arrData(ActiveTable.HeaderRowRange.Count - 1, 1)
    Dim i
    For i = 0 To ActiveTable.HeaderRowRange.Count - 1
        arrData(i, 0) = ActiveTable.HeaderRowRange.Cells(i + 1).FormulaLocal
    Next i
    Dim NFL, NF, Val, FL, Text
    For i = 0 To ActiveTable.HeaderRowRange.Count - 1
        NFL = ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).NumberFormatLocal
        NF = ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).NumberFormat
        Val = ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).Value
        Text = ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).Text
        FL = ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).FormulaLocal
        If ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).HasFormula Then
            arrData(i, 1) = FL
        Else
            arrData(i, 1) = Text
        End If
    Next i
    EditSmartTableRow.Caption = "Editing the current table row"
    EditSmartTableRow.DialogDescription.Caption = "Select line in list containing the pair: ""column"" - ""value"" of the current table row. Editing the value by double-clicking."
    EditSmartTableRow.Label1.Caption = "Key"
    EditSmartTableRow.Label2.Caption = "Value"
    EditSmartTableRow.ListBox1.List = arrData
    EditSmartTableRow.ListBox1.SetFocus
    Dim Result
    EditSmartTableRow.Show 1
    Result = EditSmartTableRow.DialogResult
    If Result = 0 Then
        Unload EditSmartTableRow
        Exit Sub
    End If
    Dim Result2
    Result2 = EditSmartTableRow.ListBox1.List
    If UBound(Result2) <> ActiveTable.HeaderRowRange.Count - 1 Then
        Unload EditSmartTableRow
        Exit Sub
    End If
    For i = 0 To ActiveTable.HeaderRowRange.Count - 1
        ActiveTable.DataBodyRange.Rows(SelectedCell.Row - SelectedCell.ListObject.DataBodyRange.Row + 1).Cells(i + 1).FormulaLocal = Result2(i, 1)
    Next i
    On Error GoTo 0
    Exit Sub
EditSmartTableRowInDlgWnd_Error:
    Unload InputStringDialog
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine EditSmartTableRowInDlgWnd, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: ExtractNote
' Purpose: Extract notes from selected range
' Procedure Kind: Function
' Procedure Access: Public
' Parameter RangeWithNotes (Range): Range of cells with notes
' Return Type: String
' Author: Petr Kovalenko
' Date: 06.02.2022
' ----------------------------------------------------------------
Public Function ExtractNote(ByVal RangeWithNotes As Range) As String
    On Error GoTo ExtractNote_Error
    Application.Volatile True
    If RangeWithNotes Is Nothing Then Exit Function
    Dim i As Range
    Dim ResultData As Variant
    Dim TargetRange As Range
    If RangeWithNotes.Count = 1 Then
        Set TargetRange = RangeWithNotes
    Else
        Set TargetRange = RangeWithNotes.SpecialCells(xlCellTypeVisible)
    End If
    For Each i In TargetRange
        If Not i.Comment Is Nothing Then
            Debug.Print i.Comment.Text
            If ResultData <> "" Then
                ResultData = ResultData & ";" & i.Comment.Text
            Else
                ResultData = i.Comment.Text
            End If
        End If
    Next
    ExtractNote = ResultData
    On Error GoTo 0
    Exit Function
ExtractNote_Error:
    ExtractNote = CVErr(xlErrValue)
End Function
