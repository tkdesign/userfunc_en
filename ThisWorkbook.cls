Option Explicit

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
 
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpString As Any, ByVal lpFileName As String) As Long

' ----------------------------------------------------------------
' Procedure Name: Workbook_Open
' Purpose: Add-in connection event handler (when starting Excel or through the add-in selection dialog)
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Petr Kovalenko
' Date: 29.01.2021
' ----------------------------------------------------------------
Private Sub Workbook_Open()
    SetupVariables
End Sub

' ----------------------------------------------------------------
' Procedure Name: Workbook_BeforeClose
' Purpose: Add-in disable event handler (when exiting Excel or through the add-in selection dialog)
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter Cancel (Boolean):
' Author: Petr Kovalenko
' Date: 29.01.2021
' ----------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SaveChangedVariables
End Sub

' ----------------------------------------------------------------
' Procedure Name: SetupVariables
' Purpose: Initializes global variables from a configuration file
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Petr Kovalenko
' Date: 29.01.2021
' ----------------------------------------------------------------
Private Sub SetupVariables()
    On Error GoTo SetupVariables_Error
    Dim IsChanged As Boolean
    Dim IntRet
    Dim StrRet As String
    Dim FilePath
    IsChanged = False
    FilePath = ThisWorkbook.Path + "\" + ThisWorkbook.Name + ".ini"
    Err.Clear
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("TextToClipboardSeparatorSelection", "Separator", ", ", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.Separator = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("TextToClipboardSeparatorSelection", "WithoutRepeat", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    MainMod.WithoutRepeat = CInt(StrRet)
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("CompareColumnsWithConditionalFormatting", "ComparedDataType", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    MainMod.ComparedDataType = CInt(StrRet)
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("MergeCellsAndText", "MergeCellsSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.MergeCellsSeparator = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("FilledCellAdressToClipboard", "CellsAddressSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.CellsAddressSeparator = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("FormulaToClipboard", "CopyFormulaSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.CopyFormulaSeparator = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpPattern", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.RegExpPattern = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "ReplacementTemplate", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.ReplacementTemplate = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpPattern2", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    MainMod.RegExpPattern2 = StrRet
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpMatchNumber", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    MainMod.RegExpMatchNumber = CInt(StrRet)
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("WrapRound", "RoundPrecision", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    MainMod.RoundPrecision = CInt(StrRet)
    On Error GoTo 0
    Exit Sub
SetupVariables_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine SetupVariables, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: SaveChangedVariables
' Purpose: Saves global variables to a configuration file
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Petr Kovalenko
' Date: 29.01.2021
' ----------------------------------------------------------------
Private Sub SaveChangedVariables()
    On Error GoTo SaveChangedVariables_Error
    Dim IsChanged As Boolean
    Dim IntRet
    Dim StrRet As String
    Dim FilePath
    IsChanged = False
    FilePath = ThisWorkbook.Path + "\" + ThisWorkbook.Name + ".ini"
    Err.Clear
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("TextToClipboardSeparatorSelection", "Separator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.Separator <> StrRet Then
        IntRet = WritePrivateProfileString("TextToClipboardSeparatorSelection", "Separator", CStr(Chr(34)) & MainMod.Separator & CStr(Chr(34)), FilePath)
    End If
    If MainMod.WithoutRepeat <> 1 And MainMod.WithoutRepeat <> 0 Then MainMod.WithoutRepeat = 1
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("TextToClipboardSeparatorSelection", "WithoutRepeat", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    If MainMod.WithoutRepeat <> CInt(StrRet) Then
        IntRet = WritePrivateProfileString("TextToClipboardSeparatorSelection", "WithoutRepeat", CStr(MainMod.WithoutRepeat), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("CompareColumnsWithConditionalFormatting", "ComparedDataType", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    If MainMod.ComparedDataType <> CInt(StrRet) Then
        IntRet = WritePrivateProfileString("CompareColumnsWithConditionalFormatting", "ComparedDataType", CStr(MainMod.ComparedDataType), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("MergeCellsAndText", "MergeCellsSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.MergeCellsSeparator <> StrRet Then
        IntRet = WritePrivateProfileString("MergeCellsAndText", "MergeCellsSeparator", CStr(Chr(34)) & MainMod.MergeCellsSeparator & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("FilledCellAdressToClipboard", "CellsAddressSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.CellsAddressSeparator <> StrRet Then
        IntRet = WritePrivateProfileString("FilledCellAdressToClipboard", "CellsAddressSeparator", CStr(Chr(34)) & MainMod.CellsAddressSeparator & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("FormulaToClipboard", "CopyFormulaSeparator", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.CopyFormulaSeparator <> StrRet Then
        IntRet = WritePrivateProfileString("FormulaToClipboard", "CopyFormulaSeparator", CStr(Chr(34)) & MainMod.CopyFormulaSeparator & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpPattern", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.RegExpPattern <> StrRet Then
        IntRet = WritePrivateProfileString("RegExp", "RegExpPattern", CStr(Chr(34)) & MainMod.RegExpPattern & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "ReplacementTemplate", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.ReplacementTemplate <> StrRet Then
        IntRet = WritePrivateProfileString("RegExp", "ReplacementTemplate", CStr(Chr(34)) & MainMod.ReplacementTemplate & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpPattern2", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If MainMod.RegExpPattern2 <> StrRet Then
        IntRet = WritePrivateProfileString("RegExp", "RegExpPattern2", CStr(Chr(34)) & MainMod.RegExpPattern2 & CStr(Chr(34)), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("RegExp", "RegExpMatchNumber", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    If MainMod.RegExpMatchNumber <> CInt(StrRet) Then
        IntRet = WritePrivateProfileString("RegExp", "RegExpMatchNumber", CStr(MainMod.RegExpMatchNumber), FilePath)
    End If
    StrRet = String(255, Chr(0))
    IntRet = GetPrivateProfileString("WrapRound", "RoundPrecision", "", StrRet, 255, FilePath)
    StrRet = Left$(StrRet, IntRet)
    If StrRet = vbNullString Then StrRet = "0"
    If MainMod.RoundPrecision <> CInt(StrRet) Then
        IntRet = WritePrivateProfileString("WrapRound", "RoundPrecision", CStr(MainMod.RoundPrecision), FilePath)
    End If
    On Error GoTo 0
    Exit Sub
SaveChangedVariables_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine SaveChangedVariables, line " & Erl & "."
End Sub

Private Sub DescribeFunction()
    On Error GoTo DescribeFunction_Error
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc()
    Dim i As Integer
    Dim j As Integer
    Dim countarg As Integer
    For i = 2 To 98
        FuncName = Application.ThisWorkbook.Worksheets(2).Cells(i, 1)
        If IsNull(FuncName) Or IsEmpty(FuncName) Or FuncName = "" Then Exit Sub
        FuncDesc = Application.ThisWorkbook.Worksheets(2).Cells(i, 2)
        Category = Application.ThisWorkbook.Worksheets(2).Cells(i, 3)
        countarg = Application.ThisWorkbook.Worksheets(2).Cells(i, 5)
        If countarg > 0 Then
            ReDim ArgDesc(1 To countarg)
            For j = 1 To countarg
                ArgDesc(j) = Application.ThisWorkbook.Worksheets(2).Cells(i, 5 + j * 2)
            Next
            Application.MacroOptions Macro:=FuncName, Description:=FuncDesc, Category:=Category, ArgumentDescriptions:=ArgDesc
        Else
            Application.MacroOptions Macro:=FuncName, Description:=FuncDesc, Category:=Category
        End If
    Next
    On Error GoTo 0
    Exit Sub
DescribeFunction_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in subroutine DescribeFunction, line " & Erl & "."
End Sub

' ----------------------------------------------------------------
' Procedure Name: QuoteParsing
' Purpose: Checks the source string (*second function argument) for quotes. If the value is enclosed in double quotes, returns the unquoted text. Otherwise, and also if source string length is less than two characters, returns the default value (**first function argument)
' Procedure Kind: Function
' Procedure Access: Private
' Parameter DefaultValue (String): Default string
' Parameter StrRet (String): Source string
' Parameter IntRet (Long): String length
' Return Type: String
' Author: Petr Kovalenko
' Date: 29.01.2021
' ----------------------------------------------------------------
Private Function QuoteParsing(ByVal DefaultValue As String, ByVal StrRet As String, ByVal IntRet As Long) As String
    On Error GoTo QuoteParsing_Error
    If IntRet > 2 Then
        If Asc(Left(StrRet, 1)) = Chr(34) And Asc(Left(StrRet, IntRet)) = Chr(34) Then
            StrRet = Mid(StrRet, 2, IntRet - 1)
        Else
            StrRet = StrRet
        End If
    ElseIf IntRet = 2 Then
        If Asc(Left(StrRet, 1)) = Chr(34) And Asc(Left(StrRet, IntRet)) = Chr(34) Then
            StrRet = ""
        Else
            StrRet = StrRet
        End If
    End If
    QuoteParsing = StrRet
    On Error GoTo 0
    Exit Function
QuoteParsing_Error:
    QuoteParsing = CVErr(xlErrValue)
End Function

Sub LogInformation(LogMessage As String)
    Dim LogFileName
    LogFileName = ThisWorkbook.Path + "\" + ThisWorkbook.Name + ".log"
    Dim FileNum As Integer
    FileNum = FreeFile
    Open LogFileName For Append As #FileNum
    Print #FileNum, LogMessage
    Close #FileNum
End Sub