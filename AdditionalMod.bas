Attribute VB_Name = "AdditionalMod"
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
Private Declare PtrSafe Function MsgBoxAlias Lib "USER32" Alias "MessageBoxW" (Optional ByVal hWnd As LongPtr, Optional ByVal PROMPT As LongPtr, Optional ByVal Caption As LongPtr, Optional ByVal BUTTONS As LongPtr) As Long

Private Declare PtrSafe Function RegisterClipboardFormat Lib "USER32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As LongPtr
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As String, ByRef pclsid As Any) As LongPtr
Private Declare PtrSafe Function OleLoadFromStream Lib "ole32" (ByVal pStm As Any, ByVal iidInterface As Any, ByRef ppvObj As Any) As LongPtr
Private Declare PtrSafe Function CreateBindCtx Lib "ole32" (ByVal reserved As LongPtr, ByRef ppbc As Any) As LongPtr
Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (Optional ByVal psz As String) As String
Private Declare PtrSafe Function CoGetMalloc Lib "ole32" (ByVal dwMemContext As LongPtr) As LongPtr

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Private Declare PtrSafe Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
 
' ----------------------------------------------------------------
' Procedure Name: KBDToENG
' Purpose: Switches to English keyboard layout
' Procedure Kind: Function
' Procedure Access: Public
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Function KBDToENG()
    On Error Resume Next
    Call LoadKeyboardLayout("00000409", &H1)
End Function
 
' ----------------------------------------------------------------
' Procedure Name: KBDToRUS
' Purpose: Switches to Russian keyboard layout
' Procedure Kind: Function
' Procedure Access: Public
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Function KBDToRUS()
    On Error Resume Next
    Call LoadKeyboardLayout("00000419", &H1)
End Function

' ----------------------------------------------------------------
' Procedure Name: GetFileName
' Purpose: Gets the filename from the system path
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strFilePath (): The full path to the file
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Function GetFileName(ByVal strFilePath) As String
    On Error GoTo GetFileName_Error
    Dim intPos%
    intPos = InStrRev(strFilePath, "/")
    GetFileName = Right(strFilePath, Len(strFilePath) - intPos)
    On Error GoTo 0
    Exit Function
GetFileName_Error:
    GetFileName = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: GetFilePath
' Purpose: Gets paths without a filename
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strFilePath (): The full path to the file
' Return Type: String
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Function GetFilePath(ByVal strFilePath) As String
    On Error GoTo GetFilePath_Error
    Dim intPos%
    intPos = InStrRev(strFilePath, "/")
    GetFilePath = Left(strFilePath, Len(strFilePath) - (Len(strFilePath) - intPos))
    On Error GoTo 0
    Exit Function
GetFilePath_Error:
    GetFilePath = CVErr(xlErrValue)
End Function

' ----------------------------------------------------------------
' Procedure Name: TransposeArray
' Purpose: Custom function for array transposition
' Procedure Kind: Function
' Procedure Access: Public
' Parameter ArraySrc (Variant): Array to transpose
' Return Type: Variant
' Author: Petr Kovalenko
' Date: 29.10.2020
' ----------------------------------------------------------------
Public Function TransposeArray(ByVal ArraySrc As Variant) As Variant
    Dim tempArray As Variant
    Dim x, y
    ReDim tempArray(LBound(ArraySrc, 2) To UBound(ArraySrc, 2), LBound(ArraySrc, 1) To UBound(ArraySrc, 1))
    For x = LBound(ArraySrc, 2) To UBound(ArraySrc, 2)
        For y = LBound(ArraySrc, 1) To UBound(ArraySrc, 1)
            tempArray(x, y) = ArraySrc(y, x)
        Next y
    Next x
    TransposeArray = tempArray
End Function

' ----------------------------------------------------------------
' Procedure Name: ClipBoard_SetData
' Purpose: Copies a string to the clipboard
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter MyString (String):
' Author: Petr Kovalenko
' Date: 23.09.2020
' ----------------------------------------------------------------
Sub ClipBoard_SetData(MyString As String)
    On Error Resume Next
    '32-bit code by Microsoft: http://msdn.microsoft.com/en-us/library/office/ff192913.aspx
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr, x As Long
    ' Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)
    ' Lock the block to get a far pointer to this memory
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    ' Copy the string to this global memory
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)
    ' Unlock the memory
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        'Debug.Print "GlobalFree returned: " & CStr(GlobalFree(hGlobalMemory))
        GoTo OutOfHere
    End If
    ' Open the Clipboard to copy data to
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Sub
    End If
    ' Clear the Clipboard
    x = EmptyClipboard()
    ' Copy the data to the Clipboard
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
OutOfHere:
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Sub

Function MsgBoxU(ByVal PROMPT As String, Optional ByVal BUTTONS As VbMsgBoxStyle, Optional ByVal TITLE As String = vbNullChar) As VbMsgBoxResult
        MsgBoxU = MsgBoxAlias(Application.hWnd, StrPtr(PROMPT), StrPtr(TITLE), BUTTONS)
End Function
