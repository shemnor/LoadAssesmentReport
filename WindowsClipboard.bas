Attribute VB_Name = "WindowsClipboard"
Option Explicit

'Handle 64-bit and 32-bit Office
#If VBA7 Then
  Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
  Public Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
  Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
  Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
  Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
  Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpFormatName As String) As Long
#Else
  Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Public Declare Function CloseClipboard Lib "user32" () As Long
  Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Public Declare Function EmptyClipboard Lib "user32" () As Long
  Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
  Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpFormatName As String) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const CF_NAME_RICHTEXTFORMAT = "Rich Text Format"
 
'Copy text to clipboard using Windows API
'Source: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx
Function ClipBoardSetText(text As String, textFormat As Long) As Boolean
#If VBA7 Then
    Dim hGlobal As LongPtr, lpGlobalMemory As LongPtr
#Else
    Dim hGlobal As Long, lpGlobalMemory As Long
#End If

    hGlobal = GlobalAlloc(GHND, Len(text) + 1)
    lpGlobalMemory = GlobalLock(hGlobal)
    If lpGlobalMemory = 0 Then
        Exit Function
    End If
    lstrcpy lpGlobalMemory, text

    If GlobalUnlock(hGlobal) <> 0 Then
        Exit Function
    End If

    If OpenClipboard(0&) = 0 Then
        Exit Function
    End If

    EmptyClipboard
    If SetClipboardData(textFormat, hGlobal) <> 0 Then
        ClipBoardSetText = True
    End If
    CloseClipboard
    
End Function

