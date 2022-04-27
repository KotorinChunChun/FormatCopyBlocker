VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncClipboard
Rem
Rem  @description   クリップボード操作系
Rem
Rem  @update        2022/04/27
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Excel.Application
Rem    Microsoft Forms 2.0 Object Library
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "User32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long

Private Declare PtrSafe Function GlobalLock Lib "Kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "Kernel32" (ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function lstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)
#End If

Rem 定数宣言
Private Const GMEM_MOVEABLE         As Long = &H2
Private Const GMEM_ZEROINIT         As Long = &H40
Private Const GHND                  As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const CF_TEXT               As Long = 1
Private Const CF_OEMTEXT            As Long = 7

Rem 指定文字列をクリップボードに保存
Public Function SetClipboardByTextAPI(strData As String, Optional wFormat As Long = CF_TEXT)
    
#If VBA7 And Win64 Then
    Dim lngHwnd As LongPtr
    Dim lngMem As LongPtr
    Dim lngDataLen As LongPtr
    Dim lngRet As LongPtr
#Else
    Dim lngHwnd As Long
    Dim lngMem As Long
    Dim lngDataLen As Long
    Dim lngRet As Long
#End If
    Const GMEM_MOVEABLE = 2
    
    Dim blnErrflg As Boolean: blnErrflg = True
    
    If OpenClipboard(0&) <> 0 Then
        If EmptyClipboard() <> 0 Then
            lngDataLen = LenB(strData) + 1
            lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
            If lngHwnd <> 0 Then
                lngMem = GlobalLock(lngHwnd)
                If lngMem <> 0 Then
                    If lstrcpy(lngMem, strData) <> 0 Then
                        lngRet = SetClipboardData(wFormat, lngHwnd)
                        blnErrflg = False
                    End If
                    lngRet = GlobalUnlock(lngHwnd)
                End If
            End If
        End If
        lngRet = CloseClipboard()
    End If
    
    SetClipboardByTextAPI = blnErrflg
    
End Function

Rem クリップボードの文字列を取得
Rem  失敗時はエラーを無視して""を返す
Public Function GetTextByClipboardTextDataObject() As String
    Dim cb As New DataObject
    cb.GetFromClipboard
    On Error Resume Next
    GetTextByClipboardTextDataObject = cb.GetText
    On Error GoTo 0
    'Debug.Print CB.GetText
End Function

Rem コピー中のセルアドレスを取得
Public Function GetAddressByClipboardCells(SheetName As Variant) As String
    
#If VBA7 And Win64 Then
    Dim lngHwnd As LongPtr
    Dim lngDataLen As LongPtr
    Dim p As LongPtr
#Else
    Dim lngHwnd As Long
    Dim lngDataLen As Long
    Dim p As Long
#End If
    
    If OpenClipboard(0&) = 0 Then Exit Function
    
    lngHwnd = GetClipboardData(RegisterClipboardFormat("Link"))
    If lngHwnd = 0 Then GoTo ExitFunctionCloseClipboard
    
    lngDataLen = GlobalSize(lngHwnd)
    p = GlobalLock(lngHwnd)
    Dim data() As Byte
    ReDim data(0 To CLng(lngDataLen) - CLng(1))
#If VBA7 And Win64 Then
    Call MoveMemory(data(0), ByVal p, lngDataLen)
#Else
    Call MoveMemory(CLng(VarPtr(data(0))), p, lngDataLen)
#End If
    Call GlobalUnlock(lngHwnd)
    
    Call CloseClipboard
    
    Dim i As Long
    For i = 0 To CLng(lngDataLen) - CLng(1)
        If data(i) = 0 Then
            data(i) = Asc(" ")
        End If
    Next
    
    Rem ※複数セルには対応していない
    Dim Address As String
    Address = Trim(StrConv(data, vbUnicode))
Rem Debug.Print "Address: " + Address
Rem     If InStr(Address, "]" & SheetName) <> 0 Then
Rem         GetAddressByClipboardCells = Trim(Replace(Mid(Address, InStr(Address, "]" & SheetName)), "]" & SheetName, ""))
Rem     Else
Rem         GetAddressByClipboardCells = ""
Rem     End If
    GetAddressByClipboardCells = Address
    Exit Function
    
ExitFunctionCloseClipboard:
    Call CloseClipboard
    GetAddressByClipboardCells = ""
End Function

Rem コピー中のセル範囲Rangeオブジェクトを取得
Public Function GetRangeByClipboardCells() As Range
    
    Rem "Excel [Book2]Sheet1 R7C16:R15C20"
    Dim ssText As String
    ssText = GetAddressByClipboardCells(Excel.ActiveSheet.Name)
    If ssText = "" Then
        Set GetRangeByClipboardCells = Nothing
        Exit Function
    End If
    
    Rem ブック名：左から[と右から]までの間
    Dim BookName As String
    ssText = Right(ssText, Len(ssText) - InStr(ssText, "["))
    BookName = Left(ssText, InStrRev(ssText, "]") - 1)
    
    Rem シート名：右から]から までの間
    Dim SheetName As Variant
    ssText = Right(ssText, Len(ssText) - 1 - Len(BookName))
    SheetName = Left(ssText, InStrRev(ssText, " ") - 1)
    
    Rem R1C1セル：右から 以降
    Dim CellName As String
    ssText = Right(ssText, Len(ssText) - 1 - Len(SheetName))
    CellName = Application.ConvertFormula(ssText, xlR1C1, xlA1)
    
    Set GetRangeByClipboardCells = Workbooks(BookName).Worksheets(SheetName).Range(CellName)
        
End Function

Rem クリップボードの内容が空か
Function IsEmptyCB() As Boolean
    IsEmptyCB = (Application.ClipboardFormats(1) = -1)
End Function

Rem クリップボードの内容がプレーンテキストのみか
Function IsTextOnlyCB() As Boolean
    If UBound(Application.ClipboardFormats, 1) = 2 Then
        IsTextOnlyCB = (Application.ClipboardFormats(1) = 0) And _
                        (Application.ClipboardFormats(2) = 44)
    End If
End Function
