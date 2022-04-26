Attribute VB_Name = "AppMain"
Rem
Rem @appname FormatCopyBlocker - 書式コピー禁止アドイン
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2022/04/26 初回版
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "書式コピー禁止アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.01"
Public Const APP_UPDATE = "2022/04/26"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/format_copy_blocker"

'--------------------------------------------------
'アドイン実行時
Sub AddinStart()
    Call Start_FormatCopyBlocker
    MsgBox "書式のコピー絶対ゆるさない", _
                vbExclamation + vbOKOnly, ThisWorkbook.Name
End Sub

'アドイン一時停止時
Sub AddinStop(): Call Stop_FormatCopyBlocker: End Sub

'アドイン設定表示
'Sub AddinConfig(): Call SettingForm.Show: End Sub
Sub AddinConfig(): MsgBox "設定画面表示用": End Sub

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'アドインを止めたい時に使うプロシージャ
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

Sub MergePrint(): MsgBox "Print": End Sub
'--------------------------------------------------

'監視開始
'Workbook_Openから呼ばれる
'他ブックの上書き保存を検知するために使用される
Sub Start_FormatCopyBlocker(): Call SubFormatCopyBlocker(False): End Sub

'監視停止
Sub Stop_FormatCopyBlocker(): Call SubFormatCopyBlocker(True):  End Sub

Sub SubFormatCopyBlocker(IsStop As Boolean)
    Static inst As FormatCopyBlocker
    
    If IsStop Then
        Set inst = Nothing
        Exit Sub
    End If
    
    Set inst = FormatCopyBlocker.Init(Application)
End Sub

Rem --------------------------------------------------------------------------------

Rem 配列の次元数を求める
Public Function GetArrayDimension_NoAPI(ByRef arr As Variant) As Long
    On Error GoTo ENDPOINT
    Dim i As Long, tmp As Long
    For i = 1 To 61
        tmp = LBound(arr, i)
    Next
    GetArrayDimension_NoAPI = 0
    Exit Function
    
ENDPOINT:
    GetArrayDimension_NoAPI = i - 1
End Function

Rem データをTSVに変換する
Function JoinTsv(v) As String
    Select Case GetArrayDimension_NoAPI(v)
        Case 0: JoinTsv = v
        Case 1: JoinTsv = Join(v, vbTab)
        Case 2: JoinTsv = Join2(v, vbTab, vbLf)
        Case Else: JoinTsv = v.Value
    End Select
End Function

Rem 二次元配列をCSV等の文字列に変換する
Public Function Join2(arr As Variant, _
                        Optional ByVal Delimiter1 As String = vbTab, _
                        Optional ByVal Delimiter2 As String = vbCrLf) As String
    Dim i As Long, j As Long
    Dim arr1() As Variant
    Dim arr2() As Variant
    
    If GetArrayDimension_NoAPI(arr) <> 2 Then Err.Raise 9999, "Join2", "Join2 : 入力変数Arrが二次元配列ではありません。"
    
    ReDim arr1(LBound(arr, 1) To UBound(arr, 1))
    ReDim arr2(LBound(arr, 2) To UBound(arr, 2))
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr2(j) = CStr(arr(i, j))
        Next
        arr1(i) = VBA.Strings.Join(arr2, Delimiter1)
    Next
    Join2 = VBA.Strings.Join(arr1, Delimiter2)
End Function
