Attribute VB_Name = "AppMain"
Rem
Rem @appname FormatCopyBlocker - �����R�s�[�֎~�A�h�C��
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2022/04/26 �����
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "�����R�s�[�֎~�A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.01"
Public Const APP_UPDATE = "2022/04/26"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/format_copy_blocker"

'--------------------------------------------------
'�A�h�C�����s��
Sub AddinStart()
    Call Start_FormatCopyBlocker
    MsgBox "�����̃R�s�[��΂�邳�Ȃ�", _
                vbExclamation + vbOKOnly, ThisWorkbook.Name
End Sub

'�A�h�C���ꎞ��~��
Sub AddinStop(): Call Stop_FormatCopyBlocker: End Sub

'�A�h�C���ݒ�\��
'Sub AddinConfig(): Call SettingForm.Show: End Sub
Sub AddinConfig(): MsgBox "�ݒ��ʕ\���p": End Sub

'�A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'�A�h�C�����~�߂������Ɏg���v���V�[�W��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub

Sub MergePrint(): MsgBox "Print": End Sub
'--------------------------------------------------

'�Ď��J�n
'Workbook_Open����Ă΂��
'���u�b�N�̏㏑���ۑ������m���邽�߂Ɏg�p�����
Sub Start_FormatCopyBlocker(): Call SubFormatCopyBlocker(False): End Sub

'�Ď���~
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

Rem �z��̎����������߂�
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

Rem �f�[�^��TSV�ɕϊ�����
Function JoinTsv(v) As String
    Select Case GetArrayDimension_NoAPI(v)
        Case 0: JoinTsv = v
        Case 1: JoinTsv = Join(v, vbTab)
        Case 2: JoinTsv = Join2(v, vbTab, vbLf)
        Case Else: JoinTsv = v.Value
    End Select
End Function

Rem �񎟌��z���CSV���̕�����ɕϊ�����
Public Function Join2(arr As Variant, _
                        Optional ByVal Delimiter1 As String = vbTab, _
                        Optional ByVal Delimiter2 As String = vbCrLf) As String
    Dim i As Long, j As Long
    Dim arr1() As Variant
    Dim arr2() As Variant
    
    If GetArrayDimension_NoAPI(arr) <> 2 Then Err.Raise 9999, "Join2", "Join2 : ���͕ϐ�Arr���񎟌��z��ł͂���܂���B"
    
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
