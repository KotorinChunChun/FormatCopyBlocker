VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatCopyBlocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem FormatCopyBlocker
Option Explicit

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Rem オブジェクトの作成
Public Function Init(pApp As Excel.Application) As FormatCopyBlocker
    If Me Is FormatCopyBlocker Then
        With New FormatCopyBlocker
            Set Init = .Init(pApp)
        End With
        Exit Function
    End If
    Set Init = Me
    Set app = pApp
End Function

Rem シートの選択セルが変化した時
Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If kccFuncClipboard.IsEmptyCB Then Exit Sub
    If kccFuncClipboard.IsTextOnlyCB Then Exit Sub
    
    Dim rng As Range
    Set rng = kccFuncClipboard.GetRangeByClipboardCells()
    If Not rng Is Nothing Then
        Rem コピー中のセルデータをTSVテキストに置換
        Debug.Print rng.Address
        Dim arr: arr = GetValuesByRangeUnmergeCellsAndFillValues(rng)
        Dim tsvText  As String: tsvText = JoinTsv(arr)
        If Application.CutCopyMode = xlCut Then
            rng.ClearContents
        End If
        Call kccFuncClipboard.SetClipboardByTextAPI(tsvText)
    Else 'If kccFuncClipboard.GetTextByClipboardText Then
        Rem コピー中のデータからテキストのみ抽出して置換
        Debug.Print "txt"
        Dim cbText As String: cbText = kccFuncClipboard.GetTextByClipboardTextDataObject()
        Call kccFuncClipboard.SetClipboardByTextAPI(cbText)
    End If
End Sub

Rem 選択範囲のセル結合を解除すると仮定して値を埋めた状態の二次元配列を返す
Rem  @param CanRowMerge : 行結合を認める
Rem  @param CanColMerge : 列結合を認める
Function GetValuesByRangeUnmergeCellsAndFillValues(SelRange As Range, _
                        Optional CanRowMerge As Boolean = True, _
                        Optional CanColMerge As Boolean = True) As Variant
    
    Rem areas非対応。本当は対応しないといけない。
    If SelRange.Areas.Count > 1 Then Stop
    
    Dim rngUsed As Range
    Set rngUsed = Intersect(SelRange, SelRange.Worksheet.UsedRange)
    If rngUsed Is Nothing Then Exit Function
    
    Dim cel As Range
    Dim area As Range
    Dim rng As Range
    Dim arr As Variant
    
    For Each area In rngUsed.Areas
        If area Is area.Item(1) Then
            arr = area.MergeArea.Value
        Else
            arr = area.Value
        End If
        For Each cel In area
            If cel.MergeCells Then
                Set rng = cel.MergeArea
                If cel.Row = rng.Row And cel.Column = rng.Column Then
                    Call FillValue(arr, cel.Value, _
                                        cel.Row - area.Row + 1, _
                                        cel.Column - area.Column + 1, _
                                        rng.Rows.Count, _
                                        rng.Columns.Count)
                End If
            End If
        Next
    Next
    
    GetValuesByRangeUnmergeCellsAndFillValues = arr
End Function

Private Sub FillValue(ByRef arr2, v, rowIndex As Long, colIndex As Long, rowHeight As Long, colWidth As Long)
    Dim rr, cc
    For rr = rowIndex To rowIndex + rowHeight - 1
        For cc = colIndex To colIndex + colWidth - 1
            arr2(rr, cc) = v
        Next
    Next
End Sub
