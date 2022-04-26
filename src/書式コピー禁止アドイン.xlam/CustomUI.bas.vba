Attribute VB_Name = "CustomUI"
Rem
Rem CustomUI
Rem
Rem 本モジュールは自作のCustomUIエディタから自動生成したイベントハンドラです。
Rem

Sub onAction_Start_FormatCopyBlocker(Control As IRibbonControl): Call Start_FormatCopyBlocker: FinalUseCommand = "Start_FormatCopyBlocker": End Sub
Sub onAction_Stop_FormatCopyBlocker(Control As IRibbonControl): Call Stop_FormatCopyBlocker: FinalUseCommand = "Stop_FormatCopyBlocker": End Sub

Sub onAction_AddinConfig(Control As IRibbonControl): Call AddinConfig: FinalUseCommand = "AddinConfig": End Sub
Sub onAction_AddinInfo(Control As IRibbonControl): Call AddinInfo: FinalUseCommand = "AddinInfo": End Sub
Sub onAction_AddinEnd(Control As IRibbonControl): Call AddinEnd: FinalUseCommand = "AddinEnd": End Sub
