Option Explicit

Public Sub ExcelSettingsTurnOn(Optional IncludeCalculation As Boolean = True)
'IncludeCalculation: 預設為 True (將運算改為手動，但如果程式需要公式輔助，應為False；注意! 更改Application.Calculation時會連帶 CutCopyMode = False)

    '當下無活頁簿會報錯
    Call ActiveWorkBookCheck
    
    '依ExcelSettingsDict內容還原設定
    If ExcelSettingsDict.Exists(TurnOnCnt & "DA") Then
        Application.DisplayAlerts = ExcelSettingsDict(TurnOnCnt & "DA")
        Application.ScreenUpdating = ExcelSettingsDict(TurnOnCnt & "SU")
        Application.EnableEvents = ExcelSettingsDict(TurnOnCnt & "EE")
        If IncludeCalculation Then Application.Calculation = ExcelSettingsDict(TurnOnCnt & "CA")
        
        '變數處理
        TurnOnCnt = TurnOnCnt - 1
    Else
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        If IncludeCalculation Then Application.Calculation = xlCalculationAutomatic
    End If
    Application.DisplayStatusBar = True
End Sub

Public Sub ExcelSettingsTurnOff(Optional in_ScreenUpdating As Boolean = False, Optional in_EnableEvents As Boolean = False, Optional IncludeCalculation As Boolean = True)
'in_ScreenUpdating:  預設為 False (設為False的時候，User將無法於DoEvents時操作Excel)
'in_EnableEvents:    預設為 False (如果活頁簿有設計事件，應設為True)
'IncludeCalculation: 預設為 True  (將運算改為手動，但如果程式需要公式輔助，應為False；注意! 更改Application.Calculation時會連帶讓 CutCopyMode = False)

    '當下無活頁簿會報錯
    Call ActiveWorkBookCheck
    
    '變數處理+儲存ExcelSettingsDict
    TurnOnCnt = TurnOnCnt + 1
    If TurnOnCnt = 1 Then
        ExcelSettingsDict(TurnOnCnt & "DA") = True
        ExcelSettingsDict(TurnOnCnt & "SU") = True
        ExcelSettingsDict(TurnOnCnt & "CA") = xlCalculationAutomatic
        ExcelSettingsDict(TurnOnCnt & "EE") = True
    Else
        ExcelSettingsDict(TurnOnCnt & "DA") = Application.DisplayAlerts
        ExcelSettingsDict(TurnOnCnt & "SU") = Application.ScreenUpdating
        ExcelSettingsDict(TurnOnCnt & "CA") = Application.Calculation
        ExcelSettingsDict(TurnOnCnt & "EE") = Application.EnableEvents
    End If
    
    '設為無效 (判斷是否自訂參數)
    Application.DisplayAlerts = False
    
    Application.ScreenUpdating = in_ScreenUpdating
    Application.EnableEvents = in_EnableEvents
    
    If IncludeCalculation Then Application.Calculation = xlCalculationManual
End Sub
