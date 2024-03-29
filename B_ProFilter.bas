Option Explicit

Dim NoQA As Boolean 'for強化篩選所用

'====Pre主要功能========================================================================================================================

Sub 篩選強化_完全比對(): Call Real篩選強化_完全比對: End Sub
Sub 使用說明_完全比對(): Call OpenLink("https://foolkidchang.wordpress.com/2019/09/30/ProFilter-ExactlyCompare/"): End Sub
Sub 使用說明_模糊比對(): Call OpenLink("https://foolkidchang.wordpress.com/2019/10/03/ProFilter-FuzzyCompare/"): End Sub

'====主要功能===========================================================================================================================

Sub Real篩選強化_完全比對(Optional SemicolonCheck As Boolean = True)
'SemicolonCheck: 分號檢查；為了區別在進階篩選時要不要詢問分號
    Call ExcelSettingsTurnOff(, , False)
    Call 篩選強化("完全比對", SemicolonCheck)
    Call ExcelSettingsTurnOn(False) ': End '為了清空記憶體(測試)；但會清空資料比對的記憶功能，故取消
End Sub

Sub 篩選強化_模糊比對()
    Call ExcelSettingsTurnOff(, , False)
    Call 篩選強化("模糊比對", True)
    Call ExcelSettingsTurnOn(False) ': End '為了清空記憶體(測試)；但會清空資料比對的記憶功能，故取消
End Sub

'====呼叫功能===========================================================================================================================

Sub 篩選強化(FilterType As String, Optional SemicolonCheck As Boolean = False) '使用進階篩選來實現多條件模糊比對

Dim mySht As Worksheet, myAutoFilter As AutoFilter, myAutoFilterRng As Range, myRng As Range
Dim LfField As Byte, RtField As Integer, myField As Byte, myFieldStr$
Dim TempSht_1 As Worksheet, TempSht_2 As Worksheet, TempSht_3 As Worksheet, TempRng As Range, TempTxt$, Answer As Byte
Dim myStr As String, myAry, MyAryRng As Range, TempTxtAry, SpcStrCnt As Long, AryCnt As Long, StrLen As Byte, HasLongStr As Boolean
Dim FormulaStr As String
Dim SnR As Long, SnC As Integer, myRowCnt As Long
Dim ChangeSplitStr As Boolean
Const SplitStr$ = ";"     '找個很容易打的字串來當分隔符號，方便將一些常用字串拆分進行搜尋
Const SplitStr_2$ = "@#@" '找個很難相同的字串來當分隔符號，避免太容易就把要查找的東西拆掉了
Const DefaultStr$ = "不修改本欄並直接按下[確定]可中止巨集"

    '==防呆1== (共用時不能新增格式化條件、不能刪除工作表；保護工作表時不能用 SpecialCells ...)
    If ActiveWorkbook.MultiUserEditing Then MsgBox "QQ" & Chr(10) & Chr(10) & "活頁簿處於共用狀態下，無法使用本功能。": GoTo OutNow
    If ActiveSheet.ProtectContents Then MsgBox "QQ" & Chr(10) & Chr(10) & "設定[保護工作表]時無法使用本功能。": GoTo OutNow
    
    '==防呆2==
    Set mySht = ActiveSheet
    If mySht Is Nothing Then GoTo OutNow
    
    '==防呆3==
    Set myAutoFilter = mySht.AutoFilter
    If myAutoFilter Is Nothing Then MsgBox "當前工作表尚未設定〔篩選範圍〕，取消執行。": GoTo OutNow
    Set myAutoFilterRng = myAutoFilter.Range
    
    '==防呆4==
    Set myRng = Selection
    LfField = myAutoFilterRng.Cells(1, 1).Column        '自動篩選範圍的第一個儲存格的欄位 = 最左端
    RtField = myAutoFilter.Filters.Count + LfField - 1  '自動篩選範圍的總篩選數 + 最左端 - 1  = 最右端
    myField = myRng.Column
    If Not (LfField <= myField And myField <= RtField) Then MsgBox "目前儲存格並未在〔篩選範圍〕內，取消執行。": GoTo OutNow
    
    myField = myField - LfField + 1 '使用
    
    '==前置處理==
    mySht.Parent.Sheets.Add After:=Sheets(Sheets.Count): Set TempSht_1 = ActiveSheet
    mySht.Parent.Sheets.Add After:=Sheets(Sheets.Count): Set TempSht_2 = ActiveSheet
    
    '==字串收集==
    If Application.CutCopyMode <> xlCut Then '非剪下才取值（包含純文字的複製也可以取值）
        With TempSht_1
            On Error Resume Next
            .Cells(1, 1).PasteSpecial Paste:=xlPasteValues  '儲存格的貼上為值，如果沒有儲存格會出現錯誤，進而跳到錯誤執行段
            .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats '儲存格的貼上格式；可以讓內容字串符合使用者看到的樣子
            If Err.Number = 0 Then: GoTo NoErr              '如果沒錯就跳過錯誤執行段
            
HasErr:     .Cells.NumberFormatLocal = "@" '強制改成文字格式
            .PasteSpecial Format:="文字"   '純文字的貼上為值
            On Error GoTo 0
            
NoErr:      .UsedRange.EntireColumn.AutoFit '避免因為數值過長而變成科學符號(還是有漏洞，但是算了，因為無法排除)
            
            '移除重複，避免條件冗長
            On Error Resume Next
            For SnC = 1 To .UsedRange.Columns.Count
                .UsedRange.Columns(SnC).RemoveDuplicates Columns:=1, Header:=xlNo
            Next
            On Error GoTo 0
            
            '==數量確認== (避免操作錯誤導致當機)
            If .UsedRange.Cells.Count > 5000 Then
                Answer = MsgBox("目前欲比對的欄位較多（ " & Format(.UsedRange.Cells.Count, "#,##0") & " 項），" & Chr(10) & _
                                "可能會造成執行時間較久，是否繼續？", vbQuestion + vbYesNo, "篩選強化_" & FilterType)
                If Answer <> vbYes Then GoTo OutNow
            End If
            
            '==字串收集_例外處理_1== 包含了分隔字串
            For Each TempRng In .UsedRange
                TempTxt = TempRng.Text
                If InStr(TempTxt, SplitStr) > 0 Then
                    If SemicolonCheck = False Then
                        Answer = 2
                    Else
                        Answer = ProQA("因複製內容包含了分隔字串（" & SplitStr & "），" & Chr(10) & _
                                       "是否要將複製內容先分隔再篩選？", _
                                 Array("是，請將複製內容先分隔再篩選", _
                                       "否，查詢內容包含分隔字串，請直接篩選"), TitleStr:="篩選強化_" & FilterType)
                    End If
                    
                    Select Case Answer
                    Case 1:     ChangeSplitStr = False '如果篩選值真的包含了原分隔字串就改用新分隔字串
                    Case 2:     ChangeSplitStr = True  '如果篩選值真的包含了原分隔字串就改用新分隔字串
                    Case Else:  GoTo OutNow
                    End Select
                    
                    Exit For
                End If
            Next
            
            '==字串收集_1==
            For Each TempRng In .UsedRange
                TempTxt = TempRng.Text
                If TempTxt <> "" Then
                    If ChangeSplitStr = True Then
                        myStr = myStr & TempTxt & SplitStr_2
                    Else
                        myStr = myStr & TempTxt & SplitStr
                    End If
                End If
            Next
            
            .Cells.Delete '先清除內容，因為模糊比對要用另一個函數來實現，需要範圍變數
        End With
    End If
    
    '==字串收集_例外處理_2== 尚未複製就執行巨集
    If myStr = "" Then
        myStr = InputBox("請輸入欲搜尋的字串，多個字串時，請以 " & SplitStr & " 隔開" & Chr(10) & _
                         "例如：AA" & SplitStr & "BB" & SplitStr & "CC" & Chr(10) & Chr(10) & _
                         "※如果不改值並按下[確定]或[Enter]，會中止巨集" & Chr(10) & _
                         "※如果直接按下[取消]或[Esc]，則會尋找空儲存格", "篩選強化_" & FilterType & "_字串輸入", DefaultStr)
        
        If myStr = DefaultStr Then
            mySht.Activate: GoTo OutNow
        ElseIf myStr <> "" Then
            myStr = myStr & SplitStr
        End If
    End If
    
    '==前置資料整理== 並同時決定使用一般篩選或是進階篩選
    If myStr = "" Then
        myAry = Array("")
    Else
        If ChangeSplitStr = True Then
            myStr = Left(myStr, Len(myStr) - Len(SplitStr_2))
            myAry = ProSplit(myStr, SplitStr_2)
        Else
            myStr = Left(myStr, Len(myStr) - Len(SplitStr))
            myAry = Split(myStr, SplitStr)
        End If
        
        '統計篩選條件數量
        AryCnt = UBound(myAry) + 1
        
        '抓取特殊字元數量
        If FilterType = "完全比對" Then '完全比對者須另外統計包含特殊字元(?、*)的篩選條件項目數
            For SnR = LBound(myAry) To UBound(myAry)
                TempTxt = myAry(SnR)
                If InStr(TempTxt, "?") + InStr(TempTxt, "*") > 0 Then SpcStrCnt = SpcStrCnt + 1
            Next
        End If
        
        '統計特殊字元數量及判斷是否有超長字串
        StrLen = 255 - IIf(FilterType = "模糊比對", 2, 0) '模糊比對者，前後要補兩個星號
        For SnR = LBound(myAry) To UBound(myAry)
            TempTxt = myAry(SnR)
            
            If Len(TempTxt) > StrLen Then
                If HasLongStr = False Then
                    Answer = ProQA("篩選條件字串超過指定長度（" & StrLen & "），" & Chr(10) & _
                                   "是否改取最大長度？", _
                                    Array("是，請將篩選條件字串改取LEFT" & StrLen & "字元", _
                                          "否，請結束功能"), TitleStr:="篩選強化_" & FilterType)
                    If Answer <> 1 Then GoTo OutNow
                    HasLongStr = True
                End If
                TempTxt = Left(TempTxt, StrLen): myAry(SnR) = TempTxt
            End If
            
            '完全比對者須另外統計包含特殊字元(?、*)的篩選條件項目數
            If FilterType = "完全比對" And InStr(TempTxt, "?") + InStr(TempTxt, "*") > 0 Then SpcStrCnt = SpcStrCnt + 1
        Next
        
        '如果是NoQA進來的，或是篩選條件在2項以下，或是完全比對模式下都沒有特殊字元...
        If NoQA = True Or AryCnt <= 2 Or (FilterType = "完全比對" And SpcStrCnt = 0) Then   '使用一般篩選
            If FilterType = "模糊比對" Then
                For SnR = LBound(myAry) To UBound(myAry)
                    myAry(SnR) = "*" & myAry(SnR) & "*"
                Next
            End If
        Else                                                                                '使用進階篩選，字元處理方式不同
            '==防呆5== 模糊比對時，如果條件大於3個，則會需要依賴欄位名稱來進行進階篩選；空白或重複的欄名會導致錯誤
            myFieldStr = myAutoFilterRng.Cells(1, myField) '取得篩選用變數
            If myFieldStr = "" Then
                MsgBox "QQ" & Chr(10) & Chr(10) & "標題列（儲存格" & myAutoFilterRng.Cells(1, myField).Address(False, False) & "）為空值" & Chr(10) & "請輸入內容後再使用本功能"
                GoTo OutNow
            Else
                For SnC = 1 To myField - 1
                    If myAutoFilterRng.Cells(1, SnC) = myFieldStr Then
                        MsgBox "QQ" & Chr(10) & Chr(10) & "標題列（儲存格" & myAutoFilterRng.Cells(1, myField).Address(False, False) & "）名稱與儲存格" & myAutoFilterRng.Cells(1, SnC).Address(False, False) & "重複" & Chr(10) & "請修改內容後再使用本功能"
                        GoTo OutNow
                    End If
                Next
            End If
            
            '==防呆6==
            If AryCnt > TempSht_1.Cells.Columns.Count Then
                MsgBox "取消執行!!" & Chr(10) & _
                       "篩選條件數量（" & Format(AryCnt, "#,##0") & "）大於工作表欄位數量（ " & Format(TempSht_1.Cells.Columns.Count, "#,##0") & " 項），無法使用本功能。"
                
                If mySht.FilterMode Then mySht.ShowAllData
                GoTo OutNow
            End If
            
            '正式處理
            For SnR = LBound(myAry) To UBound(myAry)
                TempSht_1.Cells(1, SnR + 1) = myFieldStr
                Select Case FilterType
                    Case "完全比對": TempSht_1.Cells(SnR + 2, SnR + 1) = "'=" & myAry(SnR)
                    Case "模糊比對": TempSht_1.Cells(SnR + 2, SnR + 1) = "'=*" & myAry(SnR) & "*"
                End Select
            Next
            Set MyAryRng = TempSht_1.Cells(1, 1).CurrentRegion
        End If
    End If
    
    '==篩選調整==
    mySht.Activate
    With mySht
        If .FilterMode Then .ShowAllData
        
        If myStr = "" Then
            myAutoFilterRng.AutoFilter Field:=myField, Criteria1:=myAry, Operator:=xlFilterValues
        Else
            If NoQA = True Or AryCnt <= 2 Or (FilterType = "完全比對" And SpcStrCnt = 0) Then
                myAutoFilterRng.AutoFilter Field:=myField, Criteria1:=myAry, Operator:=xlFilterValues
                
                SnR = 1
                On Error Resume Next
                Do
                    myRowCnt = myAutoFilterRng.Resize(, SnR).SpecialCells(xlCellTypeVisible).Count '避免篩選範圍首欄隱藏的狀況...
                    If Err.Number = 0 Then Exit Do
                    SnR = SnR + 1: Err.Clear
                Loop
                On Error GoTo 0
                
                For Each TempRng In myAutoFilterRng.Rows(myAutoFilterRng.Rows.Count).Cells     '因為AutoFilter會把公式 SUBTOTAL當作是範圍的一部份，且永遠Visible，故需判斷後排除
                    On Error Resume Next: FormulaStr = TempRng.Formula: On Error GoTo 0
                    If FormulaStr Like "=SUBTOTAL(*" Then myRowCnt = myRowCnt - 1: Exit For
                Next
                
                If myRowCnt = 1 Then '如果篩選沒東西... (1是標題)
                    mySht.Activate: If .FilterMode Then .ShowAllData
                    GoTo NoFinded
                End If
            Else                      '如果是兩個以上的條件，則使用進階篩選
                '==進階篩選== (先用AdvancedFilter選出要的範圍)
                myAutoFilterRng.AdvancedFilter xlFilterCopy, MyAryRng, TempSht_2.Cells(1, 1)
                
                '==重新整理== (再用[篩選強化_完全比對]來重整)
                With TempSht_2
                    myAutoFilterRng.AutoFilter
                    On Error Resume Next
                    Set myRng = Nothing: Set myRng = .Cells(2, myField).Resize(.UsedRange.Rows.Count - 1, 1)
                    On Error GoTo 0
                    If myRng Is Nothing Then '如果篩選沒東西...
                        mySht.Activate
                        GoTo NoFinded
                    End If
                    
                    For Each MyAryRng In myRng
                        If Len(MyAryRng) > 255 Then
                            Answer = ProQA("因產出結果包含〔長度大於255〕的字串，" & Chr(10) & "是否於新工作表顯示結果？", _
                                     Array("是，請於新工作表顯示結果", _
                                           "否，請取消執行"))
                            If Answer = 1 Then '如果要使用這張新工作表...
                                SnR = 1
                                Do
                                    TempTxt = "模糊比對篩選結果" & Format(SnR, "00")                                          '改名
                                    On Error Resume Next: Set TempSht_3 = .Parent.Worksheets(TempTxt): On Error GoTo 0
                                    If TempSht_3 Is Nothing Then
                                        .Name = TempTxt: .Move After:=mySht: .Tab.Color = vbYellow: Exit Do
                                    Else
                                        Set TempSht_3 = Nothing
                                    End If
                                    SnR = SnR + 1
                                Loop
                                myAutoFilterRng.Rows(1).Copy: .Activate: .Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths '複製欄寬
                                Call ProSetFilter: .UsedRange.EntireRow.AutoFit: Set TempSht_2 = Nothing                      '加上篩選、凍結窗格、調整列高...
                            End If
                            GoTo OutNow
                        End If
                    Next
                    myRng.Copy: NoQA = True: Call Real篩選強化_完全比對(False): NoQA = False
                End With
            End If
        End If
    End With
    
    GoTo OutNow

NoFinded: MsgBox "無符合資料。" & IIf(FilterType = "模糊比對", Chr(10) & Chr(10) & _
                 "小提醒：" & Chr(10) & _
                 "如果欄位的資料為數字型態，是無法被[模糊篩選]找到的喔！", "")

OutNow: On Error Resume Next: TempSht_1.Delete: TempSht_2.Delete: mySht.Activate: On Error GoTo 0

    Set MyAryRng = Nothing: Set TempSht_1 = Nothing: Set TempSht_2 = Nothing: Set TempSht_3 = Nothing: Set TempRng = Nothing
    Set mySht = Nothing: Set myAutoFilter = Nothing: Set myAutoFilterRng = Nothing: Set myAutoFilterRng = Nothing: Set myRng = Nothing

End Sub
