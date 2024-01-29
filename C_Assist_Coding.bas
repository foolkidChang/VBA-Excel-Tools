Option Explicit

Function GetDesktopPath() As String: GetDesktopPath = Environ("USERPROFILE") & "\Desktop\": End Function

Function SetDateVal(DateStr As String, Optional DateVar As Date, Optional ExtraDateStr As String) As Boolean
'ExtraDateStr: 當取得的資料只有年、月時，透過此參數補足月、日；例如來源只有2023，此參數可以補0101，最後變成20230101
    SetDateVal = True: DateVar = 0
    If DateStr = "" Then SetDateVal = False: Exit Function
    
    On Error Resume Next: DateVar = ProCDate(DateStr & ExtraDateStr): On Error GoTo 0
    If DateVar = 0 Then SetDateVal = False: Exit Function
End Function

Sub TotalHide(TrgSht As Worksheet, StartRow As Byte, StartCol As Byte)
    With TrgSht
        If StartRow <> 0 Then .Range(.Cells(StartRow, 1), .Cells(.Rows.Count, 1)).EntireRow.Hidden = True
        If StartCol <> 0 Then .Range(.Cells(1, StartCol), .Cells(1, .Columns.Count)).EntireColumn.Hidden = True
    End With
End Sub

Sub Cmmt_Add(ByVal TargetRange As Range, ByVal CommentText As String, Optional ByVal ReplaceMark As Boolean = True) '對TargetRange的每一個儲存格加上註解CommentText
'CommentText: 如果為空值→刪除註解
'ReplaceMark: 覆蓋原註解；True = 覆蓋

Dim TargetCell As Range, OldCommentText As String

    For Each TargetCell In TargetRange
        With TargetCell
            If .Comment Is Nothing Then
                .AddComment '沒有註解的需要這行；已有註解的不需要
            Else
                OldCommentText = .Comment.Text
                If InStr(OldCommentText, CommentText) > 0 Then
                    If Not ReplaceMark Then CommentText = OldCommentText & Chr(10) & CommentText '如果ReplaceMark為False則用新註解取代舊註解
                End If
            End If
            
            If CommentText = "" Then
                .Comment.Delete
            Else
                .Comment.Text Text:=CommentText
                .Comment.Visible = False '顯示為隱藏
                With .Comment.Shape '修改註解位置
                    .Placement = xlMoveAndSize '隨著儲存格移動和調整大小
                    .Left = TargetCell.Left + TargetCell.Width + 11.25
                    .Top = TargetCell.Top - 7.5
                    With .TextFrame    '修改註解窗格大小
                        .AutoSize = True
                        With .Characters.Font
                            .Name = "細明體"
                            .Size = 9
                        End With
                    End With
                End With
            End If
        End With
    Next

End Sub

Sub LiteReplace(TrgRng As Range, What$, Replacement$, Optional LookAt As XlLookAt = xlPart)
    TrgRng.Replace What:=What, Replacement:=Replacement, LookAt:=LookAt, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub ResetInteriorColor(TrgRng As Range)

    With TrgRng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
        '==Test== 1
        '.PatternColorIndex = xlAutomatic    '變成白底色
        '==Test== 2
        '.ThemeColor = xlThemeColorAccent1   '1~6；對應儲存格上的顏色
        '==Test== 3
        '.Color = vbWhite                    '變成白底色
    End With

End Sub

Sub ShtShowAllData(TrgSht As Worksheet, Optional ShowAllColumns As Boolean)
    If TrgSht.ProtectContents Then Call CrackWorkSheet(TrgSht)      '如果有加密→取消加密
    If ShowAllColumns Then TrgSht.Cells.EntireColumn.Hidden = False '如果要顯欄→取消欄位隱藏
    If TrgSht.FilterMode Then TrgSht.ShowAllData                    '如果有篩選→取消篩選，以便 AutoFit 或 Find
End Sub

Private Sub PreCrackWorkSheet(): Call CrackWorkSheet(ActiveSheet): End Sub

Sub CrackWorkSheet(TrgSht As Worksheet)

Dim I1 As Integer, I2 As Integer, I3 As Integer, I4 As Integer, I5 As Integer, I6 As Integer
Dim I7 As Integer, I8 As Integer, I9 As Integer, I10 As Integer, I11 As Integer, I12 As Integer

Dim TryPassword$
    
    On Error Resume Next
    For I1 = 65 To 66: For I2 = 65 To 66: For I3 = 65 To 66: For I4 = 65 To 66: For I5 = 65 To 66: For I6 = 65 To 66
    For I7 = 65 To 66: For I8 = 65 To 66: For I9 = 65 To 66: For I10 = 65 To 66: For I11 = 65 To 66: For I12 = 32 To 126
        TryPassword = Chr(I1) & Chr(I2) & Chr(I3) & Chr(I4) & Chr(I5) & Chr(I6) & Chr(I7) & Chr(I8) & Chr(I9) & Chr(I10) & Chr(I11) & Chr(I12)
        TrgSht.Unprotect Password:=TryPassword '測試是否可取消保護
        If ActiveSheet.ProtectContents = False Then GoTo OutNow
    Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next
    On Error GoTo 0
    
OutNow:

End Sub

Sub SetT1()
    If T1 = 0 Then T1 = Timer
End Sub
Function GetRunSec(): GetRunSec = Timer - T1: End Function
Function GetMMSS(in_Sec As Double) As String: GetMMSS = " " & IIf(in_Sec \ 60 > 0, in_Sec \ 60 & " 分 ", "") & in_Sec Mod 60 & " 秒": End Function
Function RunTime(Optional IsFinish As Boolean = True) As String '帶回執行時間，需搭配T1使用
'IsFinish: 影響最後呈現文字

Dim TmpStr As String

    If T1 = 0 Then Exit Function
    
    TmpStr = GetMMSS(Timer - T1)
    If IsFinish Then
        RunTime = "；歷時" & TmpStr & "。": T1 = 0
    Else
        RunTime = "已執行時間:" & TmpStr & "。"
    End If

End Function
Private Sub KeepRunCheckTest()
Dim BreakTime As Long, Check As Boolean
BreakTime = 2
    Do
         Check = KeepRunCheck(BreakTime)
    Loop Until Check = False
End Sub
Function KeepRunCheck(io_BreakTime As Long) As Boolean
Dim AddTime As Double
    Call SetT1 '避免使用前忘記SetT1
    
    If GetRunSec > io_BreakTime Then
        On Error Resume Next
        AddTime = InputBox("目前程式已執行" & GetMMSS(Timer - T1) & "，如需繼續執行，" & Chr(10) & _
                           "請於下方輸入「再執行時間(分鐘)」：" & Chr(10) & Chr(10) & _
                           "※如屆時程序仍未完成，系統將再次詢問。", "再執行時間確認", 5)
        On Error GoTo 0
        
        If AddTime <> 0 Then
            io_BreakTime = GetRunSec + AddTime * 60: KeepRunCheck = True
        End If
    Else
        KeepRunCheck = True
    End If

End Function

Sub CheckBar(ByVal BarNameArray) '檢查增益集指令Bar是否已存在之用

Dim CheckBar As CommandBar
Dim SnA As Byte

    For SnA = LBound(BarNameArray) To UBound(BarNameArray)
        On Error Resume Next
        Set CheckBar = CommandBars(BarNameArray(SnA))
        On Error GoTo 0
        If Not CheckBar Is Nothing Then
            CheckBar.Delete
            Set CheckBar = Nothing
        End If
    Next
    
End Sub

Function SpecialRngCheck(mySht As Worksheet, myRng As Range) As Boolean '先將工作表所有 樞紐分析表、資料表 範圍找出來，再與指定範圍進行交集判斷

Dim myPvt As PivotTable, myLOb As ListObject
Dim SpcRng As Range, IntersectRng As Range
Dim SnS As Integer, SpcCnt As Integer

    '先找樞紐分析表
    SpcCnt = mySht.PivotTables.Count
    
    If SpcCnt <> 0 Then
        For SnS = 1 To SpcCnt
            Set myPvt = mySht.PivotTables(SnS)
            On Error Resume Next                                                                                                'PageRange      = 篩選項目，好像跟下面一樣...
            If SpcRng Is Nothing Then Set SpcRng = myPvt.PageRangeCells Else Set SpcRng = Union(SpcRng, myPvt.PageRangeCells)   'PageRangeCells = 篩選項目，好像跟上面一樣...
            If SpcRng Is Nothing Then Set SpcRng = myPvt.TableRange1 Else Set SpcRng = Union(SpcRng, myPvt.TableRange1)         'TableRange1 = 不含篩選項目的樞紐分析表範圍
            On Error GoTo 0                                                                                                     'TableRange2 = 包含篩選項目的樞紐分析表範圍
        Next
    End If
    
    '再找資料表
    SpcCnt = mySht.ListObjects.Count
    If SpcCnt <> 0 Then
        For SnS = 1 To SpcCnt
            Set myLOb = mySht.ListObjects.Item(SnS)
            On Error Resume Next
            If SpcRng Is Nothing Then Set SpcRng = myLOb.Range Else Set SpcRng = Union(SpcRng, myLOb.Range)
            On Error GoTo 0
        Next
    End If
    
    '結果驗證
    On Error Resume Next
    Set IntersectRng = Intersect(myRng, SpcRng)
    On Error GoTo 0
    
    If Not IntersectRng Is Nothing Then SpecialRngCheck = True

End Function

Function CNTCheck(Str$) As Boolean  '檢查字串是否包含全形字；True = 有

Dim LineLen As Integer, LineLenB As Integer

    LineLen = Len(Str): LineLenB = ProLenB(Str)

    If LineLen <> LineLenB Then
        CNTCheck = True
    Else
        CNTCheck = False
    End If

End Function

Sub OpenLink(LinkStr$)

    On Error GoTo Err
    ActiveWorkbook.FollowHyperlink Address:=LinkStr
    Exit Sub
    On Error GoTo 0
    
Err: InputBox "因程式無法開啟連結路徑，需使用者人工操作。" & Chr(10) & Chr(10) & _
              "請複製下方內容後，至[檔案總管]或[網頁瀏覽器]試著開啟連結路徑，謝謝。", "連結開啟異常", LinkStr
    
End Sub

Sub SetUnionRng(out_UnionRng As Range, in_AddCell As Range, Optional out_TitleRng As Range, Optional in_TitleCol As Byte)
'out_UnionRng: 最終儲存的Rng
'out_TitleRng: 所有out_UnionRng列 + in_TitleCol 的集合

    If out_UnionRng Is Nothing Then
        Set out_UnionRng = in_AddCell
        If in_TitleCol <> 0 Then Set out_TitleRng = in_AddCell.Parent.Cells(in_AddCell.Row, in_TitleCol)
    Else
        Set out_UnionRng = Union(out_UnionRng, in_AddCell)
        If in_TitleCol <> 0 Then Set out_TitleRng = Union(out_TitleRng, in_AddCell.Parent.Cells(in_AddCell.Row, in_TitleCol))
    End If

End Sub

Sub ProUnion(MainRng As Range, SubRng As Range)
    If MainRng Is Nothing Then Set MainRng = SubRng Else Set MainRng = Union(MainRng, SubRng)
End Sub

Function GetFilePool(FilePool, Optional AllowMultiSelect As Boolean, Optional InitPath As String, Optional CstmExtenstion As String, Optional ShowMsgBox As Boolean = True) As Boolean
'AllowMultiSelect: 是否允許多選；預設否
'InitPath: 起始路徑
'CstmExtenstion: 是否進行副檔名篩選
'ShowMsgBox: 是否顯示對話框

    If AllowMultiSelect Then
        FilePool = Split(GetMultiFileNameStr("請選擇欲處理的檔案", ";", InitPath, CstmExtenstion), ";")
    Else
        FilePool = Split(GetSingleFileNameStr("請選擇欲處理的檔案", InitPath, CstmExtenstion), ";")
    End If
    
    GetFilePool = Not (UBound(FilePool) = 0 And FilePool(0) = "False")
    If GetFilePool = False And ShowMsgBox Then MsgBox "未正確選擇檔案，中止匯入程序。"

End Function

Private Function GetSingleFileNameStr(myTitle As String, _
                                      Optional InitPath As String, Optional CstmExtenstion As String) As String '統一回傳字串，由後端程序自行加工
'InitPath: 起始路徑

    With Application.FileDialog(msoFileDialogFilePicker)
        '參數設定
        .Title = myTitle: .AllowMultiSelect = False
        If InitPath <> "" Then .InitialFileName = InitPath
        .Filters.Clear
        If CstmExtenstion <> "" Then
            .Filters.Add "自訂篩選", "*." & CstmExtenstion & "*", 1
            .Filters.Add "所有檔案", "*.*", 2
        End If
        
        If .Show = -1 Then
            GetSingleFileNameStr = .SelectedItems(1) '因為僅允許單選，所以一定是(1)
        Else
            GetSingleFileNameStr = "False"
        End If
    End With

End Function

Private Function GetMultiFileNameStr(myTitle As String, SplitStr As String, _
                                     Optional InitPath As String, Optional CstmExtenstion As String) As String  '統一回傳字串，由後端程序自行加工
'InitPath: 起始路徑

Dim SnF As Integer

    With Application.FileDialog(msoFileDialogFilePicker)
        '參數設定
        .Title = myTitle & " (執行順序乃依據畫面看到之上下順序決定)": .AllowMultiSelect = True
        If InitPath <> "" Then .InitialFileName = InitPath
        .Filters.Clear
        If CstmExtenstion <> "" Then
            .Filters.Add "自訂篩選", "*." & CstmExtenstion & "*", 1
            .Filters.Add "所有檔案", "*.*", 2
        End If
        
        If .Show = -1 Then
            For SnF = 1 To .SelectedItems.Count
                GetMultiFileNameStr = GetMultiFileNameStr & .SelectedItems(SnF) & SplitStr
            Next
            GetMultiFileNameStr = Left(GetMultiFileNameStr, Len(GetMultiFileNameStr) - Len(SplitStr))
        Else
            GetMultiFileNameStr = "False"
        End If
        .Filters.Clear
    End With
    
End Function

Function ANSICheck(FilePool) As Boolean

Dim Answer As Byte
Dim EachFileName, Encoding As String, EncodingMsg As String

    For Each EachFileName In FilePool '逐表開啟檔案
        Encoding = FileEncodingCheck(CStr(EachFileName))
        If Encoding <> "ANSI" Then
            EncodingMsg = EncodingMsg & Chr(10) & EachFileName & " (" & Encoding & ")"
        End If
    Next
    
    ANSICheck = Not EncodingMsg <> ""
    If ANSICheck = False Then
        EncodingMsg = "偵測到非ANSI格式之文字檔如下，可能造成內容變成亂碼。" & Chr(10) & "是否繼續？" & Chr(10) & EncodingMsg
        Answer = ProQA(EncodingMsg, Array("繼續，我有難言之隱", "取消，我將另存格式"))
        
        If Answer = 1 Then ANSICheck = True
    End If

End Function

Function HasOpenedFile(FileNameAry) As Boolean
'檢查目前欲開啟的檔案是否有已開啟的同名檔案
Dim OpenedWorkbooksAry, FileName
    If Not IsArray(FileNameAry) Then
        HasOpenedFile = True
    Else
        OpenedWorkbooksAry = GetOpenedWorkbooksAry: Call ArySort(OpenedWorkbooksAry)
        For Each FileName In FileNameAry
            FileName = GetPathOrName(FileName, GetName) '取檔名
            If AryBinarySearch(OpenedWorkbooksAry, CStr(FileName)) >= 0 Then
                HasOpenedFile = True
                MsgBox "欲匯入的檔案名稱目前已開啟，請關閉該檔案後重新執行巨集" & Chr(10) & _
                       "※因為執行過程中需同時開啟" & Chr(10) & Chr(10) & _
                       "檔案名稱：" & FileName
            End If
        Next
    End If
End Function

Function GetOpenedWorkbooksAry() As Variant 'Coded By ChatGPT
'取得目前已開啟工作表的陣列
Dim WB As Workbook, arr() As String, i As Integer
    
    ReDim arr(Application.Workbooks.Count - 1)
    For Each WB In Application.Workbooks
        arr(i) = WB.Name
        i = i + 1
    Next WB
    GetOpenedWorkbooksAry = arr
End Function
