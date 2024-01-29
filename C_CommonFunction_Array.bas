Option Explicit

Sub FillAryBySingleRow(io_TrgAry, in_AddRow As Range, io_TrgRow As Long, Optional in_AutoAddRowNum As Boolean, Optional in_AllStr As Boolean)
Dim RowAry
    RowAry = Get1DAryFromRng(in_AddRow, in_AllStr)
    Call FillAryBy1DAry(io_TrgAry, RowAry, io_TrgRow, in_AutoAddRowNum)
End Sub

Sub FillAryBy1DAry(io_TrgAry, in_AddAry, io_TrgRow As Long, Optional in_AutoAddRowNum As Boolean)
'in_AutoAddRowNum: 回傳io_TrgRow時是否自動 + 1

Dim SnA As Long, SnT As Long
Dim TrgStartRow As Long, TrgEndRow As Long
Dim TrgStartCol As Long, TrgEndCol As Long, TrgTotalCol As Long
Dim AddStartCol As Long, AddEndCol As Long, AddTotalCol As Long

    'io_TrgRow Check
    TrgStartRow = LBound(io_TrgAry, 1): TrgEndRow = UBound(io_TrgAry, 1)
    If Not (TrgStartRow <= io_TrgRow And io_TrgRow <= TrgEndRow) Then MsgBox "寫入Ary列數不正確": Stop
    
    'TrgCol Check
    TrgStartCol = LBound(io_TrgAry, 2): TrgEndCol = UBound(io_TrgAry, 2): TrgTotalCol = TrgEndCol - TrgStartCol + 1
    AddStartCol = LBound(in_AddAry):    AddEndCol = UBound(in_AddAry):    AddTotalCol = AddEndCol - AddStartCol + 1
    If TrgTotalCol < AddTotalCol Then MsgBox "寫入Ary欄位數量不一致": Stop
    
    '寫入
    For SnA = AddStartCol To AddEndCol
        io_TrgAry(io_TrgRow, TrgStartCol + SnT) = in_AddAry(SnA): SnT = SnT + 1
    Next
    
    '是否自動序號
    If in_AutoAddRowNum Then io_TrgRow = io_TrgRow + 1

End Sub

Sub OutShtFromAry(in_TrgBok As Workbook, in_DataAey, in_ShtName As String, Optional io_TrgSht As Worksheet, Optional in_ShtOrder As Integer = 999)
    Call ActiveWorkBookCheck(in_TrgBok)
    Call ProShtAdd(in_TrgBok, io_TrgSht, in_ShtName, in_ShtOrder, True)
    Call FillRngFromAry(io_TrgSht.Cells(1, 1), in_DataAey)
    Call SetStdFormat(io_TrgSht, 1)
End Sub

Sub FillRngFromAry(StartCell As Range, SrcAry, Optional NeedWrapText As Boolean)  '從陣列填值到工作表；不用先選活頁簿、工作表

Dim TotalR As Long, TotalC As Integer
Dim StartR As Long, StartC As Integer
Dim EndR As Long, EndC As Integer
Dim SnR As Long, SnC As Integer
Dim TmpAry, TmpStr As String

    StartR = LBound(SrcAry, 1): EndR = UBound(SrcAry, 1)
    TotalR = EndR - StartR + 1
    
    On Error Resume Next
    StartC = LBound(SrcAry, 2): EndC = UBound(SrcAry, 2)
    TotalC = UBound(SrcAry, 2) - LBound(SrcAry, 2) + 1
    On Error GoTo 0
    
    '如果是一維陣列，則進行相關調整(一維陣列的欄列是相反的)
    If TotalC = 0 Then
        ReDim TmpAry(1 To 1, StartR To EndR)
        Call FillAryBy1DAry(TmpAry, SrcAry, 1)
        
        SrcAry = TmpAry
        StartC = StartR: EndC = EndR: TotalC = TotalR
        StartR = 1:      EndR = 1:    TotalR = 1
    End If
    
    '當Ary內容有'='開頭的內容且非公式時，如果用此法填入，會造成錯誤，需進一步逐欄貼入
    On Error Resume Next
    StartCell.Resize(TotalR, TotalC).value = SrcAry
    
    '上一階段發生不明原因錯誤時，需直接填值...
    If Err.Number <> 0 Then
        Err.Number = 0
        For SnR = 0 To TotalR - 1
        For SnC = 0 To TotalC - 1
            StartCell.Offset(SnR, SnC) = SrcAry(StartR + SnR, StartC + SnC)
            If Err.Number <> 0 Then
                Err.Number = 0
                TmpStr = SrcAry(StartR + SnR, StartC + SnC)
                If Left(TmpStr, 1) = "=" Then TmpStr = "'" & TmpStr
                StartCell.Offset(SnR, SnC) = TmpStr
            End If
        Next: Next
    End If
    On Error GoTo 0
    
    If Not NeedWrapText Then StartCell.Resize(TotalR, TotalC).WrapText = False  '避免自動換列太過嚴重 20220406 CJP

End Sub

Function BinaryAryCombine(AscendingAry1 As Variant, Ary2 As Variant, Optional Unique As Boolean) As Variant '將兩個陣列合併，可藉由參數 [Unique] 確定是否排除重複
'過程中要求Ary1需是升冪(Ascending)陣列

Dim TempAry
Dim SnA As Long, SnA2 As Long, Count As Long

    If Unique Then
        For SnA = LBound(Ary2) To UBound(Ary2)
            If AryBinarySearch(AscendingAry1, CStr(Ary2(SnA))) < 0 Then Count = Count + 1
        Next
    Else
        Count = UBound(Ary2) + 1
    End If
    
    ReDim TempAry(UBound(AscendingAry1) + Count)
    
    For SnA = LBound(AscendingAry1) To UBound(AscendingAry1)
        TempAry(SnA2) = AscendingAry1(SnA)
        SnA2 = SnA2 + 1
    Next
    
    If Unique Then
        For SnA = LBound(Ary2) To UBound(Ary2)
            If AryBinarySearch(AscendingAry1, CStr(Ary2(SnA))) < 0 Then
                TempAry(SnA2) = Ary2(SnA)
                SnA2 = SnA2 + 1
            End If
        Next
    Else
        For SnA = LBound(Ary2) To UBound(Ary2)
            TempAry(SnA2) = Ary2(SnA)
            SnA2 = SnA2 + 1
        Next
    End If
    
    BinaryAryCombine = TempAry
    
End Function

Function AryBinarySearch(SrcAry As Variant, SearchStr As String, Optional ReturnFirst As Boolean = True) As Long '使用二進位搜尋法在來源陣列找尋關鍵字在陣列裡的索引值；參數 [ReturnFirst] 可決定帶入重複項第一個或最後一個

Dim MinA As Long, MaxA As Long, NowA As Long

    MinA = LBound(SrcAry): MaxA = UBound(SrcAry)
    
    Do While MinA <= MaxA
        NowA = (MinA + MaxA) / 2 '已設定變數，所以會自動四捨五入到整數
        
        'If SrcAry(NowA) Like SearchStr Then 'Like運算子會讓中括號包夾的SearchString失效
        If SrcAry(NowA) = SearchStr Then
            If ReturnFirst Then '向上找重複字元
                Do
                    If NowA > LBound(SrcAry) Then
                        'If SrcAry(NowA - 1) Like SearchStr Then 'Like運算子會讓中括號包夾的SearchString失效
                        If SrcAry(NowA - 1) = SearchStr Then
                            NowA = NowA - 1
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            Else                '向下找重複字元
                Do
                    If NowA < UBound(SrcAry) Then
                        'If SrcAry(NowA + 1) Like SearchStr Then 'Like運算子會讓中括號包夾的SearchString失效
                        If SrcAry(NowA + 1) = SearchStr Then
                            NowA = NowA + 1
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
            AryBinarySearch = NowA: Exit Function '找到時傳回索引值
        ElseIf SrcAry(NowA) < SearchStr Then
            MinA = NowA + 1
        ElseIf SrcAry(NowA) > SearchStr Then
            MaxA = NowA - 1
        End If
    Loop
    
    AryBinarySearch = -1 '找不到時傳回-1

End Function

Function ArySearch(SrcAry As Variant, SearchStr As String, Optional Similar As Boolean = True) As Integer '為了使用Like指令，相對於AryBinarySearch(完全相符)
'Similar: 使用Like取代等於

Dim NowA As Integer

    For NowA = LBound(SrcAry) To UBound(SrcAry)
        If SearchStr Like SrcAry(NowA) And Similar Or _
           SearchStr = SrcAry(NowA) Then
            ArySearch = NowA: Exit Function '找到時傳回索引值
        End If
    Next
    
    ArySearch = -1 '找不到時傳回-1

End Function

Function ReverseAry(arr As Variant) As Variant

Dim val As Variant

    With CreateObject("System.Collections.ArrayList") '<-- create a "temporary" array list with late binding
        For Each val In arr '<--| fill arraylist
            .Add val
        Next val
        .Reverse '<--| reverse it
        ReverseAry = .Toarray '<--| write it into an array
    End With
    
End Function

Sub ArySort(StAry, Optional OriLBndStr, Optional OriUBndStr, Optional CoAry, Optional NeedReverseAry As Boolean)    'Sorts a one-dimensional VBA array from smallest to largest
                                                                                                                    'using a very fast quicksort algorithm variant.
Dim StdVal, TmpVal
Dim OriLBnd As Long, OriUBnd As Long
Dim NowLBnd As Long, NowUBnd As Long

    If IsError(OriLBndStr) Then OriLBnd = LBound(StAry) Else OriLBnd = CLng(OriLBndStr) 'Optional變數遺漏判斷
    If IsError(OriUBndStr) Then OriUBnd = UBound(StAry) Else OriUBnd = CLng(OriUBndStr) 'Optional變數遺漏判斷
    
    NowLBnd = OriLBnd: NowUBnd = OriUBnd  '從最左右端開始比較
    StdVal = StAry((OriLBnd + OriUBnd) \ 2) '基準值，抓中間項（取商數，所以會靠左項）
    
    While (NowLBnd <= NowUBnd) '一直比較，直到交會
        Do While StAry(NowLBnd) < StdVal And NowLBnd < OriUBnd '往→找 >= 基準值的項目；直到最右端
            NowLBnd = NowLBnd + 1
        Loop
        
        Do While StdVal < StAry(NowUBnd) And OriLBnd < NowUBnd '往←找 =< 基準值的項目；直到最左端
            NowUBnd = NowUBnd - 1
        Loop
        
        If (NowLBnd <= NowUBnd) Then
            TmpVal = StAry(NowLBnd): StAry(NowLBnd) = StAry(NowUBnd): StAry(NowUBnd) = TmpVal       '如果左大右小有找到就交換項目
            If Not IsError(CoAry) Then
                TmpVal = CoAry(NowLBnd): CoAry(NowLBnd) = CoAry(NowUBnd): CoAry(NowUBnd) = TmpVal   '如果左大右小有找到就交換項目
            End If
            NowLBnd = NowLBnd + 1: NowUBnd = NowUBnd - 1                                            '繼續找下一項
        End If
    Wend
    
    If (OriLBnd < NowUBnd) Then Call ArySort(StAry, OriLBnd, NowUBnd, CoAry) '交會後，如果與上限間還有項目→遞迴
    If (NowLBnd < OriUBnd) Then Call ArySort(StAry, NowLBnd, OriUBnd, CoAry) '交會後，如果與下限間還有項目→遞迴

    If NeedReverseAry = True Then
        StAry = ReverseAry(StAry)
        If Not IsError(CoAry) Then CoAry = ReverseAry(CoAry)
    End If
    
End Sub

Function GetAry(ByVal TrgSht As Worksheet, ByVal ParaTitleName As String, Optional WithoutEmpty As Boolean) '將某個欄名開頭之內容回傳為一維陣列

Dim TitleRow As Long, TitleCol As Byte, TotalRow As Long
Dim DataCnt As Long, SnA As Long, SnA_2 As Long, EmptyCnt As Long
Dim TmpAry
    
    '====基本資料撈取====
    With TrgSht
        .Cells.EntireColumn.Hidden = False: If .FilterMode Then .ShowAllData '取消欄位隱藏 + 篩選中就取消篩選，以便 AutoFit 或 Find
        
        TitleRow = GetRow(TrgSht, ParaTitleName)
        TitleCol = GetCol(TrgSht, ParaTitleName, TitleRow)
        
        If TitleRow * TitleCol = 0 Then Exit Function
        
        TotalRow = GetTotalRow(TrgSht, TitleCol)
        
        DataCnt = TotalRow - TitleRow
    End With
    
    '====初步取值====
    If DataCnt > 0 Then
        ReDim TmpAry(DataCnt - 1)
        For SnA = 0 To DataCnt - 1
            TmpAry(SnA) = TrgSht.Cells(TitleRow + 1 + SnA, TitleCol)
            If WithoutEmpty And TmpAry(SnA) = "" Then   '如果要排除空值...
                EmptyCnt = EmptyCnt + 1
            End If
        Next
    Else
        ReDim TmpAry(0)
    End If
    
    '====排除空值====
    If WithoutEmpty And EmptyCnt > 0 Then               '如果要排除空值，且確實有空值...
        ReDim TmpAry(DataCnt - 1 - EmptyCnt)
        For SnA = 0 To DataCnt - 1
            If TrgSht.Cells(TitleRow + 1 + SnA, TitleCol) <> "" Then
                TmpAry(SnA_2) = TrgSht.Cells(TitleRow + 1 + SnA, TitleCol)
                SnA_2 = SnA_2 + 1
            End If
        Next
    End If

    GetAry = TmpAry

End Function

Function GetAryFromTXT(FileName As String, SplitStr As String, Optional TitleRowAry) '將某個欄名開頭之內容回傳為一維陣列

Dim HasTitleRow As Boolean
Dim myFSO As New FileSystemObject, myText As TextStream
Dim EachLine As String
Dim TmpFullAry, TotalRow As Long, TotalCol As Byte
Dim TmpAry, CurrRow As Long, CurrCol As Byte

    '====GetTitleRowAry====
    If Not IsError(TitleRowAry) Then
        HasTitleRow = True
        TotalCol = UBound(TitleRowAry)
    End If

    '====先抓取資料總筆數====
    Set myText = myFSO.OpenTextFile(FileName, ForReading, True)
    Do Until myText.AtEndOfStream
        EachLine = ProTrim(myText.ReadLine, SplitStr)
        If EachLine <> "" Then
            TotalRow = TotalRow + 1
            If TotalCol = 0 Then TotalCol = UBound(Split(EachLine, SplitStr))
        End If
    Loop
    
    '====GetTmpAry====
    ReDim TmpFullAry(TotalRow - IIf(HasTitleRow, 0, 1), TotalCol)
    If HasTitleRow Then
        Call FillAryBy1DAry(TmpFullAry, TitleRowAry, CurrRow, True)
    End If
    
    '====正式抓取內容====
    Set myText = myFSO.OpenTextFile(FileName, ForReading, True)
    Do Until myText.AtEndOfStream
        EachLine = ProTrim(myText.ReadLine, SplitStr)
        If EachLine <> "" Then
            TmpAry = Split(EachLine, SplitStr)
            
            Call FillAryBy1DAry(TmpFullAry, TmpAry, CurrRow, True)
        End If
    Loop
    
    GetAryFromTXT = TmpFullAry

End Function

Function GetAryFromSht(in_TrgSht As Worksheet, in_ColNameCriteria As String, Optional in_NeedTitleRow As Boolean = True)

Dim TmpDT As New clsDataTable, TmpRng As Range, TmpAry

    Set TmpDT = New clsDataTable: Call TmpDT.SetRng(in_TrgSht, in_ColNameCriteria, True)
    
    Set TmpRng = TmpDT.Rng: If in_NeedTitleRow = False Then Set TmpRng = TmpRng.Offset(1).Resize(TmpRng.Rows.Count - 1)
    TmpAry = TmpRng
    
    GetAryFromSht = TmpAry

End Function

Function Get1DAryFromRng(in_TrgRng As Range, AllStr As Boolean) '引數使用Range會讓效能變差，盡量減少使用 20220906 CJP

Dim TmpAry, TmpVal, MaxRow As Integer, MaxCol As Integer
Dim SnA0 As Integer, SnA1 As Integer, SnA2 As Integer

    MaxRow = in_TrgRng.Rows.Count: MaxCol = in_TrgRng.Columns.Count
    ReDim TmpAry(1 To MaxRow * MaxCol)
    
    SnA0 = LBound(TmpAry)
    
    For SnA1 = 1 To MaxRow: For SnA2 = 1 To MaxCol
        TmpVal = in_TrgRng.Cells(SnA1, SnA2)
        If AllStr Then
            If Left(TmpVal, 1) = "0" And TmpVal <> "0" Then
                TmpVal = "'" & TmpVal
            End If
        End If
        
        TmpAry(SnA0) = TmpVal: SnA0 = SnA0 + 1
    Next: Next
    
    Get1DAryFromRng = TmpAry

End Function

Function ProSplit(OriStr, SplitStr As String)
    Do
        If Right(OriStr, Len(SplitStr)) = SplitStr Then OriStr = Left(OriStr, Len(OriStr) - Len(SplitStr)) Else Exit Do
    Loop
    
    ProSplit = Split(OriStr, SplitStr)
End Function
