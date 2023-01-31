Attribute VB_Name = "Numeric_Balloon_Check"
'vba Numeric_Balloon_Check Ver0.0.2 by Kantoku
'数値バルーンの欠番・重複をチェックする


'バルーンレポートビュー名
Private Const BALLOON_VIEW_NAME = "BALLOON"

'バルーンレポートテキスト名
Private Const REPORT_TEXT_NAME = "balloon_report"

'バルーンレポートX,Yの位置
Private Const POSITION_X = -200
Private Const POSITION_Y = 0

'作業列挙型
Private Enum OPERATION
    ALL_SHEETS_ = 0
    ACTIVE_SHEET_
    CANCEL_
End Enum

Option Explicit

Sub CATMain()

    'ドキュメントのチェック
'    If Not KCL.CanExecute("DrawingDocument") Then Exit Sub

    '処理確認
    Dim ope As OPERATION
    ope = query()
    If ope = OPERATION.CANCEL_ Then
        Exit Sub
    End If

    'バルーン文字の数値のものをソートして取得
    Dim balloonNumbers As Variant
    balloonNumbers = quick_sort( _
        get_balloon_text_numeric_array( _
            get_balloon_list(ope) _
        ) _
    )

    Dim msg As String
    If IsEmpty(balloonNumbers) Then
        msg = "数字のバルーンが見つかりませんでした"
        Exit Sub
    End If

    '重複削除・重複していたものを配列で取得
    Dim resultArray As Variant
    resultArray = get_remove_overlap_array( _
        balloonNumbers _
    )

    '使用バルーン連番でグループ分け
    Dim unique_groups As Object
    Set unique_groups = group_by_consecutive_numbers( _
        resultArray(0) _
    )

    '重複バルーン連番でグループ分け
    Dim overlap_groups As Object
    Set overlap_groups = group_by_consecutive_numbers( _
        resultArray(1) _
    )

    '結果の整頓
    msg = _
        "・使用されているバルーン番号" & vbCrLf & _
        get_result_txt(get_values(unique_groups)) & vbCrLf & _
        "・重複しているバルーン番号" & vbCrLf & _
        get_result_txt(get_values(overlap_groups))
    
    '結果出力
    dump_report msg

End Sub


'作業選択
Private Function query() _
    As OPERATION

    Dim msg As String
    msg = _
        "数値バルーンの欠番・重複をチェックします。" & vbCrLf & _
        "対象範囲を選択して下さい。" & vbCrLf & _
        " は　い：全てのシート" & vbCrLf & _
        " いいえ：アクティブシート" & vbCrLf & _
        "キャンセル：中止"
        
    Dim res As OPERATION
    Select Case MsgBox(msg, vbYesNoCancel + vbQuestion)
        Case vbYes
            res = OPERATION.ALL_SHEETS_
        Case vbNo
            res = OPERATION.ACTIVE_SHEET_
        Case Else
            res = OPERATION.CANCEL_
    End Select
    
    query = res

End Function


'テキストで出力
Private Sub dump_report( _
    ByVal msg As String)
    
    Dim view As DrawingView
    Set view = get_view_by_name(BALLOON_VIEW_NAME)

    Dim text As DrawingText
    Set text = get_text_by_name(view, REPORT_TEXT_NAME)

    text.text = _
        Date & " - " & _
        Time & vbCrLf & vbCrLf & _
        msg
    
End Sub


'テキストを名前で取得
Private Function get_text_by_name( _
    ByVal view As DrawingView, _
    ByVal name As String) _
    As DrawingText
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim texts As DrawingTexts
    Set texts = view.texts

    If texts.count < 1 Then
        Set get_text_by_name = create_text(texts, name)
        Exit Function
    End If

    Dim text As DrawingText
    For Each text In texts
        If text.name = name Then
            Set get_text_by_name = text
            Exit Function
        End If
    Next

    Set get_text_by_name = create_text(texts, name)
    
End Function


'テキストオブジェクト作成
Private Function create_text( _
    ByVal texts As DrawingTexts, _
    ByVal name As String) _
    As DrawingText
    
    Dim text As DrawingText
    Set text = texts.Add( _
        "dammy", _
        POSITION_X, _
        POSITION_Y _
    )
    text.name = name
    text.SetFontSize 0, 0, 7
    
    Set create_text = text

End Function


'ビューを名前で取得
Private Function get_view_by_name( _
    ByVal name As String) _
    As DrawingView
    
    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim views As DrawingViews
    Set views = dDoc.sheets.ActiveSheet.views

    Dim view As DrawingView
    For Each view In views
        If view.name = name Then
            Set get_view_by_name = view
            Exit Function
        End If
    Next

    Set get_view_by_name = views.Add(name)
    
End Function


'連番毎にグループ化
Private Function group_by_consecutive_numbers( _
    ByVal ary As Variant) _
    As Object
    
    Dim size As Long
    size = UBound(ary) + 1
    ReDim Preserve ary(size)
    ary(size) = -1
    
    Dim dict_groups As Object
    Set dict_groups = CreateObject("Scripting.Dictionary")

    Dim count_groups As Long
    count_groups = 0

    Dim startIdx As Long
    startIdx = 0

    Dim finishNumber As Long
    finishNumber = UBound(ary) - 1

    Dim i As Long
    For i = 0 To finishNumber
        If ary(i) + 1 <> ary(i + 1) Then
            Call dict_groups.Add( _
                count_groups, _
                get_range_ary(ary, startIdx, i) _
            )
            
            startIdx = i + 1
            count_groups = count_groups + 1
        End If
    Next
    
    Set group_by_consecutive_numbers = dict_groups
    
End Function


'スライス
Private Function get_range_ary( _
    ByVal ary As Variant, _
    ByVal startIdx As Long, _
    ByVal endIdx As Long) _
    As Variant

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = startIdx To endIdx
        dict.Add ary(i), 0
    Next
    
    get_range_ary = dict.keys()

End Function


'辞書のvaluesの代替
Private Function get_values( _
    ByVal dict As Object) _
    As Variant

    If dict.count < 1 Then
        get_values = Array()
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(UBound(dict.keys()))

    Dim key As Variant
    Dim count As Long
    count = 0
    For Each key In dict.keys()
        ary(count) = dict(key)
        count = count + 1
    Next
    
    get_values = ary

End Function


'グループ化された配列の文字化
Private Function get_result_txt( _
    ByVal ary_groups As Variant) _
    As String

    Dim msg As String

    Dim i As Long
    Dim ary As Variant, count As Long
    For i = 0 To UBound(ary_groups)
        ary = ary_groups(i)
        count = UBound(ary)
        Select Case count
            Case 0
                msg = msg & ary(0) & vbCrLf
            Case Is > 0
                msg = msg & _
                    ary(0) & " - " & _
                    ary(count) & vbCrLf
        End Select
    Next
    
    get_result_txt = msg

End Function


' 配列の重複削除
' return array(array,array) - 0:重複無し配列, 1:重複した配列
Private Function get_remove_overlap_array( _
    ByVal ary As Variant) _
    As Variant

    If IsEmpty(ary) Then
        get_remove_overlap_array = Array( _
            Array(), _
            Array() _
        )
        Exit Function
    End If

    Dim dict_unique As Object
    Set dict_unique = CreateObject("Scripting.Dictionary")

    Dim dict_overlap As Object
    Set dict_overlap = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim value As Variant
    For i = 0 To UBound(ary)
        value = ary(i)

        If Not dict_unique.exists(value) Then
            dict_unique.Add value, 0
            GoTo continue
        End If
        
        If Not dict_overlap.exists(value) Then
            dict_overlap.Add value, 0
            GoTo continue
        End If
        
continue:
    Next
    
    get_remove_overlap_array = Array( _
        dict_unique.keys(), _
        dict_overlap.keys() _
    )

End Function


'数値バルーンの文字取得
Private Function get_balloon_text_numeric_array( _
    ByVal balloonList As Collection) _
    As Variant
    
    Dim numbers As Collection
    Set numbers = New Collection
    
    Dim balloon As DrawingText
    Dim txt As String
    For Each balloon In balloonList
        txt = balloon.text
        If Not IsNumeric(txt) Then GoTo continue
        
        numbers.Add (Val(txt))

continue:
    Next

    get_balloon_text_numeric_array = collection_to_array(numbers)

End Function


'コレクション->配列
Private Function collection_to_array( _
    lst As Collection) _
    As Variant

    If lst.count < 1 Then
        collection_to_array = Empty
        Exit Function
    End If

    Dim ary() As Variant
    ReDim ary(lst.count - 1)

    Dim i As Long
    For i = 1 To lst.count
        ary(i - 1) = lst(i)
    Next

    collection_to_array = ary

End Function


'バルーンの取得
Private Function get_balloon_list( _
    ope As OPERATION) _
    As Collection

    Dim searchWord As String
    If ope = OPERATION.ACTIVE_SHEET_ Then
        searchWord = "CATDrwSearch.DrwBalloon,sel"
    Else
        searchWord = "CATDrwSearch.DrwBalloon,all"
    End If

    Dim dDoc As DrawingDocument
    Set dDoc = CATIA.ActiveDocument

    Dim sel As Selection
    Set sel = dDoc.Selection
    
    CATIA.HSOSynchronized = False
    sel.Clear
    
    sel.Add dDoc.sheets.ActiveSheet
    sel.Search searchWord

    Dim balloons As Collection
    Set balloons = New Collection

    Dim i As Long
    For i = 1 To sel.Count2
        balloons.Add sel.Item(i).value
    Next
    
    sel.Clear

    CATIA.HSOSynchronized = True

    Set get_balloon_list = balloons

End Function


'非破壊非再帰クイック＆挿入ソート
'参考 https://foolexp.wordpress.com/2011/10/29/%e3%82%af%e3%82%a4%e3%83%83%e3%82%af%e3%82%bd%e3%83%bc%e3%83%88%e3%81%a8%e6%8c%bf%e5%85%a5%e3%82%bd%e3%83%bc%e3%83%88%e3%81%ae%e3%83%8f%e3%82%a4%e3%83%96%e3%83%aa%e3%83%83%e3%83%89/
Private Function quick_sort( _
    ByVal ary As Variant) As Variant

    If IsEmpty(ary) Then
        quick_sort = Empty
        Exit Function
    End If

    Dim stack As Object
    Set stack = CreateObject("Scripting.Dictionary")
   
    Dim leftIdx As Long
    Dim rightIdx As Long
    Dim pivot As Variant
    Dim tPivot(2) As Variant
    Dim temp As Variant
   
    Dim i As Long
    Dim j As Long
    stack.Add stack.count + 1, LBound(ary)
    stack.Add stack.count + 1, UBound(ary)
    Do While stack.count > 0
               
        leftIdx = stack(stack.count - 1)
        rightIdx = stack(stack.count)
        stack.Remove stack.count
        stack.Remove stack.count

        'クイックソート
        If leftIdx < rightIdx Then
       
            pivot = ary((leftIdx + rightIdx) / 2)
           
            i = leftIdx
            j = rightIdx
           
            Do While i <= j
           
                Do While ary(i) < pivot
                    i = i + 1
                Loop
           
                Do While ary(j) > pivot
                    j = j - 1
                Loop
           
                If i <= j Then
                    temp = ary(i)
                    ary(i) = ary(j)
                    ary(j) = temp
                   
                    i = i + 1
                    j = j - 1
                End If
           
            Loop
           
            If rightIdx - i >= 0 Then
                If rightIdx - i <= 10 Then
                    insertion_sort ary, i, rightIdx
                Else
                    stack.Add stack.count + 1, i
                    stack.Add stack.count + 1, rightIdx
                End If
            End If
           
            If j - leftIdx >= 0 Then
                If j * leftIdx <= 10 Then
                    insertion_sort ary, leftIdx, j
                Else
                    stack.Add stack.count + 1, leftIdx
                    stack.Add stack.count + 1, j
                End If
            End If
        End If
   
    Loop

    quick_sort = ary
End Function


'非破壊非再帰クイック＆挿入ソートの挿入ソート
Private Function insertion_sort( _
    ary As Variant, _
    minIdx As Long, _
    maxIdx As Long)

    '挿入ソート
    Dim i As Long, j As Long
    Dim temp As Variant
    j = 1
    For j = minIdx To maxIdx
        i = j - 1
        Do While i >= 0
            If ary(i + 1) < ary(i) Then
                temp = ary(i + 1)
                ary(i + 1) = ary(i)
                ary(i) = temp
            Else
                Exit Do
            End If
            i = i - 1
        Loop
    Next
    
    insertion_sort = ary
End Function



