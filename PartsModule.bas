Attribute VB_Name = "PartsModule"
' @(h) PartsModule.Bas                ver 1.00 ( '01.10.01 SEC Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　関数パーツモジュール
' 　本モジュールは本システムで使用する関数パーツを集めた
' 　ものである。

Option Explicit
    
' @(f)
'
' 機能      : 表示用状態文字列取得
'
' 引き数    : ARG1 - nMode 0:ｽﾗﾌﾞ肌、1:ｶﾗｰﾁｪｯｸ
'          : ARG2 - 状態番号
'
' 返り値    : 状態文字列
'
' 機能説明  : 状態番号をコメント付きに変換する。
'
' 備考      :
'
Public Function ConvDpOutStat(ByVal nSysMode As Integer, nStat As Integer) As String

    Select Case nStat
        Case conDefine_SYSMODE_SKIN
            ConvDpOutStat = IIf(nSysMode = conDefine_SYSMODE_SKIN, "0:黒皮", "0:白皮")
        Case Else
            ConvDpOutStat = CStr(nStat) & ":" & CStr(nStat) & "ht後"
    End Select
    
End Function
    
' @(f)
'
' 機能      : スラブ番号変換
'
' 引き数    : ARG1 - ハイフン付きスラブ番号
'
' 返り値    : ハイフン無しスラブ番号
'
' 機能説明  : 指定したハイフン付きスラブ番号を’－’ハイフン無しスラブ番号に変換する。
'
' 備考      : アスタリスク’＊’がある場合は、パーセント’％’に変換する。
'           :COLORSYS
'
Public Function ConvSearchSlbNumber(ByVal strSearchSlbNumber As String) As String
    Dim nI As Integer
    Dim strResSearchSlbNumber As String
    
    'ハイフン’－’を取って実際の検索文字列へ変換
    For nI = 1 To Len(strSearchSlbNumber)
        If Mid(strSearchSlbNumber, nI, 1) <> "-" Then
            If Mid(strSearchSlbNumber, nI, 1) = "*" Then
                strResSearchSlbNumber = strResSearchSlbNumber & "%"
            Else
                strResSearchSlbNumber = strResSearchSlbNumber & Mid(strSearchSlbNumber, nI, 1)
            End If
        End If
    Next nI
    
    ConvSearchSlbNumber = strResSearchSlbNumber
    
End Function
    
' @(f)
'
' 機能      : デバックモード判別
'
' 引き数    : ARG1 - デバックモード文字列
'
' 返り値    : True=デバックＯＮ／False=デバックＯＦＦ
'
' 機能説明  : 指定したデバックモードの状態を判別する。
'
' 備考      :
'
Public Function IsDEBUG(ByVal strDEBUG As String) As Boolean

    IsDEBUG = False
    
    If APSysCfgData.nDEBUG_MODE <> 1 Then Exit Function
    
    Select Case strDEBUG
        Case "DISP"
            If APSysCfgData.nDISP_DEBUG = 1 Then IsDEBUG = True
        Case "FILE"
            If APSysCfgData.nFILE_DEBUG = 1 Then IsDEBUG = True
        Case "TR_SKIP"
            If APSysCfgData.nTR_SKIP = 1 Then IsDEBUG = True
        Case "DB_SKIP"
            If APSysCfgData.nDB_SKIP = 1 Then IsDEBUG = True
        Case "SOZAI_DB_SKIP"
            If APSysCfgData.nSOZAI_DB_SKIP = 1 Then IsDEBUG = True
        Case "SCAN"
            If APSysCfgData.nSCAN_SKIP = 1 Then IsDEBUG = True
        Case "HOSTDATA_DEBUG"
            If APSysCfgData.nHOSTDATA_DEBUG = 1 Then IsDEBUG = True
        Case "HOSTDATA_SKIP"
            If APSysCfgData.nHOSTDATA_SKIP = 1 Then IsDEBUG = True
        'end cho
    End Select
    
End Function
    
'' @(f)
''
'' 機能      : スラブ分割モード判別
''
'' 引き数    :
''
'' 返り値    : True=分割ＯＮ／False=分割ＯＦＦ
''
'' 機能説明  : スラブ分割モードの状態を判別する。
''
'' 備考      :
''
'Public Function IsAPSplit() As Boolean
'    'スラブは選択されているか。
'    If APSlbCont.nListSelectedIndexP1 = 0 Then
'        IsAPSplit = False
'    Else
'        '分割モードの場合。
'        If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).nSplitTotal > 1 Then
'            IsAPSplit = True
'        Else
'            IsAPSplit = False
'        End If
'    End If
'End Function

' @(f)
'
' 機能      : メッセージログの作成、表示、保存
'
' 引き数    : ARG1 - プロセス番号
'             ARG2 - メッセージ
'
' 返り値    :
'
' 機能説明  : メッセージログの作成、表示、保存を行う。
'
' 備考      : ガイダンス表示。
'
Public Sub MsgLog(ByVal nProcNum As Integer, ByVal strMessage As String)
    Dim strGuidanceMess As String
    
    Select Case nProcNum
        Case conProcNum_MAIN
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
            
            strGuidanceMess = Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision & conDefault_Separator & strMessage
        Case conProcNum_BSCONT
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & "ビジコン通信：" & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
              
        Case conProcNum_TRCONT
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & "通信サーバー通信：" & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
    
        Case conProcNum_MAINTENANCE
            strGuidanceMess = Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision & conDefault_Separator & "メンテナンス:" & strMessage
        
        Case conProcNum_WINSOCKCONT
            strGuidanceMess = strMessage
            
    End Select
    
    If IsEmpty(MainLogFileNumber) = False Then
        Print #MainLogFileNumber, strGuidanceMess
    End If

End Sub

' @(f)
'
' 機能      : INPUT MAN用入力チェック
'
' 引き数    : ARG1 -　INPUT MAN　オブジェクト
'
' 返り値    : TRUE/FALSE
'
' 機能説明  : INPUT MAN　CAPTIONのTEXTに指定されている、条件を満たすか判定し結果を返す。
'             ※KeepFocusへ設定可能
'
' 備考      : (下限値,上限値)(許可文字,...)
'
Public Function LimitCheck(ByVal obj As Object) As Boolean

    Dim obj_str As String
    Dim get_str As String
    Dim text_str, num_str As String
    Dim pos_int, pos_max_int, i, y As Integer
    Dim upper_sing, lower_sing As Single
    Dim array_str() As String
    Dim text_flag_bool As Boolean
    Dim obj_work_str As String

    On Error Resume Next

    'オブジェクトから許可文字列取得
    obj_str = obj.Caption.Text

    LimitCheck = False
    '括弧の有無判定
    If (InStr(obj_str, "(") = False) And (InStr(obj_str, ")") = False) Then
        Exit Function
    End If
    obj_work_str = obj
    If obj_work_str = "" Then
        obj_work_str = " "
    End If
    
    '数字範囲取得
    num_str = Mid(obj_str, 2, (InStr(obj_str, ")") - 2))
    '許可文字取得
    text_str = Mid(obj_str, InStr(obj_str, ")") + 2, (Len(obj_str) - Len(num_str) - 4))

    If IsNumeric(obj) Then
        '数値チェック
        If Len(num_str) <> 0 Then
            '最小値取得
            lower_sing = Mid(num_str, 1, (InStr(num_str, ",") - 1))
            '最大値取得
            upper_sing = Mid(num_str, InStr(num_str, ",") + 1, Len(num_str) - 1)
            '範囲チェック
            If (CDbl(obj_work_str) < CDbl(lower_sing)) Or (CDbl(obj_work_str) > CDbl(upper_sing)) Then
                'エラーリターン
                LimitCheck = True
                Exit Function
            Else
                Exit Function
            End If
        Else
            'エラーリターン
            LimitCheck = True
            Exit Function
        End If
    Else
        '文字チェック
        pos_int = 1
        If Len(text_str) <> 0 Then
            '文字数取得
            pos_max_int = Fix(Len(text_str) / 2) + 1
            '配列の再定義
            ReDim array_str(pos_max_int)
            '許可文字格納
            For i = 0 To pos_max_int - 1
                '文字取得
                array_str(i) = Mid(text_str, pos_int, 1)
                '取得位置更新
                pos_int = pos_int + 2
            Next
            '対象文字数分ループ
            For i = 0 To Len(obj_work_str) - 1
                '対象文字を１文字取得
                get_str = Mid(obj_work_str, i + 1, 1)
                'フラグの初期化
                text_flag_bool = False
                '文字数分ループ
                For y = 0 To pos_max_int - 1
                    '許可文字チェック
                    If (get_str = array_str(y)) Then
                        text_flag_bool = True
                    End If
                Next
                '許可文字判定
                If text_flag_bool = False Then
                    'エラーリターン
                    LimitCheck = True
                    Exit Function
                End If
            Next
        Else
            'エラーリターン
            LimitCheck = True
            Exit Function
        End If
    End If
End Function

' @(f)
'
' 機能      : システム日時から操業日付算出
'
' 引き数    :
'
' 返り値    : 操業日("YYYYMMDD")
'
' 機能説明  : システム日時から操業日付算出
'
' 備考      :
'
Public Function GetSyoGyoDate() As String
    On Error Resume Next                        'エラー処理

    Dim strSysDate              As String       'システム日付
    Dim strSysTime              As String       'システム時刻

    strSysDate = Format(Date$, "YYYY/MM/DD")
    strSysTime = Format(Time$, "HH:MM:SS")

    ':::: システム時刻が0時～7時30分以降ならば前日の日付として計算
    If "00:00:00" <= strSysTime And strSysTime <= "07:29:59" Then
        GetSyoGyoDate = Format(DateAdd("d", -1, strSysDate), "YYYYMMDD")   '操業日付セット
    Else
        GetSyoGyoDate = Format(strSysDate, "YYYYMMDD")                     '操業日付セット
    End If

    Debug.Print GetSyoGyoDate
End Function
' @(f)
'
' 機能      : INPUT MAN フォーマット設定処理
'
' 引き数    : ARG1 - INPUT MAN オフジェクト
' 　　　    ：ARG2 - Caption設定値
' 　　　    ：ARG3 - Format設定値
' 　　　    ：ARG4 - FormatMode設定値
' 　　　    ：ARG5 - bAllowSpace設定値
'
' 返り値    : 無し
'
' 機能説明  : INPUT MANのフォーマットを設定する。
'
' 備考      :
'
Public Sub SetimTextFormat(ByVal obj As Object, ByVal strCaption As String, ByVal strFormat As String, ByVal iFormatMode As Integer, ByVal bAllowSpace As Boolean)
    obj.Caption.Text = strCaption
    obj.Format = strFormat
    obj.FormatMode = iFormatMode
    obj.AllowSpace = bAllowSpace
    obj.EditMode = 3 '上書き（固定）
    obj.HighlightText = True 'テキスト選択
End Sub


Public Function cnvSplitNum(ByVal strSplitTNum As String) As Integer

    Select Case Trim(strSplitTNum)
        Case "1"
            cnvSplitNum = 1
        Case "2"
            cnvSplitNum = 2
        Case "3"
            cnvSplitNum = 3
        Case "4"
            cnvSplitNum = 4
        Case "5"
            cnvSplitNum = 5
        Case "6"
            cnvSplitNum = 6
        Case "7"
            cnvSplitNum = 7
        
        'Case "X"
        '    cnvSplitNum = 9
        'Case "Y"
        '    cnvSplitNum = 10
        '
        'Case "8"
        '    cnvSplitNum = 8
        'Case "9"
        '    cnvSplitNum = 8
        '
        Case Else
            cnvSplitNum = 0
    End Select

End Function

Public Function cnvSplitTNum(ByVal nSplitNum As Integer) As String

    Select Case nSplitNum
        Case 1
            cnvSplitTNum = "1"
        Case 2
            cnvSplitTNum = "2"
        Case 3
            cnvSplitTNum = "3"
        Case 4
            cnvSplitTNum = "4"
        Case 5
            cnvSplitTNum = "5"
        Case 6
            cnvSplitTNum = "6"
        Case 7
            cnvSplitTNum = "7"
        
        Case 8
            cnvSplitTNum = "9"
        Case 9
            cnvSplitTNum = "X"
        
        Case 10
            cnvSplitTNum = "Y"
        
        Case Else
            cnvSplitTNum = 0
    End Select

End Function

Public Function chkImgFile(ByVal strKEY As String, ByVal nSplitNum As Integer) As Boolean
        
    If Dir(App.path & "\" & conDefine_ImageDirName & "\" & strKEY & Format(nSplitNum, "00") & "(0).jpg") <> "" Then
        chkImgFile = True
    Else
        chkImgFile = False
    End If

End Function

Public Sub clrImgFile(ByVal strKEY As String)
    On Error Resume Next
    Call Kill(App.path & "\" & conDefine_ImageDirName & "\" & strKEY & ".jpg")
    On Error GoTo 0
End Sub

'COLORSYS
Public Sub init_APResData()
    Dim initData As typAPResData
    APResData = initData
End Sub

'
Public Function ReadCSV(ByVal strReadFilePath As String, ByRef strItemName() As String, ByRef strDataField() As String) As Boolean
    Dim READ_FileNumber As Variant ''ファイル番号
    Dim strBuf As String
    Dim strChk As String
    Dim pos, org_pos
    Dim nLine As Integer
    Dim nItemNum As Integer
    Dim strItem As String
    
    If Dir(strReadFilePath) = "" Then
        Call MsgLog(strReadFilePath & "が見つかりません。", True)
        ReadCSV = False 'NG
        Exit Function
    End If
    
    ReDim strItemName(0) 'クリアー
    
    READ_FileNumber = Empty
    READ_FileNumber = FreeFile               ' 未使用のファイル番号を取得します。
    Open strReadFilePath For Input As #READ_FileNumber

    'タイトル行読取と項目数、行数調査
    nLine = 0
    Do While Not EOF(READ_FileNumber)            ' ファイルの終端までループを繰り返します。
        Line Input #READ_FileNumber, strBuf      ' 行を変数に読み込みます。
        Debug.Print strBuf         ' イミディエイト ウィンドウに表示します。
        'Call MsgLog(strBuf)
        
        pos = 1
        org_pos = 1
        nItemNum = 1
        nLine = nLine + 1
    
        Do While True
             pos = InStr(org_pos, strBuf, ",", vbTextCompare)
             If org_pos <> 1 Then
                 If pos = 0 Then
                     If org_pos > Len(strBuf) + 1 Then
                         '行の最後
                         Exit Do
                     Else
                         pos = Len(strBuf) + 1
                     End If
                 End If
             Else
                 '一行なにもない
                 If pos = 0 Then
                     Exit Do
                 End If
             End If
             
            If nLine = 1 Then
                'タイトル行
                strItem = Mid(strBuf, org_pos, (pos - org_pos))
                strItemName(UBound(strItemName)) = strItem
                ReDim Preserve strItemName(UBound(strItemName) + 1)
            Else
                'データ行
            End If
             
            If pos = Len(strBuf) Then
                Exit Do
            Else
                org_pos = pos + 1
                nItemNum = nItemNum + 1
            End If
        Loop
    Loop
    
    If nLine <> 0 Then
        'データ行読取
        ReDim strDataField(UBound(strItemName), nLine - 1)
        Seek #READ_FileNumber, 1
        nLine = 0
        Do While Not EOF(READ_FileNumber)            ' ファイルの終端までループを繰り返します。
            Line Input #READ_FileNumber, strBuf      ' 行を変数に読み込みます。
            Debug.Print strBuf         ' イミディエイト ウィンドウに表示します。
            'Call MsgLog(strBuf)
            
            pos = 1
            org_pos = 1
            nItemNum = 1
            nLine = nLine + 1
        
            Do While True
                 pos = InStr(org_pos, strBuf, ",", vbTextCompare)
                 If org_pos <> 1 Then
                     If pos = 0 Then
                         If org_pos > Len(strBuf) + 1 Then
                             '行の最後
                             Exit Do
                         Else
                             pos = Len(strBuf) + 1
                         End If
                     End If
                 Else
                     '一行なにもない
                     If pos = 0 Then
                         Exit Do
                     End If
                 End If
                 
                If nLine = 1 Then
                    'タイトル行
                Else
                    'データ行
                    strItem = Mid(strBuf, org_pos, (pos - org_pos))
                    strDataField(nItemNum - 1, nLine - 2) = strItem
                End If
                 
                If pos = Len(strBuf) Then
                    Exit Do
                Else
                    org_pos = pos + 1
                    nItemNum = nItemNum + 1
                End If
            Loop
        Loop
    End If
    
    If IsEmpty(READ_FileNumber) = False Then
        Close #READ_FileNumber
        READ_FileNumber = Empty
        ReadCSV = True 'OK
    Else
        ReadCSV = False 'NG
    End If

End Function

Public Function getItemDataCSV(ByVal strTGTItemName As String, ByVal nTGTDataNumber As Integer, ByRef strItemName() As String, ByRef strDataField() As String) As String
    Dim nI As Integer
    
    For nI = 0 To UBound(strItemName) - 1
        If Trim(UCase(strItemName(nI))) = Trim(UCase(strTGTItemName)) Then
            getItemDataCSV = strDataField(nI, nTGTDataNumber - 1)
            Exit Function
        End If
    Next nI
    
    getItemDataCSV = ""
    
End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Double, ByVal iDigits As Integer) As Double
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Public Sub WaitMsgBox(ByVal callOwnerObj As Object, ByVal strMessage As String)
    Dim MsgWnd As Message
    Set MsgWnd = New Message
        
    Call MsgLog(conProcNum_MAIN, strMessage)
        
    MsgWnd.MsgText = strMessage
    MsgWnd.OK.Visible = True
'        MsgWnd.AutoDelete = True
    Do
        On Error Resume Next
        MsgWnd.Show vbModal, callOwnerObj
        If Err.Number = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    Set MsgWnd = Nothing

End Sub

' @(f)
'
' 機能      : 各種画像登録件数カウント
'
' 引き数    : ARG1 - MODE(SKIN COLOR FAIL)　スラブ肌・カラー・異常
' 　　　    ：ARG2 - CHNO   チャージＮＯ
' 　　　    ：ARG3 - AINO   合番
' 　　　    ：ARG4 - STAT   状態
' 　　　    ：ARG5 - COLOR  カラー回数
'
' 返り値    : カウントした数を文字列で戻す
'
' 機能説明  : フォルダに入っている画像件数をカウントする
'
' 備考      :
'
Public Function PhotoImgCount(ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStat As String, ByVal sColor As String) As String
    Dim objFso       As Object
    Dim iCnt         As Integer
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    
    On Error GoTo PhotoImgCount_Err
    
    ' 画像フォルダパス設定
    sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino

    ' 指定フォルダ探索
    iCnt = 0
    sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & "_??.JPG"
    sGetFileName = Dir(sPhotoPath & "\" & sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = "jpg" Or Right(sGetFileName, 3) = "JPG" Then
            iCnt = iCnt + 1
        End If

        sGetFileName = Dir()
    Loop
    Debug.Print sChkFileName
    Debug.Print iCnt
    PhotoImgCount = CStr(iCnt)
    
    Exit Function
    
PhotoImgCount_Err:
    PhotoImgCount = CStr(iCnt)
    On Error Resume Next
    
End Function


' @(f)
'
' 機能      : アップロード用フォルダ確認
'
' 引き数    : ARG1 - EXT    IMG・PDF・SCAN
' 　　　    ：ARG2 - MODE(SKIN COLOR FAIL)　スラブ肌・カラー・異常
' 　　　    ：ARG3 - CHNO   チャージＮＯ
' 　　　    ：ARG4 - AINO   合番
' 　　　    ：ARG5 - STAT   状態
' 　　　    ：ARG6 - COLOR  カラー回数
'
' 返り値    : True データ無／False データ有
'
' 機能説明  : 変更先『状態』が使用できるか確認する
'
' 備考      :
'
Public Function StatChgFoldCheck(ByVal sExtDir As String, ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStat As String, ByVal sColor As String) As Boolean
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    Dim sExt         As String
    Dim errNum       As Long
    Dim errDesc      As String
    Dim errSrc       As String
    Dim StrTmp       As String
    
    On Error GoTo StatChgFoldCheck_Err
    
    ' 画像フォルダパス設定
    If sExtDir = "IMG" Then
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & "_??.JPG"
    ElseIf sExtDir = "PDF" Then
        sExt = "PDF"
        sPhotoPath = APSysCfgData.SHARES_PDFDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & ".PDF"
    Else
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_SCNDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & ".JPG"
    End If

    StatChgFoldCheck = True

    ' 指定フォルダ探索
    On Error Resume Next
    sGetFileName = Dir(sPhotoPath & "\" & sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = sExt Then
            StatChgFoldCheck = False
            Exit Do
        End If

        sGetFileName = Dir()
    Loop
    Debug.Print sChkFileName
    
    Exit Function
    
StatChgFoldCheck_Err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "StatChgFoldCheck 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    StatChgFoldCheck = False

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : アップロード用フォルダ・ファイル変更
'
' 引き数    : ARG1 - MODE       SKINスラブ肌・COLORカラー・FAIL異常
' 　　　    ：ARG2 - CHNO       チャージＮＯ
' 　　　    ：ARG3 - AINO       合番
' 　　　    ：ARG4 - STATOLD    変更前状態
' 　　　    ：ARG4 - STATNEW    変更後状態
' 　　　    ：ARG5 - COLOROLD      カラー回数
' 　　　    ：ARG5 - COLORNEW   カラー回数
' 　　　    ：ARG6 - EXT        IMG・PDF・SCAN
'
' 返り値    : True データ無／False データ有
'
' 機能説明  : 変更先『状態』が使用できるか確認する
'
' 備考      :
'
Public Function StatChgFoldFix(ByVal sExtDir As String, ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatOld As String, ByVal sStatNew As String, ByVal sColorOld As String, ByVal sColorNew As String) As Boolean
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    Dim sNewFileName As String    ' 変更ファイル名
    Dim sExt         As String
    Dim sKeepNo      As String
    Dim errNum       As Long
    Dim errDesc      As String
    Dim errSrc       As String
    Dim StrTmp       As String
    
    On Error GoTo StatChgFoldFix_Err
    
    ' 既存画像フォルダパス設定 ********************************************************************
    If sExtDir = "IMG" Then
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & "_??.JPG"
    ElseIf sExtDir = "PDF" Then
        sExt = "PDF"
        sPhotoPath = APSysCfgData.SHARES_PDFDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & ".PDF"
    Else
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_SCNDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & ".JPG"
    End If

    StatChgFoldFix = True

    ' 指定フォルダ探索 ****************************************************************************
    sGetFileName = Dir(sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = sExt Then
            ' 変更名設定
            If sExtDir = "IMG" Then
                sKeepNo = Left(Right(sGetFileName, 6), 2)
                sNewFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColorNew & "_" & sKeepNo & "." & sExt
            Else
                sNewFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColorNew & "." & sExt
            End If
            Debug.Print "NAME " & sPhotoPath & "\" & sGetFileName & " AS " & sNewFileName
            Call MsgLog(conProcNum_MAIN, "状態変更[" & "NAME " & sPhotoPath & "\" & sGetFileName & " AS " & sNewFileName & "]") 'ガイダンス表示
            
            ' ファイル名変更実行
            Name sPhotoPath & "\" & sGetFileName As sNewFileName
        End If

        sGetFileName = Dir()
    Loop
    
    Exit Function
    
StatChgFoldFix_Err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "StatChgFoldFix 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    StatChgFoldFix = False

    On Error GoTo 0
    
End Function

