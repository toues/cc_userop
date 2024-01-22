Attribute VB_Name = "DBModule"
' @(h) DBModule.Bas                ver 1.00 ( '02.01.10 SEC Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　データベースモジュール
' 　本モジュールはデータベースアクセスの
' 　ためのものである。

'// ODBC Driver Oracle 8.01.66.00
Option Explicit

Const ORAPARM_INPUT As Integer = 1  '入力変数
Const ORAPARM_OUTPUT As Integer = 2 '出力変数
Const ORAPARM_BOTH As Integer = 3   '入力変数と出力変数の両方

Const ORATYPE_VARCHAR2 As Integer = 1
Const ORATYPE_NUMBER As Integer = 2
Const ORATYPE_SINT As Integer = 3
Const ORATYPE_FLOAT As Integer = 4
Const ORATYPE_STRING As Integer = 5
Const ORATYPE_DECIMAL As Integer = 7
Const ORATYPE_VARCHAR As Integer = 9
Const ORATYPE_DATE As Integer = 12
Const ORATYPE_REAL As Integer = 21
Const ORATYPE_DOUBLE As Integer = 22
Const ORATYPE_UNSIGNED8 As Integer = 23
Const ORATYPE_UNSIGNED16 As Integer = 25
Const ORATYPE_UNSIGNED32 As Integer = 26
Const ORATYPE_SIGNED8 As Integer = 27
Const ORATYPE_SIGNED16 As Integer = 28
Const ORATYPE_SIGNED32 As Integer = 29
Const ORATYPE_PTR As Integer = 32
Const ORATYPE_OPAQUE As Integer = 58
Const ORATYPE_UINT As Integer = 68
Const ORATYPE_RAW As Integer = 95
Const ORATYPE_CHAR As Integer = 96
Const ORATYPE_CHARZ As Integer = 97
Const ORATYPE_CURSOR As Integer = 102
Const ORATYPE_ROWID As Integer = 104
Const ORATYPE_MLSLABEL As Integer = 105
Const ORATYPE_OBJECT As Integer = 108
Const ORATYPE_REF As Integer = 110
Const ORATYPE_CLOB As Integer = 112
Const ORATYPE_BLOB As Integer = 113
Const ORATYPE_BFILE As Integer = 114
Const ORATYPE_CFILE As Integer = 115
Const ORATYPE_RSLT As Integer = 116
Const ORATYPE_NAMEDCOLLECTION As Integer = 122
Const ORATYPE_COLL As Integer = 122
Const ORATYPE_SYSFIRST As Integer = 228
Const ORATYPE_SYSLAST As Integer = 235
Const ORATYPE_OCTET As Integer = 245
Const ORATYPE_SMALLINT As Integer = 246
Const ORATYPE_VARRAY As Integer = 247
Const ORATYPE_TABLE As Integer = 248
Const ORATYPE_OTMLAST As Integer = 320

Private Const conDef_DB_ProcessName As String = "カラーチェック実績ＰＣ" ''データベース使用プロセス名

'Private Const conDef_DB_CHUNK_SIZE As Long = 16384  'チャンクサイズ

' @(f)
'
' 機能      : ＯＤＢＣ接続文字列取得
'
' 引き数    : ARG1 - 接続切り替えフラグ
'
' 返り値    : 接続文字列
'
' 機能説明  : ＯＤＢＣ接続文字列を取得する。
'
' 備考      :COLORSYS
'
Public Function DBConnectStr(ByVal sw As Integer, ByRef host As String, ByRef id As String, ByRef pass As String) As Boolean
    Select Case sw
    
    Case conDefault_DBConnect_MYUSER
        host = APSysCfgData.DB_MYUSER_DSN
        pass = APSysCfgData.DB_MYUSER_UID
        id = APSysCfgData.DB_MYUSER_PWD
        DBConnectStr = True
    
    Case conDefault_DBConnect_MYCOMN
        host = APSysCfgData.DB_MYCOMN_DSN
        pass = APSysCfgData.DB_MYCOMN_UID
        id = APSysCfgData.DB_MYCOMN_PWD
        DBConnectStr = True
    
    Case conDefault_DBConnect_SOZAI
        host = APSysCfgData.DB_SOZAI_DSN
        pass = APSysCfgData.DB_SOZAI_UID
        id = APSysCfgData.DB_SOZAI_PWD
        DBConnectStr = True
    
    Case Else
        DBConnectStr = False
    End Select

End Function

' @(f)
'
' 機能      : ＳＱＬ実行処理
'
' 引き数    : ARG1 - 接続切り替えフラグ
'             ARG2 - ＳＱＬ文字列
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定の接続を使用してＳＱＬ文字列を実行する
'
' 備考      :
'
Public Function DB_SQL_Execute(ByVal nConnectSw As Integer, ByVal strSQL As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute:ＤＢスキップモードです。") 'ガイダンス表示
        
        DB_SQL_Execute = True
        Exit Function
    End If
    
    On Error GoTo DB_SQL_Execute_err
    
    nOpen = 0
    
    ' Oracleとの接続を確立する
    '-cn.Open DBConnectStr(nConnectSw)
    bRet = DBConnectStr(nConnectSw, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0
    
    Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute 正常終了") 'ガイダンス表示

    DB_SQL_Execute = True

    On Error GoTo 0
    Exit Function

DB_SQL_Execute_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If
    
    DB_SQL_Execute = False
    
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : スラブ肌情報検索処理
'
' 引き数    : ARG1 - 検索オプション番号（未使用）
'             ARG2 - 最大検索件数
'             ARG3 - 検索スラブＮｏ．
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      :
'
Public Function DBSkinSlbSearchRead(ByVal nSearchOption As Integer, ByVal nSEARCH_MAX As Integer, ByVal nSERCH_RANGE As Integer, ByVal strSearchSlbNumber As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    Dim lRecCnt As Long
    
    Dim strSERCH_RANGE As String
    
    If nSERCH_RANGE = 9999 Then
        strSERCH_RANGE = ""
    Else
        strSERCH_RANGE = Format(DateAdd("d", -nSERCH_RANGE, Now), "YYYYMMDD")
    End If
    
    ''ＤＢオフラインで強制入力を行ったことを判断するフラグ
'    bAPInputOffline = False
'    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
        
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead:ＤＢスキップモードです。") 'ガイダンス表示
        
'******************************
        'DEMO
        ReDim APSearchTmpSlbData(0)
        
        Call DBSkinSlbSearchReadCSV
        
        Select Case nSearchOption
          Case 0 'スラブＮｏ．検索
'            For nI = 0 To 19
'                'スラブＮｏ．
'                APSearchTmpSlbData(nI).slb_chno = CStr(nI + 10000)
'                APSearchTmpSlbData(nI).slb_aino = CStr(nI + 1000)
'                APSearchTmpSlbData(nI).slb_no = APSearchTmpSlbData(nI).slb_chno & APSearchTmpSlbData(nI).slb_aino
'
'                '状態
'                APSearchTmpSlbData(nI).slb_stat = nI Mod 6
'
'                '鋼種
'                APSearchTmpSlbData(nI).slb_ksh = "AAAAAA"
'
'                '型
'                APSearchTmpSlbData(nI).slb_typ = "AAA"
'
'                '向先
'                APSearchTmpSlbData(nI).slb_uksk = "AAA"
'
'                '造塊日
'                APSearchTmpSlbData(nI).slb_zkai_dte = "20080310"
'
'                'ｽﾗﾌﾞ肌実績（初回記録日）
'                APSearchTmpSlbData(nI).sys_wrt_dte = "20080310"
'
'                'ｽﾗﾌﾞ肌ｲﾒｰｼﾞ
'                APSearchTmpSlbData(nI).bAPScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                'ｽﾗﾌﾞ肌PDF
'                APSearchTmpSlbData(nI).bAPPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'                ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
'            Next nI
        End Select
'******************************
            
        DBSkinSlbSearchRead = True
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBSkinSlbSearchRead_err
    
    Select Case nSearchOption
    Case 0 'スラブＮｏ．検索
        strSQL = "SELECT TRTS0012.*,TRTS0040.SLB_PDF_ADDR,TRTS0040.SYS_WRT_DTE AS SYS_WRT_DTE40,TRTS0050.SLB_SCAN_ADDR "
        
        strSQL = strSQL & "FROM (TRTS0012 LEFT JOIN TRTS0040 ON (TRTS0012.SLB_STAT = TRTS0040.SLB_STAT) "
        strSQL = strSQL & "AND (TRTS0012.SLB_NO = TRTS0040.SLB_NO)) "
        strSQL = strSQL & "LEFT JOIN TRTS0050 ON (TRTS0012.SLB_STAT = TRTS0050.SLB_STAT) "
        strSQL = strSQL & "AND (TRTS0012.SLB_NO = TRTS0050.SLB_NO) "
        
        strSQL = strSQL & "WHERE TRTS0012.SLB_NO LIKE '" & strSearchSlbNumber & "' "

        If strSERCH_RANGE <> "" Then
            strSQL = strSQL & _
            "And (TRTS0012.SYS_WRT_DTE >= '" & strSERCH_RANGE & "') "
        End If

        strSQL = strSQL & "ORDER BY TRTS0012.SYS_WRT_DTE DESC ,TRTS0012.SYS_WRT_TME DESC"
    End Select

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    nI = 0
    ReDim APSearchTmpSlbData(nI)
    Do While Not oDS.EOF
        
        APSearchTmpSlbData(nI).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value) ''スラブＮｏ．
        APSearchTmpSlbData(nI).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value) ''スラブチャージＮｏ．
        APSearchTmpSlbData(nI).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) ''スラブ合番
        APSearchTmpSlbData(nI).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value) ''状態
        APSearchTmpSlbData(nI).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), "", oDS.Fields("slb_zkai_dte").Value) ''造塊日
        APSearchTmpSlbData(nI).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), "", oDS.Fields("slb_ksh").Value) ''鋼種
        APSearchTmpSlbData(nI).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), "", oDS.Fields("slb_typ").Value) ''型
        APSearchTmpSlbData(nI).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), "", oDS.Fields("slb_uksk").Value) ''向先
        APSearchTmpSlbData(nI).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value) ''記録日（初回記録日）
        
        APSearchTmpSlbData(nI).sAPPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE40").Value), "", oDS.Fields("SYS_WRT_DTE40").Value) ''PDFイメージデータ記録日（初回記録日）
        
        If IsNull(oDS.Fields("SLB_SCAN_ADDR").Value) = False Then
            APSearchTmpSlbData(nI).bAPScanInput = True ''SCANデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False ''SCANデータ有りフラグ
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR").Value) = False Then
            APSearchTmpSlbData(nI).bAPPdfInput = True ''PDFイメージデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False ''PDFイメージデータ有りフラグ
        End If
        
        ' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
        APSearchTmpSlbData(nI).PhotoImgCnt1 = PhotoImgCount("SKIN", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, "00")
        
        ReDim Preserve APSearchTmpSlbData(nI + 1) 'スラブ選択画面検索リスト
    
        oDS.MoveNext

        nI = nI + 1 '格納用インデックス

        '設定０の場合制限無しとなる。
        If nSEARCH_MAX = nI Then
            Exit Do
        '最大リミッターconDefault_nSEARCH_MAX0 = 9999
        ElseIf nI > conDefault_nSEARCH_MAX0 Then
            Exit Do
        End If
    Loop

    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    DBSkinSlbSearchRead = True

    Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead 正常終了") 'ガイダンス表示

    On Error GoTo 0

    Exit Function

DBSkinSlbSearchRead_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBSkinSlbSearchRead = False

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : カラーチェック情報検索処理
'
' 引き数    : ARG1 - 検索オプション番号（0:通常検索,1:異常報告一覧検索)
'             ARG2 - 最大検索件数
'             ARG3 - 検索スラブＮｏ．
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      : 2008/09/03 異常報告一覧検索追加
'
Public Function DBColorSlbSearchRead(ByVal nSearchOption As Integer, ByVal nSEARCH_MAX As Integer, ByVal nSERCH_RANGE As Integer, ByVal strSearchSlbNumber As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    Dim lRecCnt As Long
    Dim sImageCnt As String     ' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
    
    Dim strRes_Wrt_Dte_Max As String
    Dim strNotCmp_Res_No_MIN As String
    
    Dim strSERCH_RANGE As String
    
    If nSERCH_RANGE = 9999 Then
        strSERCH_RANGE = ""
    Else
        strSERCH_RANGE = Format(DateAdd("d", -nSERCH_RANGE, Now), "YYYYMMDD")
    End If
    
    ''ＤＢオフラインで強制入力を行ったことを判断するフラグ
'    bAPInputOffline = False
'    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
        
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead:ＤＢスキップモードです。") 'ガイダンス表示
        
'******************************
        'DEMO
        ReDim APSearchTmpSlbData(0)
        
        Call DBColorSlbSearchReadCSV
        
        Select Case nSearchOption
          Case 0 'スラブＮｏ．検索
'            For nI = 0 To 19
'                'スラブＮｏ．
'                APSearchTmpSlbData(nI).slb_chno = CStr(nI + 10000)
'                APSearchTmpSlbData(nI).slb_aino = CStr(nI + 1000)
'                APSearchTmpSlbData(nI).slb_no = APSearchTmpSlbData(nI).slb_chno & APSearchTmpSlbData(nI).slb_aino
'
'                '状態
'                APSearchTmpSlbData(nI).slb_stat = nI Mod 6
'
'                '●ｶﾗｰ回数
'                APSearchTmpSlbData(nI).slb_col_cnt = "01"
'
'                '鋼種
'                APSearchTmpSlbData(nI).slb_ksh = "AAAAAA"
'
'                '型
'                APSearchTmpSlbData(nI).slb_typ = "AAA"
'
'                '向先
'                APSearchTmpSlbData(nI).slb_uksk = "AAA"
'
'                '造塊日
'                APSearchTmpSlbData(nI).slb_zkai_dte = "20080310"
'
'                'ｶﾗｰ実績（初回記録日）
'                APSearchTmpSlbData(nI).sys_wrt_dte = "20080310"
'
'                '●ビジコン送信結果
'                APSearchTmpSlbData(nI).host_send = ""
'
'                'ｶﾗｰｲﾒｰｼﾞ
'                APSearchTmpSlbData(nI).bAPScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                'ｶﾗｰPDF
'                APSearchTmpSlbData(nI).bAPPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'***********************************************************************
'                '異常報告（初回記録日）
'                APSearchTmpSlbData(nI).fail_sys_wrt_dte = "20080310"
'
'                '異常報告ビジコン送信結果
'                APSearchTmpSlbData(nI).fail_host_send = ""
'
'                '異常ｲﾒｰｼﾞ
'                APSearchTmpSlbData(nI).bAPFailScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                '異常PDF
'                APSearchTmpSlbData(nI).bAPFailPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'***********************************************************************
'                'CCNO
'                APSearchTmpSlbData(nI).slb_ccno = "10000"
'
'                '重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
'                APSearchTmpSlbData(nFirstDataIndex).slb_color_wei
'
'                '長さ
'                APSearchTmpSlbData(nFirstDataIndex).slb_lngth
'
'                '幅
'                APSearchTmpSlbData(nFirstDataIndex).slb_wdth
'
'                '厚み
'                APSearchTmpSlbData(nFirstDataIndex).slb_thkns
'
'***********************************************************************
'                '処置指示
'                APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = "20080310"
'
'***********************************************************************
'                '処置結果
'                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = "20080310"
'
'                '処置結果完了フラグ
'                APSearchTmpSlbData(nI).fail_res_cmp_flg = "1"
'
'***********************************************************************
'
'                ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
'            Next nI
        End Select
'******************************
            
        DBColorSlbSearchRead = True
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBColorSlbSearchRead_err
    
    '******************************************************
    'SQL文初期化
    strSQL = ""
    
    'SQL文前処理
    Select Case nSearchOption
        Case 0 'スラブＮｏ．検索
            '無し
        Case 1 '異常報告一覧
            strSQL = strSQL & "SELECT * FROM ("
    End Select
    
    '******************************************************
    'スラブＮｏ．検索
    strSQL = strSQL & "SELECT TRTS0014.*, "
    strSQL = strSQL & "TRTS0016.HOST_SEND AS HOST_SEND16, "
    strSQL = strSQL & "TRTS0016.HOST_WRT_DTE AS HOST_WRT_DTE16, "
    strSQL = strSQL & "TRTS0016.HOST_WRT_TME AS HOST_WRT_TME16, "
    strSQL = strSQL & "TRTS0016.SYS_WRT_DTE AS SYS_WRT_DTE16, "
    strSQL = strSQL & "TRTS0016.SYS_WRT_TME AS SYS_WRT_TME16, "
    strSQL = strSQL & "TRTS0042.SLB_PDF_ADDR AS SLB_PDF_ADDR42, "
    strSQL = strSQL & "TRTS0042.SYS_WRT_DTE AS SYS_WRT_DTE42, "
    strSQL = strSQL & "TRTS0044.SLB_PDF_ADDR AS SLB_PDF_ADDR44, "
    strSQL = strSQL & "TRTS0044.SYS_WRT_DTE AS SYS_WRT_DTE44, "
    strSQL = strSQL & "TRTS0052.SLB_SCAN_ADDR AS SLB_SCAN_ADDR52, "
    strSQL = strSQL & "TRTS0054.SLB_SCAN_ADDR AS SLB_SCAN_ADDR54, "
    strSQL = strSQL & "T20.DIR_WRT_DTE_MAX, "
    strSQL = strSQL & "T20.DIR_PRN_OUT_MAX, " '2008/09/04
    strSQL = strSQL & "T22A.NOTCMP_RES_NO_MIN, "
    strSQL = strSQL & "T22B.RES_WRT_DTE_MAX, "
    strSQL = strSQL & "T22C.HOST_SEND AS HOST_SEND22, "
    strSQL = strSQL & "T22C.HOST_WRT_DTE AS HOST_WRT_DTE22, "
    strSQL = strSQL & "T22C.HOST_WRT_TME AS HOST_WRT_TME22 "
    strSQL = strSQL & "FROM ((((((((TRTS0014 "
    
    strSQL = strSQL & "LEFT JOIN TRTS0016 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0016.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0016.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0016.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0042 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0042.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0042.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0042.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0044 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0044.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0044.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0044.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0052 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0052.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0052.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0052.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0054 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0054.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0054.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0054.SLB_COL_CNT)) "
    
    'TRTS0020
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Max(TRTS0020.DIR_NO), TRTS0020.SLB_NO, TRTS0020.SLB_STAT, "
    strSQL = strSQL & "TRTS0020.SLB_COL_CNT, Max(TRTS0020.DIR_PRN_OUT) AS DIR_PRN_OUT_MAX, Max(TRTS0020.DIR_WRT_DTE) AS DIR_WRT_DTE_MAX FROM TRTS0020 "
    strSQL = strSQL & "GROUP BY TRTS0020.SLB_NO, TRTS0020.SLB_STAT, TRTS0020.SLB_COL_CNT "
    strSQL = strSQL & "ORDER BY Max(TRTS0020.DIR_WRT_DTE)) T20 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T20.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T20.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T20.SLB_COL_CNT)) "
    
    'TRTS0022 A 未完了の存在調査
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Min(RES_NO) AS NOTCMP_RES_NO_MIN, SLB_NO, SLB_STAT, "
    strSQL = strSQL & "SLB_COL_CNT FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY RES_CMP_FLG, SLB_NO, SLB_STAT, SLB_COL_CNT HAVING (RES_CMP_FLG Is Null)) T22A "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22A.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22A.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22A.SLB_COL_CNT)) "
    
    'TRTS0022 B 完了の存在調査
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Max(TRTS0022.RES_NO), TRTS0022.SLB_NO, TRTS0022.SLB_STAT, "
    strSQL = strSQL & "TRTS0022.SLB_COL_CNT, Max(TRTS0022.RES_WRT_DTE) AS RES_WRT_DTE_MAX FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, TRTS0022.RES_CMP_FLG "
    strSQL = strSQL & "HAVING (((TRTS0022.RES_CMP_FLG)='1')) "
    strSQL = strSQL & "ORDER BY Max(TRTS0022.RES_WRT_DTE)) T22B "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22B.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22B.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22B.SLB_COL_CNT)) "
    
    'TRTS0022 C ビジコン送信関係取得
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Min(TRTS0022.RES_NO), TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, "
    strSQL = strSQL & "TRTS0022.HOST_SEND, TRTS0022.HOST_WRT_DTE, TRTS0022.HOST_WRT_TME "
    strSQL = strSQL & "FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, "
    strSQL = strSQL & "TRTS0022.HOST_SEND, TRTS0022.HOST_WRT_DTE, TRTS0022.HOST_WRT_TME) T22C "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22C.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22C.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22C.SLB_COL_CNT) "
    
    strSQL = strSQL & "WHERE TRTS0014.SLB_NO LIKE '" & strSearchSlbNumber & "' "

    If strSERCH_RANGE <> "" Then
        strSQL = strSQL & _
        "And (TRTS0014.SYS_WRT_DTE >= '" & strSERCH_RANGE & "') "
    End If

    strSQL = strSQL & "ORDER BY TRTS0014.SYS_WRT_DTE DESC ,TRTS0014.SYS_WRT_TME DESC"
    '******************************************************
    'SQL文後処理
    Select Case nSearchOption
        Case 0 'スラブＮｏ．検索
            '無し
        Case 1 '異常報告一覧:    (異常報告有り) AND ((未送信=2) OR (完無) OR (完有 AND 未完有))
            strSQL = strSQL & ") WHERE (SYS_WRT_DTE16 Is Not Null) AND ((HOST_SEND22 = '2') OR (RES_WRT_DTE_MAX Is Null) OR (RES_WRT_DTE_MAX Is Not Null AND NOTCMP_RES_NO_MIN Is Not Null))"
    End Select

    '******************************************************

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    nI = 0
    ReDim APSearchTmpSlbData(nI)
    Do While Not oDS.EOF
        
        APSearchTmpSlbData(nI).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value) ''スラブＮｏ．
        APSearchTmpSlbData(nI).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value) ''状態
        APSearchTmpSlbData(nI).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), "", oDS.Fields("slb_col_cnt").Value) ''ｶﾗｰ回数
        APSearchTmpSlbData(nI).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value) ''スラブチャージＮｏ．
        APSearchTmpSlbData(nI).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) ''スラブ合番
        
        APSearchTmpSlbData(nI).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), "", oDS.Fields("slb_ccno").Value) ''CCNO
        
        APSearchTmpSlbData(nI).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), "", oDS.Fields("slb_zkai_dte").Value) ''造塊日
        APSearchTmpSlbData(nI).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), "", oDS.Fields("slb_ksh").Value) ''鋼種
        APSearchTmpSlbData(nI).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), "", oDS.Fields("slb_typ").Value) ''型
        APSearchTmpSlbData(nI).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), "", oDS.Fields("slb_uksk").Value) ''向先
        
        APSearchTmpSlbData(nI).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), "", oDS.Fields("slb_wei").Value) ''重量
        APSearchTmpSlbData(nI).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), "", oDS.Fields("slb_lngth").Value) ''長さ
        APSearchTmpSlbData(nI).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), "", oDS.Fields("slb_wdth").Value) ''幅
        APSearchTmpSlbData(nI).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), "", oDS.Fields("slb_thkns").Value) ''厚み
        
        APSearchTmpSlbData(nI).host_send = IIf(IsNull(oDS.Fields("host_send").Value), "", oDS.Fields("host_send").Value) ''ビジコン送信結果
        APSearchTmpSlbData(nI).host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte").Value), "", oDS.Fields("host_wrt_dte").Value) ''ビジコン送信日（初回ビジコン送信日）
        APSearchTmpSlbData(nI).host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme").Value), "", oDS.Fields("host_wrt_tme").Value) ''ビジコン送信時刻（初回ビジコン送信時刻）
        APSearchTmpSlbData(nI).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value) ''記録日（初回記録日）
        APSearchTmpSlbData(nI).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), "", oDS.Fields("sys_wrt_tme").Value) ''記録時刻（初回記録時刻）
        
        '異常一覧リスト表示専用 '2008/09/04
        APSearchTmpSlbData(nI).slb_fault_e_judg = IIf(IsNull(oDS.Fields("slb_fault_e_judg").Value), "", oDS.Fields("slb_fault_e_judg").Value)  ''欠陥E面判定
        APSearchTmpSlbData(nI).slb_fault_w_judg = IIf(IsNull(oDS.Fields("slb_fault_w_judg").Value), "", oDS.Fields("slb_fault_w_judg").Value)  ''欠陥W面判定
        APSearchTmpSlbData(nI).slb_fault_s_judg = IIf(IsNull(oDS.Fields("slb_fault_s_judg").Value), "", oDS.Fields("slb_fault_s_judg").Value)  ''欠陥S面判定
        APSearchTmpSlbData(nI).slb_fault_n_judg = IIf(IsNull(oDS.Fields("slb_fault_n_judg").Value), "", oDS.Fields("slb_fault_n_judg").Value)  ''欠陥N面判定
        
        '******************
        'スラブ異常
        APSearchTmpSlbData(nI).fail_host_send = IIf(IsNull(oDS.Fields("host_send16").Value), "", oDS.Fields("host_send16").Value) ''ビジコン送信結果
        APSearchTmpSlbData(nI).fail_host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte16").Value), "", oDS.Fields("host_wrt_dte16").Value) ''ビジコン送信日（初回ビジコン送信日）
        APSearchTmpSlbData(nI).fail_host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme16").Value), "", oDS.Fields("host_wrt_tme16").Value) ''ビジコン送信時刻（初回ビジコン送信時刻）
        APSearchTmpSlbData(nI).fail_sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte16").Value), "", oDS.Fields("sys_wrt_dte16").Value) ''記録日（初回記録日）
        APSearchTmpSlbData(nI).fail_sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme16").Value), "", oDS.Fields("sys_wrt_tme16").Value) ''記録時刻（初回記録時刻）
        '******************
        
        APSearchTmpSlbData(nI).sAPPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE42").Value), "", oDS.Fields("SYS_WRT_DTE42").Value) ''PDFイメージデータ記録日（初回記録日）
        APSearchTmpSlbData(nI).sAPFailPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE44").Value), "", oDS.Fields("SYS_WRT_DTE44").Value) ''PDFイメージデータ記録日（初回記録日）
        
        If IsNull(oDS.Fields("SLB_SCAN_ADDR52").Value) = False Then
            APSearchTmpSlbData(nI).bAPScanInput = True ''SCANデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False ''SCANデータ有りフラグ
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR42").Value) = False Then
            APSearchTmpSlbData(nI).bAPPdfInput = True ''PDFイメージデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False ''PDFイメージデータ有りフラグ
        End If
            
        '******************
        'スラブ異常
        If IsNull(oDS.Fields("SLB_SCAN_ADDR54").Value) = False Then
            APSearchTmpSlbData(nI).bAPFailScanInput = True ''SCANデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPFailScanInput = False ''SCANデータ有りフラグ
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR44").Value) = False Then
            APSearchTmpSlbData(nI).bAPFailPdfInput = True ''PDFイメージデータ有りフラグ
        Else
            APSearchTmpSlbData(nI).bAPFailPdfInput = False ''PDFイメージデータ有りフラグ
        End If
        '******************
            
        '処置指示
        APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = IIf(IsNull(oDS.Fields("dir_wrt_dte_max").Value), "", oDS.Fields("dir_wrt_dte_max").Value)
        '2008/09/04 印刷済みフラグ
        APSearchTmpSlbData(nI).fail_dir_prn_out_max = IIf(IsNull(oDS.Fields("DIR_PRN_OUT_MAX").Value), "", oDS.Fields("DIR_PRN_OUT_MAX").Value)
            
        '処置結果
        strRes_Wrt_Dte_Max = IIf(IsNull(oDS.Fields("res_wrt_dte_max").Value), "", oDS.Fields("res_wrt_dte_max").Value)
        strNotCmp_Res_No_MIN = IIf(IsNull(oDS.Fields("notcmp_res_no_min").Value), "", oDS.Fields("notcmp_res_no_min").Value)
        
        '2016/04/20 - TAI - S
        '作業場
        APSearchTmpSlbData(nI).slb_works_sky_tok = IIf(IsNull(oDS.Fields("slb_works_sky_tok").Value), "", oDS.Fields("slb_works_sky_tok").Value) ''作業場
        '2016/04/20 - TAI - E
        
        APSearchTmpSlbData(nI).fail_res_host_send = IIf(IsNull(oDS.Fields("host_send22").Value), "", oDS.Fields("host_send22").Value) ''ビジコン送信結果
        APSearchTmpSlbData(nI).fail_res_host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte22").Value), "", oDS.Fields("host_wrt_dte22").Value) ''ビジコン送信日（初回ビジコン送信日）
        APSearchTmpSlbData(nI).fail_res_host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme22").Value), "", oDS.Fields("host_wrt_tme22").Value) ''ビジコン送信時刻（初回ビジコン送信時刻）
        
        APSearchTmpSlbData(nI).fail_res_cmp_flg = "0"
        APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
        
        If strRes_Wrt_Dte_Max <> "" Then
            '登録日付有り
            If strNotCmp_Res_No_MIN <> "" Then
                '未完了レコード有り
                APSearchTmpSlbData(nI).fail_res_cmp_flg = "0"
                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
            Else
                '全て完了
                APSearchTmpSlbData(nI).fail_res_cmp_flg = "1"
                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = strRes_Wrt_Dte_Max
            End If
        Else
            '登録無し
            APSearchTmpSlbData(nI).fail_res_cmp_flg = ""
            APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
        End If
        
        ReDim Preserve APSearchTmpSlbData(nI + 1) 'スラブ選択画面検索リスト
    
        ' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
'        If nSearchOption = 0 Then
            sImageCnt = PhotoImgCount("COLOR", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
            APSearchTmpSlbData(nI).PhotoImgCnt1 = sImageCnt
            sImageCnt = PhotoImgCount("SLBFAIL", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
            APSearchTmpSlbData(nI).PhotoImgCnt2 = sImageCnt
'        Else
'            sImageCnt = PhotoImgCount("SLBFAIL", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
'            APSearchTmpSlbData(nI).PhotoImgCnt1 = sImageCnt
'        End If
        
        oDS.MoveNext

        nI = nI + 1 '格納用インデックス

        '設定０の場合制限無しとなる。
        If nSEARCH_MAX = nI Then
            Exit Do
        '最大リミッターconDefault_nSEARCH_MAX0 = 9999
        ElseIf nI > conDefault_nSEARCH_MAX0 Then
            Exit Do
        End If
    Loop

    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    DBColorSlbSearchRead = True

    Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead 正常終了") 'ガイダンス表示

    On Error GoTo 0

    Exit Function

DBColorSlbSearchRead_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBColorSlbSearchRead = False

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : TRTS0012読込処理
'
' 引き数    : ARG1 - スラブ番号
'           : ARG2 - 状態
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してTRTS0012のレコードを読込
'
' 備考      : スラブ肌実績入力データ読込
'           :COLORSYS
'
Public Function TRTS0012_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String) As Boolean
'slb_chno          VARCHAR2(5)          /* スラブチャージNO */
'slb_aino          VARCHAR2(4)          /* スラブ合番 */
'slb_stat          VARCHAR2(1)          /* 状態 */
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "TRTS0012_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        ReDim APResTmpData(1)
        APResTmpData(0).slb_no = strSlb_No
        APResTmpData(0).slb_chno = Mid(strSlb_No, 1, 5)
        APResTmpData(0).slb_aino = Mid(strSlb_No, 6)
        APResTmpData(0).slb_stat = strSlb_Stat
        TRTS0012_Read = True
        Exit Function
    End If

    On Error GoTo TRTS0012_Read_err

    nOpen = 0

    ' Oracleとの接続を確立する
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT * FROM TRTS0012 WHERE slb_no='" & strSlb_No & "' AND slb_stat='" & strSlb_Stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APResTmpData(0)
    If Not oDS.EOF Then
        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), conDefault_slb_no, oDS.Fields("slb_no").Value)
        APResTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), conDefault_slb_chno, oDS.Fields("slb_chno").Value)
        APResTmpData(0).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), conDefault_slb_aino, oDS.Fields("slb_aino").Value)
        APResTmpData(0).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), conDefault_slb_stat, oDS.Fields("slb_stat").Value)
        
        APResTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), conDefault_slb_ccno, oDS.Fields("slb_ccno").Value)
        APResTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), conDefault_slb_zkai_dte, oDS.Fields("slb_zkai_dte").Value)
        APResTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), conDefault_slb_ksh, oDS.Fields("slb_ksh").Value)
        APResTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), conDefault_slb_typ, oDS.Fields("slb_typ").Value)
        APResTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), conDefault_slb_uksk, oDS.Fields("slb_uksk").Value)
        APResTmpData(0).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), conDefault_slb_wei, oDS.Fields("slb_wei").Value)
        APResTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), conDefault_slb_lngth, oDS.Fields("slb_lngth").Value)
        APResTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), conDefault_slb_wdth, oDS.Fields("slb_wdth").Value)
        APResTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), conDefault_slb_thkns, oDS.Fields("slb_thkns").Value)
        APResTmpData(0).slb_nxt_prcs = IIf(IsNull(oDS.Fields("slb_nxt_prcs").Value), conDefault_slb_nxt_prcs, oDS.Fields("slb_nxt_prcs").Value)
        APResTmpData(0).slb_cmt1 = IIf(IsNull(oDS.Fields("slb_cmt1").Value), conDefault_slb_cmt1, oDS.Fields("slb_cmt1").Value)
        APResTmpData(0).slb_cmt2 = IIf(IsNull(oDS.Fields("slb_cmt2").Value), conDefault_slb_cmt2, oDS.Fields("slb_cmt2").Value)
        
        APResTmpData(0).slb_fault_cd_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s1").Value), conDefault_slb_fault_cd_e_s1, oDS.Fields("slb_fault_cd_e_s1").Value)
        APResTmpData(0).slb_fault_cd_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s2").Value), conDefault_slb_fault_cd_e_s2, oDS.Fields("slb_fault_cd_e_s2").Value)
        APResTmpData(0).slb_fault_cd_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s3").Value), conDefault_slb_fault_cd_e_s3, oDS.Fields("slb_fault_cd_e_s3").Value)
        APResTmpData(0).slb_fault_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_e_s1").Value), conDefault_slb_fault_e_s1, oDS.Fields("slb_fault_e_s1").Value)
        APResTmpData(0).slb_fault_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_e_s2").Value), conDefault_slb_fault_e_s2, oDS.Fields("slb_fault_e_s2").Value)
        APResTmpData(0).slb_fault_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_e_s3").Value), conDefault_slb_fault_e_s3, oDS.Fields("slb_fault_e_s3").Value)
        APResTmpData(0).slb_fault_e_n1 = IIf(IsNull(oDS.Fields("slb_fault_e_n1").Value), conDefault_slb_fault_e_n1, oDS.Fields("slb_fault_e_n1").Value)
        APResTmpData(0).slb_fault_e_n2 = IIf(IsNull(oDS.Fields("slb_fault_e_n2").Value), conDefault_slb_fault_e_n2, oDS.Fields("slb_fault_e_n2").Value)
        APResTmpData(0).slb_fault_e_n3 = IIf(IsNull(oDS.Fields("slb_fault_e_n3").Value), conDefault_slb_fault_e_n3, oDS.Fields("slb_fault_e_n3").Value)
        
        APResTmpData(0).slb_fault_cd_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s1").Value), conDefault_slb_fault_cd_w_s1, oDS.Fields("slb_fault_cd_w_s1").Value)
        APResTmpData(0).slb_fault_cd_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s2").Value), conDefault_slb_fault_cd_w_s2, oDS.Fields("slb_fault_cd_w_s2").Value)
        APResTmpData(0).slb_fault_cd_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s3").Value), conDefault_slb_fault_cd_w_s3, oDS.Fields("slb_fault_cd_w_s3").Value)
        APResTmpData(0).slb_fault_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_w_s1").Value), conDefault_slb_fault_w_s1, oDS.Fields("slb_fault_w_s1").Value)
        APResTmpData(0).slb_fault_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_w_s2").Value), conDefault_slb_fault_w_s2, oDS.Fields("slb_fault_w_s2").Value)
        APResTmpData(0).slb_fault_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_w_s3").Value), conDefault_slb_fault_w_s3, oDS.Fields("slb_fault_w_s3").Value)
        APResTmpData(0).slb_fault_w_n1 = IIf(IsNull(oDS.Fields("slb_fault_w_n1").Value), conDefault_slb_fault_w_n1, oDS.Fields("slb_fault_w_n1").Value)
        APResTmpData(0).slb_fault_w_n2 = IIf(IsNull(oDS.Fields("slb_fault_w_n2").Value), conDefault_slb_fault_w_n2, oDS.Fields("slb_fault_w_n2").Value)
        APResTmpData(0).slb_fault_w_n3 = IIf(IsNull(oDS.Fields("slb_fault_w_n3").Value), conDefault_slb_fault_w_n3, oDS.Fields("slb_fault_w_n3").Value)
        
        APResTmpData(0).slb_fault_cd_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s1").Value), conDefault_slb_fault_cd_s_s1, oDS.Fields("slb_fault_cd_s_s1").Value)
        APResTmpData(0).slb_fault_cd_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s2").Value), conDefault_slb_fault_cd_s_s2, oDS.Fields("slb_fault_cd_s_s2").Value)
        APResTmpData(0).slb_fault_cd_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s3").Value), conDefault_slb_fault_cd_s_s3, oDS.Fields("slb_fault_cd_s_s3").Value)
        APResTmpData(0).slb_fault_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_s_s1").Value), conDefault_slb_fault_s_s1, oDS.Fields("slb_fault_s_s1").Value)
        APResTmpData(0).slb_fault_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_s_s2").Value), conDefault_slb_fault_s_s2, oDS.Fields("slb_fault_s_s2").Value)
        APResTmpData(0).slb_fault_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_s_s3").Value), conDefault_slb_fault_s_s3, oDS.Fields("slb_fault_s_s3").Value)
        APResTmpData(0).slb_fault_s_n1 = IIf(IsNull(oDS.Fields("slb_fault_s_n1").Value), conDefault_slb_fault_s_n1, oDS.Fields("slb_fault_s_n1").Value)
        APResTmpData(0).slb_fault_s_n2 = IIf(IsNull(oDS.Fields("slb_fault_s_n2").Value), conDefault_slb_fault_s_n2, oDS.Fields("slb_fault_s_n2").Value)
        APResTmpData(0).slb_fault_s_n3 = IIf(IsNull(oDS.Fields("slb_fault_s_n3").Value), conDefault_slb_fault_s_n3, oDS.Fields("slb_fault_s_n3").Value)
        
        APResTmpData(0).slb_fault_cd_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s1").Value), conDefault_slb_fault_cd_n_s1, oDS.Fields("slb_fault_cd_n_s1").Value)
        APResTmpData(0).slb_fault_cd_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s2").Value), conDefault_slb_fault_cd_n_s2, oDS.Fields("slb_fault_cd_n_s2").Value)
        APResTmpData(0).slb_fault_cd_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s3").Value), conDefault_slb_fault_cd_n_s3, oDS.Fields("slb_fault_cd_n_s3").Value)
        APResTmpData(0).slb_fault_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_n_s1").Value), conDefault_slb_fault_n_s1, oDS.Fields("slb_fault_n_s1").Value)
        APResTmpData(0).slb_fault_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_n_s2").Value), conDefault_slb_fault_n_s2, oDS.Fields("slb_fault_n_s2").Value)
        APResTmpData(0).slb_fault_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_n_s3").Value), conDefault_slb_fault_n_s3, oDS.Fields("slb_fault_n_s3").Value)
        APResTmpData(0).slb_fault_n_n1 = IIf(IsNull(oDS.Fields("slb_fault_n_n1").Value), conDefault_slb_fault_n_n1, oDS.Fields("slb_fault_n_n1").Value)
        APResTmpData(0).slb_fault_n_n2 = IIf(IsNull(oDS.Fields("slb_fault_n_n2").Value), conDefault_slb_fault_n_n2, oDS.Fields("slb_fault_n_n2").Value)
        APResTmpData(0).slb_fault_n_n3 = IIf(IsNull(oDS.Fields("slb_fault_n_n3").Value), conDefault_slb_fault_n_n3, oDS.Fields("slb_fault_n_n3").Value)
        
        APResTmpData(0).slb_fault_cd_bs_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bs_s").Value), conDefault_slb_fault_cd_bs_s, oDS.Fields("slb_fault_cd_bs_s").Value)
        APResTmpData(0).slb_fault_cd_bm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bm_s").Value), conDefault_slb_fault_cd_bm_s, oDS.Fields("slb_fault_cd_bm_s").Value)
        APResTmpData(0).slb_fault_cd_bn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bn_s").Value), conDefault_slb_fault_cd_bn_s, oDS.Fields("slb_fault_cd_bn_s").Value)
        APResTmpData(0).slb_fault_bs_s = IIf(IsNull(oDS.Fields("slb_fault_bs_s").Value), conDefault_slb_fault_bs_s, oDS.Fields("slb_fault_bs_s").Value)
        APResTmpData(0).slb_fault_bm_s = IIf(IsNull(oDS.Fields("slb_fault_bm_s").Value), conDefault_slb_fault_bm_s, oDS.Fields("slb_fault_bm_s").Value)
        APResTmpData(0).slb_fault_bn_s = IIf(IsNull(oDS.Fields("slb_fault_bn_s").Value), conDefault_slb_fault_bn_s, oDS.Fields("slb_fault_bn_s").Value)
        APResTmpData(0).slb_fault_bs_n = IIf(IsNull(oDS.Fields("slb_fault_bs_n").Value), conDefault_slb_fault_bs_n, oDS.Fields("slb_fault_bs_n").Value)
        APResTmpData(0).slb_fault_bm_n = IIf(IsNull(oDS.Fields("slb_fault_bm_n").Value), conDefault_slb_fault_bm_n, oDS.Fields("slb_fault_bm_n").Value)
        APResTmpData(0).slb_fault_bn_n = IIf(IsNull(oDS.Fields("slb_fault_bn_n").Value), conDefault_slb_fault_bn_n, oDS.Fields("slb_fault_bn_n").Value)
        
        APResTmpData(0).slb_fault_cd_ts_s = IIf(IsNull(oDS.Fields("slb_fault_cd_ts_s").Value), conDefault_slb_fault_cd_ts_s, oDS.Fields("slb_fault_cd_ts_s").Value)
        APResTmpData(0).slb_fault_cd_tm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tm_s").Value), conDefault_slb_fault_cd_tm_s, oDS.Fields("slb_fault_cd_tm_s").Value)
        APResTmpData(0).slb_fault_cd_tn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tn_s").Value), conDefault_slb_fault_cd_tn_s, oDS.Fields("slb_fault_cd_tn_s").Value)
        APResTmpData(0).slb_fault_ts_s = IIf(IsNull(oDS.Fields("slb_fault_ts_s").Value), conDefault_slb_fault_ts_s, oDS.Fields("slb_fault_ts_s").Value)
        APResTmpData(0).slb_fault_tm_s = IIf(IsNull(oDS.Fields("slb_fault_tm_s").Value), conDefault_slb_fault_tm_s, oDS.Fields("slb_fault_tm_s").Value)
        APResTmpData(0).slb_fault_tn_s = IIf(IsNull(oDS.Fields("slb_fault_tn_s").Value), conDefault_slb_fault_tn_s, oDS.Fields("slb_fault_tn_s").Value)
        APResTmpData(0).slb_fault_ts_n = IIf(IsNull(oDS.Fields("slb_fault_ts_n").Value), conDefault_slb_fault_ts_n, oDS.Fields("slb_fault_ts_n").Value)
        APResTmpData(0).slb_fault_tm_n = IIf(IsNull(oDS.Fields("slb_fault_tm_n").Value), conDefault_slb_fault_tm_n, oDS.Fields("slb_fault_tm_n").Value)
        APResTmpData(0).slb_fault_tn_n = IIf(IsNull(oDS.Fields("slb_fault_tn_n").Value), conDefault_slb_fault_tn_n, oDS.Fields("slb_fault_tn_n").Value)
        
        APResTmpData(0).slb_wrt_nme = IIf(IsNull(oDS.Fields("slb_wrt_nme").Value), conDefault_slb_wrt_nme, oDS.Fields("slb_wrt_nme").Value)
        APResTmpData(0).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), conDefault_sys_wrt_dte, oDS.Fields("sys_wrt_dte").Value)
        APResTmpData(0).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), conDefault_sys_wrt_tme, oDS.Fields("sys_wrt_tme").Value)
'        APResTmpData(0).sys_rwrt_dte = IIf(IsNull(oDS.Fields("sys_rwrt_dte").Value), conDefault_sys_rwrt_dte, oDS.Fields("sys_rwrt_dte").Value)
'        APResTmpData(0).sys_rwrt_tme = IIf(IsNull(oDS.Fields("sys_rwrt_tme").Value), conDefault_sys_rwrt_tme, oDS.Fields("sys_rwrt_tme").Value)
'        APResTmpData(0).sys_acs_pros = IIf(IsNull(oDS.Fields("sys_acs_pros").Value), conDefault_sys_acs_pros, oDS.Fields("sys_acs_pros").Value)
'        APResTmpData(0).sys_acs_enum = IIf(IsNull(oDS.Fields("sys_acs_enum").Value), conDefault_sys_acs_enum, oDS.Fields("sys_acs_enum").Value)
       
    End If

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Read 正常終了") 'ガイダンス表示

    TRTS0012_Read = True

    On Error GoTo 0
    Exit Function

TRTS0012_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0012_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0012書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 実績入力のカレントデータを書込
'
' 備考      : 実績入力データ書き込み
'
Public Function TRTS0012_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0012_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0012_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0012_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0012 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0012 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_ccno,"       ''/* スラブCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* 造塊日 */
        strSQL = strSQL & "slb_ksh,"        ''/* 鋼種 */
        strSQL = strSQL & "slb_typ,"        ''/* 型 */
        strSQL = strSQL & "slb_uksk,"       ''/* 向先 */
        strSQL = strSQL & "slb_wei,"        ''/* 重量 */
        strSQL = strSQL & "slb_lngth,"      ''/* 長さ */
        strSQL = strSQL & "slb_wdth,"       ''/* 幅 */
        strSQL = strSQL & "slb_thkns,"      ''/* 厚み */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* 次工程 */
        strSQL = strSQL & "slb_cmt1,"       ''/* コメント1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* コメント2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* 欠陥E面CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* 欠陥E面CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* 欠陥E面CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* 欠陥E面種類1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* 欠陥E面種類2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* 欠陥E面種類3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* 欠陥E面個数1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* 欠陥E面個数2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* 欠陥W面CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* 欠陥W面CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* 欠陥W面CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* 欠陥W面種類1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* 欠陥W面種類2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* 欠陥W面種類3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* 欠陥W面個数1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* 欠陥W面個数2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* 欠陥S面CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* 欠陥S面CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* 欠陥S面CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* 欠陥S面種類1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* 欠陥S面種類2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* 欠陥S面種類3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* 欠陥S面個数1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* 欠陥S面個数2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* 欠陥N面CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* 欠陥N面CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* 欠陥N面CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* 欠陥N面種類1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* 欠陥N面種類2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* 欠陥N面種類3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* 欠陥N面個数1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* 欠陥N面個数2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* 欠陥N面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* 内部割れBSCD */
        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* 内部割れBMCD */
        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* 内部割れBNCD */
        strSQL = strSQL & "slb_fault_bs_s,"     ''/* 内部割れBS種類 */
        strSQL = strSQL & "slb_fault_bm_s,"     ''/* 内部割れBM種類 */
        strSQL = strSQL & "slb_fault_bn_s,"     ''/* 内部割れBN種類 */
        strSQL = strSQL & "slb_fault_bs_n,"     ''/* 内部割れBS個数 */
        strSQL = strSQL & "slb_fault_bm_n,"     ''/* 内部割れBM個数 */
        strSQL = strSQL & "slb_fault_bn_n,"     ''/* 内部割れBN個数 */
        
        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* 内部割れTSCD */
        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* 内部割れTMCD */
        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* 内部割れTNCD */
        strSQL = strSQL & "slb_fault_ts_s,"     ''/* 内部割れTS種類 */
        strSQL = strSQL & "slb_fault_tm_s,"     ''/* 内部割れTM種類 */
        strSQL = strSQL & "slb_fault_tn_s,"     ''/* 内部割れTN種類 */
        strSQL = strSQL & "slb_fault_ts_n,"     ''/* 内部割れTS個数 */
        strSQL = strSQL & "slb_fault_tm_n,"     ''/* 内部割れTM個数 */
        strSQL = strSQL & "slb_fault_tn_n,"     ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* 欠陥E面判定 */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* 欠陥W面判定 */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* 欠陥S面判定 */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* 欠陥N面判定 */
        strSQL = strSQL & "slb_fault_b_judg,"       ''/* 欠陥B面判定 */
        strSQL = strSQL & "slb_fault_t_judg,"       ''/* 欠陥T面判定 */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* スタッフ名 */
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* スラブCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* 造塊日 */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* 鋼種 */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* 型 */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* 向先 */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* 重量 */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* 長さ */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* 幅 */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* 厚み */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* 次工程 */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* コメント1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* コメント2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* 欠陥E面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* 欠陥E面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* 欠陥E面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* 欠陥E面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* 欠陥E面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* 欠陥E面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* 欠陥E面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* 欠陥E面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* 欠陥W面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* 欠陥W面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* 欠陥W面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* 欠陥W面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* 欠陥W面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* 欠陥W面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* 欠陥W面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* 欠陥W面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* 欠陥S面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* 欠陥S面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* 欠陥S面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* 欠陥S面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* 欠陥S面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* 欠陥S面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* 欠陥S面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* 欠陥S面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* 欠陥N面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* 欠陥N面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* 欠陥N面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* 欠陥N面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* 欠陥N面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* 欠陥N面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* 欠陥N面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* 欠陥N面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* 欠陥N面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* 内部割れBSCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* 内部割れBMCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* 内部割れBNCD */
        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* 内部割れBS種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* 内部割れBM種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* 内部割れBN種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* 内部割れBS個数 */
        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* 内部割れBM個数 */
        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* 内部割れBN個数 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* 内部割れTSCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* 内部割れTMCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* 内部割れTNCD */
        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* 内部割れTS種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* 内部割れTM種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* 内部割れTN種類 */
        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* 内部割れTS個数 */
        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* 内部割れTM個数 */
        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* 欠陥E面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* 欠陥W面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* 欠陥S面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* 欠陥N面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* 欠陥B面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* 欠陥T面判定 */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* スタッフ名 */
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* 登録日 */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Write 正常終了") 'ガイダンス表示

    TRTS0012_Write = True

    On Error GoTo 0
    Exit Function

TRTS0012_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0012_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0014読込処理
'
' 引き数    : ARG1 - スラブ番号
'           : ARG2 - 状態
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してTRTS0014のレコードを読込
'
' 備考      : カラーチェック実績入力データ読込
'           :COLORSYS
'
Public Function TRTS0014_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String, ByVal strSlb_Col_Cnt As String) As Boolean
'slb_chno          VARCHAR2(5)          /* スラブチャージNO */
'slb_aino          VARCHAR2(4)          /* スラブ合番 */
'slb_stat          VARCHAR2(1)          /* 状態 */
'slb_col_cnt       VARCHAR2(2)          /* ｶﾗｰ回数 */
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0014_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        ReDim APResTmpData(0)
        
        Call TRTS0014_ReadCSV
'        APResTmpData(0).slb_no = strSlb_No
'        APResTmpData(0).slb_stat = strSlb_Stat
'        APResTmpData(0).slb_col_cnt = Format(CInt(strSlb_Col_Cnt), "00")
        TRTS0014_Read = True
        Exit Function
    End If

    On Error GoTo TRTS0014_Read_err

    nOpen = 0

    ' Oracleとの接続を確立する
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT * FROM TRTS0014 WHERE slb_no='" & strSlb_No & "' AND slb_stat='" & strSlb_Stat & "' AND slb_col_cnt='" & Format(CInt(strSlb_Col_Cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APResTmpData(0)
    If Not oDS.EOF Then
        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), conDefault_slb_no, oDS.Fields("slb_no").Value)
        APResTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), conDefault_slb_chno, oDS.Fields("slb_chno").Value)
        APResTmpData(0).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), conDefault_slb_aino, oDS.Fields("slb_aino").Value)
        APResTmpData(0).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), conDefault_slb_stat, oDS.Fields("slb_stat").Value)
        
        APResTmpData(0).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), conDefault_slb_col_cnt, oDS.Fields("slb_col_cnt").Value)
        
        APResTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), conDefault_slb_ccno, oDS.Fields("slb_ccno").Value)
        APResTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), conDefault_slb_zkai_dte, oDS.Fields("slb_zkai_dte").Value)
        APResTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), conDefault_slb_ksh, oDS.Fields("slb_ksh").Value)
        APResTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), conDefault_slb_typ, oDS.Fields("slb_typ").Value)
        APResTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), conDefault_slb_uksk, oDS.Fields("slb_uksk").Value)
        APResTmpData(0).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), conDefault_slb_wei, oDS.Fields("slb_wei").Value)
        APResTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), conDefault_slb_lngth, oDS.Fields("slb_lngth").Value)
        APResTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), conDefault_slb_wdth, oDS.Fields("slb_wdth").Value)
        APResTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), conDefault_slb_thkns, oDS.Fields("slb_thkns").Value)
        APResTmpData(0).slb_nxt_prcs = IIf(IsNull(oDS.Fields("slb_nxt_prcs").Value), conDefault_slb_nxt_prcs, oDS.Fields("slb_nxt_prcs").Value)
        APResTmpData(0).slb_cmt1 = IIf(IsNull(oDS.Fields("slb_cmt1").Value), conDefault_slb_cmt1, oDS.Fields("slb_cmt1").Value)
        APResTmpData(0).slb_cmt2 = IIf(IsNull(oDS.Fields("slb_cmt2").Value), conDefault_slb_cmt2, oDS.Fields("slb_cmt2").Value)
        
        APResTmpData(0).slb_fault_cd_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s1").Value), conDefault_slb_fault_cd_e_s1, oDS.Fields("slb_fault_cd_e_s1").Value)
        APResTmpData(0).slb_fault_cd_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s2").Value), conDefault_slb_fault_cd_e_s2, oDS.Fields("slb_fault_cd_e_s2").Value)
        APResTmpData(0).slb_fault_cd_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s3").Value), conDefault_slb_fault_cd_e_s3, oDS.Fields("slb_fault_cd_e_s3").Value)
        APResTmpData(0).slb_fault_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_e_s1").Value), conDefault_slb_fault_e_s1, oDS.Fields("slb_fault_e_s1").Value)
        APResTmpData(0).slb_fault_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_e_s2").Value), conDefault_slb_fault_e_s2, oDS.Fields("slb_fault_e_s2").Value)
        APResTmpData(0).slb_fault_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_e_s3").Value), conDefault_slb_fault_e_s3, oDS.Fields("slb_fault_e_s3").Value)
        APResTmpData(0).slb_fault_e_n1 = IIf(IsNull(oDS.Fields("slb_fault_e_n1").Value), conDefault_slb_fault_e_n1, oDS.Fields("slb_fault_e_n1").Value)
        APResTmpData(0).slb_fault_e_n2 = IIf(IsNull(oDS.Fields("slb_fault_e_n2").Value), conDefault_slb_fault_e_n2, oDS.Fields("slb_fault_e_n2").Value)
        APResTmpData(0).slb_fault_e_n3 = IIf(IsNull(oDS.Fields("slb_fault_e_n3").Value), conDefault_slb_fault_e_n3, oDS.Fields("slb_fault_e_n3").Value)
        
        APResTmpData(0).slb_fault_cd_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s1").Value), conDefault_slb_fault_cd_w_s1, oDS.Fields("slb_fault_cd_w_s1").Value)
        APResTmpData(0).slb_fault_cd_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s2").Value), conDefault_slb_fault_cd_w_s2, oDS.Fields("slb_fault_cd_w_s2").Value)
        APResTmpData(0).slb_fault_cd_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s3").Value), conDefault_slb_fault_cd_w_s3, oDS.Fields("slb_fault_cd_w_s3").Value)
        APResTmpData(0).slb_fault_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_w_s1").Value), conDefault_slb_fault_w_s1, oDS.Fields("slb_fault_w_s1").Value)
        APResTmpData(0).slb_fault_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_w_s2").Value), conDefault_slb_fault_w_s2, oDS.Fields("slb_fault_w_s2").Value)
        APResTmpData(0).slb_fault_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_w_s3").Value), conDefault_slb_fault_w_s3, oDS.Fields("slb_fault_w_s3").Value)
        APResTmpData(0).slb_fault_w_n1 = IIf(IsNull(oDS.Fields("slb_fault_w_n1").Value), conDefault_slb_fault_w_n1, oDS.Fields("slb_fault_w_n1").Value)
        APResTmpData(0).slb_fault_w_n2 = IIf(IsNull(oDS.Fields("slb_fault_w_n2").Value), conDefault_slb_fault_w_n2, oDS.Fields("slb_fault_w_n2").Value)
        APResTmpData(0).slb_fault_w_n3 = IIf(IsNull(oDS.Fields("slb_fault_w_n3").Value), conDefault_slb_fault_w_n3, oDS.Fields("slb_fault_w_n3").Value)
        
        APResTmpData(0).slb_fault_cd_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s1").Value), conDefault_slb_fault_cd_s_s1, oDS.Fields("slb_fault_cd_s_s1").Value)
        APResTmpData(0).slb_fault_cd_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s2").Value), conDefault_slb_fault_cd_s_s2, oDS.Fields("slb_fault_cd_s_s2").Value)
        APResTmpData(0).slb_fault_cd_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s3").Value), conDefault_slb_fault_cd_s_s3, oDS.Fields("slb_fault_cd_s_s3").Value)
        APResTmpData(0).slb_fault_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_s_s1").Value), conDefault_slb_fault_s_s1, oDS.Fields("slb_fault_s_s1").Value)
        APResTmpData(0).slb_fault_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_s_s2").Value), conDefault_slb_fault_s_s2, oDS.Fields("slb_fault_s_s2").Value)
        APResTmpData(0).slb_fault_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_s_s3").Value), conDefault_slb_fault_s_s3, oDS.Fields("slb_fault_s_s3").Value)
        APResTmpData(0).slb_fault_s_n1 = IIf(IsNull(oDS.Fields("slb_fault_s_n1").Value), conDefault_slb_fault_s_n1, oDS.Fields("slb_fault_s_n1").Value)
        APResTmpData(0).slb_fault_s_n2 = IIf(IsNull(oDS.Fields("slb_fault_s_n2").Value), conDefault_slb_fault_s_n2, oDS.Fields("slb_fault_s_n2").Value)
        APResTmpData(0).slb_fault_s_n3 = IIf(IsNull(oDS.Fields("slb_fault_s_n3").Value), conDefault_slb_fault_s_n3, oDS.Fields("slb_fault_s_n3").Value)
        
        APResTmpData(0).slb_fault_cd_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s1").Value), conDefault_slb_fault_cd_n_s1, oDS.Fields("slb_fault_cd_n_s1").Value)
        APResTmpData(0).slb_fault_cd_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s2").Value), conDefault_slb_fault_cd_n_s2, oDS.Fields("slb_fault_cd_n_s2").Value)
        APResTmpData(0).slb_fault_cd_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s3").Value), conDefault_slb_fault_cd_n_s3, oDS.Fields("slb_fault_cd_n_s3").Value)
        APResTmpData(0).slb_fault_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_n_s1").Value), conDefault_slb_fault_n_s1, oDS.Fields("slb_fault_n_s1").Value)
        APResTmpData(0).slb_fault_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_n_s2").Value), conDefault_slb_fault_n_s2, oDS.Fields("slb_fault_n_s2").Value)
        APResTmpData(0).slb_fault_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_n_s3").Value), conDefault_slb_fault_n_s3, oDS.Fields("slb_fault_n_s3").Value)
        APResTmpData(0).slb_fault_n_n1 = IIf(IsNull(oDS.Fields("slb_fault_n_n1").Value), conDefault_slb_fault_n_n1, oDS.Fields("slb_fault_n_n1").Value)
        APResTmpData(0).slb_fault_n_n2 = IIf(IsNull(oDS.Fields("slb_fault_n_n2").Value), conDefault_slb_fault_n_n2, oDS.Fields("slb_fault_n_n2").Value)
        APResTmpData(0).slb_fault_n_n3 = IIf(IsNull(oDS.Fields("slb_fault_n_n3").Value), conDefault_slb_fault_n_n3, oDS.Fields("slb_fault_n_n3").Value)
        
'        APResTmpData(0).slb_fault_cd_bs_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bs_s").Value), conDefault_slb_fault_cd_bs_s, oDS.Fields("slb_fault_cd_bs_s").Value)
'        APResTmpData(0).slb_fault_cd_bm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bm_s").Value), conDefault_slb_fault_cd_bm_s, oDS.Fields("slb_fault_cd_bm_s").Value)
'        APResTmpData(0).slb_fault_cd_bn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bn_s").Value), conDefault_slb_fault_cd_bn_s, oDS.Fields("slb_fault_cd_bn_s").Value)
'        APResTmpData(0).slb_fault_bs_s = IIf(IsNull(oDS.Fields("slb_fault_bs_s").Value), conDefault_slb_fault_bs_s, oDS.Fields("slb_fault_bs_s").Value)
'        APResTmpData(0).slb_fault_bm_s = IIf(IsNull(oDS.Fields("slb_fault_bm_s").Value), conDefault_slb_fault_bm_s, oDS.Fields("slb_fault_bm_s").Value)
'        APResTmpData(0).slb_fault_bn_s = IIf(IsNull(oDS.Fields("slb_fault_bn_s").Value), conDefault_slb_fault_bn_s, oDS.Fields("slb_fault_bn_s").Value)
'        APResTmpData(0).slb_fault_bs_n = IIf(IsNull(oDS.Fields("slb_fault_bs_n").Value), conDefault_slb_fault_bs_n, oDS.Fields("slb_fault_bs_n").Value)
'        APResTmpData(0).slb_fault_bm_n = IIf(IsNull(oDS.Fields("slb_fault_bm_n").Value), conDefault_slb_fault_bm_n, oDS.Fields("slb_fault_bm_n").Value)
'        APResTmpData(0).slb_fault_bn_n = IIf(IsNull(oDS.Fields("slb_fault_bn_n").Value), conDefault_slb_fault_bn_n, oDS.Fields("slb_fault_bn_n").Value)
'
'        APResTmpData(0).slb_fault_cd_ts_s = IIf(IsNull(oDS.Fields("slb_fault_cd_ts_s").Value), conDefault_slb_fault_cd_ts_s, oDS.Fields("slb_fault_cd_ts_s").Value)
'        APResTmpData(0).slb_fault_cd_tm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tm_s").Value), conDefault_slb_fault_cd_tm_s, oDS.Fields("slb_fault_cd_tm_s").Value)
'        APResTmpData(0).slb_fault_cd_tn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tn_s").Value), conDefault_slb_fault_cd_tn_s, oDS.Fields("slb_fault_cd_tn_s").Value)
'        APResTmpData(0).slb_fault_ts_s = IIf(IsNull(oDS.Fields("slb_fault_ts_s").Value), conDefault_slb_fault_ts_s, oDS.Fields("slb_fault_ts_s").Value)
'        APResTmpData(0).slb_fault_tm_s = IIf(IsNull(oDS.Fields("slb_fault_tm_s").Value), conDefault_slb_fault_tm_s, oDS.Fields("slb_fault_tm_s").Value)
'        APResTmpData(0).slb_fault_tn_s = IIf(IsNull(oDS.Fields("slb_fault_tn_s").Value), conDefault_slb_fault_tn_s, oDS.Fields("slb_fault_tn_s").Value)
'        APResTmpData(0).slb_fault_ts_n = IIf(IsNull(oDS.Fields("slb_fault_ts_n").Value), conDefault_slb_fault_ts_n, oDS.Fields("slb_fault_ts_n").Value)
'        APResTmpData(0).slb_fault_tm_n = IIf(IsNull(oDS.Fields("slb_fault_tm_n").Value), conDefault_slb_fault_tm_n, oDS.Fields("slb_fault_tm_n").Value)
'        APResTmpData(0).slb_fault_tn_n = IIf(IsNull(oDS.Fields("slb_fault_tn_n").Value), conDefault_slb_fault_tn_n, oDS.Fields("slb_fault_tn_n").Value)
        
        APResTmpData(0).slb_wrt_nme = IIf(IsNull(oDS.Fields("slb_wrt_nme").Value), conDefault_slb_wrt_nme, oDS.Fields("slb_wrt_nme").Value)
        
        APResTmpData(0).host_send = IIf(IsNull(oDS.Fields("host_send").Value), conDefault_host_send, oDS.Fields("host_send").Value)
        APResTmpData(0).host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte").Value), conDefault_host_wrt_dte, oDS.Fields("host_wrt_dte").Value)
        APResTmpData(0).host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme").Value), conDefault_host_wrt_tme, oDS.Fields("host_wrt_tme").Value)
        
        APResTmpData(0).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), conDefault_sys_wrt_dte, oDS.Fields("sys_wrt_dte").Value)
        APResTmpData(0).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), conDefault_sys_wrt_tme, oDS.Fields("sys_wrt_tme").Value)
'        APResTmpData(0).sys_rwrt_dte = IIf(IsNull(oDS.Fields("sys_rwrt_dte").Value), conDefault_sys_rwrt_dte, oDS.Fields("sys_rwrt_dte").Value)
'        APResTmpData(0).sys_rwrt_tme = IIf(IsNull(oDS.Fields("sys_rwrt_tme").Value), conDefault_sys_rwrt_tme, oDS.Fields("sys_rwrt_tme").Value)
'        APResTmpData(0).sys_acs_pros = IIf(IsNull(oDS.Fields("sys_acs_pros").Value), conDefault_sys_acs_pros, oDS.Fields("sys_acs_pros").Value)
'        APResTmpData(0).sys_acs_enum = IIf(IsNull(oDS.Fields("sys_acs_enum").Value), conDefault_sys_acs_enum, oDS.Fields("sys_acs_enum").Value)
       
    End If

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Read 正常終了") 'ガイダンス表示

    TRTS0014_Read = True

    On Error GoTo 0
    Exit Function

TRTS0014_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0014_Read = False

    On Error GoTo 0
End Function

Private Sub TRTS0014_ReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APResTmpData(0)

    bRet = ReadCSV(App.path & "\" & "TRTS0014_Read.csv", strItem(), strData())

    For nI = 0 To UBound(strData, 2) - 1

        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APResTmpData(0).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APResTmpData(0).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_col_cnt = getItemDataCSV("slb_col_cnt", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_ccno = getItemDataCSV("slb_ccno", nI + 1, strItem(), strData())
        APResTmpData(0).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())
        APResTmpData(0).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())
        APResTmpData(0).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())
        APResTmpData(0).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        APResTmpData(0).slb_wei = getItemDataCSV("slb_wei", nI + 1, strItem(), strData())
        APResTmpData(0).slb_lngth = getItemDataCSV("slb_lngth", nI + 1, strItem(), strData())
        APResTmpData(0).slb_wdth = getItemDataCSV("slb_wdth", nI + 1, strItem(), strData())
        APResTmpData(0).slb_thkns = getItemDataCSV("slb_thkns", nI + 1, strItem(), strData())
        APResTmpData(0).slb_nxt_prcs = getItemDataCSV("slb_nxt_prcs", nI + 1, strItem(), strData())
        APResTmpData(0).slb_cmt1 = getItemDataCSV("slb_cmt1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_cmt2 = getItemDataCSV("slb_cmt2", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_e_s1 = getItemDataCSV("slb_fault_cd_e_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_e_s2 = getItemDataCSV("slb_fault_cd_e_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_e_s3 = getItemDataCSV("slb_fault_cd_e_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s1 = getItemDataCSV("slb_fault_e_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s2 = getItemDataCSV("slb_fault_e_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s3 = getItemDataCSV("slb_fault_e_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n1 = getItemDataCSV("slb_fault_e_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n2 = getItemDataCSV("slb_fault_e_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n3 = getItemDataCSV("slb_fault_e_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_w_s1 = getItemDataCSV("slb_fault_cd_w_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_w_s2 = getItemDataCSV("slb_fault_cd_w_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_w_s3 = getItemDataCSV("slb_fault_cd_w_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s1 = getItemDataCSV("slb_fault_w_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s2 = getItemDataCSV("slb_fault_w_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s3 = getItemDataCSV("slb_fault_w_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n1 = getItemDataCSV("slb_fault_w_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n2 = getItemDataCSV("slb_fault_w_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n3 = getItemDataCSV("slb_fault_w_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_s_s1 = getItemDataCSV("slb_fault_cd_s_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_s_s2 = getItemDataCSV("slb_fault_cd_s_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_s_s3 = getItemDataCSV("slb_fault_cd_s_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s1 = getItemDataCSV("slb_fault_s_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s2 = getItemDataCSV("slb_fault_s_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s3 = getItemDataCSV("slb_fault_s_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n1 = getItemDataCSV("slb_fault_s_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n2 = getItemDataCSV("slb_fault_s_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n3 = getItemDataCSV("slb_fault_s_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_n_s1 = getItemDataCSV("slb_fault_cd_n_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_n_s2 = getItemDataCSV("slb_fault_cd_n_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_n_s3 = getItemDataCSV("slb_fault_cd_n_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s1 = getItemDataCSV("slb_fault_n_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s2 = getItemDataCSV("slb_fault_n_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s3 = getItemDataCSV("slb_fault_n_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n1 = getItemDataCSV("slb_fault_n_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n2 = getItemDataCSV("slb_fault_n_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n3 = getItemDataCSV("slb_fault_n_n3", nI + 1, strItem(), strData())
        
'        APResTmpData(0).slb_fault_cd_bs_s = getItemDataCSV("slb_fault_cd_bs_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_bm_s = getItemDataCSV("slb_fault_cd_bm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_bn_s = getItemDataCSV("slb_fault_cd_bn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bs_s = getItemDataCSV("slb_fault_bs_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bm_s = getItemDataCSV("slb_fault_bm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bn_s = getItemDataCSV("slb_fault_bn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bs_n = getItemDataCSV("slb_fault_bs_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bm_n = getItemDataCSV("slb_fault_bm_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bn_n = getItemDataCSV("slb_fault_bn_n", nI + 1, strItem(), strData())
'
'        APResTmpData(0).slb_fault_cd_ts_s = getItemDataCSV("slb_fault_cd_ts_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_tm_s = getItemDataCSV("slb_fault_cd_tm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_tn_s = getItemDataCSV("slb_fault_cd_tn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_ts_s = getItemDataCSV("slb_fault_ts_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tm_s = getItemDataCSV("slb_fault_tm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tn_s = getItemDataCSV("slb_fault_tn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_ts_n = getItemDataCSV("slb_fault_ts_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tm_n = getItemDataCSV("slb_fault_tm_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tn_n = getItemDataCSV("slb_fault_tn_n", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_wrt_nme = getItemDataCSV("slb_wrt_nme", nI + 1, strItem(), strData())
        
        APResTmpData(0).host_send = getItemDataCSV("host_send", nI + 1, strItem(), strData())
        APResTmpData(0).host_wrt_dte = getItemDataCSV("host_wrt_dte", nI + 1, strItem(), strData())
        APResTmpData(0).host_wrt_tme = getItemDataCSV("host_wrt_tme", nI + 1, strItem(), strData())
        
        APResTmpData(0).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())
        APResTmpData(0).sys_wrt_tme = getItemDataCSV("sys_wrt_tme", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_rwrt_dte = getItemDataCSV("sys_rwrt_dte", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_rwrt_tme = getItemDataCSV("sys_rwrt_tme", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_acs_pros = getItemDataCSV("sys_acs_pros", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_acs_enum = getItemDataCSV("sys_acs_enum", nI + 1, strItem(), strData())
        
        Exit For '１件のみ有効
        

    Next nI

End Sub





' @(f)
'
' 機能      : TRTS0014書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 実績入力のカレントデータを書込
'
' 備考      : 実績入力データ書き込み
'
Public Function TRTS0014_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0014_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0014_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0014_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0014 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0014 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        
        strSQL = strSQL & "slb_col_cnt,"       ''/* ｶﾗｰ回数 */
        
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_ccno,"       ''/* スラブCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* 造塊日 */
        strSQL = strSQL & "slb_ksh,"        ''/* 鋼種 */
        strSQL = strSQL & "slb_typ,"        ''/* 型 */
        strSQL = strSQL & "slb_uksk,"       ''/* 向先 */
        strSQL = strSQL & "slb_wei,"        ''/* 重量 */
        strSQL = strSQL & "slb_lngth,"      ''/* 長さ */
        strSQL = strSQL & "slb_wdth,"       ''/* 幅 */
        strSQL = strSQL & "slb_thkns,"      ''/* 厚み */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* 次工程 */
        strSQL = strSQL & "slb_cmt1,"       ''/* コメント1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* コメント2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* 欠陥E面CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* 欠陥E面CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* 欠陥E面CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* 欠陥E面種類1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* 欠陥E面種類2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* 欠陥E面種類3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* 欠陥E面個数1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* 欠陥E面個数2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* 欠陥W面CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* 欠陥W面CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* 欠陥W面CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* 欠陥W面種類1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* 欠陥W面種類2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* 欠陥W面種類3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* 欠陥W面個数1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* 欠陥W面個数2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* 欠陥S面CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* 欠陥S面CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* 欠陥S面CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* 欠陥S面種類1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* 欠陥S面種類2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* 欠陥S面種類3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* 欠陥S面個数1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* 欠陥S面個数2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* 欠陥N面CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* 欠陥N面CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* 欠陥N面CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* 欠陥N面種類1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* 欠陥N面種類2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* 欠陥N面種類3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* 欠陥N面個数1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* 欠陥N面個数2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* 欠陥N面個数3 */
        
'        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* 内部割れBSCD */
'        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* 内部割れBMCD */
'        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* 内部割れBNCD */
'        strSQL = strSQL & "slb_fault_bs_s,"     ''/* 内部割れBS種類 */
'        strSQL = strSQL & "slb_fault_bm_s,"     ''/* 内部割れBM種類 */
'        strSQL = strSQL & "slb_fault_bn_s,"     ''/* 内部割れBN種類 */
'        strSQL = strSQL & "slb_fault_bs_n,"     ''/* 内部割れBS個数 */
'        strSQL = strSQL & "slb_fault_bm_n,"     ''/* 内部割れBM個数 */
'        strSQL = strSQL & "slb_fault_bn_n,"     ''/* 内部割れBN個数 */
'
'        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* 内部割れTSCD */
'        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* 内部割れTMCD */
'        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* 内部割れTNCD */
'        strSQL = strSQL & "slb_fault_ts_s,"     ''/* 内部割れTS種類 */
'        strSQL = strSQL & "slb_fault_tm_s,"     ''/* 内部割れTM種類 */
'        strSQL = strSQL & "slb_fault_tn_s,"     ''/* 内部割れTN種類 */
'        strSQL = strSQL & "slb_fault_ts_n,"     ''/* 内部割れTS個数 */
'        strSQL = strSQL & "slb_fault_tm_n,"     ''/* 内部割れTM個数 */
'        strSQL = strSQL & "slb_fault_tn_n,"     ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* 欠陥E面判定 */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* 欠陥W面判定 */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* 欠陥S面判定 */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* 欠陥N面判定 */
'        strSQL = strSQL & "slb_fault_b_judg,"       ''/* 欠陥B面判定 */
'        strSQL = strSQL & "slb_fault_t_judg,"       ''/* 欠陥T面判定 */
        strSQL = strSQL & "slb_fault_u_judg,"       ''/* 欠陥U面判定 */
        strSQL = strSQL & "slb_fault_d_judg,"       ''/* 欠陥D面判定 */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* 検査員名 */
        
        '2016/04/20 - TAI - S
        strSQL = strSQL & "slb_fault_total_judg,"     ''/* 検査結果 */
        strSQL = strSQL & "slb_works_sky_tok,"        ''/* 作業場所 */
        '2016/04/20 - TAI - E
        
        strSQL = strSQL & "host_send,"          ''/* ビジコン送信結果 */
        strSQL = strSQL & "host_wrt_dte,"       ''/* ビジコン登録日 */
        strSQL = strSQL & "host_wrt_tme,"       ''/* ビジコン登録時刻 */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","     ''/* ｶﾗｰ回数 */
        
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* スラブCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* 造塊日 */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* 鋼種 */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* 型 */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* 向先 */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* 重量 */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* 長さ */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* 幅 */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* 厚み */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* 次工程 */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* コメント1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* コメント2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* 欠陥E面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* 欠陥E面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* 欠陥E面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* 欠陥E面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* 欠陥E面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* 欠陥E面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* 欠陥E面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* 欠陥E面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* 欠陥W面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* 欠陥W面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* 欠陥W面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* 欠陥W面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* 欠陥W面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* 欠陥W面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* 欠陥W面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* 欠陥W面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* 欠陥S面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* 欠陥S面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* 欠陥S面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* 欠陥S面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* 欠陥S面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* 欠陥S面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* 欠陥S面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* 欠陥S面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* 欠陥N面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* 欠陥N面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* 欠陥N面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* 欠陥N面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* 欠陥N面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* 欠陥N面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* 欠陥N面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* 欠陥N面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* 欠陥N面個数3 */
        
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* 内部割れBSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* 内部割れBMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* 内部割れBNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* 内部割れBS種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* 内部割れBM種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* 内部割れBN種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* 内部割れBS個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* 内部割れBM個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* 内部割れBN個数 */
'
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* 内部割れTSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* 内部割れTMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* 内部割れTNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* 内部割れTS種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* 内部割れTM種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* 内部割れTN種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* 内部割れTS個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* 内部割れTM個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* 欠陥E面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* 欠陥W面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* 欠陥S面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* 欠陥N面判定 */
'        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* 欠陥B面判定 */
'        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* 欠陥T面判定 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_u_judg & "'" & ","      ''/* 欠陥U面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_d_judg & "'" & ","      ''/* 欠陥D面判定 */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* 検査員名 */
        
        '2016/04/20 - TAI - S
        strSQL = strSQL & "'" & APResData.slb_fault_total_judg & "'" & ","        ''/* 検査結果 */
        strSQL = strSQL & "'" & APResData.slb_works_sky_tok & "'" & ","           ''/* 作業場所 */
        '2016/04/20 - TAI - E
        
        strSQL = strSQL & "'" & APResData.host_send & "'" & ","             ''/* ビジコン送信結果 */
        strSQL = strSQL & "'" & APResData.host_wrt_dte & "'" & ","          ''/* ビジコン登録日 */
        strSQL = strSQL & "'" & APResData.host_wrt_tme & "'" & ","          ''/* ビジコン登録時刻 */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* 登録日 */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Write 正常終了") 'ガイダンス表示

    TRTS0014_Write = True

    On Error GoTo 0
    Exit Function

TRTS0014_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0014_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0016書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : スラブ異常報告書入力のカレントデータを書込
'
' 備考      : スラブ異常報告書入力データ書き込み
'
Public Function TRTS0016_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0016_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0016_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0016_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0016 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0016 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        
        strSQL = strSQL & "slb_col_cnt,"       ''/* ｶﾗｰ回数 */
        
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_ccno,"       ''/* スラブCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* 造塊日 */
        strSQL = strSQL & "slb_ksh,"        ''/* 鋼種 */
        strSQL = strSQL & "slb_typ,"        ''/* 型 */
        strSQL = strSQL & "slb_uksk,"       ''/* 向先 */
        strSQL = strSQL & "slb_wei,"        ''/* 重量 */
        strSQL = strSQL & "slb_lngth,"      ''/* 長さ */
        strSQL = strSQL & "slb_wdth,"       ''/* 幅 */
        strSQL = strSQL & "slb_thkns,"      ''/* 厚み */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* 次工程 */
        strSQL = strSQL & "slb_cmt1,"       ''/* コメント1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* コメント2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* 欠陥E面CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* 欠陥E面CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* 欠陥E面CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* 欠陥E面種類1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* 欠陥E面種類2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* 欠陥E面種類3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* 欠陥E面個数1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* 欠陥E面個数2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* 欠陥W面CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* 欠陥W面CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* 欠陥W面CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* 欠陥W面種類1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* 欠陥W面種類2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* 欠陥W面種類3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* 欠陥W面個数1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* 欠陥W面個数2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* 欠陥S面CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* 欠陥S面CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* 欠陥S面CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* 欠陥S面種類1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* 欠陥S面種類2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* 欠陥S面種類3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* 欠陥S面個数1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* 欠陥S面個数2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* 欠陥N面CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* 欠陥N面CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* 欠陥N面CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* 欠陥N面種類1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* 欠陥N面種類2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* 欠陥N面種類3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* 欠陥N面個数1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* 欠陥N面個数2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* 欠陥N面個数3 */
        
'        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* 内部割れBSCD */
'        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* 内部割れBMCD */
'        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* 内部割れBNCD */
'        strSQL = strSQL & "slb_fault_bs_s,"     ''/* 内部割れBS種類 */
'        strSQL = strSQL & "slb_fault_bm_s,"     ''/* 内部割れBM種類 */
'        strSQL = strSQL & "slb_fault_bn_s,"     ''/* 内部割れBN種類 */
'        strSQL = strSQL & "slb_fault_bs_n,"     ''/* 内部割れBS個数 */
'        strSQL = strSQL & "slb_fault_bm_n,"     ''/* 内部割れBM個数 */
'        strSQL = strSQL & "slb_fault_bn_n,"     ''/* 内部割れBN個数 */
'
'        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* 内部割れTSCD */
'        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* 内部割れTMCD */
'        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* 内部割れTNCD */
'        strSQL = strSQL & "slb_fault_ts_s,"     ''/* 内部割れTS種類 */
'        strSQL = strSQL & "slb_fault_tm_s,"     ''/* 内部割れTM種類 */
'        strSQL = strSQL & "slb_fault_tn_s,"     ''/* 内部割れTN種類 */
'        strSQL = strSQL & "slb_fault_ts_n,"     ''/* 内部割れTS個数 */
'        strSQL = strSQL & "slb_fault_tm_n,"     ''/* 内部割れTM個数 */
'        strSQL = strSQL & "slb_fault_tn_n,"     ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* 欠陥E面判定 */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* 欠陥W面判定 */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* 欠陥S面判定 */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* 欠陥N面判定 */
'        strSQL = strSQL & "slb_fault_b_judg,"       ''/* 欠陥B面判定 */
'        strSQL = strSQL & "slb_fault_t_judg,"       ''/* 欠陥T面判定 */
        strSQL = strSQL & "slb_fault_u_judg,"       ''/* 欠陥U面判定 */
        strSQL = strSQL & "slb_fault_d_judg,"       ''/* 欠陥D面判定 */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* 検査員名 */
        
        strSQL = strSQL & "host_send,"          ''/* ビジコン送信結果 */
        strSQL = strSQL & "host_wrt_dte,"       ''/* ビジコン登録日 */
        strSQL = strSQL & "host_wrt_tme,"       ''/* ビジコン登録時刻 */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","     ''/* ｶﾗｰ回数 */
        
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* スラブCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* 造塊日 */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* 鋼種 */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* 型 */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* 向先 */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* 重量 */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* 長さ */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* 幅 */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* 厚み */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* 次工程 */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* コメント1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* コメント2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* 欠陥E面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* 欠陥E面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* 欠陥E面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* 欠陥E面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* 欠陥E面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* 欠陥E面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* 欠陥E面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* 欠陥E面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* 欠陥E面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* 欠陥W面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* 欠陥W面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* 欠陥W面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* 欠陥W面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* 欠陥W面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* 欠陥W面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* 欠陥W面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* 欠陥W面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* 欠陥W面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* 欠陥S面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* 欠陥S面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* 欠陥S面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* 欠陥S面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* 欠陥S面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* 欠陥S面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* 欠陥S面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* 欠陥S面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* 欠陥S面個数3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* 欠陥N面CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* 欠陥N面CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* 欠陥N面CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* 欠陥N面種類1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* 欠陥N面種類2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* 欠陥N面種類3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* 欠陥N面個数1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* 欠陥N面個数2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* 欠陥N面個数3 */
        
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* 内部割れBSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* 内部割れBMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* 内部割れBNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* 内部割れBS種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* 内部割れBM種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* 内部割れBN種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* 内部割れBS個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* 内部割れBM個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* 内部割れBN個数 */
'
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* 内部割れTSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* 内部割れTMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* 内部割れTNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* 内部割れTS種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* 内部割れTM種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* 内部割れTN種類 */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* 内部割れTS個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* 内部割れTM個数 */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* 内部割れTN個数 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* 欠陥E面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* 欠陥W面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* 欠陥S面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* 欠陥N面判定 */
'        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* 欠陥B面判定 */
'        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* 欠陥T面判定 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_u_judg & "'" & ","      ''/* 欠陥U面判定 */
        strSQL = strSQL & "'" & APResData.slb_fault_d_judg & "'" & ","      ''/* 欠陥D面判定 */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* 検査員名 */
        
        strSQL = strSQL & "'" & APResData.fail_host_send & "'" & ","             ''/* スラブ異常報告　ビジコン送信結果 */
        strSQL = strSQL & "'" & APResData.fail_host_wrt_dte & "'" & ","          ''/* スラブ異常報告　ビジコン登録日 */
        strSQL = strSQL & "'" & APResData.fail_host_wrt_tme & "'" & ","          ''/* スラブ異常報告　ビジコン登録時刻 */
        
        strSQL = strSQL & "'" & APResData.fail_sys_wrt_dte & "'" & ","           ''/* スラブ異常報告　登録日 */
        strSQL = strSQL & "'" & APResData.fail_sys_wrt_tme & "'" & ","           ''/* スラブ異常報告　登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0016_Write 正常終了") 'ガイダンス表示

    TRTS0016_Write = True

    On Error GoTo 0
    Exit Function

TRTS0016_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0016_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0016_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0022書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 処置結果入力のデータを書込
'
' 備考      : 処置結果入力データ書き込み
'
Public Function TRTS0022_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0022_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0022_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0022_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0022 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        For nI = 0 To UBound(APDirResData) - 1

            If APDirResData(nI).res_sys_wrt_dte = "" Then
                APDirResData(nI).res_sys_wrt_dte = Format(Now, "YYYYMMDD")
                APDirResData(nI).res_sys_wrt_tme = Format(Now, "HHMMSS")
            End If

            '********** レコード追加 **********
            strSQL = "INSERT INTO TRTS0022 ("
    '        '-------------------------------------
            strSQL = strSQL & "slb_no," ''/* スラブＮＯ */
            strSQL = strSQL & "slb_stat,"   ''/* 状態 */
            strSQL = strSQL & "slb_col_cnt,"    ''/* カラー回数 */
            strSQL = strSQL & "res_no," ''/* 実績番号 */
            strSQL = strSQL & "slb_chno,"   ''/* スラブチャージNO */
            strSQL = strSQL & "slb_aino,"   ''/* スラブ合番 */
            strSQL = strSQL & "res_nme1,"   ''/* 実績項目1 */
            strSQL = strSQL & "res_val1,"   ''/* 実績値1 */
            strSQL = strSQL & "res_uni1,"   ''/* 実績単位1 */
            strSQL = strSQL & "res_nme2,"   ''/* 実績項目2 */
            strSQL = strSQL & "res_val2,"   ''/* 実績値2 */
            strSQL = strSQL & "res_uni2,"   ''/* 実績単位2 */
            strSQL = strSQL & "res_cmt1,"   ''/* コメント1 */
            strSQL = strSQL & "res_cmt2,"   ''/* コメント2 */
            strSQL = strSQL & "res_cmp_flg,"    ''/* 処置完了フラグ */
            strSQL = strSQL & "res_aft_stat,"   ''/* 処置後状態 */
            strSQL = strSQL & "res_wrt_dte,"    ''/* 入力日 */
            strSQL = strSQL & "res_wrt_nme,"    ''/* 入力者名 */
            strSQL = strSQL & "host_send,"          ''/* ビジコン送信結果 */
            strSQL = strSQL & "host_wrt_dte,"       ''/* ビジコン登録日 */
            strSQL = strSQL & "host_wrt_tme,"       ''/* ビジコン登録時刻 */
            strSQL = strSQL & "sys_wrt_dte,"    ''/* 登録日 */
            strSQL = strSQL & "sys_wrt_tme,"    ''/* 登録時刻 */
            strSQL = strSQL & "sys_rwrt_dte,"   ''/* 更新日 */
            strSQL = strSQL & "sys_rwrt_tme,"   ''/* 更新時刻 */
            strSQL = strSQL & "sys_acs_pros,"   ''/* アクセスプロセス名 */
            strSQL = strSQL & "sys_acs_enum"   ''/* アクセス社員ＮＯ */
            
            '---
    
            strSQL = strSQL & ") VALUES ("
    
            strSQL = strSQL & "'" & APDirResData(nI).slb_no & "'" & ","        ''/* スラブＮＯ */
            strSQL = strSQL & "'" & APDirResData(nI).slb_stat & "'" & ","      ''/* 状態 */
            
            strSQL = strSQL & "'" & Format(CInt(APDirResData(nI).slb_col_cnt), "00") & "'" & ","     ''/* ｶﾗｰ回数 */
        
            strSQL = strSQL & "'" & APDirResData(nI).dir_no & "'" & "," ''/* 実績番号 */
            strSQL = strSQL & "'" & APDirResData(nI).slb_chno & "'" & ","   ''/* スラブチャージNO */
            strSQL = strSQL & "'" & APDirResData(nI).slb_aino & "'" & ","   ''/* スラブ合番 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_nme1 & "'" & ","   ''/* 実績項目1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_val1 & "'" & ","   ''/* 実績値1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_uni1 & "'" & ","   ''/* 実績単位1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_nme2 & "'" & ","   ''/* 実績項目2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_val2 & "'" & ","   ''/* 実績値2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_uni2 & "'" & ","   ''/* 実績単位2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_cmt1 & "'" & ","   ''/* コメント1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_cmt2 & "'" & ","   ''/* コメント2 */
            strSQL = strSQL & "'" & APDirResData(nI).res_cmp_flg & "'" & ","    ''/* 処置完了フラグ */
            strSQL = strSQL & "'" & APDirResData(nI).res_aft_stat & "'" & ","   ''/* 処置後状態 */
            strSQL = strSQL & "'" & APDirResData(nI).res_wrt_dte & "'" & ","    ''/* 入力日 */
            strSQL = strSQL & "'" & APDirResData(nI).res_wrt_nme & "'" & ","    ''/* 入力者名 */
            strSQL = strSQL & "'" & APResData.fail_res_host_send & "'" & ","        ''/* ビジコン送信結果 */
            strSQL = strSQL & "'" & APResData.fail_res_host_wrt_dte & "'" & ","     ''/* ビジコン登録日 */
            strSQL = strSQL & "'" & APResData.fail_res_host_wrt_tme & "'" & ","     ''/* ビジコン登録時刻 */
            strSQL = strSQL & "'" & APDirResData(nI).res_sys_wrt_dte & "'" & ","     ''/* 登録日 */
            strSQL = strSQL & "'" & APDirResData(nI).res_sys_wrt_tme & "'" & ","    ''/* 登録時刻 */
            strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","   ''/* 更新日 */
            strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","   ''/* 更新時刻 */
            strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","   ''/* アクセスプロセス名 */
            strSQL = strSQL & "'" & "" & "'" & ")"   ''/* アクセス社員ＮＯ */
        
    '        '-------------------------------------
    '
            Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
            '-cn.Execute (strSQL)
            oDB.ExecuteSql (strSQL)
        
        Next nI
    
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0022_Write 正常終了") 'ガイダンス表示

    TRTS0022_Write = True

    On Error GoTo 0
    Exit Function

TRTS0022_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0022_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0022_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : DBDirResData_Read処理
'
' 引き数    : ARG1 - スラブ番号
'           : ARG2 - 状態
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してTRTS0020,TRTS0022のレコードを読込
'
' 備考      : カラーチェック異常処置指示データ読込
'           :COLORSYS
'
Public Function DBDirResData_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String, ByVal strSlb_Col_Cnt As String) As Boolean
'slb_chno          VARCHAR2(5)          /* スラブチャージNO */
'slb_aino          VARCHAR2(4)          /* スラブ合番 */
'slb_stat          VARCHAR2(1)          /* 状態 */
'slb_col_cnt       VARCHAR2(2)          /* ｶﾗｰ回数 */
    ' ADOのオブジェクト変数を宣言する
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBDirResData_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        ReDim APDirResTmpData(1)
        APDirResTmpData(0).slb_no = strSlb_No
        APDirResTmpData(0).slb_stat = strSlb_Stat
        APDirResTmpData(0).slb_col_cnt = Format(CInt(strSlb_Col_Cnt), "00")
        APDirResTmpData(0).dir_no = "01"
        APDirResTmpData(0).dir_nme1 = "指示項目1"
        APDirResTmpData(0).dir_val1 = "指示値1"
        APDirResTmpData(0).dir_uni1 = "指示単位1"
        APDirResTmpData(0).dir_nme2 = "指示項目2"
        APDirResTmpData(0).dir_val2 = "指示値2"
        APDirResTmpData(0).dir_uni2 = "指示単位2"
        APDirResTmpData(0).dir_cmt1 = "コメント1"
        APDirResTmpData(0).dir_cmt2 = "コメント2"
        APDirResTmpData(0).dir_wrt_dte = "20080505"
        APDirResTmpData(0).dir_wrt_nme = "指示者名"
        DBDirResData_Read = True
        Exit Function
    End If

    On Error GoTo DBDirResData_Read_err

    nOpen = 0

    ' Oracleとの接続を確立する
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT TRTS0020.*, "
    strSQL = strSQL & "TRTS0022.RES_CMP_FLG, "
    strSQL = strSQL & "TRTS0022.RES_AFT_STAT, "
    strSQL = strSQL & "TRTS0022.RES_WRT_DTE, "
    strSQL = strSQL & "TRTS0022.RES_WRT_NME, "
    strSQL = strSQL & "TRTS0022.SYS_WRT_DTE AS SYS_WRT_DTE22, "
    strSQL = strSQL & "TRTS0022.SYS_WRT_TME AS SYS_WRT_TME22 "
    strSQL = strSQL & "FROM TRTS0020 LEFT JOIN TRTS0022 ON (TRTS0020.DIR_NO = TRTS0022.RES_NO) "
    strSQL = strSQL & "AND (TRTS0020.SLB_COL_CNT = TRTS0022.SLB_COL_CNT) "
    strSQL = strSQL & "AND (TRTS0020.SLB_STAT = TRTS0022.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0020.SLB_NO = TRTS0022.SLB_NO) "
    strSQL = strSQL & "WHERE TRTS0020.slb_no='" & strSlb_No & "' AND TRTS0020.slb_stat='" & strSlb_Stat & "' AND TRTS0020.slb_col_cnt='" & Format(CInt(strSlb_Col_Cnt), "00") & "' "
    strSQL = strSQL & "ORDER BY TRTS0020.dir_no"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APDirResTmpData(0)
    Do While Not oDS.EOF
        APDirResTmpData(UBound(APDirResTmpData)).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value)  '' スラブＮＯ
        APDirResTmpData(UBound(APDirResTmpData)).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value)  '' 状態
        APDirResTmpData(UBound(APDirResTmpData)).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), "", oDS.Fields("slb_col_cnt").Value) '' カラー回数
        APDirResTmpData(UBound(APDirResTmpData)).dir_no = IIf(IsNull(oDS.Fields("dir_no").Value), "", oDS.Fields("dir_no").Value)  '' 指示番号
        APDirResTmpData(UBound(APDirResTmpData)).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value)  '' スラブチャージNO
        APDirResTmpData(UBound(APDirResTmpData)).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) '' スラブ合番
        APDirResTmpData(UBound(APDirResTmpData)).dir_nme1 = IIf(IsNull(oDS.Fields("dir_nme1").Value), "", oDS.Fields("dir_nme1").Value)  '' 指示項目1
        APDirResTmpData(UBound(APDirResTmpData)).dir_val1 = IIf(IsNull(oDS.Fields("dir_val1").Value), "", oDS.Fields("dir_val1").Value)  '' 指示値1
        APDirResTmpData(UBound(APDirResTmpData)).dir_uni1 = IIf(IsNull(oDS.Fields("dir_uni1").Value), "", oDS.Fields("dir_uni1").Value)  '' 指示単位1
        APDirResTmpData(UBound(APDirResTmpData)).dir_nme2 = IIf(IsNull(oDS.Fields("dir_nme2").Value), "", oDS.Fields("dir_nme2").Value)  '' 指示項目2
        APDirResTmpData(UBound(APDirResTmpData)).dir_val2 = IIf(IsNull(oDS.Fields("dir_val2").Value), "", oDS.Fields("dir_val2").Value)  '' 指示値2
        APDirResTmpData(UBound(APDirResTmpData)).dir_uni2 = IIf(IsNull(oDS.Fields("dir_uni2").Value), "", oDS.Fields("dir_uni2").Value)  '' 指示単位2
        APDirResTmpData(UBound(APDirResTmpData)).dir_cmt1 = IIf(IsNull(oDS.Fields("dir_cmt1").Value), "", oDS.Fields("dir_cmt1").Value)  '' コメント1
        APDirResTmpData(UBound(APDirResTmpData)).dir_cmt2 = IIf(IsNull(oDS.Fields("dir_cmt2").Value), "", oDS.Fields("dir_cmt2").Value)  '' コメント2
        APDirResTmpData(UBound(APDirResTmpData)).dir_wrt_dte = IIf(IsNull(oDS.Fields("dir_wrt_dte").Value), "", oDS.Fields("dir_wrt_dte").Value) '' 指示日
        APDirResTmpData(UBound(APDirResTmpData)).dir_wrt_nme = IIf(IsNull(oDS.Fields("dir_wrt_nme").Value), "", oDS.Fields("dir_wrt_nme").Value) '' 指示者名
    APDirResTmpData(UBound(APDirResTmpData)).dir_sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value)            ''登録日
    APDirResTmpData(UBound(APDirResTmpData)).dir_sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), "", oDS.Fields("sys_wrt_tme").Value)           ''登録時刻
        
    APDirResTmpData(UBound(APDirResTmpData)).res_cmp_flg = IIf(IsNull(oDS.Fields("res_cmp_flg").Value), "", oDS.Fields("res_cmp_flg").Value)           ''処置完了フラグ 1:完了
    APDirResTmpData(UBound(APDirResTmpData)).res_aft_stat = IIf(IsNull(oDS.Fields("res_aft_stat").Value), "", oDS.Fields("res_aft_stat").Value)          ''処置後状態 1:不適合有り（割れ、疵有り）
    APDirResTmpData(UBound(APDirResTmpData)).res_wrt_dte = IIf(IsNull(oDS.Fields("res_wrt_dte").Value), "", oDS.Fields("res_wrt_dte").Value)           ''入力日
    APDirResTmpData(UBound(APDirResTmpData)).res_wrt_nme = IIf(IsNull(oDS.Fields("res_wrt_nme").Value), "", oDS.Fields("res_wrt_nme").Value)           ''入力者名
    APDirResTmpData(UBound(APDirResTmpData)).res_sys_wrt_dte = IIf(IsNull(oDS.Fields("SYS_WRT_DTE22").Value), "", oDS.Fields("SYS_WRT_DTE22").Value)            ''登録日
    APDirResTmpData(UBound(APDirResTmpData)).res_sys_wrt_tme = IIf(IsNull(oDS.Fields("SYS_WRT_TME22").Value), "", oDS.Fields("SYS_WRT_TME22").Value)           ''登録時刻
        
        ReDim Preserve APDirResTmpData(UBound(APDirResTmpData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBDirResData_Read 正常終了") 'ガイダンス表示

    DBDirResData_Read = True

    On Error GoTo 0
    Exit Function

DBDirResData_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBDirResData_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBDirResData_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0050書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : スラブ肌用SCANLOC情報の書込
'
' 備考      : スラブ肌用SCANLOC情報書き込み
'
Public Function TRTS0050_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0050_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0050_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0050_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0050 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        'ファイル名作成
        strDestination = conDefault_DEFINE_SCNDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SKIN" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_00.JPG"

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0050 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCANアドレス */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCANアドレス */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* 登録日 */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0050_Write 正常終了") 'ガイダンス表示

    TRTS0050_Write = True

    On Error GoTo 0
    Exit Function

TRTS0050_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0050_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0050_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0052書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : カラーチェック用SCANLOC情報の書込
'
' 備考      : カラーチェック用SCANLOC情報書き込み
'
Public Function TRTS0052_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0052_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0052_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0052_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0052 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        'ファイル名作成
        strDestination = conDefault_DEFINE_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0052 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        strSQL = strSQL & "slb_col_cnt,"       ''/* カラー回数 */
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCANアドレス */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","      ''/* カラー回数 */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCANアドレス */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* 登録日 */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0052_Write 正常終了") 'ガイダンス表示

    TRTS0052_Write = True

    On Error GoTo 0
    Exit Function

TRTS0052_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0052_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0052_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0054書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : スラブ異常報告用SCANLOC情報の書込
'
' 備考      : スラブ異常報告用SCANLOC情報書き込み
'
Public Function TRTS0054_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0054_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        TRTS0054_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0054_Write_err

    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0054 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        'ファイル名作成
        strDestination = conDefault_DEFINE_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0054 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* スラブＮＯ */
        strSQL = strSQL & "slb_stat,"       ''/* 状態 */
        strSQL = strSQL & "slb_col_cnt,"       ''/* カラー回数 */
        strSQL = strSQL & "slb_chno,"       ''/* スラブチャージNO */
        strSQL = strSQL & "slb_aino,"       ''/* スラブ合番 */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCANアドレス */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros,"       ''/* アクセスプロセス名 */
        strSQL = strSQL & "sys_acs_enum"        ''/* アクセス社員ＮＯ */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* スラブＮＯ */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* 状態 */
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","      ''/* カラー回数 */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* スラブチャージNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* スラブ合番 */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCANアドレス */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* 登録日 */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_prosアクセスプロセス名 */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enumアクセス社員ＮＯ */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* 登録日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    'セッショントランザクションコミット
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0054_Write 正常終了") 'ガイダンス表示

    TRTS0054_Write = True

    On Error GoTo 0
    Exit Function

TRTS0054_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0054_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0054_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0060読込処理
'
' 引き数    :
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : TRTS0060のレコードを読込
'
' 備考      : スタッフ情報マスタ読込
'           :COLORSYS
'
Public Function TRTS0060_Read() As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0060_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        If UBound(APStaffData) < 1 Then
            ReDim APStaffData(3)
            APStaffData(0).inp_StaffName = "NAME1"
            APStaffData(1).inp_StaffName = "NAME2"
            APStaffData(2).inp_StaffName = "NAME3"
        End If
        TRTS0060_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0060_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0060.* From TRTS0060 ORDER BY TRTS0060.STAFF_NME", cn, adOpenStatic, adLockReadOnly
    
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset("SELECT TRTS0060.* From TRTS0060 ORDER BY TRTS0060.STAFF_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APStaffData(0)
    Do While Not oDS.EOF
        APStaffData(UBound(APStaffData)).inp_StaffName = oDS.Fields("Staff_Nme").Value
        ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Read 正常終了") 'ガイダンス表示

    TRTS0060_Read = True

    On Error GoTo 0
    
    ''TRTS0060レジストリ書込処理
    Call TRTS0060_Reg_Write
    
    Exit Function

TRTS0060_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0060_Read = False
    On Error GoTo 0
    
    ''TRTS0060レジストリ読込処理
    Call TRTS0060_Reg_Read

End Function

' @(f)
'
' 機能      : TRTS0060レジストリ書込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに書込
'
' 備考      : 保持中のデータをレジストリに書込
'           :COLORSYS
'
Public Sub TRTS0060_Reg_Write()
    Dim nI As Integer
    
    ''レジストリに保存
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPStaffDataCount", UBound(APStaffData)
    For nI = 1 To UBound(APStaffData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APStaffName" & CStr(nI), APStaffData(nI - 1).inp_StaffName
    Next nI

End Sub

' @(f)
'
' 機能      : TRTS0060レジストリ読込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに読込
'
' 備考      : 保持中のデータをレジストリに読込
'           :COLORSYS
'
Public Sub TRTS0060_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''レジストリから読み込み
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPStaffDataCount", 0)
    If nCount = 0 Then
        '社員マスタ読み込みエリア初期化
        ReDim APStaffData(1)
        APStaffData(0).inp_StaffName = "guest"
    Else
        'レジストから読み込み
        ReDim APStaffData(0)
        For nI = 1 To nCount
            APStaffData(nI - 1).inp_StaffName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APStaffName" & CStr(nI), "")
            ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' 機能      : TRTS0060書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'             ARG2 - スタッフ名
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスタッフ名情報を書込
'
' 備考      : スタッフ名マスタ書込
'           :COLORSYS
'
Public Function TRTS0060_Write(ByVal bDeleteOnly As Boolean, ByVal strStaffName As String) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0060_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        If bDeleteOnly = False Then
            APStaffData(UBound(APStaffData)).inp_StaffName = Left(strStaffName, 32)
            ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        End If
        TRTS0060_Write = True
        Exit Function
    End If
    
    ''データ追加の場合はTRTS0060レジストリ書込処理
    If bDeleteOnly = False Then
        APStaffData(UBound(APStaffData)).inp_StaffName = Left(strStaffName, 32)
        ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        Call TRTS0060_Reg_Write
    End If
    
    On Error GoTo TRTS0060_Write_err
    
    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0060 WHERE Staff_Nme='" & strStaffName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0060 ("
        strSQL = strSQL & "Staff_Nme,"          'VARCHAR2(32)         /* スタッフ名 */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* アクセスプロセス名 */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strStaffName & "'" & ","                'VARCHAR2(32)   /* スタッフ名 */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 登録日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0060", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    'セッショントランザクションコミット
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Write 正常終了") 'ガイダンス表示

    TRTS0060_Write = True

    On Error GoTo 0
    Exit Function

TRTS0060_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0060_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0062読込処理
'
' 引き数    :
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : TRTS0062のレコードを読込
'
' 備考      : 検査員情報マスタ読込
'           :COLORSYS
'
Public Function TRTS0062_Read() As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0062_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        If UBound(APInspData) < 1 Then
            ReDim APInspData(3)
            APInspData(0).inp_InspName = "NAME1"
            APInspData(1).inp_InspName = "NAME2"
            APInspData(2).inp_InspName = "NAME3"
        End If
        TRTS0062_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0062_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0062.* From TRTS0062 ORDER BY TRTS0062.INSP_NME", cn, adOpenStatic, adLockReadOnly
    
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset("SELECT TRTS0062.* From TRTS0062 ORDER BY TRTS0062.INSP_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APInspData(0)
    Do While Not oDS.EOF
        APInspData(UBound(APInspData)).inp_InspName = oDS.Fields("Insp_Nme").Value
        ReDim Preserve APInspData(UBound(APInspData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Read 正常終了") 'ガイダンス表示

    TRTS0062_Read = True

    On Error GoTo 0
    
    ''TRTS0062レジストリ書込処理
    Call TRTS0062_Reg_Write
    
    Exit Function

TRTS0062_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0062_Read = False
    On Error GoTo 0
    
    ''TRTS0062レジストリ読込処理
    Call TRTS0062_Reg_Read

End Function

' @(f)
'
' 機能      : TRTS0062レジストリ書込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに書込
'
' 備考      : 保持中のデータをレジストリに書込
'           :COLORSYS
'
Public Sub TRTS0062_Reg_Write()
    Dim nI As Integer
    
    ''レジストリに保存
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPInspDataCount", UBound(APInspData)
    For nI = 1 To UBound(APInspData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APInspName" & CStr(nI), APInspData(nI - 1).inp_InspName
    Next nI

End Sub

' @(f)
'
' 機能      : TRTS0062レジストリ読込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに読込
'
' 備考      : 保持中のデータをレジストリに読込
'           :COLORSYS
'
Public Sub TRTS0062_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''レジストリから読み込み
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPInspDataCount", 0)
    If nCount = 0 Then
        '社員マスタ読み込みエリア初期化
        ReDim APInspData(1)
        APInspData(0).inp_InspName = "guest"
    Else
        'レジストから読み込み
        ReDim APInspData(0)
        For nI = 1 To nCount
            APInspData(nI - 1).inp_InspName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APInspName" & CStr(nI), "")
            ReDim Preserve APInspData(UBound(APInspData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' 機能      : TRTS0062書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'             ARG2 - 検査員名
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定の検査員名情報を書込
'
' 備考      : 検査員名マスタ書込
'           :COLORSYS
'
Public Function TRTS0062_Write(ByVal bDeleteOnly As Boolean, ByVal strInspName As String) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0062_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        If bDeleteOnly = False Then
            APInspData(UBound(APInspData)).inp_InspName = Left(strInspName, 32)
            ReDim Preserve APInspData(UBound(APInspData) + 1)
        End If
        TRTS0062_Write = True
        Exit Function
    End If
    
    ''データ追加の場合はTRTS0062レジストリ書込処理
    If bDeleteOnly = False Then
        APInspData(UBound(APInspData)).inp_InspName = Left(strInspName, 32)
        ReDim Preserve APInspData(UBound(APInspData) + 1)
        Call TRTS0062_Reg_Write
    End If
    
    On Error GoTo TRTS0062_Write_err
    
    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0062 WHERE Insp_Nme='" & strInspName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0062 ("
        strSQL = strSQL & "Insp_Nme,"          'VARCHAR2(32)         /* スタッフ名 */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* アクセスプロセス名 */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strInspName & "'" & ","                'VARCHAR2(32)   /* スタッフ名 */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 登録日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0062", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    'セッショントランザクションコミット
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Write 正常終了") 'ガイダンス表示

    TRTS0062_Write = True

    On Error GoTo 0
    Exit Function

TRTS0062_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0062_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : TRTS0066読込処理
'
' 引き数    :
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : TRTS0066のレコードを読込
'
' 備考      : 入力者情報マスタ読込
'           :COLORSYS
'
Public Function TRTS0066_Read() As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0066_Read:ＤＢスキップモードです。") 'ガイダンス表示
        
        If UBound(APInpData) < 1 Then
            ReDim APInpData(3)
            APInpData(0).inp_InpName = "NAME1"
            APInpData(1).inp_InpName = "NAME2"
            APInpData(2).inp_InpName = "NAME3"
        End If
        TRTS0066_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0066_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0066.* From TRTS0066 ORDER BY TRTS0066.INSP_NME", cn, adOpenStatic, adLockReadOnly
    
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset("SELECT TRTS0066.* From TRTS0066 ORDER BY TRTS0066.INP_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APInpData(0)
    Do While Not oDS.EOF
        APInpData(UBound(APInpData)).inp_InpName = oDS.Fields("Inp_Nme").Value
        ReDim Preserve APInpData(UBound(APInpData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Read 正常終了") 'ガイダンス表示

    TRTS0066_Read = True

    On Error GoTo 0
    
    ''TRTS0066レジストリ書込処理
    Call TRTS0066_Reg_Write
    
    Exit Function

TRTS0066_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0066_Read = False
    On Error GoTo 0
    
    ''TRTS0066レジストリ読込処理
    Call TRTS0066_Reg_Read

End Function

' @(f)
'
' 機能      : TRTS0066レジストリ書込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに書込
'
' 備考      : 保持中のデータをレジストリに書込
'           :COLORSYS
'
Public Sub TRTS0066_Reg_Write()
    Dim nI As Integer
    
    ''レジストリに保存
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPInpDataCount", UBound(APInpData)
    For nI = 1 To UBound(APInpData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APInpName" & CStr(nI), APInpData(nI - 1).inp_InpName
    Next nI

End Sub

' @(f)
'
' 機能      : TRTS0066レジストリ読込処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 保持中のデータをレジストリに読込
'
' 備考      : 保持中のデータをレジストリに読込
'           :COLORSYS
'
Public Sub TRTS0066_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''レジストリから読み込み
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPInpDataCount", 0)
    If nCount = 0 Then
        '社員マスタ読み込みエリア初期化
        ReDim APInpData(1)
        APInpData(0).inp_InpName = "guest"
    Else
        'レジストから読み込み
        ReDim APInpData(0)
        For nI = 1 To nCount
            APInpData(nI - 1).inp_InpName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APInpName" & CStr(nI), "")
            ReDim Preserve APInpData(UBound(APInpData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' 機能      : TRTS0066書込処理
'
' 引き数    : ARG1 - 削除のみ実行フラグ
'             ARG2 - 入力者名
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定の入力者名情報を書込
'
' 備考      : 入力者名マスタ書込
'           :COLORSYS
'
Public Function TRTS0066_Write(ByVal bDeleteOnly As Boolean, ByVal strInpName As String) As Boolean
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0066_Write:ＤＢスキップモードです。") 'ガイダンス表示
        
        If bDeleteOnly = False Then
            APInpData(UBound(APInpData)).inp_InpName = Left(strInpName, 32)
            ReDim Preserve APInpData(UBound(APInpData) + 1)
        End If
        TRTS0066_Write = True
        Exit Function
    End If
    
    ''データ追加の場合はTRTS0066レジストリ書込処理
    If bDeleteOnly = False Then
        APInpData(UBound(APInpData)).inp_InpName = Left(strInpName, 32)
        ReDim Preserve APInpData(UBound(APInpData) + 1)
        Call TRTS0066_Reg_Write
    End If
    
    On Error GoTo TRTS0066_Write_err
    
    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    '********** レコード削除 **********
    strSQL = "DELETE From TRTS0066 WHERE Inp_Nme='" & strInpName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** レコード追加 **********
        strSQL = "INSERT INTO TRTS0066 ("
        strSQL = strSQL & "Inp_Nme,"          'VARCHAR2(32)         /* スタッフ名 */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* 登録日 */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* 登録時刻 */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* 更新日 */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* 更新時刻 */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* アクセスプロセス名 */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strInpName & "'" & ","                'VARCHAR2(32)   /* スタッフ名 */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 登録日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 登録時刻 */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* 更新日 */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* 更新時刻 */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* アクセスプロセス名 */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0066", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    'セッショントランザクションコミット
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Write 正常終了") 'ガイダンス表示

    TRTS0066_Write = True

    On Error GoTo 0
    Exit Function

TRTS0066_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Write 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    TRTS0066_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' 機能      : 素材統括ＤＢ－NCHTAISL読込処理
'
' 引き数    : ARG1 - スラブ番号
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブ番号を使用してNCHTAISLのレコードを読込
'
' 備考      :
'           :COLORSYS
'
Public Function SOZAI_NCHTAISL_Read(ByVal strSlb_No As String) As Boolean
    
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("SOZAI_DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read:素材統括ＤＢスキップモードです。") 'ガイダンス表示
        
        ReDim APSozaiTmpData(1)
        '**********************************************************'
        'nchtaisl
        APSozaiTmpData(0).slb_no = "123451234"      ''スラブNO"
        APSozaiTmpData(0).slb_ksh = "ABCDEF"        ''鋼種
        APSozaiTmpData(0).slb_uksk = "AB"          ''向先（熱延向先）
        APSozaiTmpData(0).slb_lngth = "12345"       ''長さ
        APSozaiTmpData(0).slb_color_wei = "12345"   ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
        APSozaiTmpData(0).slb_typ = "ABC"           ''型
        APSozaiTmpData(0).slb_skin_wei = "12345"    ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
        APSozaiTmpData(0).slb_wdth = "1234"         ''幅
        APSozaiTmpData(0).slb_thkns = "123.12"      ''厚み
        APSozaiTmpData(0).slb_zkai_dte = "20080101" ''造塊日（造塊年月日）
        '**********************************************************'
'        'skjchjdtテーブル
'        APSozaiTmpData(0).slb_chno = "12345"        ''チャージNO
'        APSozaiTmpData(0).slb_ccno = "12345"        ''CCNO
        '**********************************************************'
        SOZAI_NCHTAISL_Read = True
        Exit Function
    End If

    On Error GoTo SOZAI_NCHTAISL_Read_err

    ReDim APSozaiTmpData(0)
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_SOZAI, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '2008/08/30 A.K NCHTAISL最新データ抽出バージョン（システム日付よりも未来データを除く）
    'strSQL = "SELECT * FROM NCHTAISL WHERE slbno='" & strSlb_No & "'"
    
    strSQL = "SELECT * FROM "
    strSQL = strSQL & "(SELECT NCHTAISL.SLBNO,NCHTAISL.鋼種,NCHTAISL.熱延向先,NCHTAISL.長さ,"
    strSQL = strSQL & "NCHTAISL.SEG出側重量,NCHTAISL.型,NCHTAISL.黒皮重量,NCHTAISL.幅,"
    strSQL = strSQL & "NCHTAISL.厚み,NCHTAISL.造塊日付年,NCHTAISL.造塊日付月,NCHTAISL.造塊日付日,"
    strSQL = strSQL & "((NCHTAISL.造塊日付年 * 10000) + (NCHTAISL.造塊日付月 * 100) + NCHTAISL.造塊日付日) as nYYYYMMDD "
    strSQL = strSQL & "FROM NCHTAISL "
    strSQL = strSQL & "WHERE (((NCHTAISL.SLBNO)='" & strSlb_No & "'))) "
    strSQL = strSQL & "WHERE nYYYYMMDD <= '" & Format(Now, "YYYYMMDD") & "' "
    strSQL = strSQL & "ORDER BY nYYYYMMDD DESC"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    If Not oDS.EOF Then
        ReDim APSozaiTmpData(1)

        '**********************************************************'
        'nchtaisl
        APSozaiTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slbno").Value), "", oDS.Fields("slbno").Value)                      ''スラブNO"
        APSozaiTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("鋼種").Value), "", oDS.Fields("鋼種").Value)                         ''鋼種
        APSozaiTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("熱延向先").Value), "", Left(oDS.Fields("熱延向先").Value, 2))       ''向先（熱延向先）⇒左端から２桁に丸める。
        APSozaiTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("長さ").Value), "", oDS.Fields("長さ").Value)                       ''長さ
        APSozaiTmpData(0).slb_color_wei = IIf(IsNull(oDS.Fields("SEG出側重量").Value), "", oDS.Fields("SEG出側重量").Value)     ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
        APSozaiTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("型").Value), "", oDS.Fields("型").Value)                             ''型
        APSozaiTmpData(0).slb_skin_wei = IIf(IsNull(oDS.Fields("黒皮重量").Value), "", oDS.Fields("黒皮重量").Value)            ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
        APSozaiTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("幅").Value), "", oDS.Fields("幅").Value)                            ''幅
        APSozaiTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("厚み").Value), "", oDS.Fields("厚み").Value)                       ''厚み
        APSozaiTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("造塊日付年").Value), "0000", Format(oDS.Fields("造塊日付年").Value, "0000")) & _
                                         IIf(IsNull(oDS.Fields("造塊日付月").Value), "00", Format(oDS.Fields("造塊日付月").Value, "00")) & _
                                         IIf(IsNull(oDS.Fields("造塊日付日").Value), "00", Format(oDS.Fields("造塊日付日").Value, "00"))     ''造塊日（造塊年月日）
        '**********************************************************'

        '厚みをXXX.XXへ丸め
        If APSozaiTmpData(0).slb_thkns <> "" Then
            If IsNumeric(APSozaiTmpData(0).slb_thkns) Then
                APSozaiTmpData(0).slb_thkns = ToHalfAdjust(CDbl(APSozaiTmpData(0).slb_thkns), 2)
            End If
        End If

    End If

    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read 正常終了") 'ガイダンス表示

    SOZAI_NCHTAISL_Read = True

    On Error GoTo 0
    Exit Function

SOZAI_NCHTAISL_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next
    
    Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    SOZAI_NCHTAISL_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' 機能      : 素材統括ＤＢ－SKJCHJDT読込処理
'
' 引き数    : ARG1 - スラブチャージ番号
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : 指定のスラブチャージ番号を使用してSKJCHJDTのレコードを読込
'
' 備考      :
'           :COLORSYS
'
Public Function SOZAI_SKJCHJDT_Read(ByVal strSlb_Chno As String) As Boolean
    
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    'デバックモード時
    If IsDEBUG("SOZAI_DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read:素材統括ＤＢスキップモードです。") 'ガイダンス表示
        
        ReDim APSozaiTmpData(1)
'        '**********************************************************'
'        'nchtaisl
'        APSozaiTmpData(0).slb_no = "123451234"      ''スラブNO"
'        APSozaiTmpData(0).slb_ksh = "ABCDEF"        ''鋼種
'        APSozaiTmpData(0).slb_uksk = "AB"          ''向先（熱延向先）
'        APSozaiTmpData(0).slb_lngth = "12345"       ''長さ
'        APSozaiTmpData(0).slb_color_wei = "12345"   ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
'        APSozaiTmpData(0).slb_typ = "ABC"           ''型
'        APSozaiTmpData(0).slb_skin_wei = "12345"    ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
'        APSozaiTmpData(0).slb_wdth = "1234"         ''幅
'        APSozaiTmpData(0).slb_thkns = "123.12"      ''厚み
'        APSozaiTmpData(0).slb_zkai_dte = "20080101" ''造塊日（造塊年月日）
        '**********************************************************'
        'skjchjdtテーブル
        APSozaiTmpData(0).slb_chno = "12345"        ''チャージNO
        APSozaiTmpData(0).slb_ccno = "12345"        ''CCNO
        '**********************************************************'
        SOZAI_SKJCHJDT_Read = True
        Exit Function
    End If

    On Error GoTo SOZAI_SKJCHJDT_Read_err

    ReDim APSozaiTmpData(0)
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_SOZAI, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '2008/08/30 A.K SKJCHJDT最新データ抽出バージョン（システム日付よりも未来データを除く）
    'strSQL = "SELECT * FROM SKJCHJDT WHERE chno='" & strSlb_Chno & "'"
    
    strSQL = "SELECT * FROM "
    strSQL = strSQL & "(SELECT SKJCHJDT.CHNO,SKJCHJDT.CCNO,"
    strSQL = strSQL & "SKJCHJDT.鋼種,SKJCHJDT.型,"
    strSQL = strSQL & "SKJCHJDT.LS時刻_1,SKJCHJDT.LS時刻_2,SKJCHJDT.LS時刻_3,"
    strSQL = strSQL & "((SKJCHJDT.LS時刻_1 * 10000) + (SKJCHJDT.LS時刻_2 * 100) + SKJCHJDT.LS時刻_3) as nYYYYMMDD "
    strSQL = strSQL & "FROM SKJCHJDT "
    strSQL = strSQL & "WHERE (((SKJCHJDT.CHNO)='" & strSlb_Chno & "'))) "
    strSQL = strSQL & "WHERE nYYYYMMDD <= '" & Format(Now, "YYYYMMDD") & "' "
    strSQL = strSQL & "ORDER BY nYYYYMMDD DESC"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示

    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    If Not oDS.EOF Then
        ReDim APSozaiTmpData(1)

        '**********************************************************'
        'skjchjdtテーブル
        APSozaiTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("chno").Value), "", oDS.Fields("chno").Value)        ''チャージNO
        APSozaiTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("ccno").Value), "", oDS.Fields("ccno").Value)        ''CCNO
        
        '2008/08/30 A.K NCHTAISLに該当レコードがない場合は上位画面で採用する項目を一時保存
        APSozaiTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("鋼種").Value), "", oDS.Fields("鋼種").Value)        ''鋼種
        APSozaiTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("型").Value), "", oDS.Fields("型").Value)           ''型
        '**********************************************************'

    End If

    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read 正常終了") 'ガイダンス表示

    SOZAI_SKJCHJDT_Read = True

    On Error GoTo 0
    Exit Function

SOZAI_SKJCHJDT_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next
    
    Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示

    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    SOZAI_SKJCHJDT_Read = False

    On Error GoTo 0
End Function

Private Sub DBSkinSlbSearchReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APSearchTmpSlbData(0)

    bRet = ReadCSV(App.path & "\" & "DBSkinSlbSearchRead.csv", strItem(), strData())
    
    For nI = 0 To UBound(strData, 2) - 1
        APSearchTmpSlbData(nI).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())
                
        '状態
        APSearchTmpSlbData(nI).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())

        '鋼種
        APSearchTmpSlbData(nI).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())

        '型
        APSearchTmpSlbData(nI).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())

        '向先
        APSearchTmpSlbData(nI).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        '造塊日
        APSearchTmpSlbData(nI).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())

        'ｽﾗﾌﾞ肌実績（初回記録日）
        APSearchTmpSlbData(nI).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())

        'ｽﾗﾌﾞ肌ｲﾒｰｼﾞ
        If getItemDataCSV("bAPScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False
        End If

        'ｽﾗﾌﾞ肌PDF
        If getItemDataCSV("bAPPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False
        End If

        ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
    
    Next nI

End Sub

Private Sub DBColorSlbSearchReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APSearchTmpSlbData(0)

    bRet = ReadCSV(App.path & "\" & "DBColorSlbSearchRead.csv", strItem(), strData())

    For nI = 0 To UBound(strData, 2) - 1


        APSearchTmpSlbData(nI).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())

        '状態
        APSearchTmpSlbData(nI).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())

        '●ｶﾗｰ回数
        APSearchTmpSlbData(nI).slb_col_cnt = getItemDataCSV("slb_col_cnt", nI + 1, strItem(), strData())

        '鋼種
        APSearchTmpSlbData(nI).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())

        '型
        APSearchTmpSlbData(nI).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())

        '向先
        APSearchTmpSlbData(nI).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        '造塊日
        APSearchTmpSlbData(nI).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())

        'ｶﾗｰ実績（初回記録日）
        APSearchTmpSlbData(nI).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).sys_wrt_tme = getItemDataCSV("sys_wrt_tme", nI + 1, strItem(), strData())

        '●ビジコン送信結果
        APSearchTmpSlbData(nI).host_send = getItemDataCSV("host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).host_wrt_dte = getItemDataCSV("host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).host_wrt_tme = getItemDataCSV("host_wrt_tme", nI + 1, strItem(), strData())

        'ｶﾗｰｲﾒｰｼﾞ
        If getItemDataCSV("bAPScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False
        End If

        'ｶﾗｰPDF
        If getItemDataCSV("bAPPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False
        End If

'***********************************************************************
        '異常報告（初回記録日）
        APSearchTmpSlbData(nI).fail_sys_wrt_dte = getItemDataCSV("fail_sys_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_sys_wrt_tme = getItemDataCSV("fail_sys_wrt_tme", nI + 1, strItem(), strData())

        '異常報告ビジコン送信結果
        APSearchTmpSlbData(nI).fail_host_send = getItemDataCSV("fail_host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_host_wrt_dte = getItemDataCSV("fail_host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_host_wrt_tme = getItemDataCSV("fail_host_wrt_tme", nI + 1, strItem(), strData())

        '異常ｲﾒｰｼﾞ
        If getItemDataCSV("bAPFailScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPFailScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPFailScanInput = False
        End If

        '異常PDF
        If getItemDataCSV("bAPFailPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPFailPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPFailPdfInput = False
        End If

'***********************************************************************
        'CCNO
        APSearchTmpSlbData(nI).slb_ccno = getItemDataCSV("slb_ccno", nI + 1, strItem(), strData())

        '重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量 sozai="slb_color_wei"）
        APSearchTmpSlbData(nI).slb_wei = getItemDataCSV("slb_wei", nI + 1, strItem(), strData())

        '長さ
        APSearchTmpSlbData(nI).slb_lngth = getItemDataCSV("slb_lngth", nI + 1, strItem(), strData())

        '幅
        APSearchTmpSlbData(nI).slb_wdth = getItemDataCSV("slb_wdth", nI + 1, strItem(), strData())

        '厚み
        APSearchTmpSlbData(nI).slb_thkns = getItemDataCSV("slb_thkns", nI + 1, strItem(), strData())

'***********************************************************************
        '処置指示
        APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = getItemDataCSV("fail_dir_sys_wrt_dte", nI + 1, strItem(), strData())

'***********************************************************************
        '処置結果
        APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = getItemDataCSV("fail_res_sys_wrt_dte", nI + 1, strItem(), strData())

        '処置結果完了フラグ
        APSearchTmpSlbData(nI).fail_res_cmp_flg = getItemDataCSV("fail_res_cmp_flg", nI + 1, strItem(), strData())

        '処置結果ビジコン送信結果
        APSearchTmpSlbData(nI).fail_res_host_send = getItemDataCSV("fail_res_host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_res_host_wrt_dte = getItemDataCSV("fail_res_host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_res_host_wrt_tme = getItemDataCSV("fail_res_host_wrt_tme", nI + 1, strItem(), strData())

        ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)

    Next nI

End Sub


' @(f)
'
' 機能      : 状態キー変更先データＤＢ確認
'
' 引き数    : ARG1 - 検索スラブＮｏ．
'
' 返り値    : True データ無／False データ有
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      :
'
Public Function DBStatChgCheckSKIN(ByVal sSlbno As String, ByVal sSlbStat As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN:ＤＢスキップモードです。") 'ガイダンス表示
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBStatChgCheckSKIN_err
    
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    ' 変更先データ確認クエリ
    strSQL = "SELECT TRTS0012.SLB_NO "
    strSQL = strSQL & "FROM TRTS0012 "
    strSQL = strSQL & "WHERE TRTS0012.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0012.SLB_STAT = '" & sSlbStat & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
        ' データ無
        DBStatChgCheckSKIN = True
    Else
        ' データ有
        DBStatChgCheckSKIN = False
    End If
    oDS.Close
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN 正常終了") 'ガイダンス表示

    On Error GoTo 0

    Exit Function

DBStatChgCheckSKIN_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBStatChgCheckSKIN = False

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : 状態キー変更先データＤＢ確認
'
' 引き数    : ARG1 - 検索スラブＮｏ．
'
' 返り値    : 0 データ無／1 TRTS0012データ有／2 TRTS0020データ有
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      :
'
Public Function DBStatChgCheckCOLOR(ByVal sSlbno As String, ByVal sSlbStat As String, ByVal sSlbStatNow As String, ByVal sSlbColCnt As String) As Integer
    ' ADOのオブジェクト変数を宣言する
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR:ＤＢスキップモードです。") 'ガイダンス表示
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBStatChgCheckCOLOR_err
    
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    ' 変更先データ確認クエリ
    strSQL = "SELECT TRTS0014.SLB_NO "
    strSQL = strSQL & "FROM TRTS0014 "
    strSQL = strSQL & "WHERE TRTS0014.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0014.SLB_STAT = '" & sSlbStat & "' "
    strSQL = strSQL & "AND TRTS0014.SLB_COL_CNT = '" & sSlbColCnt & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
        ' データ無
        DBStatChgCheckCOLOR = 0
    Else
        ' データ有
        DBStatChgCheckCOLOR = 1
    End If
    oDS.Close
    
    ' 指示データ確認クエリ・指示データが存在したら変更しない
    strSQL = "SELECT TRTS0020.SLB_NO "
    strSQL = strSQL & "FROM TRTS0020 "
    strSQL = strSQL & "WHERE TRTS0020.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0020.SLB_STAT = '" & sSlbStatNow & "' "
    strSQL = strSQL & "AND TRTS0020.SLB_COL_CNT = '" & sSlbColCnt & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    'オラクルダイナセットオブジェクトの作成
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
    Else
        ' データ有
        DBStatChgCheckCOLOR = 2
    End If
    oDS.Close
    
    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR 正常終了") 'ガイダンス表示

    On Error GoTo 0

    Exit Function

DBStatChgCheckCOLOR_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBStatChgCheckCOLOR = 1

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : 状態キー変更先データＤＢ確認
'
' 引き数    : ARG1 - 検索スラブＮｏ．
'
' 返り値    : True データ無／False データ有
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      :
'
Public Function DBStatChgFixSKIN(ByVal sSlbno As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatNow As String, ByVal sStatNew As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim sScanAddr As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN:ＤＢスキップモードです。") 'ガイダンス表示
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBStatChgFixSKIN_err
    
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    ' TRTS0012 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0012 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0040 DELETE *****************************************************************************
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    strSQL = "DELETE FROM TRTS0040 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0040 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    sScanAddr = conDefault_DEFINE_PDFDIR & "\SKIN" & "\" & sChno & "\" & sAino & _
                                           "\SKIN" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_00.PDF"
    strSQL = "UPDATE TRTS0040 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0050 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0050 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0050 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    sScanAddr = conDefault_DEFINE_SCNDIR & "\SKIN" & "\" & sChno & "\" & sAino & _
                                           "\SKIN" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_00.JPG"
    strSQL = "UPDATE TRTS0050 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    'セッショントランザクションコミット
    oSess.CommitTrans

    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN 正常終了") 'ガイダンス表示
    DBStatChgFixSKIN = True

    On Error GoTo 0

    Exit Function

DBStatChgFixSKIN_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        oSess.RollbackTrans
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBStatChgFixSKIN = False

    On Error GoTo 0

End Function

' @(f)
'
' 機能      : 状態キー変更先データＤＢ確認
'
' 引き数    : ARG1 - 検索スラブＮｏ．
'
' 返り値    : True データ無／False データ有
'
' 機能説明  : 指定のスラブ番号を使用してスラブ情報を検索する
'
' 備考      :
'
Public Function DBStatChgFixCOLOR(ByVal sSlbno As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatNow As String, ByVal sStatNew As String, ByVal sColCntOld As String, ByVal sColCntNew As String) As Boolean
    ' ADOのオブジェクト変数を宣言する
    Dim oSess As Object     'オラクルセッションオブジェクト
    Dim oDB As Object       'オラクルデータベースオブジェクト
    Dim oDS As Object       'オラクルダイナセットオブジェクト
    Dim sId As String       'ユーザ名
    Dim sPass As String     'パスワード
    Dim sHost As String     'ホスト接続文字列
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim sScanAddr As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    'デバックモード時
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR:ＤＢスキップモードです。") 'ガイダンス表示
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "ＤＢオンラインモードです。") 'ガイダンス表示
    
    On Error GoTo DBStatChgFixCOLOR_err
    
    nOpen = 0

    ' Oracleとの接続を確立する
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    'オラクルセッションオブジェクトの作成
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    'オラクルデータベースオブジェクトの作成
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    'セッショントランザクション開始
    oSess.BeginTrans

    ' TRTS0014 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0014 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0016 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0016 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0020 DELETE *****************************************************************************
'    strSQL = "DELETE FROM TRTS0020 "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0020 UPDATE *****************************************************************************
'    strSQL = "UPDATE TRTS0020 "
'    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0022 DELETE *****************************************************************************
'    strSQL = "DELETE FROM TRTS0022 "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0022 UPDATE *****************************************************************************
'    strSQL = "UPDATE TRTS0022 "
'    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
'    oDB.ExecuteSql (strSQL)

    ' TRTS0042 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0042 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0042 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    sScanAddr = conDefault_DEFINE_PDFDIR & "\COLOR" & "\" & sChno & "\" & sAino & _
                                           "\COLOR" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".PDF"
    strSQL = "UPDATE TRTS0042 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0044 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0044 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0044 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    sScanAddr = conDefault_DEFINE_PDFDIR & "\SLBFAIL" & "\" & sChno & "\" & sAino & _
                                           "\SLBFAIL" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".PDF"
    strSQL = "UPDATE TRTS0044 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0052 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0052 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    ' TRTS0052 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    sScanAddr = conDefault_DEFINE_SCNDIR & "\COLOR" & "\" & sChno & "\" & sAino & _
                                           "\COLOR" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".JPG"
    strSQL = "UPDATE TRTS0052 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0054 DELETE *****************************************************************************
    strSQL = "DELETE TRTS0054 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0054 UPDATE *****************************************************************************
    '参照ディレクトリ・ファイル名作成
    sScanAddr = conDefault_DEFINE_SCNDIR & "\SLBFAIL" & "\" & sChno & "\" & sAino & _
                                           "\SLBFAIL" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".JPG"
    strSQL = "UPDATE TRTS0054 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") 'ガイダンス表示
    oDB.ExecuteSql (strSQL)

    'セッショントランザクションコミット
    oSess.CommitTrans

    Set oDB = Nothing    'データベースオブジェクトを解放
    Set oSess = Nothing  'セッションオブジェクトを解放
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR 正常終了") 'ガイダンス表示
    DBStatChgFixCOLOR = True

    On Error GoTo 0

    Exit Function

DBStatChgFixCOLOR_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR 異常終了") 'ガイダンス表示
    Call MsgLog(conProcNum_MAIN, StrTmp) 'ガイダンス表示
    
    If nOpen >= 2 Then
        Set oDB = Nothing    'データベースオブジェクトを解放
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  'セッションオブジェクトを解放
    End If

    DBStatChgFixCOLOR = False

    On Error GoTo 0

End Function

