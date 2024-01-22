VERSION 5.00
Object = "{B6B49C41-8023-4CA6-BDF0-FC5291FC6D71}#18.0#0"; "WCSockControl.ocx"
Begin VB.Form frmTRSend 
   BackColor       =   &H00C0FFFF&
   Caption         =   "TRサーバー通信"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   8865
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Timer timTimeOut 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   7620
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSKIP 
      Caption         =   "スキップ"
      Height          =   435
      Left            =   7620
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin WCSocket.WCSockControl WCSockControl1 
      Height          =   1275
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   2249
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  '透明
      Caption         =   "あああああああああああああ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmTRSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmTRSend.Frm                ver 1.00

' @(s)
' カラーチェック実績ＰＣ　通信サーバー送信表示フォーム
' 　本モジュールは通信サーバー送信表示フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object       ''コールバックオブジェクト格納
Private iCallBackID As Integer          ''コールバックＩＤ格納
Private sCmdID As String                ''送信コマンドID指定 '2008/09/04
Private strResultError As String    ''通信データ上のエラー電文
Private bTimeOutFlag As Boolean         ''通信Timeoutフラグ　True：Timeout発生　False：Timeout発生しない
                
' @(f)
'
' 機能      : コールバック設定
'
' 引き数    : ARG1 - コールバックオブジェクト
'             ARG2 - コールバックＩＤ
'
' 返り値    :
'
' 機能説明  : 戻り先コールバック情報を設定する。
'
' 備考      :2008/09/04 CmdID 追加
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer, ByVal CmdID As String)
    iCallBackID = ObjctID
    sCmdID = CmdID '2008/09/04 CmdID 追加
    Set cCallBackObject = callBackObj
End Sub

' @(f)
'
' 機能      : キャンセルの応答
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : キャンセルを応答する。
'
' 備考      : フォームをアンロードする。
'
Private Sub cmdCancelClose()
    
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResCANCEL
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' 機能      : ＯＫの応答
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＯＫを応答する。
'
' 備考      : フォームをアンロードする。
'
Private Sub cmdOKClose()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' 機能      : スキップの応答
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキップを応答する。
'
' 備考      : フォームをアンロードする。
'
Private Sub cmdSKIPClose()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResSKIP
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' 機能      : ＯＫボタン/再送信ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＯＫボタン処理/再送信ボタン処理
'
' 備考      : エラー発生内容確認ＯＫ（処理はキャンセル）
'
Private Sub cmdOK_Click()
    Call cmdCancelClose 'エラー発生時
End Sub

' @(f)
'
' 機能      : スキップボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキップボタン処理。
'
' 備考      : エラー発生内容確認、通信サーバー送信スキップ（処理は一時ＯＫ扱い）
'
Private Sub cmdSKIP_Click()
    Call cmdSKIPClose 'エラー発生時
End Sub

' @(f)
'
' 機能      : 通信サーバー送信結果分析
'
' 引き数    : ARG1 - 送信結果データ
'
' 返り値    :
'
' 機能説明  : 通信サーバー送信結果の処理を行う。
'
' 備考      : ソケット通信対応
'
Private Sub TrSendResult(ByVal strRetData As String)
    Dim nI As Integer
    Dim strResult As String
    Dim strMIL_TITLE As String
    Dim strLBLINFO As String
    Dim nErrNo As Integer
    
    nErrNo = SetResultToAPRegistSlbData(strRetData)

    If nErrNo = 0 Then
        'エラー無しＯＫ
        Call cmdOKClose
        Exit Sub
    ElseIf nErrNo > 0 Then
        'ビジコンからエラー有り
        lblinfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
        "通信サーバーからエラーが通知されました。" & vbCrLf & _
        "内容:"
        
        'COLORSYS
'        For nI = 0 To 9
        strResult = strResultError
'        If Trim(strResult) = "" Then Exit For
        lblinfo.Caption = lblinfo.Caption & Trim(strResult) & vbCrLf & "     "
'        Next nI
        
        strLBLINFO = lblinfo.Caption
        
    Else
        
        If nErrNo = -999 Then
            'ＯＣＸ応答無し（タイムアウト）のエラー
            lblinfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "通信サーバーの応答がありません。タイムアウトしました。" & vbCrLf & "通信環境が正しく設定されているか確認してください。"
            
            strLBLINFO = lblinfo.Caption
                
        ElseIf nErrNo = -888 Then
            '通信エラー有り
            lblinfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "通信サーバーの応答がありません。" & vbCrLf & "通信環境が正しく設定されているか確認してください。"
            
            strLBLINFO = lblinfo.Caption
        
        ElseIf nErrNo = -777 Then
            '通信エラー有り
            lblinfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "受信データ先頭に埋め込むデータ長さ" & vbCrLf & "と受信データバイト数が合いません。"
            
            strLBLINFO = lblinfo.Caption
        
        Else
            'その他（受信データフォーマット）エラー有り
            lblinfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "通信サーバーからの受信データに無効な文字があります。" & vbCrLf & "[" & strRetData & "]"
            
            strLBLINFO = lblinfo.Caption
            
        End If
    End If
    
    '2002-03-08 LOGに保存
    Call MsgLog(conProcNum_TRCONT, strLBLINFO)
        
    'ボタン制御
'    If nErrNo > 0 And iCallBackID = CALLBACK_TRSEND Then  'ＰＤＦ作成要求、通信サーバーからエラー通知がある場合
        cmdOK.Visible = True
        cmdSKIP.Visible = False 'スキップ不可
'    Else
'        cmdOK.Visible = True
'        cmdSKIP.Visible = True
'    End If

End Sub

' @(f)
'
' 機能      : 登録応答データの分解
'
' 引き数    : ARG1 - データ
'
' 返り値    : 0=正常終了／他=異常終了（-999：通信タイムアウト　-888：通信できません　-1：受信データフォーマットエラー）
'
' 機能説明  : 登録応答データの分解を行う。
'
' 備考      : COLORSYS
'
Private Function SetResultToAPRegistSlbData(ByVal strRetData As String) As Integer
    Dim strBuf As String
    Dim nRet As Integer
    Dim nI As Integer
    Dim sStr As String
    
    sStr = strRetData

    nRet = 0

    'Debugモード又はSKIPモード
    If IsDEBUG("TR_SKIP") Then
        Call MsgLog(conProcNum_TRCONT, "受信データ(ＰＤＦ作成要求)TRスキップ:Err No.[0] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    '通信モード
'    If strRetData = "" And IsDEBUG("HOSTDATA_DEBUG") = False Then
    If strRetData = "" And IsDEBUG("TR_SKIP") = False Then
        If bTimeOutFlag = False Then
            nRet = -888    '通信モード：通信エラーが発生
        Else
            nRet = -999    '通信モード：通信Timeoutが発生
        End If
        Call MsgLog(conProcNum_TRCONT, "受信データ(実績登録):Err No.[" & Format(nRet, "#0") & "] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If

     'コード転換
    strRetData = StrConv(strRetData, vbFromUnicode)
    
    'エラーコード取得
    strBuf = Mid(sStr, 27, 2)
    If Trim(strBuf) = "" Then
        nRet = -1                        ' フォーマットエラー
    Else
        If IsNumeric(strBuf) = False Then
            nRet = -1                    ' フォーマットエラー
        Else
            nRet = CInt(strBuf)
        End If
    End If
    
    'LOGに保存
'    If IsDEBUG("HOSTDATA_DEBUG") Then
'        Call MsgLog(conProcNum_TRCONT, "受信データ(ＰＤＦ作成要求)HOSTデバック:Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
'    Else
        Call MsgLog(conProcNum_TRCONT, "受信データ(ＰＤＦ作成要求):Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
'    End If
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "エラーコード[" & strBuf & "]")
    End If
    
    'H1 データ数チェック
    strBuf = StrConv(MidB(strRetData, 1, 4), vbUnicode)
    
    If strBuf <> "0000" Then
        nRet = -777                    ' 受信データ数が合わない場合
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    'H2 送信年月日時分
    strBuf = StrConv(MidB(strRetData, 5, 12), vbUnicode)
    
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "送信年月日時分[" & strBuf & "]")
    End If
    
    'H3 メッセージＩＤ
    strBuf = StrConv(MidB(strRetData, 17, 4), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "メッセージＩＤ[" & strBuf & "]")
    End If

    'H4 処理ＩＤ
    strBuf = StrConv(MidB(strRetData, 21, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "処理ＩＤ[" & strBuf & "]")
    End If

    'H5 伝文種別
    strBuf = StrConv(MidB(strRetData, 26, 1), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "伝文種別[" & strBuf & "]")
    End If

    'エラーコード
    strBuf = StrConv(MidB(strRetData, 27, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "エラーコード[" & strBuf & "]")
    End If

    If IsNumeric(strBuf) Then
        nRet = CInt(strBuf)
    Else
        nRet = -1                        ' フォーマットエラー
    End If

    'エラー内容
    strBuf = StrConv(MidB(strRetData, 29, 144), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "エラー内容[" & strBuf & "]")
    End If
    
    strResultError = strBuf

    'チャージＮＯ
    strBuf = StrConv(MidB(strRetData, 173, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "チャージＮＯ[" & strBuf & "]")
    End If

    '合番
    strBuf = StrConv(MidB(strRetData, 178, 4), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "合番[" & strBuf & "]")
    End If

    '状態
    strBuf = StrConv(MidB(strRetData, 182, 1), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "状態[" & strBuf & "]")
    End If

    'カラー回数
    strBuf = StrConv(MidB(strRetData, 183, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "カラー回数[" & strBuf & "]")
    End If

    SetResultToAPRegistSlbData = nRet

End Function


' @(f)
'
' 機能      : TR通信ログ出力イベント
'
' 引き数    : ARG1 - 戻り出力ログ
'
' 返り値    :
'
' 機能説明  : HOST通信ログ出力イベント時の処理を行う。
'
' 備考      :
'
Private Sub WCSockControl1_ProcessLog(ByVal strBuf As String)
    Call MsgLog(conProcNum_WINSOCKCONT, strBuf)
End Sub

' @(f)
'
' 機能      : フォームActive
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームActive時の処理を行う。
'
' 備考      :
'
Private Sub Form_Activate()
    
    Select Case sCmdID
        Case "COL01"
            Me.Caption = "ＰＤＦ作成要求送信"
        Case "COL02"
            Me.Caption = "指示印刷要求送信"
    End Select
    
    cmdOK.Caption = "OK"
    
   'iHostSendCount = 1
    
    WCSockControl1.RemotePort = APSysCfgData.nTR_PORT '通信サーバーのPortNo
    WCSockControl1.RemoteHost = APSysCfgData.TR_IP 'ビジコンのIP
    WCSockControl1.ConnectTimeOut = APSysCfgData.nTR_TOUT(1) 'オープン時のタイムアウト 秒で指定
    WCSockControl1.SendTimeOut = APSysCfgData.nTR_TOUT(2) 'データ通信時のタイムアウト 秒で指定
    WCSockControl1.RetryTimes = APSysCfgData.nTR_RETRY '通信時リトライ回数
        
    Call TrSend
    
End Sub

' @(f)
'
' 機能      : 通信サーバー送信処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 通信サーバー送信処理を行う。
'
' 備考      :
'
Private Sub TrSend()
    Dim strRet As String
    Dim strSendString As String
    Dim sLocalPath As String
    Dim lLen As Long
    
    '送信
    sLocalPath = App.path & "\"
    
    'If IsDEBUG("DISP") Then
    '    Me.Height = 7635
    '    Me.Width = 14250
    'Else
    '    Me.Height = 2715
    '    Me.Width = 8490
    'End If
    
    strSendString = GetAPResTrData()
    lLen = 38 'COLORSYS
    
    lblinfo.Caption = "送信中です。" & vbCrLf & "しばらくお待ちください。"
    
    strRet = ""
    
'    If IsDEBUG("HOSTDATA_DEBUG") Or IsDEBUG("HOSTDATA_SKIP") Then
'        WriteHostData sLocalPath & "ReqHostData.txt", strSendString
'
'        If IsDEBUG("HOSTDATA_DEBUG") Then
'            strRet = ReadHostData(sLocalPath & "RegistResult.txt")
'        End If
'        Call ApSendResult(strRet)
'        Exit Sub
'    End If
    
    If IsDEBUG("TR_SKIP") Then
        Select Case sCmdID
            Case "COL01"
                WriteTrData sLocalPath & "ReqTrDataCOL01.txt", strSendString
                strRet = "0000YYYYMMDDhhmmPC01COL01000" & Space(144) & "123451234112"
            Case "COL02"
                WriteTrData sLocalPath & "ReqTrDataCOL02.txt", strSendString
                strRet = "0000YYYYMMDDhhmmPC01COL02000" & Space(144) & "123451234112"
        End Select
        
        Call TrSendResult(strRet)
        Exit Sub
    End If
    
    bTimeOutFlag = False
    
    timTimeOut.Enabled = False ''監視タイマーＯＦＦ
    If APSysCfgData.nTR_TOUT(0) <> 0 Then
        timTimeOut.Interval = APSysCfgData.nTR_TOUT(0) * 1000 '全体監視のタイムアウト 秒からmSに変換
        timTimeOut.Enabled = True '監視タイマーＯＮ
    End If
     
    strRet = WCSockControl1.WCSSingleSendRec(strSendString, 1, lLen)
     
    timTimeOut.Enabled = False ''監視タイマーＯＦＦ
     
    Call TrSendResult(strRet)
    
End Sub

' @(f)
'
' 機能      : TR送信用データ作成
'
' 引き数    :
'
' 返り値    : TR送信用データ
'
' 機能説明  : TR送信用データの作成を行う。
'
' 備考      :COLORSYS
'
Private Function GetAPResTrData() As String
    Dim strSendString As String
    Dim nI As Integer
    Dim nJ As Integer
    
    strSendString = "XXXX"     'データ長(1300Bytes) Winsock OCXで埋め込む
    strSendString = strSendString & Format(Now, "YYYYMMDDHHMM")    'H2 送信年月日時分
    strSendString = strSendString & "PC01"   'H3 メッセージＩＤ
    strSendString = strSendString & sCmdID   'H4 処理ＩＤ
    
    'H5 伝文種別
    Select Case sCmdID
        Case "COL01"
            'Me.Caption = "ＰＤＦ作成要求送信"
            '呼出元により、処理分岐
            Select Case cCallBackObject.Name
                '*************************************************************
                Case "frmSkinScanWnd" ''スラブ肌調査表入力
                    strSendString = strSendString & CStr(conDefine_SYSMODE_SKIN)     'H5 伝文種別
                '*************************************************************
                Case "frmColorScanWnd" ''カラーチェック検査表入力
                    strSendString = strSendString & CStr(conDefine_SYSMODE_COLOR)     'H5 伝文種別
                '*************************************************************
                Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
                    strSendString = strSendString & CStr(conDefine_SYSMODE_SLBFAIL)     'H5 伝文種別
                '*************************************************************
                Case Else
                    Call WaitMsgBox(Me, "frmTRSend:GetAPResTrData:呼出元エラー")
                    Call MsgLog(conProcNum_TRCONT, "frmTRSend:GetAPResTrData:呼出元エラー")
                    GetAPResTrData = ""
                    Exit Function
            End Select
        Case "COL02"
            'Me.Caption = "指示印刷要求送信"
            strSendString = strSendString & "0" 'H5 伝文種別 未使用のため、0:固定
    End Select
    
    
    strSendString = strSendString & _
    Format(Left(APResData.slb_no, 9), "!@@@@@@@@@") 'スラブNo
  
    strSendString = strSendString & _
    Format(Left(APResData.slb_stat, 1), "!@") '状態
  
    strSendString = strSendString & Format(CInt(APResData.slb_col_cnt), "00") ''カラー回数
  
    GetAPResTrData = strSendString

    Debug.Print "SendData:[" & strSendString & "]"
    Call MsgLog(conProcNum_TRCONT, "送信データ(ＰＤＦ作成要求):[" & strSendString & "]")

End Function


' @(f)
'
' 機能      : TR送信用タイムアウトイベント
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : TR送信用タイムアウトイベント時の処理を行う。
'
' 備考      : ＶＢ側ＯＣＸタイムアウト監視用
'           :
'

Private Sub timTimeOut_Timer()

    bTimeOutFlag = True
    timTimeOut.Enabled = False '監視タイマーＯＦＦ
    WCSockControl1.WCSForceEnd   'ビジコン通信強制終了
End Sub

' @(f)
'
' 機能      : TR受信用ダミー書込
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : TR無しでのデバッグ用データ書込処理
'
' 備考      :
'
Private Sub WriteTrData(ByVal sFileName As String, ByVal sSendData As String)
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Output Access Write As #fp
    Print #fp, sSendData
    Close #fp
End Sub

