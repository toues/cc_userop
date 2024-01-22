VERSION 5.00
Object = "{B6B49C41-8023-4CA6-BDF0-FC5291FC6D71}#18.0#0"; "WCSockControl.ocx"
Begin VB.Form frmHostSend 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "実績データビジコン登録"
   ClientHeight    =   3540
   ClientLeft      =   825
   ClientTop       =   1050
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdSKIP 
      Caption         =   "スキップ"
      Height          =   435
      Left            =   7620
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   7620
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer timTimeOut 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "frmHostSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmHostSend.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　ビジコン送信表示フォーム
' 　本モジュールはビジコン送信表示フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object       ''コールバックオブジェクト格納
Private iCallBackID As Integer          ''コールバックＩＤ格納
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
' 備考      :
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer)
    iCallBackID = ObjctID
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
    If iCallBackID = CALLBACK_HOSTSEND_QUERY Then
        Call HostSend       '再送信
    Else
        Call cmdCancelClose 'エラー発生時
    End If
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
' 備考      : エラー発生内容確認、ＨＯＳＴ送信スキップ（処理は一時ＯＫ扱い）
'
Private Sub cmdSKIP_Click()
    Call cmdSKIPClose 'エラー発生時
End Sub

' @(f)
'
' 機能      : HOST送信結果分析
'
' 引き数    : ARG1 - 送信結果データ
'
' 返り値    :
'
' 機能説明  : HOST送信結果の処理を行う。
'
' 備考      : ソケット通信対応
'
Private Sub HostSendResult(ByVal strRetData As String)
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
        lblInfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
        "ビジコンからエラーが通知されました。" & vbCrLf & _
        "内容:"
        
        'COLORSYS
'        For nI = 0 To 9
        strResult = strResultError
'        If Trim(strResult) = "" Then Exit For
        lblInfo.Caption = lblInfo.Caption & Trim(strResult) & vbCrLf & "     "
'        Next nI
        
        strLBLINFO = lblInfo.Caption
        
    Else
        
        If nErrNo = -999 Then
            'ＯＣＸ応答無し（タイムアウト）のエラー
            lblInfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "ビジコンの応答がありません。タイムアウトしました。" & vbCrLf & "通信環境が正しく設定されているか確認してください。"
            
            strLBLINFO = lblInfo.Caption
                
        ElseIf nErrNo = -888 Then
            '通信エラー有り
            lblInfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "ビジコンの応答がありません。" & vbCrLf & "通信環境が正しく設定されているか確認してください。"
            
            strLBLINFO = lblInfo.Caption
        
        ElseIf nErrNo = -777 Then
            '通信エラー有り
            lblInfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "受信データ先頭に埋め込むデータ長さ" & vbCrLf & "と受信データバイト数が合いません。"
            
            strLBLINFO = lblInfo.Caption
        
        Else
            'その他（受信データフォーマット）エラー有り
            lblInfo.Caption = "エラー番号:" & CStr(nErrNo) & "　" & _
            "通信エラーが発生しました。" & vbCrLf & _
            "内容:" & "ビジコンからの受信データに無効な文字があります。" & vbCrLf & "[" & strRetData & "]"
            
            strLBLINFO = lblInfo.Caption
            
        End If
    End If
    
    '2002-03-08 LOGに保存
    Call MsgLog(conProcNum_BSCONT, strLBLINFO)
        
    'ボタン制御
    If nErrNo > 0 And iCallBackID = CALLBACK_HOSTSEND Then  '実績登録、ビジコンからエラー通知がある場合
        cmdOK.Visible = True
        cmdSKIP.Visible = False 'スキップ不可
    Else
        cmdOK.Visible = True
        cmdSKIP.Visible = True
    End If

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
    If IsDEBUG("HOSTDATA_SKIP") Then
        Call MsgLog(conProcNum_BSCONT, "受信データ(実績登録)HOSTスキップ:Err No.[0] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    '通信モード
    If strRetData = "" And IsDEBUG("HOSTDATA_DEBUG") = False Then
        If bTimeOutFlag = False Then
            nRet = -888    '通信モード：通信エラーが発生
        Else
            nRet = -999    '通信モード：通信Timeoutが発生
        End If
        Call MsgLog(conProcNum_BSCONT, "受信データ(実績登録):Err No.[" & Format(nRet, "#0") & "] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If

     'コード転換
    strRetData = StrConv(strRetData, vbFromUnicode)
    
    'エラーコード取得
    strBuf = Mid(sStr, 24, 2)
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
    If IsDEBUG("HOSTDATA_DEBUG") Then
        Call MsgLog(conProcNum_BSCONT, "受信データ(実績登録)HOSTデバック:Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
    Else
        Call MsgLog(conProcNum_BSCONT, "受信データ(実績登録):Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
    End If
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "エラーコード[" & strBuf & "]")
    End If
    
    'H1 データ数チェック
    strBuf = StrConv(MidB(strRetData, 1, 4), vbUnicode)
    
    If strBuf <> "0000" Then
        nRet = -777                    ' 受信データ数が合わない場合
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    'H2 端末名
    strBuf = StrConv(MidB(strRetData, 5, 3), vbUnicode)
    
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "端末名[" & strBuf & "]")
    End If
    
    'H3 トランザクション
    strBuf = StrConv(MidB(strRetData, 8, 8), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "トランザクション[" & strBuf & "]")
    End If

    '空白
    strBuf = StrConv(MidB(strRetData, 16, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "空白[" & strBuf & "]")
    End If

    '種別
    strBuf = StrConv(MidB(strRetData, 21, 3), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "種別[" & strBuf & "]")
    End If

    'エラーコード
    strBuf = StrConv(MidB(strRetData, 24, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "エラーコード[" & strBuf & "]")
    End If

    If IsNumeric(strBuf) Then
        nRet = CInt(strBuf)
    Else
        nRet = -1                        ' フォーマットエラー
    End If

    'エラー内容
    strBuf = StrConv(MidB(strRetData, 26, 50), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "エラー内容[" & strBuf & "]")
    End If
    
    strResultError = strBuf

    '空白
    strBuf = StrConv(MidB(strRetData, 76, 25), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "空白[" & strBuf & "]")
    End If


    SetResultToAPRegistSlbData = nRet
End Function


' @(f)
'
' 機能      : HOST通信ログ出力イベント
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
    
    If iCallBackID = CALLBACK_HOSTSEND_QUERY Then
        Me.Caption = "スラブ情報ビジコン問い合わせ"
        cmdOK.Caption = "再送信"
    Else
        Me.Caption = "実績データビジコン登録"
        cmdOK.Caption = "OK"
   End If
    
   'iHostSendCount = 1
    
    WCSockControl1.RemotePort = APSysCfgData.nHOST_PORT 'ビジコンのPortNo
    WCSockControl1.RemoteHost = APSysCfgData.HOST_IP 'ビジコンのIP
    WCSockControl1.ConnectTimeOut = APSysCfgData.nHOST_TOUT(1) 'オープン時のタイムアウト 秒で指定
    WCSockControl1.SendTimeOut = APSysCfgData.nHOST_TOUT(2) 'データ通信時のタイムアウト 秒で指定
    WCSockControl1.RetryTimes = APSysCfgData.nHOST_RETRY '通信時リトライ回数
        
    Call HostSend
    
End Sub

' @(f)
'
' 機能      : ビジコン送信処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ビジコン送信処理を行う。
'
' 備考      :
'
Private Sub HostSend()
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
    
    strSendString = GetAPResHostData()
    lLen = 150 'COLORSYS
    
    lblInfo.Caption = "送信中です。" & vbCrLf & "しばらくお待ちください。"
    
    strRet = ""
    
    If IsDEBUG("HOSTDATA_DEBUG") Or IsDEBUG("HOSTDATA_SKIP") Then
        WriteHostData sLocalPath & "ReqHostData.txt", strSendString

        If IsDEBUG("HOSTDATA_DEBUG") Then
            strRet = ReadHostData(sLocalPath & "RegistResult.txt")
        End If
        Call HostSendResult(strRet)
        Exit Sub
    End If
    
    bTimeOutFlag = False
    
    timTimeOut.Enabled = False ''監視タイマーＯＦＦ
    If APSysCfgData.nHOST_TOUT(0) <> 0 Then
        timTimeOut.Interval = APSysCfgData.nHOST_TOUT(0) * 1000 '全体監視のタイムアウト 秒からmSに変換
        timTimeOut.Enabled = True '監視タイマーＯＮ
    End If
     
    strRet = WCSockControl1.WCSSingleSendRec(strSendString, 1, lLen)
     
    timTimeOut.Enabled = False ''監視タイマーＯＦＦ
     
    Call HostSendResult(strRet)
    
End Sub

' @(f)
'
' 機能      : HOST送信用データ作成
'
' 引き数    :
'
' 返り値    : HOST送信用データ
'
' 機能説明  : HOST送信用データの作成を行う。
'
' 備考      :COLORSYS
'
Private Function GetAPResHostData() As String
    Dim strSendString As String
    Dim nI As Integer
    Dim nJ As Integer
    
    strSendString = "XXXX"     'データ長(1300Bytes) Winsock OCXで埋め込む
    strSendString = strSendString & "A96"     'H2 端末名
    strSendString = strSendString & "EA96"   'H3 トランザクション名
    strSendString = strSendString & Space(4)    'H3 余り
    strSendString = strSendString & Space(5)    '1 空白
    strSendString = strSendString & "A96"   '2 種別
  
  
    strSendString = strSendString & _
    Format(Left(APResData.slb_no, 9), "!@@@@@@@@@") 'スラブNo
  
  
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        '*************************************************************
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            strSendString = strSendString & APResData.host_send_flg
        
            If APResData.host_wrt_dte = "" Then
                APResData.host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
  
        
        '*************************************************************
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            strSendString = strSendString & APResData.host_send_flg
        
            If APResData.fail_host_wrt_dte = "" Then
                APResData.fail_host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.fail_host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
        
  
        '*************************************************************
        Case "frmDirResWnd" ''処置結果入力(完了フラグ以外、スラブ異常報告と同じ）
            strSendString = strSendString & APResData.host_send_flg
            
            If APResData.fail_host_wrt_dte = "" Then
                APResData.fail_host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.fail_host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
        
  
        '*************************************************************
        Case "frmColorSlbSelWnd" ''カラーチェック欠陥判定情報削除処理（カラーチェック－スラブ選択画面）
            
            strSendString = strSendString & "9" '取消コード
            
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_dte <> "" Then
                'カラーチェック検査表で送信済み
                strSendString = strSendString & _
                Format(Left(APResData.host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
            
            ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte <> "" Then
                'スラブ異常報告で送信済み
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
                
            ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_cmp_flg = "1" Then
                '処置結果報告で送信済み
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '作業日（年＋月＋日）YYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '作業日（時＋分）HHMM
                
            End If

        Case Else
            Call WaitMsgBox(Me, "frmHostSend:GetAPResHostData:呼出元エラー")
            Call MsgLog(conProcNum_BSCONT, "frmHostSend:GetAPResHostData:呼出元エラー")
    End Select
  
    If APResData.slb_fault_u_judg = "9" Then
        strSendString = strSendString & "*"
    Else
        strSendString = strSendString & Format(Left(APResData.slb_fault_u_judg & " ", 1), "!@")   '上面判定
    End If
  
    If APResData.slb_fault_d_judg = "9" Then
        strSendString = strSendString & "*"
    Else
        strSendString = strSendString & Format(Left(APResData.slb_fault_d_judg & " ", 1), "!@")   '下面判定
    End If
  
    strSendString = strSendString & Space(103)    '空白
  
  
    GetAPResHostData = strSendString

    Debug.Print "SendData:[" & strSendString & "]"
    Call MsgLog(conProcNum_BSCONT, "送信データ(実績登録):[" & strSendString & "]")

End Function

' @(f)
'
' 機能      : HOST受信用ダミー読出
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : HOST無しでのデバッグ用データ読出処理
'
' 備考      :
'
Private Function ReadHostData(ByVal sFileName As String)
    Dim strReadData As String
    Dim StrTmp As String
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Input As #fp
    Do While Not EOF(fp)
        Line Input #fp, StrTmp
        strReadData = strReadData & StrTmp
    Loop
    Close #fp
    ReadHostData = strReadData
End Function

' @(f)
'
' 機能      : HOST受信用ダミー書込
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : HOST無しでのデバッグ用データ書込処理
'
' 備考      :
'
Private Sub WriteHostData(ByVal sFileName As String, ByVal sSendData As String)
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Output Access Write As #fp
    Print #fp, sSendData
    Close #fp
End Sub

' @(f)
'
' 機能      : HOST送信用タイムアウトイベント
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : HOST送信用タイムアウトイベント時の処理を行う。
'
' 備考      : ＶＢ側ＯＣＸタイムアウト監視用
'           :
'

Private Sub timTimeOut_Timer()

    bTimeOutFlag = True
    timTimeOut.Enabled = False '監視タイマーＯＦＦ
    WCSockControl1.WCSForceEnd   'ビジコン通信強制終了
End Sub

