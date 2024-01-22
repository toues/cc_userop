VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm BaseWnd 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "カラーチェック情報電子化システム"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   -3210
   ClientWidth     =   19080
   Icon            =   "BaseWnd.frx":0000
   LinkTopic       =   "BasedWnd"
   Visible         =   0   'False
   WindowState     =   2  '最大化
   Begin VB.PictureBox MainControl 
      Align           =   2  '下揃え
      BackColor       =   &H00808080&
      Height          =   11355
      Index           =   0
      Left            =   0
      ScaleHeight     =   11295
      ScaleWidth      =   19020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -1275
      Width           =   19080
      Begin VB.CommandButton cmdColorIn_Tok 
         BackColor       =   &H0080FFFF&
         Caption         =   "      特鋼         ｶﾗｰﾁｪｯｸ   検査表入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   300
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   2
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdWEBURL_Color_Result_Tok 
         BackColor       =   &H0080FFFF&
         Caption         =   "特鋼  カラー結果一覧　(WEB)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   4920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   5
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdWEBURL_Color_Result 
         BackColor       =   &H00FFFF80&
         Caption         =   "SKY  カラー結果一覧　(WEB)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   9540
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   4
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdColorSlbFail 
         BackColor       =   &H00FFFF80&
         Caption         =   "   異常報告   一覧"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   14160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   3
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdColorIn 
         BackColor       =   &H00FFFF80&
         Caption         =   "      SKY         ｶﾗｰﾁｪｯｸ   検査表入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   4920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   1
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdSysCfg 
         BackColor       =   &H0080FF80&
         Caption         =   "システム設定"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   9540
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   6
         ToolTipText     =   "System setting"
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdSkinIn 
         BackColor       =   &H00FFFF80&
         Caption         =   "スラブ肌調査入力"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   300
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   0
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton ShutButton 
         BackColor       =   &H0080FF80&
         Caption         =   "終了"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   14160
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   7
         ToolTipText     =   "System shut down"
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  '下揃え
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   10590
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   28019
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "2016/05/02"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "13:12"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox MainControl 
      Align           =   2  '下揃え
      Height          =   510
      Index           =   1
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   19020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10080
      Width           =   19080
      Begin VB.ListBox lstGuidance 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         ItemData        =   "BaseWnd.frx":030A
         Left            =   1140
         List            =   "BaseWnd.frx":030C
         TabIndex        =   8
         Top             =   60
         Width           =   17775
      End
      Begin VB.Label Label4 
         Caption         =   "ガイダンス"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '上揃え
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   13
      Top             =   0
      Width           =   19080
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '上揃え
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   14
      Top             =   0
      Width           =   19080
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  '上揃え
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   15
      Top             =   0
      Width           =   19080
   End
   Begin VB.Menu mnuSkinIn 
      Caption         =   "スラブ肌調査入力"
   End
   Begin VB.Menu mnuDummy0 
      Caption         =   "          "
   End
   Begin VB.Menu mnuColorIn 
      Caption         =   "SKYカラーチェック検査表入力"
   End
   Begin VB.Menu mnuDummy1 
      Caption         =   "          "
   End
   Begin VB.Menu mnuColorIn_Tok 
      Caption         =   "特鋼カラーチェック検査表入力"
   End
   Begin VB.Menu mnuDummy5 
      Caption         =   ""
   End
   Begin VB.Menu mnuColorSlbFail 
      Caption         =   "異常報告一覧"
   End
   Begin VB.Menu mnuDummy2 
      Caption         =   ""
   End
   Begin VB.Menu mnuWEBURL_Color_Result 
      Caption         =   "SKYカラー結果一覧(WEB)"
   End
   Begin VB.Menu mnuDummy3 
      Caption         =   ""
   End
   Begin VB.Menu mnuWEBURL_Color_Result_Tok 
      Caption         =   "特鋼カラー結果一覧(WEB)"
   End
   Begin VB.Menu mnuDummy6 
      Caption         =   ""
   End
   Begin VB.Menu mnuSysCfg 
      Caption         =   "システム設定"
   End
   Begin VB.Menu mnuDummy4 
      Caption         =   ""
   End
   Begin VB.Menu mnuShutDown 
      Caption         =   "終了"
   End
End
Attribute VB_Name = "BaseWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) BaseWnd.Frm                ver 1.00
' @(s)
' カラーチェック実績ＰＣ　ＭＤＩベースフォーム
' 　本モジュールはＭＤＩベースフォームで使用する
' 　ためのものである。

Option Explicit

Public fMDIWnd As Object ''ＭＤＩ子フォーム格納

Dim m_shutDownFlag As Boolean ''終了フラグ格納

Dim WSRecFlag As Boolean ''受信電文有りフラグ
Dim WST1OutFlag As Boolean ''オープン時タイムアウト
Dim WST2OutFlag As Boolean  ''応答受信時タイムアウト
Dim WSRetryCount As Integer ''コネクト用リトライカウント



' @(f)
'
' 機能      : ＭＤＩフォームアクティブイベント
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＭＤＩフォームアクティブ時のイベント。
'
' 備考      :
'
Private Sub MDIForm_Activate()
    Me.Caption = "カラーチェック情報電子化システム" & " Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

' @(f)
'
' 機能      : ＭＤＩフォーム読込みイベント
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＭＤＩフォーム読込み時のイベント。
'
' 備考      :
'
Private Sub MDIForm_Load()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim varAppStartLog As Variant

    varAppStartLog = Empty
    varAppStartLog = FreeFile
    Open App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt" For Append Access Write As #varAppStartLog
    If IsEmpty(varAppStartLog) = False Then
        Print #varAppStartLog, Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision
    End If
    Close #varAppStartLog
    varAppStartLog = Empty

    MainLogFileNumber = Empty
    MainLogFileNumber = FreeFile               ' 未使用のファイル番号を取得します。
        
    'ログファイルを開く
    Open App.path & "\" & conDefine_LogDirName & "\" & "MAIN_LOG.txt" For Append Access Write As #MainLogFileNumber
    Call MsgLog(conProcNum_MAIN, "*********************************************************")
    Call MsgLog(conProcNum_MAIN, "******************** ＭＡＩＮログ開始 ********************")
    Call MsgLog(conProcNum_MAIN, "*********************************************************")

    '
    'For nI = 0 To 1
    '    LEAD_CAP(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    '    LEAD_CAP(nI).EnableMethodErrors = False 'False   システムエラーイベントを発生させない
    '    LEAD_CAP(nI).EnableTwainEvent = True
    '    'LEAD_LIST(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    '    'LEAD_LIST(nI).EnableMethodErrors = False 'False   システムエラーイベントを発生させない
    '    'LEAD_LIST(nI).EnableTwainEvent = True
    'Next nI
    
    Debug.Print "MDIForm_Load"
    Call InputDataClear

   '次工程マスタ情報初期化
    ReDim APNextProcDataSkin(0)
    ReDim APNextProcDataColor(0)

    Call LoadAPSysCfgDataSetting
    'Call LoadAPResDataSetting
    
    'ＣＳＯＫＥＴ開始
    'Call CSTRAN_START
    
    Call MenuUnLock

    'スタッフ名マスタ情報初期化
    ReDim APStaffData(0)

    '検査員名マスタ情報初期化
    ReDim APInspData(0)

    '入力者名マスタ情報初期化
    ReDim APInpData(0)


    'スタッフ名マスタ読込み
    bRet = TRTS0060_Read()
    
    '検査員名マスタ読込み
    bRet = TRTS0062_Read()
    
    '入力者名マスタ読込み
    bRet = TRTS0066_Read()
    
End Sub

' @(f)
'
' 機能      : ＤＢの再読み込み
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＤＢの再読み込みを行う。
'
' 備考      :
'
Public Sub ReLoad()
    Dim bRet As Boolean

    Call MsgLog(conProcNum_MAIN, "ＤＢの再読み込みを行います。")   'ガイダンス表示

    Debug.Print "MDIForm_ReLoad"
    Call InputDataClear

    Call MenuUnLock

    'スタッフ名マスタ情報初期化
    ReDim APStaffData(0)

    '検査員名マスタ情報初期化
    ReDim APInspData(0)

    '入力者名マスタ情報初期化
    ReDim APInpData(0)


    'スタッフ名マスタ読込み
    bRet = TRTS0060_Read()
    
    '検査員名マスタ読込み
    bRet = TRTS0062_Read()
    
    '入力者名マスタ読込み
    bRet = TRTS0066_Read()
    
End Sub

' @(f)
'
' 機能      : スラブ肌調査入力ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ肌調査入力ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdSkinIn_Click()
    Call mnuSkinIn_Click 'スラブ肌調査入力メニュー
End Sub

' @(f)
'
' 機能      : ｶﾗｰﾁｪｯｸ検査表入力ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ｶﾗｰﾁｪｯｸ検査表入力ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdColorIn_Click()
    Call mnuColorIn_Click 'ｶﾗｰﾁｪｯｸ検査表入力メニュー
End Sub

' @(f)
'
' 機能      : 異常報告一覧ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 異常報告一覧ボタンでメニューを開く。
'
' 備考      :COLORSYS 2008/09/03
'
Private Sub cmdColorSlbFail_Click()
    Call mnuColorSlbFail_Click '異常報告一覧メニュー
End Sub

' @(f)
'
' 機能      : カラー結果一覧(WEB)ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : カラー結果一覧(WEB)ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdWEBURL_Color_Result_Click()
    Call mnuWEBURL_Color_Result_Click 'カラー結果一覧(WEB)メニュー
End Sub

' @(f)
'
' 機能      : システム設定ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : システム設定ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdSysCfg_Click()
    Call mnuSysCfg_Click 'システム設定メニュー
End Sub

' @(f)
'
' 機能      : デバックモードボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : デバックモードボタンでメニューを開く。
'
' 備考      : レジストリのnDEBUG_MODEが1の時のみメニューを開く。
'          ：COLORSYS
'
Private Sub mnuDummy0_Click()
    '2002-03-22
    If APSysCfgData.nDEBUG_MODE = 1 Then
        Call mnuDebug_Click
    End If
End Sub


' @(f)
'
' 機能      : 終了ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 終了ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub ShutButton_Click()
    Call mnuShutDown_Click '終了メニュー
End Sub

' @(f)
'
' 機能      : メニュー使用不可切替
'
' 引き数    : ARG1 - 実行メニュー名
'
' 返り値    :
'
' 機能説明  : 実行メニューに応じて他のメニュー
'             状態を使用不可にする。
'
' 備考      :COLORSYS
'
Private Sub MenuLock(ByVal strMenuName As String)
    
    Select Case strMenuName
        
        Case "mnuSkinIn", "mnuColorIn", "mnuSysCfg", "mnuColorSlbFail", "mnuColorIn_Tok"
            
            mnuSkinIn.Enabled = False
            cmdSkinIn.Enabled = False
            
            mnuColorIn.Enabled = False
            cmdColorIn.Enabled = False
            
            '2008/09/04
            mnuColorSlbFail.Enabled = False
            cmdColorSlbFail.Enabled = False
            
            mnuSysCfg.Enabled = False
            cmdSysCfg.Enabled = False
        
            '2016/04/20 - TAI - S
            mnuColorIn_Tok.Enabled = False
            cmdColorIn_Tok.Enabled = False
            '2016/04/20 - TAI - E
        
    End Select

End Sub

' @(f)
'
' 機能      : メニュー使用可切替
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : メニュー状態を使用可にする。
'
' 備考      :COLORSYS
'
Private Sub MenuUnLock()

    mnuSkinIn.Enabled = True
    cmdSkinIn.Enabled = True
    
    mnuColorIn.Enabled = True
    cmdColorIn.Enabled = True
    
    '2008/09/04
    mnuColorSlbFail.Enabled = True
    cmdColorSlbFail.Enabled = True
    
    mnuSysCfg.Enabled = True
    cmdSysCfg.Enabled = True
    
    '2016/04/20 - TAI - S
    mnuColorIn_Tok.Enabled = True
    cmdColorIn_Tok.Enabled = True
    '2016/04/20 - TAI - E
    
End Sub

' @(f)
'
' 機能      : スラブ肌調査入力メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ肌調査入力画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuSkinIn_Click()
    Call MenuLock("mnuSkinIn")
        
    'スラブ検索リストクリア
    ReDim APSearchListSlbData(0)
    
    ' 20090115 modify by M.Aoyagi キー変更処理追加画面
    frmSkinSlbSelWnd.Show vbModeless, Me 'スラブ肌調査入力用－スラブ選択画面
End Sub

' @(f)
'
' 機能      : ｶﾗｰﾁｪｯｸ検査表入力メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ｶﾗｰﾁｪｯｸ検査表入力画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuColorIn_Click()
    Call MenuLock("mnuColorIn")
    
    '2016/04/20 - TAI - S
    '作業場所を"SKY"にする
    works_sky_tok = WORKS_SKY
    '2016/04/20 - TAI - E

    'スラブ検索リストクリア
    ReDim APSearchListSlbData(0)
    
    ' 20090115 modify by M.Aoyagi キー変更処理追加画面
    frmColorSlbSelWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力用－スラブ選択画面
End Sub

' @(f)
'
' 機能      : 異常報告一覧メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 異常報告一覧画面を開く。
'
' 備考      :COLORSYS 2008/09/03
'
Private Sub mnuColorSlbFail_Click()
    Call MenuLock("mnuColorSlbFail")
    
    'スラブ検索リストクリア
    ReDim APSearchListSlbData(0)
    
    frmColorSlbFailWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力用－異常報告一覧画面
End Sub

' @(f)
'
' 機能      : カラー結果一覧(WEB)メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : カラー結果一覧(WEB)をIEで開く。
'
' 備考      :COLORSYS 2008/09/03
'
Private Sub mnuWEBURL_Color_Result_Click()
    Dim RetVal
    RetVal = Shell(APSysCfgData.WEBURL_Color_Result, 3)
End Sub

' @(f)
'
' 機能      : デバックモードメニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : デバックモード画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuDebug_Click()
    frmDEBUG.Show vbModeless, Me
End Sub

' @(f)
'
' 機能      : システム設定メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : システム設定画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuSysCfg_Click()
    Dim Result As Boolean
    
    Call MenuLock("mnuSysCfg")

    'Change user
    Result = cUser.ChangeUser
    If Result = False Then
        'ログオン不可
        Call MsgLog(conProcNum_MAIN, "ログオン不可")
        Call MenuUnLock
    Else
        'オグオン許可
        Call MsgLog(conProcNum_MAIN, "ログオン許可")
    
        'If Not fMDIWnd Is Nothing Then
        '    Unload fMDIWnd
        '    Set fMDIWnd = Nothing
        'End If
        'Set fMDIWnd = frmSysCfgWnd
        'fMDIWnd.Show
        
        frmSysCfgWnd.SetCallBack Me, CALLBACK_MAIN_RETSYSCFGWND
        frmSysCfgWnd.Show vbModeless, Me 'システム設定画面
    End If
End Sub

' @(f)
'
' 機能      : 終了メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 終了問い合わせ画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuShutDown_Click()
    Unload fMainWnd '終了問い合わせ
End Sub

' @(f)
'
' 機能      : ＭＤＩフォーム破棄時処理
'
' 引き数    : ARG1 - キャンセルフラグ（戻り）
'             ARG2 - 破棄時モード（戻り）
'
' 返り値    :
'
' 機能説明  : ＭＤＩフォーム破棄時の処理。
'
' 備考      : システムの終了問い合わせ前の場合は、
'             システムの終了問い合わせ画面を表示する。
'             (コールバック有り)
'
Private Sub MDIForm_QueryUnload(CANCEL As Integer, UnloadMode As Integer)

    If m_shutDownFlag = False Then
        CANCEL = 1
        Dim fmessage As Object
        Set fmessage = New MessageYN
        fmessage.MsgText = "システムを終了します。" & vbCrLf & "よろしいですか？"
        fmessage.AutoDelete = True
        fmessage.SetCallBack Me, CALLBACK_MAIN_SHUTDOWN, True
            Do
                On Error Resume Next
                fmessage.Show vbModeless, Me
                If Err.Number = 0 Then
                    Exit Do
                End If
                DoEvents
            Loop
        Set fmessage = Nothing
    Else
        'frmViewMain.WindowState = 2
        'frmViewMain.Visible = True
        'frmViewMain.ZOrder 0
        fMainWnd.fMDIWnd.WindowState = 2
        fMainWnd.fMDIWnd.Visible = True
        fMainWnd.fMDIWnd.ZOrder 0
        
        Call MsgLog(conProcNum_MAIN, "*********************************************************")
        Call MsgLog(conProcNum_MAIN, "******************** ＭＡＩＮログ終了 ********************")
        Call MsgLog(conProcNum_MAIN, "*********************************************************")
        
        'ＭＡＩＮログファイルのクローズ
        Close #MainLogFileNumber
        MainLogFileNumber = Empty
        
        If Dir(App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt") <> "" Then
            Call Kill(App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt")
        End If
        
        MainModule.UnloadAll
    End If
    
End Sub

' @(f)
'
' 機能      : コールバック処理
'
' 引き数    : ARG1 - コールバック番号
'             ARG2 - コールバック戻り
'
' 返り値    :
'
' 機能説明  : コールバック番号と戻りに応じて、次処理を行う。
'
' 備考      :
'
Public Sub CallBackMessage(ByVal CallNo As Integer, ByVal Result As Integer)
    Dim cnt As Integer
    Dim nI As Integer
    Dim nJ As Integer
    Dim strImageFileName As String
    Dim strMIL_TITLE As String
    Dim strLBLINFO As String
    Dim bRet As Boolean
    Dim strWork As String
    
    Select Case CallNo
    
    'システム終了OK
    Case CALLBACK_MAIN_SHUTDOWN
        If Result = CALLBACK_ncResOK Then          'OK
            m_shutDownFlag = True
            On Error Resume Next
            On Error GoTo 0
            Unload Me
        End If
    
    'スラブ肌調査入力－スラブ選択画面からOK
    Case CALLBACK_MAIN_RETSKINSLBSELWND
        If Result = CALLBACK_ncResOK Then          'OK
            frmSkinScanWnd.Show vbModeless, Me 'スラブ肌調査入力画面へ
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETSKINSLBSELWND CANCEL"
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If
    
    'ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択画面からOK
    Case CALLBACK_MAIN_RETCOLORSLBSELWND
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND（処置ボタン実行）
            frmDirResWnd.SetCallBack Me, CALLBACK_MAIN_RETDIRRESWND1
            frmDirResWnd.Show vbModeless, Me '処置内容確認／結果登録画面へ移行
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND1
            frmColorScanWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力画面へ
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETCOLORSLBSELWND CANCEL"
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If

    'ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧画面からOK 2008/09/03
    Case CALLBACK_MAIN_RETCOLORSLBFAILWND
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND（処置ボタン実行）
            frmDirResWnd.SetCallBack Me, CALLBACK_MAIN_RETDIRRESWND2
            frmDirResWnd.Show vbModeless, Me '処置内容確認／結果登録画面へ移行
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND2
            frmColorScanWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力画面へ
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETCOLORSLBFAILWND CANCEL"
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If

    'スラブ肌調査入力画面からOK
    Case CALLBACK_MAIN_RETSKINSCANWND
        If Result = CALLBACK_ncResOK Then          'OK
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdSkinIn_Click '2008/08/04 スラブ肌調査入力開始ボタン実行（繰返し対応）
        Else                                        'CANCEL
            ' 20090115 modify by M.Aoyagi キー変更処理追加画面
            frmSkinSlbSelWnd.Show vbModeless, Me 'スラブ肌調査入力－スラブ選択画面へ
        End If
    
    'ｶﾗｰﾁｪｯｸ検査表入力画面からOK（1.スラブ選択シナリオ）
    Case CALLBACK_MAIN_RETCOLORSCANWND1
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND（保留ボタン実行）
            frmSlbFailScanWnd.SetCallBack Me, CALLBACK_MAIN_RETSLBFAILSCANWND1
            frmSlbFailScanWnd.Show vbModeless, Me 'スラブ異常報告書入力画面へ移行
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 ｶﾗｰﾁｪｯｸ検査表入力ボタン実行（繰返し対応）
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '特鋼
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            ' 20090115 modify by M.Aoyagi    キー変更処理追加画面
            frmColorSlbSelWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択画面へ
        End If
    
    'ｶﾗｰﾁｪｯｸ検査表入力画面からOK（2.異常報告一覧シナリオ）
    Case CALLBACK_MAIN_RETCOLORSCANWND2
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND（保留ボタン実行）
            frmSlbFailScanWnd.SetCallBack Me, CALLBACK_MAIN_RETSLBFAILSCANWND2
            frmSlbFailScanWnd.Show vbModeless, Me 'スラブ異常報告書入力画面へ移行
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 異常報告一覧ボタン実行（繰返し対応）
        Else                                        'CANCEL
            frmColorSlbFailWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧画面へ
        End If
    
    'スラブ異常報告書入力画面からOK（1.スラブ選択シナリオ）
    Case CALLBACK_MAIN_RETSLBFAILSCANWND1
        If Result = CALLBACK_ncResOK Then          'OK
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 ｶﾗｰﾁｪｯｸ検査表入力ボタン実行（繰返し対応）
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '特鋼
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND1
            frmColorScanWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力画面へ
        End If
    
    'スラブ異常報告書入力画面からOK（2.異常報告一覧シナリオ）
    Case CALLBACK_MAIN_RETSLBFAILSCANWND2
        If Result = CALLBACK_ncResOK Then          'OK
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 異常報告一覧ボタン実行（繰返し対応）
        Else                                        'CANCEL
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND2
            frmColorScanWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力画面へ
        End If
    
    'システム設定画面から戻り
    Case CALLBACK_MAIN_RETSYSCFGWND
        If Result = CALLBACK_ncResOK Then
            Call MsgLog(conProcNum_MAIN, "システム設定画面からOK")  'ガイダンス表示
            Call ReLoad
        End If
        
        Call MenuUnLock
        Call RefreshViewMain
        'If Not fMDIWnd Is Nothing Then
        '    Unload fMDIWnd
        '    Set fMDIWnd = Nothing
        'End If
        'Set fMDIWnd = frmViewMain
        'fMDIWnd.Show
    
    'ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択画面 -> スラブ異常処置指示／結果入力からOK
    Case CALLBACK_MAIN_RETDIRRESWND1 '（1.スラブ選択シナリオ）
        If Result = CALLBACK_ncResOK Then
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 ｶﾗｰﾁｪｯｸ検査表入力ボタン実行（繰返し対応）
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '特鋼
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            frmColorSlbSelWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択画面へ
        End If
    
    'ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧画面 -> スラブ異常処置指示／結果入力からOK
    Case CALLBACK_MAIN_RETDIRRESWND2 '（2.異常報告一覧シナリオ）
        If Result = CALLBACK_ncResOK Then
            '//入力データクリア実行
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 異常報告一覧ボタン実行（繰返し対応）
        Else                                        'CANCEL
            frmColorSlbFailWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧画面へ
        End If
    
    End Select
End Sub

' @(f)
'
' 機能      : データクリア処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : データクリア処理を行う。
'
' 備考      :
'
Private Sub InputDataClear()
    Dim nI As Integer

    Debug.Print "InputDataClear"

    APSlbCont.bProcessing = False
    APSlbCont.nListSelectedIndexP1 = 0
    APSlbCont.nSearchInputModeSelectedIndex = 0
    APSlbCont.strSearchInputSlbNumber = ""

    'スラブリストデータ初期化
    ReDim APSearchListSlbData(0)

    ReDim APSearchTmpSlbData(0)
    
End Sub

' @(f)
'
' 機能      : ＭＤＩフォーム破棄処理
'
' 引き数    : ARG1 - キャンセルフラグ
'
' 返り値    :
'
' 機能説明  : ＭＤＩフォーム破棄を行う。
'
' 備考      :
'
Public Sub MDIForm_Unload(CANCEL As Integer)
    'Unload LogoWnd
    Call SaveAPSysCfgDataSetting
    'Call SaveAPResDataSetting
    If Not fMDIWnd Is Nothing Then
        Unload fMDIWnd
        Set fMDIWnd = Nothing
    End If
End Sub

' @(f)
'
' 機能      : システム情報読込み
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : システム情報をレジストリから読込む。
'
' 備考      :
'
Public Sub LoadAPSysCfgDataSetting()
    Dim nI As Integer
    Dim nCnt As Integer
    
    APSysCfgData.nDEBUG_MODE = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDEBUG_MODE", conDefault_DEBUG_MODE)
    APSysCfgData.nDISP_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDISP_DEBUG", 0)
    APSysCfgData.nFILE_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nFILE_DEBUG", 0)
    APSysCfgData.nTR_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_SKIP", 0)
    APSysCfgData.nDB_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDB_SKIP", 0)
    APSysCfgData.nSOZAI_DB_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nSOZAI_DB_SKIP", 0)
    APSysCfgData.nSCAN_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nSCAN_SKIP", 0)
    APSysCfgData.nHOSTDATA_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_DEBUG", 0)
    APSysCfgData.nHOSTDATA_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_SKIP", 0)
    
    '************ COLORSYS
    APSysCfgData.DB_MYUSER_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_DSN", conDefault_DB_MYUSER_DSN)
    APSysCfgData.DB_MYUSER_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_UID", conDefault_DB_MYUSER_UID)
    APSysCfgData.DB_MYUSER_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_PWD", conDefault_DB_MYUSER_PWD)
    APSysCfgData.DB_MYCOMN_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_DSN", conDefault_DB_MYCOMN_DSN)
    APSysCfgData.DB_MYCOMN_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_UID", conDefault_DB_MYCOMN_UID)
    APSysCfgData.DB_MYCOMN_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_PWD", conDefault_DB_MYCOMN_PWD)
    APSysCfgData.DB_SOZAI_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_DSN", conDefault_DB_SOZAI_DSN)
    APSysCfgData.DB_SOZAI_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_UID", conDefault_DB_SOZAI_UID)
    APSysCfgData.DB_SOZAI_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_PWD", conDefault_DB_SOZAI_PWD)
    
    APSysCfgData.SHARES_SCNDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_SCNDIR", conDefault_SHARES_SCNDIR)
    APSysCfgData.SHARES_IMGDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_IMGDIR", conDefault_SHARES_IMGDIR)
    ' 20090116 add by M.Aoyagi
    APSysCfgData.SHARES_PDFDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_PDFDIR", conDefault_SHARES_PDFDIR)
    
    APSysCfgData.PHOTOIMG_DIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DIR", conDefault_PHOTOIMG_DIR)
    APSysCfgData.PHOTOIMG_DELCHK = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DELCHK", conDefault_PHOTOIMG_DELCHK)
    APSysCfgData.PHOTOIMG_ALLFILES = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_ALLFILES", conDefault_PHOTOIMG_ALLFILES)
    
    '2008/09/01 SystEx. A.K
    APSysCfgData.NowStaffName(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName0", "")
    APSysCfgData.NowStaffName(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName1", "")
    APSysCfgData.NowStaffName(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName2", "")
    
    APSysCfgData.NowNextProcess(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess0", "")
    APSysCfgData.NowNextProcess(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess1", "")
    APSysCfgData.NowNextProcess(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess2", "")
    
    '2008/09/03 カラー結果一覧のWEB-URL
    APSysCfgData.WEBURL_Color_Result = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result", conDefault_WEBURL_Color_Result)
    
    '2015/09/15 特鋼カラー結果一覧のWEB-URL
    APSysCfgData.WEBURL_Color_Result_Tok = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result_Tok", conDefault_WEBURL_Color_Result_Tok)
    '************
    
    ' ソケット通信対応
    APSysCfgData.HOST_IP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "HOST_IP", conDefault_HOST_IP) 'ビジコンIP
    APSysCfgData.nHOST_PORT = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_PORT", conDefault_nHOST_PORT) 'ビジコンPORT
    APSysCfgData.nHOST_TOUT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT0", conDefault_nHOST_TOUT0) 'ビジコン通信タイムアウト（全体）
    APSysCfgData.nHOST_TOUT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT1", conDefault_nHOST_TOUT1) 'ビジコン通信タイムアウト（オープン時）
    APSysCfgData.nHOST_TOUT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT2", conDefault_nHOST_TOUT2) 'ビジコン通信タイムアウト（データ通信）
    APSysCfgData.nHOST_RETRY = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_RETRY", conDefault_nHOST_RETRY) '通信リトライ回数

    APSysCfgData.TR_IP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "TR_IP", conDefault_TR_IP) 'ＦＴＰ通信ＩＰアドレス
    APSysCfgData.nTR_PORT = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_PORT", conDefault_nTR_PORT) 'ＦＴＰ通信ポート番号
    APSysCfgData.nTR_TOUT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT0", conDefault_nTR_TOUT0) 'ＦＴＰ通信タイムアウト（全体）
    APSysCfgData.nTR_TOUT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT1", conDefault_nTR_TOUT1) 'ＦＴＰ通信タイムアウト（オープン時）
    APSysCfgData.nTR_TOUT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT2", conDefault_nTR_TOUT2) 'ＦＴＰ通信タイムアウト（データ通信）
    APSysCfgData.nTR_RETRY = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_RETRY", conDefault_nTR_RETRY) '通信リトライ回数
    
    APSysCfgData.nIMAGE_SIZE(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE0", conDefault_nIMAGE_SIZE0)
    APSysCfgData.nIMAGE_SIZE(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE1", conDefault_nIMAGE_SIZE1)
    APSysCfgData.nIMAGE_SIZE(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE2", conDefault_nIMAGE_SIZE2)
    
    APSysCfgData.nIMAGE_ROTATE(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE0", conDefault_nIMAGE_ROTATE0)
    APSysCfgData.nIMAGE_ROTATE(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE1", conDefault_nIMAGE_ROTATE1)
    APSysCfgData.nIMAGE_ROTATE(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE2", conDefault_nIMAGE_ROTATE2)
    
    APSysCfgData.nIMAGE_LEFT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT0", conDefault_nIMAGE_LEFT0)
    APSysCfgData.nIMAGE_TOP(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP0", conDefault_nIMAGE_TOP0)
    APSysCfgData.nIMAGE_WIDTH(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH0", conDefault_nIMAGE_WIDTH0)
    APSysCfgData.nIMAGE_HEIGHT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT0", conDefault_nIMAGE_HEIGHT0)
    
    APSysCfgData.nIMAGE_LEFT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT1", conDefault_nIMAGE_LEFT1)
    APSysCfgData.nIMAGE_TOP(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP1", conDefault_nIMAGE_TOP1)
    APSysCfgData.nIMAGE_WIDTH(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH1", conDefault_nIMAGE_WIDTH1)
    APSysCfgData.nIMAGE_HEIGHT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT1", conDefault_nIMAGE_HEIGHT1)
    
    APSysCfgData.nIMAGE_LEFT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT2", conDefault_nIMAGE_LEFT2)
    APSysCfgData.nIMAGE_TOP(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP2", conDefault_nIMAGE_TOP2)
    APSysCfgData.nIMAGE_WIDTH(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH2", conDefault_nIMAGE_WIDTH2)
    APSysCfgData.nIMAGE_HEIGHT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT2", conDefault_nIMAGE_HEIGHT2)
    
    If IsDEBUG("SCAN") Then
        APSysCfgData.nIMAGE_LEFT(0) = conDefault_nIMAGE_DEB_LEFT0
        APSysCfgData.nIMAGE_TOP(0) = conDefault_nIMAGE_DEB_TOP0
        APSysCfgData.nIMAGE_WIDTH(0) = conDefault_nIMAGE_DEB_WIDTH0
        APSysCfgData.nIMAGE_HEIGHT(0) = conDefault_nIMAGE_DEB_HEIGHT0
        
        APSysCfgData.nIMAGE_LEFT(1) = conDefault_nIMAGE_DEB_LEFT1
        APSysCfgData.nIMAGE_TOP(1) = conDefault_nIMAGE_DEB_TOP1
        APSysCfgData.nIMAGE_WIDTH(1) = conDefault_nIMAGE_DEB_WIDTH1
        APSysCfgData.nIMAGE_HEIGHT(1) = conDefault_nIMAGE_DEB_HEIGHT1
    
        APSysCfgData.nIMAGE_LEFT(2) = conDefault_nIMAGE_DEB_LEFT2
        APSysCfgData.nIMAGE_TOP(2) = conDefault_nIMAGE_DEB_TOP2
        APSysCfgData.nIMAGE_WIDTH(2) = conDefault_nIMAGE_DEB_WIDTH2
        APSysCfgData.nIMAGE_HEIGHT(2) = conDefault_nIMAGE_DEB_HEIGHT2
    End If
    
    '次工程マスター読込(SKIN)
    nCnt = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", 0)
    ReDim APNextProcDataSkin(0)
    For nI = 1 To nCnt
        APNextProcDataSkin(nI - 1).inp_NextProc = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), "")
        ReDim Preserve APNextProcDataSkin(UBound(APNextProcDataSkin) + 1)
    Next nI

    If UBound(APNextProcDataSkin) = 0 Then
        ReDim APNextProcDataSkin(7)
        APNextProcDataSkin(0).inp_NextProc = ""
        APNextProcDataSkin(1).inp_NextProc = "受入送り"
        APNextProcDataSkin(2).inp_NextProc = "SLG研削"
        APNextProcDataSkin(3).inp_NextProc = "特鋼研削"
        APNextProcDataSkin(4).inp_NextProc = "SKY切断"
        APNextProcDataSkin(5).inp_NextProc = "ソーキング"
        APNextProcDataSkin(6).inp_NextProc = "指示書にて別途指示(保留扱い)"
    End If

    '次工程マスター読込(COLOR)
    nCnt = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", 0)
    ReDim APNextProcDataColor(0)
    For nI = 1 To nCnt
        APNextProcDataColor(nI - 1).inp_NextProc = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), "")
        ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
    Next nI

    If UBound(APNextProcDataColor) = 0 Then
        ReDim APNextProcDataColor(6)
        APNextProcDataColor(0).inp_NextProc = ""
        APNextProcDataColor(1).inp_NextProc = "受入送り"
        APNextProcDataColor(2).inp_NextProc = "部分取り再カラー"
        APNextProcDataColor(3).inp_NextProc = "＃６０仕上研削"
        APNextProcDataColor(4).inp_NextProc = "SKY切断"
        APNextProcDataColor(5).inp_NextProc = "指示書にて別途指示(保留扱い)"
    End If

    ''面欠陥リスト情報（スラブ肌）
    ReDim APFaultFaceSkin(13)
    APFaultFaceSkin(0).strCode = ""
    APFaultFaceSkin(0).strName = ""
    APFaultFaceSkin(1).strCode = "ﾀﾃﾜﾚ"
    APFaultFaceSkin(1).strName = "縦割れ"
    APFaultFaceSkin(2).strCode = "ﾖｺﾜﾚ"
    APFaultFaceSkin(2).strName = "横割れ"
    APFaultFaceSkin(3).strCode = "ｺｰﾅｰ"
'    APFaultFaceSkin(3).strName = "コーナー横割れ"
    APFaultFaceSkin(3).strName = "ｺｰﾅｰ横割れ"
    APFaultFaceSkin(4).strCode = "ﾉﾛｶﾐ"
    APFaultFaceSkin(4).strName = "ノロカミ"
    APFaultFaceSkin(5).strCode = "ｽﾃｯｷ"
'    APFaultFaceSkin(5).strName = "スティッキング"
    APFaultFaceSkin(5).strName = "ｽﾃｨｯｷﾝｸﾞ"
    APFaultFaceSkin(6).strCode = "ﾌﾞﾘｰ"
'    APFaultFaceSkin(6).strName = "ブリーディング"
    APFaultFaceSkin(6).strName = "ﾌﾞﾘｰﾃﾞｨﾝｸﾞ"
    APFaultFaceSkin(7).strCode = "ﾃﾞﾌﾟ"
'    APFaultFaceSkin(7).strName = "デプレッション"
    APFaultFaceSkin(7).strName = "ﾃﾞﾌﾟﾚｯｼｮﾝ"
    APFaultFaceSkin(8).strCode = "ﾆｼﾞｭ"
    APFaultFaceSkin(8).strName = "2重肌"
    APFaultFaceSkin(9).strCode = "ﾀﾞﾝﾂ"
    APFaultFaceSkin(9).strName = "段継"
    APFaultFaceSkin(10).strCode = "ﾀﾃﾍｺ"
    APFaultFaceSkin(10).strName = "縦凹み"
    APFaultFaceSkin(11).strCode = "ｻﾞｸﾜ"
    APFaultFaceSkin(11).strName = "ザク割れ"
    APFaultFaceSkin(12).strCode = "ﾋｹﾞﾜ"
    APFaultFaceSkin(12).strName = "ひげ割れ"

    ''内部欠陥リスト情報（スラブ肌）
    ReDim APFaultInsideSkin(5)
    APFaultInsideSkin(0).strCode = ""
    APFaultInsideSkin(0).strName = ""
    APFaultInsideSkin(1).strCode = "ﾅｲﾌﾞ"
    APFaultInsideSkin(1).strName = "内部割れ"
    APFaultInsideSkin(2).strCode = "ﾋｹﾞﾜ"
    APFaultInsideSkin(2).strName = "ひげ割れ"
    APFaultInsideSkin(3).strCode = "ｾﾝﾀｰ"
'    APFaultInsideSkin(3).strName = "センターポロシティor中心偏析"
    APFaultInsideSkin(3).strName = "ｾﾝﾀｰﾎﾟﾛｼﾃｨ"
    APFaultInsideSkin(4).strCode = "ﾜﾆｸﾁ"
'    APFaultInsideSkin(4).strName = "ワニくち割れ"
    APFaultInsideSkin(4).strName = "ﾜﾆくち割れ"

    ''面欠陥リスト情報（カラーチェック）
    ReDim APFaultFaceColor(11)
    APFaultFaceColor(0).strCode = ""
    APFaultFaceColor(0).strName = ""
    APFaultFaceColor(1).strCode = "ﾀﾃﾜﾚ"
    APFaultFaceColor(1).strName = "縦割れ"
    APFaultFaceColor(2).strCode = "ﾖｺﾜﾚ"
    APFaultFaceColor(2).strName = "横割れ"
    APFaultFaceColor(3).strCode = "ｺｰﾅｰ"
    'APFaultFaceColor(3).strName = "コーナー横割れ"
    APFaultFaceColor(3).strName = "ｺｰﾅｰ横割れ"
    APFaultFaceColor(4).strCode = "ﾋﾟﾝﾎ"
    APFaultFaceColor(4).strName = "ピンホール"
    APFaultFaceColor(5).strCode = "ﾓﾔ"
    APFaultFaceColor(5).strName = "モヤ疵"
    APFaultFaceColor(6).strCode = "ｸﾛｶﾜ"
    APFaultFaceColor(6).strName = "黒皮残り"
    APFaultFaceColor(7).strCode = "ﾘｭｳｶ"
    APFaultFaceColor(7).strName = "粒界酸化"
    APFaultFaceColor(8).strCode = "ｻﾞｸﾜ"
    APFaultFaceColor(8).strName = "ザク割れ"
    APFaultFaceColor(9).strCode = "ﾋｹﾞﾜ"
    APFaultFaceColor(9).strName = "ひげ割れ"
    '2016/04/20 - TAI - S
    APFaultFaceColor(10).strCode = "ﾌｶﾎﾞ"
    APFaultFaceColor(10).strName = "深堀り"
    '2016/04/20 - TAI - E

    ''処置状態リスト
    ReDim APDirRes_Stat(2)
    APDirRes_Stat(0).inp_DirRes_StatCode = ""
    APDirRes_Stat(0).inp_DirRes_Stat = ""
    
    APDirRes_Stat(1).inp_DirRes_StatCode = "1"
    APDirRes_Stat(1).inp_DirRes_Stat = "1:完了"


    ''処置結果リスト
    ReDim APDirRes_Res(2)
    APDirRes_Res(0).inp_DirRes_ResCode = ""
    APDirRes_Res(0).inp_DirRes_Res = ""
    
    APDirRes_Res(1).inp_DirRes_ResCode = "1"
    APDirRes_Res(1).inp_DirRes_Res = "1:不適合有り"

End Sub

' @(f)
'
' 機能      : システム情報保存
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : システム情報をレジストリに保存する。
'
' 備考      :
'
Public Sub SaveAPSysCfgDataSetting()
    Dim nI As Integer
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDEBUG_MODE", APSysCfgData.nDEBUG_MODE
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDISP_DEBUG", APSysCfgData.nDISP_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nFILE_DEBUG", APSysCfgData.nFILE_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_SKIP", APSysCfgData.nTR_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDB_SKIP", APSysCfgData.nDB_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nSOZAI_DB_SKIP", APSysCfgData.nSOZAI_DB_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nSCAN_SKIP", APSysCfgData.nSCAN_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_DEBUG", APSysCfgData.nHOSTDATA_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_SKIP", APSysCfgData.nHOSTDATA_SKIP
    
    '************ COLORSYS
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_IP", APSysCfgData.DB_MYUSER_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_UID", APSysCfgData.DB_MYUSER_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_PWD", APSysCfgData.DB_MYUSER_PWD
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_IP", APSysCfgData.DB_MYCOMN_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_UID", APSysCfgData.DB_MYCOMN_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_PWD", APSysCfgData.DB_MYCOMN_PWD
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_IP", APSysCfgData.DB_SOZAI_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_UID", APSysCfgData.DB_SOZAI_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_PWD", APSysCfgData.DB_SOZAI_PWD
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_SCNDIR", APSysCfgData.SHARES_SCNDIR
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_IMGDIR", APSysCfgData.SHARES_IMGDIR
    ' 20090116 add by M.Aoyagi
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_PDFDIR", APSysCfgData.SHARES_PDFDIR
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DIR", APSysCfgData.PHOTOIMG_DIR
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DELCHK", APSysCfgData.PHOTOIMG_DELCHK
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_ALLFILES", APSysCfgData.PHOTOIMG_ALLFILES
    
    '2008/09/01 SystEx. A.K
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName0", APSysCfgData.NowStaffName(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName1", APSysCfgData.NowStaffName(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName2", APSysCfgData.NowStaffName(2)
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess0", APSysCfgData.NowNextProcess(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess1", APSysCfgData.NowNextProcess(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess2", APSysCfgData.NowNextProcess(2)
    
    '2008/09/03 カラー結果一覧のWEB-URL
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result", APSysCfgData.WEBURL_Color_Result
    
    '2015/09/15 カラー結果一覧のWEB-URL
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result_Tok", APSysCfgData.WEBURL_Color_Result_Tok
    '************
    
    ' ソケット通信対応
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "HOST_IP", APSysCfgData.HOST_IP 'ビジコンIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_PORT", APSysCfgData.nHOST_PORT 'ビジコンPORT
     For nI = 0 To 2
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT" & CStr(nI), APSysCfgData.nHOST_TOUT(nI) 'ビジコン通信タイムアウト
    Next nI
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_RETRY", APSysCfgData.nHOST_RETRY '通信リトライ回数
 
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "TR_IP", APSysCfgData.TR_IP 'ＦＴＰ通信ＩＰアドレス
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_PORT", APSysCfgData.nTR_PORT 'ＦＴＰ通信ポート番号
    For nI = 0 To 2
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT" & CStr(nI), APSysCfgData.nTR_TOUT(nI) 'ＦＴＰ通信タイムアウト
    Next nI
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_RETRY", APSysCfgData.nTR_RETRY '通信リトライ回数
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE0", APSysCfgData.nIMAGE_SIZE(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE1", APSysCfgData.nIMAGE_SIZE(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE2", APSysCfgData.nIMAGE_SIZE(2)
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE0", APSysCfgData.nIMAGE_ROTATE(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE1", APSysCfgData.nIMAGE_ROTATE(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE2", APSysCfgData.nIMAGE_ROTATE(2)

    If IsDEBUG("SCAN") = False Then
        For nI = 0 To 2
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT" & CStr(nI), APSysCfgData.nIMAGE_LEFT(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP" & CStr(nI), APSysCfgData.nIMAGE_TOP(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH" & CStr(nI), APSysCfgData.nIMAGE_WIDTH(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT" & CStr(nI), APSysCfgData.nIMAGE_HEIGHT(nI)
        Next nI
    End If

    '次工程マスター保存(SKIN)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", UBound(APNextProcDataSkin)
    For nI = 1 To UBound(APNextProcDataSkin)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), APNextProcDataSkin(nI - 1).inp_NextProc
    Next nI

    '次工程マスター保存(COLOR)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
    For nI = 1 To UBound(APNextProcDataColor)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
    Next nI

End Sub

' @(f)
'
' 機能      : メイン画面再描画要求
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : メイン画面の再描画を要求する。
'
' 備考      :
'
Public Sub ReqRefreshViewMain()
    Call RefreshViewMain
End Sub

' @(f)
'
' 機能      : メイン画面再描画
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : メイン画面の再描画を要求する。
'
' 備考      : ＭＤＩ子フォームが存在する時のみ要求する。
'
Private Sub RefreshViewMain()
    If Not fMDIWnd Is Nothing Then
        If fMDIWnd.Name = "frmViewMain" Then
            Call fMainWnd.fMDIWnd.RefreshViewMain
        End If
    End If
End Sub

'2016/04/20 - TAI - S

' @(f)
'
' 機能      : 特鋼カラー結果一覧(WEB)ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 特鋼カラー結果一覧(WEB)ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdWEBURL_Color_Result_Tok_Click()
    Call mnuWEBURL_Color_Result_Tok_Click 'カラー結果一覧(WEB)メニュー
End Sub


' @(f)
'
' 機能      : 特鋼カラー結果一覧(WEB)メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 特鋼カラー結果一覧(WEB)をIEで開く。
'
' 備考      :COLORSYS
'
Private Sub mnuWEBURL_Color_Result_Tok_Click()
    Dim RetVal
    RetVal = Shell(APSysCfgData.WEBURL_Color_Result_Tok, 3)
End Sub

' @(f)
'
' 機能      : 特鋼ｶﾗｰﾁｪｯｸ検査表入力ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 特鋼ｶﾗｰﾁｪｯｸ検査表入力ボタンでメニューを開く。
'
' 備考      :COLORSYS
'
Private Sub cmdColorIn_Tok_Click()
    Call mnuColorIn_Tok_Click 'ｶﾗｰﾁｪｯｸ検査表入力メニュー
End Sub


' @(f)
'
' 機能      : 特鋼ｶﾗｰﾁｪｯｸ検査表入力メニュー
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 特鋼ｶﾗｰﾁｪｯｸ検査表入力画面を開く。
'
' 備考      :COLORSYS
'
Private Sub mnuColorIn_Tok_Click()
    Call MenuLock("mnuColorIn_Tok")
    
    '2016/04/20 - TAI - S
    '作業場所を"特鋼"にする
    works_sky_tok = WORKS_TOK
    '2016/04/20 - TAI - E

    'スラブ検索リストクリア
    ReDim APSearchListSlbData(0)
    
    frmColorSlbSelWnd.Show vbModeless, Me 'ｶﾗｰﾁｪｯｸ検査表入力用－スラブ選択画面
End Sub



'2016/04/20 - TAI - E

