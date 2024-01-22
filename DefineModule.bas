Attribute VB_Name = "DefineModule"
' @(h) DefineModule.Bas                ver 1.00 ( '08 SEC Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　システムメイン（変数／定数）定義モジュール
' 　本モジュールはシステムで使用する変数／定数を定義する
' 　ためのものである。

Option Explicit

''プロセス番号
Public Const conProcNum_MAIN As Integer = 1 ''メイン処理
Public Const conProcNum_BSCONT As Integer = 2 ''ビジコン通信処理
Public Const conProcNum_TRCONT As Integer = 3 ''ソケット通信処理
Public Const conProcNum_MAINTENANCE As Integer = 4 ''メンテナンス処理
'Public Const conProcNum_SELPARSLB As Integer = 5 ''親スラブ選択処理
Public Const conProcNum_WINSOCKCONT As Integer = 6 ''Winsock イベントLOG

''グリットコントロール定義
Public Const FlexAlignCenter As Long = 2 ''表示位置センター

''コールバック
Public Const CALLBACK_ncResOK = 1 ''応答ＯＫ
Public Const CALLBACK_ncResCANCEL = 2 ''応答キャンセル
Public Const CALLBACK_ncResSKIP = 3 ''応答スキップ
Public Const CALLBACK_ncResEXTEND = 4 ''応答拡張

Public Const CALLBACK_MAIN_SHUTDOWN = 1 ''メイン－システム終了
Public Const CALLBACK_MAIN_RETSKINSCANWND = 2 ''メイン－スラブ肌調査入力
Public Const CALLBACK_MAIN_RETCOLORSCANWND1 = 3 ''メイン－ｶﾗｰﾁｪｯｸ検査表入力 -> ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択へ戻り
Public Const CALLBACK_MAIN_RETCOLORSCANWND2 = 4 ''メイン－ｶﾗｰﾁｪｯｸ検査表入力 -> ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧へ戻り
Public Const CALLBACK_MAIN_RETSLBFAILSCANWND1 = 5 ''メイン－スラブ異常報告書入力 -> ｶﾗｰﾁｪｯｸ検査表入力 -> ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択へ戻り
Public Const CALLBACK_MAIN_RETSLBFAILSCANWND2 = 6 ''メイン－スラブ異常報告書入力 -> ｶﾗｰﾁｪｯｸ検査表入力 -> ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧へ戻り

Public Const CALLBACK_MAIN_RETSKINSLBSELWND = 7 ''メイン－スラブ肌調査入力－スラブ選択
Public Const CALLBACK_MAIN_RETCOLORSLBSELWND = 8 ''メイン－ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択
Public Const CALLBACK_MAIN_RETCOLORSLBFAILWND = 9 ''メイン－ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧

Public Const CALLBACK_MAIN_RETSYSCFGWND = 10 ''メイン－システム設定
Public Const CALLBACK_MAIN_RETDIRRESWND1 = 11 ''スラブ異常処置指示／結果入力 -> ｶﾗｰﾁｪｯｸ検査表入力－スラブ選択へ戻り
Public Const CALLBACK_MAIN_RETDIRRESWND2 = 12 ''スラブ異常処置指示／結果入力 -> ｶﾗｰﾁｪｯｸ検査表入力－異常報告一覧へ戻り

Public Const CALLBACK_RES_DBSNDDATA_SKIN = 20 ''SKIN－ＤＢ登録
Public Const CALLBACK_RES_SKINDATA_DBDEL_REQ = 21 ''SKIN－データ削除問合せ

Public Const CALLBACK_RES_DBSNDDATA_COLOR = 22 ''COLOR－ＤＢ登録
Public Const CALLBACK_RES_DBSNDDATA_SLBFAIL = 23 ''SLBFAIL－ＤＢ登録
Public Const CALLBACK_RES_DBSNDDATA_DIRRES = 24 ''DIRRES－ＤＢ登録
Public Const CALLBACK_RES_HOSTSNDDATA_DIRRES = 25 ''DIRRES－HOST登録
Public Const CALLBACK_RES_COLORDATA_DBDEL_REQ = 26 ''COLOR－データ削除問合せ
Public Const CALLBACK_RES_COLORDATA_HOSTDEL_REQ = 27 ''COLOR－データ削除問合せ（ビジコン削除シナリオ）
Public Const CALLBACK_RES_COLORDATA_HOSTDEL_REQ2 = 28 ''COLOR－データ削除問合せ(ビジコン削除シナリオ⇒ＤＢ削除シナリオ）

Public Const CALLBACK_RES_DIRPRN_REQ = 30 '指示印刷問合せ 2008/09/04
Public Const CALLBACK_RES_DIRPRN_SND = 31 '指示印刷要求結果 2008/09/04

Public Const CALLBACK_RES_STATECHANGE_SKIN = 50  'SKIN－状態変更問合せ 2009/01/28
Public Const CALLBACK_RES_STATECHANGE_COLOR = 50 'SKIN－状態変更問合せ 2009/01/28

Public Const CALLBACK_OPREGWND = 100 ''操作員登録
Public Const CALLBACK_NEXTPROCWND = 101 ''次工程登録
Public Const CALLBACK_FULLSCANIMAGEWND = 102 ''フルスキャンイメージ表示画面
Public Const CALLBACK_PHOTOIMGUPWND = 103 ''写真添付画面
Public Const CALLBACK_PHOTOIMG_DELETE = 104 '削除'
Public Const CALLBACK_PHOTOIMG_UPLOAD = 105 'アップロード

Public Const CALLBACK_HOSTSEND = 110 ''ホスト送信
'Public Const CALLBACK_HOSTSEND_RESDELETE = 111 ''ホスト送信－実績削除
'Public Const CALLBACK_HOSTSEND_SLBDELETE = 112 ''ホスト送信－スラブ削除
'Public Const CALLBACK_HOSTSEND_SLBDELETE2 = 113 ''ホスト送信－スラブ削除２
Public Const CALLBACK_HOSTSEND_QUERY = 114 ''ホスト送信－スラブ情報問い合わせ

Public Const CALLBACK_TRSEND = 115 ''TR送信

Public Const CALLBACK_USEIMGDATA = 200 ''イメージデータ使用（ＤＢ存在）
Public Const CALLBACK_GETIMGDATA = 201 ''イメージデータ取得（読込）

''レジストリ用定義
''アプリケーション名
Public Const conReg_APPNAME As String = "COLORSYS" ''本システム名

''セクション名
Public Const conReg_APSYSCFG As String = "SYSCFG DATA" ''システム設定情報
'Public Const conReg_APSLB As String = "SLB DATA" ''スラブ情報
Public Const conReg_APRESULT As String = "RESULT DATA" ''実績入力情報

''レジストリ初期値

''デバッグ用
Public Const conDefault_DEBUG_MODE As Integer = 1 ''デバッグＯＮ

''DB
Public Const conDefault_DBConnect_MYUSER As Integer = 0
Public Const conDefault_DBConnect_MYCOMN As Integer = 1
Public Const conDefault_DBConnect_SOZAI As Integer = 2

Public Const conDefault_DB_MYUSER_DSN As String = "ORAM_COL" ''データソース名
Public Const conDefault_DB_MYUSER_UID As String = "UCOL" ''ユーザーＩＤ
Public Const conDefault_DB_MYUSER_PWD As String = "UCOL" ''パスワード
'Public Const conDefault_DB_MYUSER_UID As String = "UCOLTEST" ''テスト　ユーザーＩＤ
'Public Const conDefault_DB_MYUSER_PWD As String = "UCOLTEST" ''テスト　パスワード

Public Const conDefault_DB_MYCOMN_DSN As String = "ORAM_COL" ''データソース名
Public Const conDefault_DB_MYCOMN_UID As String = "NYKCOMN" ''ユーザーＩＤ
Public Const conDefault_DB_MYCOMN_PWD As String = "NYKCOMN" ''パスワード
'Public Const conDefault_DB_MYCOMN_UID As String = "NYKCOMNTEST" ''テスト　ユーザーＩＤ
'Public Const conDefault_DB_MYCOMN_PWD As String = "NYKCOMNTEST" ''テスト　パスワード

Public Const conDefault_DB_SOZAI_DSN As String = "ORAM_SOZAI" ''データソース名
Public Const conDefault_DB_SOZAI_UID As String = "JISSEKI1" ''ユーザーＩＤ
Public Const conDefault_DB_SOZAI_PWD As String = "JISSEKI1" ''パスワード

Public Const conDefault_SHARES_SCNDIR As String = "\\COLDBSRV\shares\SCAN" ''スキャナーイメージファイル保存先パス名
Public Const conDefault_SHARES_IMGDIR As String = "\\COLDBSRV\shares\IMG" ''写真イメージファイル保存先パス名
Public Const conDefault_SHARES_PDFDIR As String = "\\COLDBSRV\shares\PDF" '' 20090124 add by M.Aoyagi    PDFファイル保存先パス名

Public Const conDefault_DEFINE_SCNDIR As String = "$SCNDIR"
Public Const conDefault_DEFINE_PDFDIR As String = "$PDFDIR"               '' 20090124 add by M.Aoyagi    状態変更時使用

Public Const conDefault_PHOTOIMG_DIR As String = "c:"       ''写真ローカルパス
Public Const conDefault_PHOTOIMG_DELCHK As Integer = 0      ''写真コピー元削除フラグ
Public Const conDefault_PHOTOIMG_ALLFILES As Integer = 0    ''写真全てのファイル指定フラグ

'-------------------------------------
'Public Const conDefault_TRN_MSG_NO As String = "" ''トランザクションメッセージ番号
'-------------------------------------
'2008-04-28 TRTS0012 Define
Public Const conDefault_slb_no As String = ""
Public Const conDefault_slb_chno As String = ""
Public Const conDefault_slb_aino As String = ""
Public Const conDefault_slb_stat As String = ""

Public Const conDefault_slb_col_cnt As String = ""

Public Const conDefault_slb_ccno As String = ""
Public Const conDefault_slb_zkai_dte As String = ""
Public Const conDefault_slb_ksh As String = ""
Public Const conDefault_slb_typ As String = ""
Public Const conDefault_slb_uksk As String = ""
Public Const conDefault_slb_wei As String = ""
Public Const conDefault_slb_lngth As String = ""
Public Const conDefault_slb_wdth As String = ""
Public Const conDefault_slb_thkns As String = ""
Public Const conDefault_slb_nxt_prcs As String = ""
Public Const conDefault_slb_cmt1 As String = ""
Public Const conDefault_slb_cmt2 As String = ""

Public Const conDefault_slb_fault_cd_e_s1 As String = ""
Public Const conDefault_slb_fault_cd_e_s2 As String = ""
Public Const conDefault_slb_fault_cd_e_s3 As String = ""
Public Const conDefault_slb_fault_e_s1 As String = ""
Public Const conDefault_slb_fault_e_s2 As String = ""
Public Const conDefault_slb_fault_e_s3 As String = ""
Public Const conDefault_slb_fault_e_n1 As String = ""
Public Const conDefault_slb_fault_e_n2 As String = ""
Public Const conDefault_slb_fault_e_n3 As String = ""

Public Const conDefault_slb_fault_cd_w_s1 As String = ""
Public Const conDefault_slb_fault_cd_w_s2 As String = ""
Public Const conDefault_slb_fault_cd_w_s3 As String = ""
Public Const conDefault_slb_fault_w_s1 As String = ""
Public Const conDefault_slb_fault_w_s2 As String = ""
Public Const conDefault_slb_fault_w_s3 As String = ""
Public Const conDefault_slb_fault_w_n1 As String = ""
Public Const conDefault_slb_fault_w_n2 As String = ""
Public Const conDefault_slb_fault_w_n3 As String = ""

Public Const conDefault_slb_fault_cd_s_s1 As String = ""
Public Const conDefault_slb_fault_cd_s_s2 As String = ""
Public Const conDefault_slb_fault_cd_s_s3 As String = ""
Public Const conDefault_slb_fault_s_s1 As String = ""
Public Const conDefault_slb_fault_s_s2 As String = ""
Public Const conDefault_slb_fault_s_s3 As String = ""
Public Const conDefault_slb_fault_s_n1 As String = ""
Public Const conDefault_slb_fault_s_n2 As String = ""
Public Const conDefault_slb_fault_s_n3 As String = ""

Public Const conDefault_slb_fault_cd_n_s1 As String = ""
Public Const conDefault_slb_fault_cd_n_s2 As String = ""
Public Const conDefault_slb_fault_cd_n_s3 As String = ""
Public Const conDefault_slb_fault_n_s1 As String = ""
Public Const conDefault_slb_fault_n_s2 As String = ""
Public Const conDefault_slb_fault_n_s3 As String = ""
Public Const conDefault_slb_fault_n_n1 As String = ""
Public Const conDefault_slb_fault_n_n2 As String = ""
Public Const conDefault_slb_fault_n_n3 As String = ""

Public Const conDefault_slb_fault_cd_bs_s As String = ""
Public Const conDefault_slb_fault_cd_bm_s As String = ""
Public Const conDefault_slb_fault_cd_bn_s As String = ""
Public Const conDefault_slb_fault_bs_s As String = ""
Public Const conDefault_slb_fault_bm_s As String = ""
Public Const conDefault_slb_fault_bn_s As String = ""
Public Const conDefault_slb_fault_bs_n As String = ""
Public Const conDefault_slb_fault_bm_n As String = ""
Public Const conDefault_slb_fault_bn_n As String = ""

Public Const conDefault_slb_fault_cd_ts_s As String = ""
Public Const conDefault_slb_fault_cd_tm_s As String = ""
Public Const conDefault_slb_fault_cd_tn_s As String = ""
Public Const conDefault_slb_fault_ts_s As String = ""
Public Const conDefault_slb_fault_tm_s As String = ""
Public Const conDefault_slb_fault_tn_s As String = ""
Public Const conDefault_slb_fault_ts_n As String = ""
Public Const conDefault_slb_fault_tm_n As String = ""
Public Const conDefault_slb_fault_tn_n As String = ""

Public Const conDefault_slb_wrt_nme As String = ""

Public Const conDefault_host_send As String = ""
Public Const conDefault_host_wrt_dte As String = ""
Public Const conDefault_host_wrt_tme As String = ""

Public Const conDefault_sys_wrt_dte As String = ""
Public Const conDefault_sys_wrt_tme As String = ""
Public Const conDefault_sys_rwrt_dte As String = ""
Public Const conDefault_sys_rwrt_tme As String = ""
Public Const conDefault_sys_acs_pros As String = ""
Public Const conDefault_sys_acs_enum As String = ""

'Public Const conDefault_LINE_NAME As String = "col"
'Public Const conDefault_LINE_NUMBER As String = ""

Public Const conDefault_nSEARCH_MAX0 As Integer = 9999 ''スラブＮｏ．
'Public Const conDefault_nSEARCH_MAX1 As Integer = 1 ''直近
'Public Const conDefault_nSEARCH_MAX2 As Integer = 10 ''直近過去
'Public Const conDefault_nSEARCH_MAX3 As Integer = 1 ''強制入力
'Public Const conDefault_nSEARCH_RANGE As Integer = 90 ''検索有効範囲　過去？日

'Public Const conDefault_HOST_NAME As String = "QVCB89" ''ホスト名称
Public Const conDefault_HOST_IP As String = "172.18.192.19" ''ホスト IP
Public Const conDefault_nHOST_PORT As String = "15025" ''ホストPort
Public Const conDefault_nHOST_TOUT0 As Long = 35 ''ホスト通信タイムアウト（全体）（秒）
Public Const conDefault_nHOST_TOUT1 As Long = 5 ''ホスト通信タイムアウト（オープン時）（秒）
Public Const conDefault_nHOST_TOUT2 As Long = 10 ''ホスト通信タイムアウト（データ通信）（秒）
Public Const conDefault_nHOST_RETRY As Integer = 2 ''ホスト通信リトライ回数

Public Const conDefault_TR_IP As String = "172.18.128.254" ''通信サーバーＩＰアドレス
Public Const conDefault_nTR_PORT As Integer = 15032 ''通信サーバーポート番号
Public Const conDefault_nTR_TOUT0 As Long = 35 ''通信サーバータイムアウト（全体）（秒）
Public Const conDefault_nTR_TOUT1 As Long = 5 ''通信サーバータイムアウト（オープン時）（秒）
Public Const conDefault_nTR_TOUT2 As Long = 10 ''通信サーバータイムアウト（データ通信）（秒）
Public Const conDefault_nTR_RETRY As Integer = 2 ''通信サーバーリトライ回数

Public Const conDefault_nIMAGE_SIZE0 As Integer = 30 ''イメージ表示率
Public Const conDefault_nIMAGE_SIZE1 As Integer = 30 ''イメージ表示率
Public Const conDefault_nIMAGE_SIZE2 As Integer = 30 ''イメージ表示率
Public Const conDefault_nIMAGE_ROTATE0 As Integer = 90 ''イメージ回転
Public Const conDefault_nIMAGE_ROTATE1 As Integer = 90 ''イメージ回転
Public Const conDefault_nIMAGE_ROTATE2 As Integer = 90 ''イメージ回転

'DEMO時設定
Public Const conDefault_nIMAGE_DEB_LEFT0 As Integer = 0 ''イメージ0左座標（デモ）
Public Const conDefault_nIMAGE_DEB_TOP0 As Integer = 0 ''イメージ0上座標（デモ）
Public Const conDefault_nIMAGE_DEB_WIDTH0 As Integer = 3467 ''イメージ0幅（デモ）
Public Const conDefault_nIMAGE_DEB_HEIGHT0 As Integer = 2475 ''イメージ0高さ（デモ）

Public Const conDefault_nIMAGE_DEB_LEFT1 As Integer = 0 ''イメージ1左座標（デモ）
Public Const conDefault_nIMAGE_DEB_TOP1 As Integer = 0 ''イメージ1上座標（デモ）
Public Const conDefault_nIMAGE_DEB_WIDTH1 As Integer = 3467 ''イメージ1幅（デモ）
Public Const conDefault_nIMAGE_DEB_HEIGHT1 As Integer = 2475 ''イメージ1高さ（デモ）

Public Const conDefault_nIMAGE_DEB_LEFT2 As Integer = 0 ''イメージ2左座標（デモ）
Public Const conDefault_nIMAGE_DEB_TOP2 As Integer = 0 ''イメージ2上座標（デモ）
Public Const conDefault_nIMAGE_DEB_WIDTH2 As Integer = 3467 ''イメージ2幅（デモ）
Public Const conDefault_nIMAGE_DEB_HEIGHT2 As Integer = 2475 ''イメージ2高さ（デモ）

'本番設定
Public Const conDefault_nIMAGE_LEFT0 As Integer = 0 ''イメージ0左座標（本番）
Public Const conDefault_nIMAGE_TOP0 As Integer = 0 ''イメージ0上座標（本番）
Public Const conDefault_nIMAGE_WIDTH0 As Integer = 3467 ''イメージ0幅（本番）
Public Const conDefault_nIMAGE_HEIGHT0 As Integer = 2475 ''イメージ0高さ（本番）

Public Const conDefault_nIMAGE_LEFT1 As Integer = 0 ''イメージ1左座標（本番）
Public Const conDefault_nIMAGE_TOP1 As Integer = 0 ''イメージ1上座標（本番）
Public Const conDefault_nIMAGE_WIDTH1 As Integer = 3467 ''イメージ1幅（本番）
Public Const conDefault_nIMAGE_HEIGHT1 As Integer = 2475 ''イメージ1高さ（本番）

Public Const conDefault_nIMAGE_LEFT2 As Integer = 0 ''イメージ2左座標（本番）
Public Const conDefault_nIMAGE_TOP2 As Integer = 0 ''イメージ2上座標（本番）
Public Const conDefault_nIMAGE_WIDTH2 As Integer = 3467 ''イメージ2幅（本番）
Public Const conDefault_nIMAGE_HEIGHT2 As Integer = 2475 ''イメージ2高さ（本番）

Public Const conAccessLevel_Users As Integer = 0 ''アクセスレベル（ユーザー）
Public Const conAccessLevel_Administrators As Integer = 1 ''アクセスレベル（管理）


Public Const conDefault_Separator As String = ":" ''ログ用区切り文字

Public Const conDefine_lGuidanceListMAX As Long = 1000 ''ガイダンス表示　リスト最大件数

Public Const conDefine_ImageDirName As String = "TEMP" ''イメージファイル格納フォルダ
Public Const conDefine_LogDirName As String = "LOGS" ''ＬＯＧファイル格納フォルダ

Public Const conDefine_SYSMODE_SKIN As Integer = 0
Public Const conDefine_SYSMODE_COLOR As Integer = 1
Public Const conDefine_SYSMODE_SLBFAIL As Integer = 2

Public Const conDefine_ColorActive As Long = &H80000005 ''ユーザー定義（ウインドの背景）
Public Const conDefine_ColorNotActive As Long = &H80000013 ''非アクティブタイトル文字色
Public Const conDefine_ColorBKLostFocus As Long = &H80000005 ''ユーザー定義（ウインドの背景）
Public Const conDefine_ColorBKGotFocus As Long = &HFFFF& ''背景黄色
Public Const conDefine_Color_ForColor_HOST_ERROR As Long = &HFF& ''通信エラー　文字色赤

Public MainLogFileNumber As Variant ''ログファイル用 ファイル番号

'2008/09/03 カラー結果一覧のWEB-URL
Public Const conDefault_WEBURL_Color_Result As String = "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://COLDBSRV/CC/jsp/JumpRsltList.jsp?sCall1=sky&sCall2=sky"

'2015/09/15 特鋼カラー結果一覧のWEB-URL
Public Const conDefault_WEBURL_Color_Result_Tok As String = "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://COLDBSRV/CC/jsp/JumpTokRsltList.jsp?sCall1=sky&sCall2=sky"

''システム情報
Public Type typAPSysCfgData
    nDEBUG_MODE As Integer ''デバックモード
    nDISP_DEBUG As Integer ''画面デバック表示
    nFILE_DEBUG As Integer ''LOGファイルデバック表示
    nHOSTDATA_DEBUG As Integer ''実績データ登録ホスト通信デバックモード（戻り値を埋め込みます。）
    nTR_SKIP As Integer ''通信サーバー系スキップ
    nHOSTDATA_SKIP As Integer ''実績データ登録ホスト通信系スキップ
    nDB_SKIP As Integer ''ＤＢスキップ
    nSOZAI_DB_SKIP As Integer ''素材統括ＤＢスキップ
    nSCAN_SKIP As Integer ''スキャナー系スキップ
    
    DB_MYUSER_DSN As String  ''データソース名
    DB_MYUSER_UID As String ''ユーザーＩＤ
    DB_MYUSER_PWD As String ''パスワード
    DB_MYCOMN_DSN As String  ''データソース名
    DB_MYCOMN_UID As String ''ユーザーＩＤ
    DB_MYCOMN_PWD As String ''パスワード
    DB_SOZAI_DSN As String  ''データソース名
    DB_SOZAI_UID As String ''ユーザーＩＤ
    DB_SOZAI_PWD As String ''パスワード
    
    SHARES_SCNDIR As String ''スキャナーイメージ保存先パス
    SHARES_IMGDIR As String ''写真イメージ保存先パス
    SHARES_PDFDIR As String ''20090116 add by M.Aoyagi PDF保存先パス
    
    PHOTOIMG_DIR As String          ''写真ローカルパス
    PHOTOIMG_DELCHK As Integer      ''写真コピー元削除フラグ
    PHOTOIMG_ALLFILES As Integer    ''写真全てのファイル指定フラグ
    
'    TRN_MSG_NO As String ''2004-12-01 トランザクションメッセージ番号
    
'    nUSE_OFFICE As Integer ''事務所設置 0:OFF 1:ON
'    nSEARCH_MAX(0 To 3) As Integer ''検索件数設定
    
'    nSEARCH_RANGE As Integer ''検索有効範囲　過去？日
    
   ' ソケット通信対応
    'HOST_NAME As String ''ホスト名称
    'nHOST_TOUT(0 To 1) As Integer ''通信タイムアウト (0)=ALL (1)=IVT
    HOST_IP As String ''ビジコンＩＰアドレス
    nHOST_PORT As Integer ''ビジコンポート番号
    nHOST_TOUT(0 To 2) As Long ''通信タイムアウト (0)=ALL (1)=OPEN時　(2)=データ通信時
    nHOST_RETRY As Integer ''通信リトライ回数
    
'    nUSE_EMAIL As Integer ''電子ﾒｰﾙを使用してｴﾗｰ通知 0:OFF 1:ON
    
    TR_IP As String ''サーバーＩＰアドレス
    nTR_PORT As Integer ''ポート番号
    'nTR_TOUT(0) As Integer ''通信タイムアウト (0)=ALL
    nTR_TOUT(0 To 2) As Long ''通信タイムアウト (0)=ALL (1)=OPEN時　(2)=データ通信時
    nTR_RETRY As Integer ''通信リトライ回数
    
'    SMTP As String ''送信ﾒｰﾙ (SMTP) ｻｰﾊﾞｰ
'    AP_EMAIL As String ''ｴﾗｰ通知用　電子ﾒｰﾙ ｱﾄﾞﾚｽ
'    USER_EMAIL(1 To 20) As String ''ｴﾗｰ通知先 電子ﾒｰﾙ ｱﾄﾞﾚｽ
    nIMAGE_SIZE(0 To 2) As Integer ''イメージ表示サイズ 10,20,30,40,50,60,70,80,90,100
    nIMAGE_ROTATE(0 To 2) As Integer ''スキャナ読込時回転　0,90,180,270
    nIMAGE_LEFT(0 To 2) As Integer ''切り出しイメージ左座標（Ｐｉｘｅｌｓ）
    nIMAGE_TOP(0 To 2) As Integer ''切り出しイメージ上座標（Ｐｉｘｅｌｓ）
    nIMAGE_WIDTH(0 To 2) As Integer ''切り出しイメージ幅（Ｐｉｘｅｌｓ）
    nIMAGE_HEIGHT(0 To 2) As Integer ''切り出しイメージ高さ（Ｐｉｘｅｌｓ）
    
'    NowLineName As String ''現在選択中のライン名
'    NowLineNumber As String ''現在選択中のライン番号
'    NowLineType As String ''2002-07-11 現在選択中のラインタイプ
    
'    nLineNumberCount As Integer ''ライン番号リストカウント
'    LineNumber() As String ''ライン番号
'    LineType() As String ''2002-07-11 ラインタイプ
    
'    NowStaffNumber As String ''現在選択中の社員番号
    
    NowStaffName(0 To 2) As String ''現在選択中の氏名（保持用）
    NowNextProcess(0 To 2) As String ''現在選択中の次工程（保持用）
    
    WEBURL_Color_Result As String ''カラー結果一覧のWEB-URL
    WEBURL_Color_Result_Tok As String ''特鋼カラー結果一覧のWEB-URL
    
    'NowOperator As String ''現在選択中の操作員
'    NowGroup As String ''現在選択中の班
    'nOperatorCount As Integer ''操作員名リストカウント
    'Operator() As String ''操作員名リスト
'    nGroupCount As Integer ''操作員（班）リストカウント
'    Group() As String ''操作員（班）リスト
    'nStaffCount As Integer ''社員リストカウント
    'nStaffAccessLevel() As Integer ''社員アクセスレベル
    'StaffNumber() As String ''社員番号
    'StaffName() As String ''社員氏名
End Type

'''システムコントロールデータ
'Public Type typAPSysCont
'    bNewEntry As Boolean ''True:新規 False:修正
'End Type

''スラブ情報コントロールデータ
Public Type typAPSlbCont
    bProcessing As Boolean ''スラブ選択ロック用処理中フラグ
    strSearchInputSlbNumber As String ''検索スラブＮｏ．
    nSearchInputModeSelectedIndex As Integer ''検索オプション（入力モード）指定インデックス番号
    nSearchInputStatusSelectedIndex As Integer ''検索オプション（状態入力）指定インデックス番号
    nListSelectedIndexP1 As Integer ''スラブリスト指定インデックス+1番号 0は未指定
'    nChildSelectedIndexP1 As Integer ''子スラブ指定インデックス+1番号 0は未指定
End Type

''スラブ情報
Public Type typAPSlbData
''''------------旧データ
    '検索リスト
    bWorkSelected As Boolean    ''ワーク用
    slb_no As String            ''スラブＮｏ．
    slb_chno As String          ''スラブチャージＮｏ．
    slb_aino As String          ''スラブ合番
    slb_stat As String          ''状態
    slb_zkai_dte As String      ''造塊日
    slb_ksh As String           ''鋼種
    slb_typ As String           ''型
    slb_uksk As String          ''向先
    sys_wrt_dte As String       ''記録日（初回記録日）
    
    '*********************************************
    'カラーチェック
    slb_ccno As String          ''CCNO
    slb_wei As String           ''重量
    slb_lngth As String         ''長さ
    slb_wdth As String          ''幅
    slb_thkns As String         ''厚み
    
    slb_col_cnt As String       ''ｶﾗｰ回数
    host_send As String         ''ビジコン送信 結果フラグ
    host_wrt_dte As String      ''ビジコン送信 記録日
    host_wrt_tme As String      ''ビジコン送信 記録時刻
    sys_wrt_tme As String       ''記録時刻（初回記録時刻）
    
    '異常一覧リスト表示専用 '2008/09/04
    slb_fault_e_judg As String  ''欠陥E面判定
    slb_fault_w_judg As String  ''欠陥W面判定
    slb_fault_s_judg As String  ''欠陥S面判定
    slb_fault_n_judg As String  ''欠陥N面判定
    
    '*********************************************
    'スラブ異常
    fail_host_send As String    ''スラブ異常用　ビジコン送信結果フラグ
    fail_host_wrt_dte As String ''スラブ異常用　ビジコン送信 記録日
    fail_host_wrt_tme As String ''スラブ異常用　ビジコン送信 記録時刻
    fail_sys_wrt_dte As String  ''スラブ異常用　記録日（初回記録日）
    fail_sys_wrt_tme As String  ''スラブ異常用　記録時刻（初回記録時刻）
    '*********************************************
    '処置指示
    fail_dir_sys_wrt_dte As String  ''処置指示用　記録日（初回記録日）
    fail_dir_prn_out_max As String  ''指示印刷済みフラグ
    '*********************************************
    '処置結果
    fail_res_sys_wrt_dte As String  ''処置結果用　記録日（初回記録日）
    fail_res_cmp_flg As String      ''処置結果用　完了フラグ（全体）
    fail_res_host_send As String    ''処置結果用　ビジコン送信結果フラグ
    fail_res_host_wrt_dte As String ''処置結果用　ビジコン送信 記録日
    fail_res_host_wrt_tme As String ''処置結果用　ビジコン送信 記録時刻
    '*********************************************
    
    bAPScanInput As Boolean ''SCANイメージデータ有りフラグ
    bAPFailScanInput As Boolean ''スラブ異常用SCANイメージデータ有りフラグ
    
    bAPPdfInput As Boolean ''PDFイメージデータ有りフラグ
    sAPPdfInput_ReqDate As String
    bAPFailPdfInput As Boolean ''スラブ異常用PDFイメージデータ有りフラグ
    sAPFailPdfInput_ReqDate As String
    
    PhotoImgCnt1 As String '' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
    PhotoImgCnt2 As String '' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
    
    '2016/04/20 - TAI - S
    slb_works_sky_tok As String         '作業場
    '2016/04/20 - TAI - E
End Type

''実績データ（スラブ肌、カラーチェック共用）
''COLOR
Public Type typAPResData
    slb_no As String            ''スラブNO
    slb_chno As String          ''スラブチャージNO
    slb_aino As String          ''スラブ合番
    slb_stat As String          ''状態
    slb_col_cnt As String       ''カラー回数
    slb_ccno As String          ''スラブCCNO
    slb_zkai_dte As String      ''造塊日
    slb_ksh As String           ''鋼種
    slb_typ As String           ''型
    slb_uksk As String          ''向先
    slb_wei As String           ''重量
    slb_lngth As String         ''長さ
    slb_wdth As String          ''幅
    slb_thkns As String         ''厚み
    slb_nxt_prcs As String      ''次工程
    slb_cmt1 As String          ''コメント1
    slb_cmt2 As String          ''コメント2
    
    slb_fault_cd_e_s1 As String ''欠陥E面CD1
    slb_fault_cd_e_s2 As String ''欠陥E面CD2
    slb_fault_cd_e_s3 As String ''欠陥E面CD3
    slb_fault_e_s1 As String    ''欠陥E面種類1
    slb_fault_e_s2 As String    ''欠陥E面種類2
    slb_fault_e_s3 As String    ''欠陥E面種類3
    slb_fault_e_n1 As String    ''欠陥E面個数1
    slb_fault_e_n2 As String    ''欠陥E面個数2
    slb_fault_e_n3 As String    ''欠陥E面個数3
    
    slb_fault_cd_w_s1 As String ''欠陥W面CD1
    slb_fault_cd_w_s2 As String ''欠陥W面CD2
    slb_fault_cd_w_s3 As String ''欠陥W面CD3
    slb_fault_w_s1 As String    ''欠陥W面種類1
    slb_fault_w_s2 As String    ''欠陥W面種類2
    slb_fault_w_s3 As String    ''欠陥W面種類3
    slb_fault_w_n1 As String    ''欠陥W面個数1
    slb_fault_w_n2 As String    ''欠陥W面個数2
    slb_fault_w_n3 As String    ''欠陥W面個数3
    
    slb_fault_cd_s_s1 As String ''欠陥S面CD1
    slb_fault_cd_s_s2 As String ''欠陥S面CD2
    slb_fault_cd_s_s3 As String ''欠陥S面CD3
    slb_fault_s_s1 As String    ''欠陥S面種類1
    slb_fault_s_s2 As String    ''欠陥S面種類2
    slb_fault_s_s3 As String    ''欠陥S面種類3
    slb_fault_s_n1 As String    ''欠陥S面個数1
    slb_fault_s_n2 As String    ''欠陥S面個数2
    slb_fault_s_n3 As String    ''欠陥S面個数3
    
    slb_fault_cd_n_s1 As String ''欠陥N面CD1
    slb_fault_cd_n_s2 As String ''欠陥N面CD2
    slb_fault_cd_n_s3 As String ''欠陥N面CD3
    slb_fault_n_s1 As String    ''欠陥N面種類1
    slb_fault_n_s2 As String    ''欠陥N面種類2
    slb_fault_n_s3 As String    ''欠陥N面種類3
    slb_fault_n_n1 As String    ''欠陥N面個数1
    slb_fault_n_n2 As String    ''欠陥N面個数2
    slb_fault_n_n3 As String    ''欠陥N面個数3
    
    slb_fault_cd_bs_s As String ''内部割れBSCD
    slb_fault_cd_bm_s As String ''内部割れBMCD
    slb_fault_cd_bn_s As String ''内部割れBNCD
    slb_fault_bs_s As String    ''内部割れBS種類
    slb_fault_bm_s As String    ''内部割れBM種類
    slb_fault_bn_s As String    ''内部割れBN種類
    slb_fault_bs_n As String    ''内部割れBS個数
    slb_fault_bm_n As String    ''内部割れBM個数
    slb_fault_bn_n As String    ''内部割れBN個数
    
    slb_fault_cd_ts_s As String ''内部割れTSCD
    slb_fault_cd_tm_s As String ''内部割れTMCD
    slb_fault_cd_tn_s As String ''内部割れTNCD
    slb_fault_ts_s As String    ''内部割れTS種類
    slb_fault_tm_s As String    ''内部割れTM種類
    slb_fault_tn_s As String    ''内部割れTN種類
    slb_fault_ts_n As String    ''内部割れTS個数
    slb_fault_tm_n As String    ''内部割れTM個数
    slb_fault_tn_n As String    ''内部割れTN個数
    
    slb_fault_e_judg As String  ''欠陥E面判定
    slb_fault_w_judg As String  ''欠陥W面判定
    slb_fault_s_judg As String  ''欠陥S面判定
    slb_fault_n_judg As String  ''欠陥N面判定
    slb_fault_b_judg As String  ''欠陥B面判定
    slb_fault_t_judg As String  ''欠陥T面判定
    
    slb_fault_u_judg As String  ''欠陥U面判定
    slb_fault_d_judg As String  ''欠陥D面判定
    
    slb_wrt_nme As String       ''検査員名
    host_send As String         ''ビジコン送信結果
    host_wrt_dte As String      ''記録日
    host_wrt_tme As String      ''記録時刻
    sys_wrt_dte As String       ''登録日
    sys_wrt_tme As String       ''登録時刻
    
    fail_host_send As String         ''スラブ異常報告用　ビジコン送信結果
    fail_host_wrt_dte As String      ''スラブ異常報告用　記録日
    fail_host_wrt_tme As String      ''スラブ異常報告用　記録時刻
    fail_sys_wrt_dte As String       ''スラブ異常報告用　登録日
    fail_sys_wrt_tme As String       ''スラブ異常報告用　登録時刻
    
    '処置指示
    fail_dir_sys_wrt_dte As String  ''処置指示用　記録日（初回記録日）
    
    '処置結果
    fail_res_host_send As String         ''処置結果用　ビジコン送信結果
    fail_res_host_wrt_dte As String      ''処置結果用　記録日
    fail_res_host_wrt_tme As String      ''処置結果用　記録時刻
    
    '共通フラグ
    host_send_flg As String ''ビジコン送信フラグ（各画面で送信前にセット）削除系は未使用
    
    PhotoImgCnt As String '' 20090115 add by M.Aoyagi    画像登録件数表示の為追加
    
    '2016/04/20 - TAI - S
    '検査結果
    slb_fault_total_judg As String
    '作業場所
    slb_works_sky_tok As String
    '2016/04/20 - TAI - E

End Type

''素材統括データ（スラブ肌、カラーチェック共用）
''COLOR
Public Type typAPSozaiData
    '**********************************************************'
    'nchtaisl
    slb_no As String            ''スラブNO
    slb_ksh As String           ''鋼種
    slb_uksk As String          ''向先（熱延向先）
    slb_lngth As String         ''長さ
    slb_color_wei As String     ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
    slb_typ As String           ''型
    slb_skin_wei As String      ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
    slb_wdth As String          ''幅
    slb_thkns As String         ''厚み
    slb_zkai_dte As String      ''造塊日（造塊年月日）
    '**********************************************************'
    'skjchjdtテーブル
    slb_chno As String          ''チャージNO
    slb_ccno As String          ''CCNO
    '**********************************************************'
End Type

''処置内容指示確認／結果登録用データ
''COLOR
Public Type typAPDirResData
    slb_no As String            ''スラブNO
    slb_chno As String          ''スラブチャージNO
    slb_aino As String          ''スラブ合番
    slb_stat As String          ''状態
    slb_col_cnt As String       ''カラー回数
    dir_no As String            ''指示番号
    
    dir_nme1 As String            ''指示項目1
    dir_val1 As String            ''指示値1
    dir_uni1 As String            ''指示単位1
    dir_nme2 As String            ''指示項目2
    dir_val2 As String            ''指示値2
    dir_uni2 As String            ''指示単位2
    dir_cmt1 As String            ''コメント1
    dir_cmt2 As String            ''コメント2
    dir_wrt_dte  As String            ''指示日
    dir_wrt_nme As String            ''指示者名
    dir_sys_wrt_dte As String            ''登録日
    dir_sys_wrt_tme As String            ''登録時刻
    
    res_cmt1 As String            ''コメント1（未使用／予約）
    res_cmt2 As String            ''コメント2（未使用／予約）
    res_cmp_flg As String           ''処置完了フラグ 1:完了
    res_aft_stat As String          ''処置後状態 1:不適合有り（割れ、疵有り）
    res_wrt_dte  As String          ''入力日
    res_wrt_nme As String           ''入力者名
    res_sys_wrt_dte As String            ''登録日
    res_sys_wrt_tme As String            ''登録時刻
    
End Type

'欠陥情報
''COLORSYS
Public Type typAPFaultList
    strCode As String
    strName As String
End Type

''スタッフ情報
''COLORSYS
Public Type typAPStaffData
    inp_StaffName As String ''スタッフ名
End Type

''検査員情報
''COLORSYS
Public Type typAPInspData
    inp_InspName As String ''検査員名
End Type

''入力者情報
''COLORSYS
Public Type typAPInpData
    inp_InpName As String ''入力者名
End Type

''次工程情報
Public Type typAPNextProcData
    inp_NextProc As String ''次工程
End Type

''処置状態
Public Type typAPDirRes_Stat
    inp_DirRes_StatCode As String
    inp_DirRes_Stat As String
End Type

''処置結果
Public Type typAPDirRes_Res
    inp_DirRes_ResCode As String
    inp_DirRes_Res As String
End Type

''システム情報
Public APSysCfgData As typAPSysCfgData ''システム情報

'''システムコントロールデータ
'Public APSysCont As typAPSysCont ''2001-11-09 システムコントロールデータ

''スラブ情報コントロールデータ
Public APSlbCont As typAPSlbCont ''2001-11-08 スラブ情報コントロールデータ

''面欠陥リスト情報（スラブ肌）
Public APFaultFaceSkin() As typAPFaultList
''内部欠陥リスト情報（スラブ肌）
Public APFaultInsideSkin() As typAPFaultList

''面欠陥リスト情報（カラーチェック）
Public APFaultFaceColor() As typAPFaultList

''処理中実績データ（画面表示＆レジストリ保存用）
Public APResData As typAPResData ''処理中実績データ（画面表示＆レジストリ保存用）
Public APResDataBK As typAPResData ''処理中実績データ（処理用バックアップエリア）

Public APSozaiData As typAPSozaiData ''素材統括問合せデータ

Public APDirResData() As typAPDirResData ''処置内容指示確認／結果登録用データ

''検索スラブリスト
Public APSearchListSlbData() As typAPSlbData ''検索スラブリスト

''スラブ検索用ＴＭＰ
Public APSearchTmpSlbData() As typAPSlbData ''スラブ検索用ＴＭＰ

''実績データ読込み用ＴＭＰ
Public APResTmpData() As typAPResData ''実績データ読込み用ＴＭＰ
Public APSozaiTmpData() As typAPSozaiData ''素材統括問合せデータ
Public APDirResTmpData() As typAPDirResData ''処置内容指示確認／結果登録用データ

''スタッフ名マスタ情報
''COLORSYS
Public APStaffData() As typAPStaffData ''スタッフ名マスタ情報

''検査員名マスタ情報
''COLORSYS
Public APInspData() As typAPInspData ''検査員名マスタ情報

''入力者名マスタ情報
''COLORSYS
Public APInpData() As typAPInpData ''入力者名マスタ情報

''次工程マスタ情報
Public APNextProcDataSkin() As typAPNextProcData ''次工程マスタ情報
Public APNextProcDataColor() As typAPNextProcData ''次工程マスタ情報

Public APDirRes_Stat() As typAPDirRes_Stat ''処置状態
Public APDirRes_Res() As typAPDirRes_Res ''処置結果

''ＤＢオフラインで強制入力を行ったことを判断するフラグ
'Public bAPInputOffline As Boolean

'2016/04/20 - TAI - S
'作業場所情報
Public works_sky_tok As String
Public Const WORKS_SKY As String = "SKY"       'SKY
Public Const WORKS_TOK As String = "TOK"       '特鋼
'2016/04/20 - TAI - E

