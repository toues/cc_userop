VERSION 5.00
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX12N.OCX"
Begin VB.Form frmFullImage 
   BackColor       =   &H00C0FFC0&
   Caption         =   "イメージ全体表示"
   ClientHeight    =   9855
   ClientLeft      =   855
   ClientTop       =   1125
   ClientWidth     =   14985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   15500
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdOK 
      Caption         =   "戻る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13680
      TabIndex        =   0
      Top             =   9360
      Width           =   1215
   End
   Begin LEADLib.LEAD LEAD1 
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14955
      _Version        =   65539
      _ExtentX        =   26379
      _ExtentY        =   16325
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   613
      ScaleWidth      =   993
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
End
Attribute VB_Name = "frmFullImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmFullImage.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　イメージ全体表示フォーム
' 　本モジュールはイメージ全体表示フォームで使用する
' 　ためのものである。

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

Option Explicit

' @(f)
'
' 機能      : ＯＫボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＯＫボタン処理。
'
' 備考      :
'
Private Sub cmdOK_Click()
    Unload Me
    
    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' 機能      : 印刷ボタン
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 印刷ボタン処理。
'
' 備考      :
'
Private Sub cmdPrint_Click(Index As Integer)
End Sub

' @(f)
'
' 機能      : フォームロード
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームロード時の処理を行う。
'
' 備考      :
'
Private Sub Form_Load()

    LEAD1.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD1.EnableMethodErrors = False 'False   システムエラーイベントを発生させない
    LEAD1.EnableTwainEvent = True
    
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''スラブ肌調査入力
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
        
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_COLOR)
        
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SLBFAIL)
    
    End Select
    
End Sub

' @(f)
'
' 機能      : フォームリサイズ
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームリサイズ時の処理を行う。
'
' 備考      :
'
Private Sub Form_Resize()
    If (Me.Height - 1100) > 0 Then
        LEAD1.Height = Me.Height - 1000
        LEAD1.Width = Me.Width
        LEAD1.Left = 150
        LEAD1.Top = 0
        cmdOK.Top = LEAD1.Height + 100
        cmdOK.Left = (Me.Width - 1100)
'        cmdPrint(0).Top = LEAD1.Height + 100
'        cmdPrint(1).Top = LEAD1.Height + 100
'        cmdPrint(0).Left = (cmdOK.Left - cmdPrint(0).Width) - 100
'        cmdPrint(1).Left = (cmdPrint(0).Left - cmdPrint(1).Width) - 100
    End If
End Sub

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

