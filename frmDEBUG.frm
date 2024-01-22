VERSION 5.00
Begin VB.Form frmDEBUG 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "デバック設定画面"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5295
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CheckBox chkDEBUG_MODE 
      Caption         =   "DEBUGMODE"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame fraDEBUG 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3975
      Begin VB.CheckBox chkSOZAI_DB_SKIP 
         Caption         =   "素材統括DBスキップ"
         Height          =   375
         Left            =   300
         TabIndex        =   12
         Top             =   1920
         Width           =   2955
      End
      Begin VB.Frame Frame3 
         Caption         =   "実績データ登録"
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   2820
         Width           =   3495
         Begin VB.CheckBox chkHOSTDATA_SKIP 
            Caption         =   "HOSTスキップ"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   765
            Width           =   3015
         End
         Begin VB.CheckBox chkHOSTDATA_DEBUG 
            Caption         =   "HOSTデバック（戻り値埋め込み）"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CheckBox chkFILE_DEBUG 
         Caption         =   "ログファイルデバック"
         Height          =   375
         Left            =   300
         TabIndex        =   8
         Top             =   780
         Width           =   2955
      End
      Begin VB.CheckBox chkDISP_DEBUG 
         Caption         =   "画面デバック表示"
         Height          =   375
         Left            =   300
         TabIndex        =   7
         Top             =   400
         Width           =   2955
      End
      Begin VB.CheckBox chkSCAN_SKIP 
         Caption         =   "SCANスキップ"
         Height          =   375
         Left            =   300
         TabIndex        =   6
         Top             =   2340
         Width           =   2955
      End
      Begin VB.CheckBox chkDB_SKIP 
         Caption         =   "DBスキップ"
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   1500
         Width           =   2955
      End
      Begin VB.CheckBox chkTRAN_SKIP 
         Caption         =   "通信サーバー要求スキップ"
         Height          =   375
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   4140
      TabIndex        =   0
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   3435
   End
End
Attribute VB_Name = "frmDEBUG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmDEBUG.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　システムデバック画面表示フォーム
' 　本モジュールはシステムデバック画面表示フォームで使用する
' 　ためのものである。

Option Explicit

Private Sub chkHOSTDATA_DEBUG_Click()
    If chkHOSTDATA_DEBUG Then
        chkHOSTDATA_SKIP.Enabled = False
    Else
        chkHOSTDATA_SKIP.Enabled = True
    End If
End Sub

Private Sub chkHOSTDATA_SKIP_Click()
    If chkHOSTDATA_SKIP Then
        chkHOSTDATA_DEBUG.Enabled = False
    Else
        chkHOSTDATA_DEBUG.Enabled = True
    End If
End Sub

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
    APSysCfgData.nDEBUG_MODE = chkDEBUG_MODE.Value
    APSysCfgData.nDISP_DEBUG = chkDISP_DEBUG.Value
    APSysCfgData.nFILE_DEBUG = chkFILE_DEBUG.Value
    APSysCfgData.nHOSTDATA_DEBUG = chkHOSTDATA_DEBUG.Value
    APSysCfgData.nTR_SKIP = chkTRAN_SKIP.Value
    APSysCfgData.nHOSTDATA_SKIP = chkHOSTDATA_SKIP.Value
    APSysCfgData.nDB_SKIP = chkDB_SKIP.Value
    APSysCfgData.nSOZAI_DB_SKIP = chkSOZAI_DB_SKIP.Value
    APSysCfgData.nSCAN_SKIP = chkSCAN_SKIP.Value
    
    Unload Me
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
    chkDEBUG_MODE.Value = APSysCfgData.nDEBUG_MODE
    chkDISP_DEBUG.Value = APSysCfgData.nDISP_DEBUG
    chkFILE_DEBUG.Value = APSysCfgData.nFILE_DEBUG
    chkHOSTDATA_DEBUG.Value = APSysCfgData.nHOSTDATA_DEBUG
    chkTRAN_SKIP.Value = APSysCfgData.nTR_SKIP
    chkHOSTDATA_SKIP.Value = APSysCfgData.nHOSTDATA_SKIP
    chkDB_SKIP.Value = APSysCfgData.nDB_SKIP
    chkSOZAI_DB_SKIP.Value = APSysCfgData.nSOZAI_DB_SKIP
    chkSCAN_SKIP.Value = APSysCfgData.nSCAN_SKIP
    
    lblinfo.Caption = App.path
End Sub

