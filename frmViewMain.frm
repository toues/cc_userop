VERSION 5.00
Begin VB.Form frmViewMain 
   BackColor       =   &H80000004&
   Caption         =   "View"
   ClientHeight    =   14955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14955
   ScaleWidth      =   19080
   WindowState     =   2  '最大化
   Begin VB.TextBox txtDummy 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   14580
      Width           =   405
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "カラーチェック情報電子化システム"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   48
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   19035
   End
End
Attribute VB_Name = "frmViewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmViewMain.Frm                ver 1.00 ( '2008.04.17 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　メイン表示フォーム
' 　本モジュールはメイン表示フォームで使用する
' 　ためのものである。

Option Explicit

' @(f)
'
' 機能      : 画面表示リフレッシュ
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 画面表示リフレッシュ処理。
'
' 備考      :
'
Public Sub RefreshViewMain()
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
End Sub

' @(f)
'
' 機能      : フォームの初期化
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームの初期化処理。
'
' 備考      :
'
Private Sub InitForm()
End Sub

' @(f)
'
' 機能      : フォームリサイズ
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームリサイズ処理
'
' 備考      :
'
Private Sub Form_Resize()
End Sub

' @(f)
'
' 機能      : ダミー入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : ダミー入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub txtDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc(vbTab) Then
'        fMainWnd.cmdOpChg.SetFocus
        fMainWnd.cmdSkinIn.SetFocus
    End If
End Sub

