VERSION 5.00
Begin VB.Form MessageYN 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "カラーチェック情報電子化システム－ＰＣシステムメッセージ"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   3345
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Visible         =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2580
      Width           =   1500
   End
   Begin VB.CommandButton CANCEL 
      Caption         =   "キャンセル"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   2580
      Width           =   1500
   End
   Begin VB.TextBox MsgText 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'なし
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2190
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "MessageYN.frx":0000
      Top             =   240
      Width           =   5235
   End
End
Attribute VB_Name = "MessageYN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) MessageYN.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　メッセージＹ／Ｎ表示フォーム
' 　本モジュールはメッセージＹ／Ｎ表示フォームで使用する
' 　ためのものである。

Option Explicit

Public AutoDelete As Boolean ''画面クローズ自動フラグ
Public Yes As Boolean ''問い合わせフラグ

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納
Private bCallBackFlag As Boolean ''コールバックフラグ格納

' @(f)
'
' 機能      : キャンセルボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : キャンセルボタン処理。
'
' 備考      :
'
Private Sub Cancel_Click()
    Yes = False
    
    fMainWnd.Enabled = True
    
    If AutoDelete = False Then
        Me.Hide
    Else
        Unload Me
    End If
    
    If bCallBackFlag = True Then
        cCallBackObject.CallBackMessage iCallBackID, 0
        Set cCallBackObject = Nothing
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
Private Sub OK_Click()
    Yes = True
    
    fMainWnd.Enabled = True
    
    If AutoDelete = False Then
        Me.Hide
    Else
        Unload Me
    End If
        
    If bCallBackFlag = True Then
        cCallBackObject.CallBackMessage iCallBackID, 1
        Set cCallBackObject = Nothing
    End If

End Sub

' @(f)
'
' 機能      : コールバック設定
'
' 引き数    : ARG1 - コールバックオブジェクト
'             ARG2 - コールバックＩＤ
'             ARG3 - 画面クローズ自動フラグ
'
' 返り値    :
'
' 機能説明  : 戻り先コールバック情報を設定する。
'
' 備考      :
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer, ByVal AutDel As Boolean)
    AutoDelete = AutDel
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
    bCallBackFlag = True
End Sub


