VERSION 5.00
Begin VB.Form Message 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "カラーチェック情報電子化システム－ＰＣシステムメッセージ"
   ClientHeight    =   3555
   ClientLeft      =   5085
   ClientTop       =   4860
   ClientWidth     =   6480
   ControlBox      =   0   'False
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  '矢印
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   3555
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox MsgText 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   2280
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Message.frx":030A
      Top             =   120
      Width           =   6225
   End
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
      Left            =   2280
      TabIndex        =   0
      Top             =   2700
      Width           =   1815
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) Message.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　メッセージ表示フォーム
' 　本モジュールはメッセージ表示フォームで使用する
' 　ためのものである。

Option Explicit

Public AutoDelete As Boolean ''画面クローズ自動フラグ
Public Yes As Boolean ''問い合わせフラグ

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納
Private bCallBackFlag As Boolean ''コールバックフラグ格納

' @(f)
'
' 機能      : 画面クローズ
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 画面クローズ処理。
'
' 備考      :
'
Public Sub OK_Close()
    Call OK_Click
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
    
    Select Case MsgText.Text
    End Select

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

