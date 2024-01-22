VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSetLogin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "システム設定用　パスワード設定"
   ClientHeight    =   2460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin imText6Ctl.imText imtxtUserName 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "frmSetLogin.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSetLogin.frx":006E
      Key             =   "frmSetLogin.frx":008C
      BackColor       =   -2147483643
      EditMode        =   3
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   "A9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   256
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin imText6Ctl.imText imtxtPassword 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "frmSetLogin.frx":00D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSetLogin.frx":013E
      Key             =   "frmSetLogin.frx":015C
      BackColor       =   -2147483643
      EditMode        =   3
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   "A9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   256
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imText6Ctl.imText imtxtPassword2 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "frmSetLogin.frx":01A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSetLogin.frx":020E
      Key             =   "frmSetLogin.frx":022C
      BackColor       =   -2147483643
      EditMode        =   3
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   "A9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   256
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label UserLabel 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FF8080&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1440
      Width           =   2325
   End
   Begin VB.Label LevLabel 
      BackColor       =   &H00FF8080&
      Caption         =   "Supervisor"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3180
      TabIndex        =   7
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label UserLabel 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FF8080&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   525
      Width           =   2385
   End
   Begin VB.Label UserLabel 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   975
      Width           =   2385
   End
End
Attribute VB_Name = "frmSetLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSetLogin.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　ログイン登録フォーム
' 　本モジュールはログイン登録フォームで使用する
' 　ためのものである。

Option Explicit

Public LoginSucceeded As Boolean ''登録完了フラグ

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
Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
    Unload Me
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
    If imtxtUserName = "" Then
        MsgBox "ユーザー名が入力されていません。もう一度入力してください!", , "ﾛｸﾞｵﾝ"
        imtxtPassword.SetFocus
        imtxtPassword.SelStart = 0
        imtxtPassword.SelLength = Len(imtxtPassword.Text)
    End If
    If imtxtPassword <> "" And imtxtPassword = imtxtPassword2 Then
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "ﾊﾟｽﾜｰﾄﾞが正しくありません。もう一度入力してください !", , "ﾛｸﾞｵﾝ"
        imtxtPassword.SetFocus
        imtxtPassword.SelStart = 0
        imtxtPassword.SelLength = Len(imtxtPassword.Text)
    End If
End Sub

' @(f)
'
' 機能      : パスワード入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : パスワード入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

' @(f)
'
' 機能      : パスワード２入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : パスワード２入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtPassword2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

' @(f)
'
' 機能      : ユーザー名入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : ユーザー名入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub
