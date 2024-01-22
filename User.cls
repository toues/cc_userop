VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) User.cls                ver 1.00 ( '01.10.01 SEC Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　ユーザークラスモジュール
' 　本モジュールはユーザーログイン処理を目的とした
' 　クラスモジュールである。

Option Explicit

Dim sUser As String ''ユーザー

' @(f)
'
' 機能      : ユーザーログイン処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ログイン処理を行う。
'
' 備考      :
'
Public Sub SetUser()
    Dim SupervisorName As String
    Dim SupervisorPassword As String
    Dim UserName As String
    Dim UserPassword As String
    
    SupervisorName = GetSetting(conReg_APPNAME, "Password", "Supervisor Name", "")
    SupervisorPassword = GetSetting(conReg_APPNAME, "Password", "Supervisor Password", "")
    UserName = GetSetting(conReg_APPNAME, "Password", "User Name", "")
    UserPassword = GetSetting(conReg_APPNAME, "Password", "User Password", "")
    
    '空白だったら設定する。
    If SupervisorName = "" Then
        Dim fSetLogin As New frmSetLogin
        fSetLogin.LevLabel = "Supervisor"
        fSetLogin.Show vbModal
        If Not fSetLogin.LoginSucceeded Then
            'キャンセルの場合
            Unload fSetLogin
            End
        Else
            SaveSetting conReg_APPNAME, "Password", "Supervisor Name", fSetLogin.imtxtUserName
            SaveSetting conReg_APPNAME, "Password", "Supervisor Password", fSetLogin.imtxtPassword
            'If UserName <> "" Then
            '    MsgBox "新しいPasswordが設定されました。再起動してください"
            '    Unload fSetLogin
            '   End
            'End If
        End If
        MsgBox "新しいPasswordが設定されました。"
        Unload fSetLogin
    End If
    
    'If UserName = "" Then
    '    Set fSetLogin = New frmSetLogin
    '    fSetLogin.LevLabel = "User"
    '    fSetLogin.Show vbModal
    '    If Not fSetLogin.LoginSucceeded Then
    '        Unload fSetLogin
    '        End
    '    Else
    '    MsgBox "新しいPasswordが設定されました。再起動してください"
    '    SaveSetting conReg_APPNAME, "Password", "User Name", fSetLogin.imtxtUserName
    '    SaveSetting conReg_APPNAME, "Password", "User Password", fSetLogin.imtxtPassword
    '    Unload fSetLogin
    '    End
    '    End If
    'End If
End Sub

' @(f)
'
' 機能      : ユーザー変更状況取得
'
' 引き数    :
'
' 返り値    : True=ログイン成功／False=ログイン失敗
'
' 機能説明  : ユーザー変更ログイン処理を行う。
'
' 備考      :
'
Property Get ChangeUser() As Boolean
    Debug.Print "CALL USER"
    Dim SupervisorName As String
    Dim SupervisorPassword As String
    Dim UserName As String
    Dim UserPassword As String
    SupervisorName = GetSetting(conReg_APPNAME, "Password", "Supervisor Name", "")
    SupervisorPassword = GetSetting(conReg_APPNAME, "Password", "Supervisor Password", "")
    UserName = GetSetting(conReg_APPNAME, "Password", "User Name", "")
    UserPassword = GetSetting(conReg_APPNAME, "Password", "User Password", "")
    
    Dim fLogin As New frmLogin
    fLogin.UserName = UserName
    fLogin.UserPassword = UserPassword
    fLogin.SupervisorName = SupervisorName
    fLogin.SupervisorPassword = SupervisorPassword
    fLogin.Show vbModal
    If Not fLogin.LoginSucceeded Then
        'ﾛｸﾞｲﾝに失敗。
        ChangeUser = False
        Unload fLogin
    Else
        'set user level
        User = fLogin.User
        ChangeUser = True
        Unload fLogin
    End If
End Property

' @(f)
'
' 機能      : ユーザー取得
'
' 引き数    :
'
' 返り値    : ユーザー
'
' 機能説明  : ユーザー取得処理を行う。
'
' 備考      :
'
Property Get User() As String
    User = sUser
End Property

' @(f)
'
' 機能      : ユーザー設定
'
' 引き数    : 新しいユーザー
'
' 返り値    :
'
' 機能説明  : ユーザー設定処理を行う。
'
' 備考      :
'
Property Let User(ByVal newUser As String)
    sUser = newUser
End Property


