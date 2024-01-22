Attribute VB_Name = "MainModule"
' @(h) MainModule.Bas                ver 1.00 ( '01.10.01 SEC Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　メインモジュール
' 　本モジュールはシステムを起動する
' 　ためのものである。

Option Explicit

Public cUser As User ''ユーザークラス
Public fMainWnd As BaseWnd ''ベースウインド

' @(f)
'
' 機能      : メイン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : システムを起動する。
'
' 備考      :
'
Sub Main()

    If App.PrevInstance = True Then
        '重複起動の禁止
        End
    End If
    
    'ＬＯＧファイル格納フォルダ作成
    If Dir(App.Path & "\" & conDefine_LogDirName, vbDirectory) = "" Then
        Call MkDir(App.Path & "\" & conDefine_LogDirName)
    End If
    
    'イメージファイル格納フォルダ作成
    If Dir(App.Path & "\" & conDefine_ImageDirName, vbDirectory) = "" Then
        Call MkDir(App.Path & "\" & conDefine_ImageDirName)
    End If
    
    'Create User class
    Set cUser = New User
    Dim Result As Boolean
    cUser.SetUser
    
    'Change user
    'result = cUser.ChangeUser
    'If result = False Then
    '    End
    'End If

    
    frmSplash.Show
'***** Add custom code ********************************************************

'******************************************************************************
    Set fMainWnd = BaseWnd
        
End Sub

' @(f)
'
' 機能      : 全アンロード
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 全アンロード処理。
'
' 備考      : 各制御クラスの回線クローズ処理及び、クラスの破棄
'
Public Sub UnloadAll()
    '各制御クラスの回線クローズ処理及び、クラスの破棄
    Set cUser = Nothing
    
    'ＣＳＯＫＥＴ終了
    'Call CSTRAN_END
    
End Sub
