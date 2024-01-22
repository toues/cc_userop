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
' �J���[�`�F�b�N���тo�b�@���[�U�[�N���X���W���[��
' �@�{���W���[���̓��[�U�[���O�C��������ړI�Ƃ���
' �@�N���X���W���[���ł���B

Option Explicit

Dim sUser As String ''���[�U�[

' @(f)
'
' �@�\      : ���[�U�[���O�C������
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���O�C���������s���B
'
' ���l      :
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
    
    '�󔒂�������ݒ肷��B
    If SupervisorName = "" Then
        Dim fSetLogin As New frmSetLogin
        fSetLogin.LevLabel = "Supervisor"
        fSetLogin.Show vbModal
        If Not fSetLogin.LoginSucceeded Then
            '�L�����Z���̏ꍇ
            Unload fSetLogin
            End
        Else
            SaveSetting conReg_APPNAME, "Password", "Supervisor Name", fSetLogin.imtxtUserName
            SaveSetting conReg_APPNAME, "Password", "Supervisor Password", fSetLogin.imtxtPassword
            'If UserName <> "" Then
            '    MsgBox "�V����Password���ݒ肳��܂����B�ċN�����Ă�������"
            '    Unload fSetLogin
            '   End
            'End If
        End If
        MsgBox "�V����Password���ݒ肳��܂����B"
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
    '    MsgBox "�V����Password���ݒ肳��܂����B�ċN�����Ă�������"
    '    SaveSetting conReg_APPNAME, "Password", "User Name", fSetLogin.imtxtUserName
    '    SaveSetting conReg_APPNAME, "Password", "User Password", fSetLogin.imtxtPassword
    '    Unload fSetLogin
    '    End
    '    End If
    'End If
End Sub

' @(f)
'
' �@�\      : ���[�U�[�ύX�󋵎擾
'
' ������    :
'
' �Ԃ�l    : True=���O�C�������^False=���O�C�����s
'
' �@�\����  : ���[�U�[�ύX���O�C���������s���B
'
' ���l      :
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
        '۸޲݂Ɏ��s�B
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
' �@�\      : ���[�U�[�擾
'
' ������    :
'
' �Ԃ�l    : ���[�U�[
'
' �@�\����  : ���[�U�[�擾�������s���B
'
' ���l      :
'
Property Get User() As String
    User = sUser
End Property

' @(f)
'
' �@�\      : ���[�U�[�ݒ�
'
' ������    : �V�������[�U�[
'
' �Ԃ�l    :
'
' �@�\����  : ���[�U�[�ݒ菈�����s���B
'
' ���l      :
'
Property Let User(ByVal newUser As String)
    sUser = newUser
End Property

