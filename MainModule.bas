Attribute VB_Name = "MainModule"
' @(h) MainModule.Bas                ver 1.00 ( '01.10.01 SEC Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@���C�����W���[��
' �@�{���W���[���̓V�X�e�����N������
' �@���߂̂��̂ł���B

Option Explicit

Public cUser As User ''���[�U�[�N���X
Public fMainWnd As BaseWnd ''�x�[�X�E�C���h

' @(f)
'
' �@�\      : ���C��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �V�X�e�����N������B
'
' ���l      :
'
Sub Main()

    If App.PrevInstance = True Then
        '�d���N���̋֎~
        End
    End If
    
    '�k�n�f�t�@�C���i�[�t�H���_�쐬
    If Dir(App.Path & "\" & conDefine_LogDirName, vbDirectory) = "" Then
        Call MkDir(App.Path & "\" & conDefine_LogDirName)
    End If
    
    '�C���[�W�t�@�C���i�[�t�H���_�쐬
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
' �@�\      : �S�A�����[�h
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �S�A�����[�h�����B
'
' ���l      : �e����N���X�̉���N���[�Y�����y�сA�N���X�̔j��
'
Public Sub UnloadAll()
    '�e����N���X�̉���N���[�Y�����y�сA�N���X�̔j��
    Set cUser = Nothing
    
    '�b�r�n�j�d�s�I��
    'Call CSTRAN_END
    
End Sub
