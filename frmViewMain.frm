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
      Name            =   "�l�r �o�S�V�b�N"
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
   WindowState     =   2  '�ő剻
   Begin VB.TextBox txtDummy 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   14580
      Width           =   405
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�J���[�`�F�b�N���d�q���V�X�e��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
' �J���[�`�F�b�N���тo�b�@���C���\���t�H�[��
' �@�{���W���[���̓��C���\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

' @(f)
'
' �@�\      : ��ʕ\�����t���b�V��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ��ʕ\�����t���b�V�������B
'
' ���l      :
'
Public Sub RefreshViewMain()
End Sub

' @(f)
'
' �@�\      : �t�H�[�����[�h
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[�����[�h���̏������s���B
'
' ���l      :
'
Private Sub Form_Load()
End Sub

' @(f)
'
' �@�\      : �t�H�[���̏�����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[���̏����������B
'
' ���l      :
'
Private Sub InitForm()
End Sub

' @(f)
'
' �@�\      : �t�H�[�����T�C�Y
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[�����T�C�Y����
'
' ���l      :
'
Private Sub Form_Resize()
End Sub

' @(f)
'
' �@�\      : �_�~�[����BOX�L�[����
'
' ������    : ARG1 - �L�[�R�[�h
'             ARG2 - �V�t�g�t���O
'
' �Ԃ�l    :
'
' �@�\����  : �_�~�[����BOX�L�[�������̏������s���B
'
' ���l      :
'
Private Sub txtDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc(vbTab) Then
'        fMainWnd.cmdOpChg.SetFocus
        fMainWnd.cmdSkinIn.SetFocus
    End If
End Sub

