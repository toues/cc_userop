VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  '�Œ��޲�۸�
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Splash"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Timer timSplash 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2280
   End
   Begin VB.Label lblCompanyProduct 
      Alignment       =   2  '��������
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      Caption         =   "������Ђx�`�j�h�m���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   1620
      Width           =   2730
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   1650
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   2  '��������
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4920
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      Caption         =   "�J���[�`�F�b�N���d�q���V�X�e��"
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   75
      TabIndex        =   2
      Top             =   900
      Width           =   5445
   End
   Begin VB.Label lblPlatform 
      Alignment       =   2  '��������
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      Caption         =   "�J���[�`�F�b�N���ѓ��͂o�b"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   5475
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   135
      TabIndex        =   0
      Top             =   2250
      Width           =   5280
   End
   Begin VB.Image imgLogo 
      Height          =   570
      Left            =   180
      Stretch         =   -1  'True
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSplash.Frm                ver 1.00 ( '2008.04.17 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�X�v���b�V���\���t�H�[��
' �@�{���W���[���̓X�v���b�V���\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

' @(f)
'
' �@�\      : �t�H�[���_�u���N���b�N
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[���_�u���N���b�N���̏������s���B
'
' ���l      :
'
Private Sub Form_DblClick()
    Unload Me
End Sub

' @(f)
'
' �@�\      : �t�H�[���L�[��
'
' ������    : ARG1 - ASCII�R�[�h
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[���L�[�����̏������s���B
'
' ���l      :
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
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
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    'lblProductName.Caption = App.Title
    'lblLicenseTo.Caption = App.LegalCopyright
    'lblComment.Caption = App.Comments
    'lblWarning = ""

    timSplash.Enabled = True

End Sub

' @(f)
'
' �@�\      : �t�H�[���A�����[�h
'
' ������    : ARG1 - �L�����Z���t���O
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[���A�����[�h���̏������s���B
'
' ���l      :
'
Private Sub Form_Unload(CANCEL As Integer)
    Set fMainWnd.fMDIWnd = frmViewMain
    fMainWnd.fMDIWnd.Show
End Sub

' @(f)
'
' �@�\      : �\������p�^�C�}�[�C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �\������p�^�C�}�[�C�x���g���̏������s���B
'
' ���l      :
'
Private Sub timSplash_Timer()
    timSplash.Enabled = False
    Unload Me
End Sub
