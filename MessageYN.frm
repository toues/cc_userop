VERSION 5.00
Begin VB.Form MessageYN 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�J���[�`�F�b�N���d�q���V�X�e���|�o�b�V�X�e�����b�Z�[�W"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   3345
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Visible         =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "�l�r �o����"
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
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �o����"
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
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '�Ȃ�
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �o����"
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
' �J���[�`�F�b�N���тo�b�@���b�Z�[�W�x�^�m�\���t�H�[��
' �@�{���W���[���̓��b�Z�[�W�x�^�m�\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Public AutoDelete As Boolean ''��ʃN���[�Y�����t���O
Public Yes As Boolean ''�₢���킹�t���O

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[
Private bCallBackFlag As Boolean ''�R�[���o�b�N�t���O�i�[

' @(f)
'
' �@�\      : �L�����Z���{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �L�����Z���{�^�������B
'
' ���l      :
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
' �@�\      : �n�j�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �n�j�{�^�������B
'
' ���l      :
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
' �@�\      : �R�[���o�b�N�ݒ�
'
' ������    : ARG1 - �R�[���o�b�N�I�u�W�F�N�g
'             ARG2 - �R�[���o�b�N�h�c
'             ARG3 - ��ʃN���[�Y�����t���O
'
' �Ԃ�l    :
'
' �@�\����  : �߂��R�[���o�b�N����ݒ肷��B
'
' ���l      :
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer, ByVal AutDel As Boolean)
    AutoDelete = AutDel
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
    bCallBackFlag = True
End Sub


