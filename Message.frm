VERSION 5.00
Begin VB.Form Message 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�J���[�`�F�b�N���d�q���V�X�e���|�o�b�V�X�e�����b�Z�[�W"
   ClientHeight    =   3555
   ClientLeft      =   5085
   ClientTop       =   4860
   ClientWidth     =   6480
   ControlBox      =   0   'False
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  '���
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   3555
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox MsgText 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00C0FFFF&
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
         Name            =   "�l�r �o����"
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
' �J���[�`�F�b�N���тo�b�@���b�Z�[�W�\���t�H�[��
' �@�{���W���[���̓��b�Z�[�W�\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Public AutoDelete As Boolean ''��ʃN���[�Y�����t���O
Public Yes As Boolean ''�₢���킹�t���O

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[
Private bCallBackFlag As Boolean ''�R�[���o�b�N�t���O�i�[

' @(f)
'
' �@�\      : ��ʃN���[�Y
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ��ʃN���[�Y�����B
'
' ���l      :
'
Public Sub OK_Close()
    Call OK_Click
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

