VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�V�X�e���ݒ�p�@���O�C��"
   ClientHeight    =   1890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4560
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin imText6Ctl.imText imtxtUserName 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   180
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "frmLogin.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLogin.frx":006E
      Key             =   "frmLogin.frx":008C
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
   Begin imText6Ctl.imText imtxtPassword 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   660
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "frmLogin.frx":00D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLogin.frx":013E
      Key             =   "frmLogin.frx":015C
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
      Alignment       =   1  '�E����
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
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
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label UserLabel 
      Alignment       =   1  '�E����
      BackColor       =   &H00FF8080&
      Caption         =   "User Name"
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
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   255
      Width           =   1620
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmLogin.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@���O�C���\���t�H�[��
' �@�{���W���[���̓��O�C���\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Public UserName As String ''���[�U�[���i�[
Public UserPassword As String ''���[�U�[�p�X���[�h�i�[
Public SupervisorName As String ''�Ǘ��Җ��i�[
Public SupervisorPassword As String ''�Ǘ��҃p�X���[�h�i�[

Public LoginSucceeded As Boolean ''���O�C�����t���O�i�[

Public User As String ''���[�U�[���i�[

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
Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
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
Private Sub cmdOK_Click()
    'If UserName = imtxtUserName And UserPassword = imtxtPassword Then
    '    LoginSucceeded = True
    '    User = "User"
    '    Me.Hide
    'ElseIf SupervisorName = imtxtUserName And SupervisorPassword = imtxtPassword Then
    If SupervisorName = imtxtUserName And SupervisorPassword = imtxtPassword Then
        LoginSucceeded = True
        User = "Supervisor"
        Me.Hide
    Else
        Dim fMsg As Object
        Set fMsg = New Message
        fMsg.MsgText = "�߽ܰ�ނ��Ⴂ�܂��B �ēx���͂��Ă��������B"
        fMsg.AutoDelete = True
        Do
            On Error Resume Next
            fMsg.Show vbModal
            If Err.Number = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        Set fMsg = Nothing
        
        imtxtPassword.SetFocus
        imtxtPassword.SelStart = 0
        imtxtPassword.SelLength = Len(imtxtPassword.Text)
    End If
    
End Sub

' @(f)
'
' �@�\      : �p�X���[�h����BOX�L�[����
'
' ������    : ARG1 - �L�[�R�[�h
'             ARG2 - �V�t�g�t���O
'
' �Ԃ�l    :
'
' �@�\����  : �p�X���[�h����BOX�L�[�������̏������s���B
'
' ���l      :
'
Private Sub imtxtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

' @(f)
'
' �@�\      : ���[�U�[������BOX�L�[����
'
' ������    : ARG1 - �L�[�R�[�h
'             ARG2 - �V�t�g�t���O
'
' �Ԃ�l    :
'
' �@�\����  : ���[�U�[������BOX�L�[�������̏������s���B
'
' ���l      :
'
Private Sub imtxtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub
