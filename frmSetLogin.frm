VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSetLogin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�V�X�e���ݒ�p�@�p�X���[�h�ݒ�"
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
   StartUpPosition =   2  '��ʂ̒���
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�L�����Z��"
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
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H00FF8080&
      Caption         =   "Confirm Password"
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
      Left            =   3180
      TabIndex        =   7
      Top             =   180
      Width           =   1335
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
      Left            =   135
      TabIndex        =   6
      Top             =   525
      Width           =   2385
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
' �J���[�`�F�b�N���тo�b�@���O�C���o�^�t�H�[��
' �@�{���W���[���̓��O�C���o�^�t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Public LoginSucceeded As Boolean ''�o�^�����t���O

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
    Unload Me
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
    If imtxtUserName = "" Then
        MsgBox "���[�U�[�������͂���Ă��܂���B������x���͂��Ă�������!", , "۸޵�"
        imtxtPassword.SetFocus
        imtxtPassword.SelStart = 0
        imtxtPassword.SelLength = Len(imtxtPassword.Text)
    End If
    If imtxtPassword <> "" And imtxtPassword = imtxtPassword2 Then
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "�߽ܰ�ނ�����������܂���B������x���͂��Ă������� !", , "۸޵�"
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
' �@�\      : �p�X���[�h�Q����BOX�L�[����
'
' ������    : ARG1 - �L�[�R�[�h
'             ARG2 - �V�t�g�t���O
'
' �Ԃ�l    :
'
' �@�\����  : �p�X���[�h�Q����BOX�L�[�������̏������s���B
'
' ���l      :
'
Private Sub imtxtPassword2_KeyDown(KeyCode As Integer, Shift As Integer)
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
