VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSrvLocation 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�u��}�X�^"
   ClientHeight    =   4185
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5535
   Begin imText6Ctl.imText imtxtLocName 
      Height          =   315
      Left            =   2700
      TabIndex        =   1
      Top             =   180
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "frmSrvLocation.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSrvLocation.frx":006E
      Key             =   "frmSrvLocation.frx":008C
      BackColor       =   -2147483643
      EditMode        =   0
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   "A9�y"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   16
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imText6Ctl.imText imtxtLocCode 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   180
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   556
      Caption         =   "frmSrvLocation.frx":00D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSrvLocation.frx":013E
      Key             =   "frmSrvLocation.frx":015C
      BackColor       =   -2147483643
      EditMode        =   0
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   "9A"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   4
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelLocation 
      Caption         =   "�폜"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   3060
      Width           =   795
   End
   Begin VB.ListBox lstLocation 
      Height          =   2760
      Left            =   1020
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdAddLocation 
      Caption         =   "�ǉ�"
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "�L��"
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmSrvLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSrvLocation.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�u��}�X�^�ǉ��^�폜�t�H�[��
' �@�{���W���[���͒u��}�X�^�ǉ��^�폜�t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[

' @(f)
'
' �@�\      : �ǉ��{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ǉ��{�^�������B
'
' ���l      :
'
Private Sub cmdAddLocation_Click()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim bSearch As Boolean
    
    If Trim(imtxtLocCode.Text) = "" Or Trim(imtxtLocName.Text) = "" Then Exit Sub
    
    cmdAddLocation.Enabled = False
    
    bSearch = False
    For nI = 1 To UBound(APLocData)
        If Trim(imtxtLocCode.Text) = APLocData(nI - 1).inp_LocCode Then
            bSearch = True
            Exit For
        End If
    Next nI
    
    If bSearch = False Then
        bRet = TRTS0040_Write(False, Trim(imtxtLocCode.Text), Trim(imtxtLocName.Text))
        bRet = TRTS0040_Read()
        Call InitForm
    End If
    
    cmdAddLocation.Enabled = True

End Sub

' @(f)
'
' �@�\      : �폜�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �폜�{�^�������B
'
' ���l      :
'
Private Sub cmdDelLocation_Click()
    Dim bRet As Boolean
    
    If lstLocation.ListIndex > -1 Then
        bRet = TRTS0040_Write(True, APLocData(lstLocation.ListIndex).inp_LocCode, APLocData(lstLocation.ListIndex).inp_LocName)
        bRet = TRTS0040_Read()
        Call InitForm
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
Private Sub cmdOK_Click()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
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
    
    Call InitForm

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
    Dim nI As Integer
    
    lstLocation.Clear
    For nI = 1 To UBound(APLocData)
        lstLocation.AddItem APLocData(nI - 1).inp_LocCode & conDefault_Separator & APLocData(nI - 1).inp_LocName
    Next nI

End Sub

' @(f)
'
' �@�\      : �u��R�[�h����BOX�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u��R�[�h����BOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub imtxtLocCode_GotFocus()
    imtxtLocCode.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' �@�\      : �u��R�[�h����BOX�t�H�[�J�X����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u��R�[�h����BOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub imtxtLocCode_LostFocus()
    imtxtLocCode.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' �@�\      : �u�ꖼ�̓���BOX�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u�ꖼ�̓���BOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub imtxtLocName_GotFocus()
    imtxtLocName.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' �@�\      : �u�ꖼ�̓���BOX�t�H�[�J�X����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u�ꖼ�̓���BOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub imtxtLocName_LostFocus()
    imtxtLocName.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' �@�\      : �u�ꃊ�X�gBOX�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u�ꃊ�X�gBOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub lstLocation_GotFocus()
'    lstLocation.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' �@�\      : �u�ꃊ�X�gBOX�t�H�[�J�X����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �u�ꃊ�X�gBOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub lstLocation_LostFocus()
'    lstLocation.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' �@�\      : �R�[���o�b�N�ݒ�
'
' ������    : ARG1 - �R�[���o�b�N�I�u�W�F�N�g
'             ARG2 - �R�[���o�b�N�h�c
'
' �Ԃ�l    :
'
' �@�\����  : �߂��R�[���o�b�N����ݒ肷��B
'
' ���l      :
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer)
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
End Sub



