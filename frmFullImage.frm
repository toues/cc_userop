VERSION 5.00
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX12N.OCX"
Begin VB.Form frmFullImage 
   BackColor       =   &H00C0FFC0&
   Caption         =   "�C���[�W�S�̕\��"
   ClientHeight    =   9855
   ClientLeft      =   855
   ClientTop       =   1125
   ClientWidth     =   14985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   15500
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   Begin VB.CommandButton cmdOK 
      Caption         =   "�߂�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13680
      TabIndex        =   0
      Top             =   9360
      Width           =   1215
   End
   Begin LEADLib.LEAD LEAD1 
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14955
      _Version        =   65539
      _ExtentX        =   26379
      _ExtentY        =   16325
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   613
      ScaleWidth      =   993
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
End
Attribute VB_Name = "frmFullImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmFullImage.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�C���[�W�S�̕\���t�H�[��
' �@�{���W���[���̓C���[�W�S�̕\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[

Option Explicit

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
' �@�\      : ����{�^��
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ����{�^�������B
'
' ���l      :
'
Private Sub cmdPrint_Click(Index As Integer)
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

    LEAD1.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD1.EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    LEAD1.EnableTwainEvent = True
    
    '�ďo���ɂ��A��������
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''�X���u����������
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
        
        Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_COLOR)
        
        Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SLBFAIL)
    
    End Select
    
End Sub

' @(f)
'
' �@�\      : �t�H�[�����T�C�Y
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[�����T�C�Y���̏������s���B
'
' ���l      :
'
Private Sub Form_Resize()
    If (Me.Height - 1100) > 0 Then
        LEAD1.Height = Me.Height - 1000
        LEAD1.Width = Me.Width
        LEAD1.Left = 150
        LEAD1.Top = 0
        cmdOK.Top = LEAD1.Height + 100
        cmdOK.Left = (Me.Width - 1100)
'        cmdPrint(0).Top = LEAD1.Height + 100
'        cmdPrint(1).Top = LEAD1.Height + 100
'        cmdPrint(0).Left = (cmdOK.Left - cmdPrint(0).Width) - 100
'        cmdPrint(1).Left = (cmdPrint(0).Left - cmdPrint(1).Width) - 100
    End If
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

