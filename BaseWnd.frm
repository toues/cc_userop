VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm BaseWnd 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "�J���[�`�F�b�N���d�q���V�X�e��"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   -3210
   ClientWidth     =   19080
   Icon            =   "BaseWnd.frx":0000
   LinkTopic       =   "BasedWnd"
   Visible         =   0   'False
   WindowState     =   2  '�ő剻
   Begin VB.PictureBox MainControl 
      Align           =   2  '������
      BackColor       =   &H00808080&
      Height          =   11355
      Index           =   0
      Left            =   0
      ScaleHeight     =   11295
      ScaleWidth      =   19020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -1275
      Width           =   19080
      Begin VB.CommandButton cmdColorIn_Tok 
         BackColor       =   &H0080FFFF&
         Caption         =   "      ���|         �װ����   �����\����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   300
         Style           =   1  '���̨���
         TabIndex        =   2
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdWEBURL_Color_Result_Tok 
         BackColor       =   &H0080FFFF&
         Caption         =   "���|  �J���[���ʈꗗ�@(WEB)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   4920
         Style           =   1  '���̨���
         TabIndex        =   5
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdWEBURL_Color_Result 
         BackColor       =   &H00FFFF80&
         Caption         =   "SKY  �J���[���ʈꗗ�@(WEB)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   9540
         Style           =   1  '���̨���
         TabIndex        =   4
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdColorSlbFail 
         BackColor       =   &H00FFFF80&
         Caption         =   "   �ُ��   �ꗗ"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   14160
         Style           =   1  '���̨���
         TabIndex        =   3
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdColorIn 
         BackColor       =   &H00FFFF80&
         Caption         =   "      SKY         �װ����   �����\����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   4920
         Style           =   1  '���̨���
         TabIndex        =   1
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdSysCfg 
         BackColor       =   &H0080FF80&
         Caption         =   "�V�X�e���ݒ�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   9540
         Style           =   1  '���̨���
         TabIndex        =   6
         ToolTipText     =   "System setting"
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton cmdSkinIn 
         BackColor       =   &H00FFFF80&
         Caption         =   "�X���u����������"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   300
         Style           =   1  '���̨���
         TabIndex        =   0
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
      Begin VB.CommandButton ShutButton 
         BackColor       =   &H0080FF80&
         Caption         =   "�I��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   36
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         Left            =   14160
         Style           =   1  '���̨���
         TabIndex        =   7
         ToolTipText     =   "System shut down"
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   4500
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  '������
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   10590
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   28019
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "2016/05/02"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "13:12"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox MainControl 
      Align           =   2  '������
      Height          =   510
      Index           =   1
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   19020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10080
      Width           =   19080
      Begin VB.ListBox lstGuidance 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         ItemData        =   "BaseWnd.frx":030A
         Left            =   1140
         List            =   "BaseWnd.frx":030C
         TabIndex        =   8
         Top             =   60
         Width           =   17775
      End
      Begin VB.Label Label4 
         Caption         =   "�K�C�_���X"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�㑵��
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   13
      Top             =   0
      Width           =   19080
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '�㑵��
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   14
      Top             =   0
      Width           =   19080
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  '�㑵��
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   15
      Top             =   0
      Width           =   19080
   End
   Begin VB.Menu mnuSkinIn 
      Caption         =   "�X���u����������"
   End
   Begin VB.Menu mnuDummy0 
      Caption         =   "          "
   End
   Begin VB.Menu mnuColorIn 
      Caption         =   "SKY�J���[�`�F�b�N�����\����"
   End
   Begin VB.Menu mnuDummy1 
      Caption         =   "          "
   End
   Begin VB.Menu mnuColorIn_Tok 
      Caption         =   "���|�J���[�`�F�b�N�����\����"
   End
   Begin VB.Menu mnuDummy5 
      Caption         =   ""
   End
   Begin VB.Menu mnuColorSlbFail 
      Caption         =   "�ُ�񍐈ꗗ"
   End
   Begin VB.Menu mnuDummy2 
      Caption         =   ""
   End
   Begin VB.Menu mnuWEBURL_Color_Result 
      Caption         =   "SKY�J���[���ʈꗗ(WEB)"
   End
   Begin VB.Menu mnuDummy3 
      Caption         =   ""
   End
   Begin VB.Menu mnuWEBURL_Color_Result_Tok 
      Caption         =   "���|�J���[���ʈꗗ(WEB)"
   End
   Begin VB.Menu mnuDummy6 
      Caption         =   ""
   End
   Begin VB.Menu mnuSysCfg 
      Caption         =   "�V�X�e���ݒ�"
   End
   Begin VB.Menu mnuDummy4 
      Caption         =   ""
   End
   Begin VB.Menu mnuShutDown 
      Caption         =   "�I��"
   End
End
Attribute VB_Name = "BaseWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) BaseWnd.Frm                ver 1.00
' @(s)
' �J���[�`�F�b�N���тo�b�@�l�c�h�x�[�X�t�H�[��
' �@�{���W���[���͂l�c�h�x�[�X�t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Public fMDIWnd As Object ''�l�c�h�q�t�H�[���i�[

Dim m_shutDownFlag As Boolean ''�I���t���O�i�[

Dim WSRecFlag As Boolean ''��M�d���L��t���O
Dim WST1OutFlag As Boolean ''�I�[�v�����^�C���A�E�g
Dim WST2OutFlag As Boolean  ''������M���^�C���A�E�g
Dim WSRetryCount As Integer ''�R�l�N�g�p���g���C�J�E���g



' @(f)
'
' �@�\      : �l�c�h�t�H�[���A�N�e�B�u�C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �l�c�h�t�H�[���A�N�e�B�u���̃C�x���g�B
'
' ���l      :
'
Private Sub MDIForm_Activate()
    Me.Caption = "�J���[�`�F�b�N���d�q���V�X�e��" & " Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

' @(f)
'
' �@�\      : �l�c�h�t�H�[���Ǎ��݃C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �l�c�h�t�H�[���Ǎ��ݎ��̃C�x���g�B
'
' ���l      :
'
Private Sub MDIForm_Load()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim varAppStartLog As Variant

    varAppStartLog = Empty
    varAppStartLog = FreeFile
    Open App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt" For Append Access Write As #varAppStartLog
    If IsEmpty(varAppStartLog) = False Then
        Print #varAppStartLog, Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision
    End If
    Close #varAppStartLog
    varAppStartLog = Empty

    MainLogFileNumber = Empty
    MainLogFileNumber = FreeFile               ' ���g�p�̃t�@�C���ԍ����擾���܂��B
        
    '���O�t�@�C�����J��
    Open App.path & "\" & conDefine_LogDirName & "\" & "MAIN_LOG.txt" For Append Access Write As #MainLogFileNumber
    Call MsgLog(conProcNum_MAIN, "*********************************************************")
    Call MsgLog(conProcNum_MAIN, "******************** �l�`�h�m���O�J�n ********************")
    Call MsgLog(conProcNum_MAIN, "*********************************************************")

    '
    'For nI = 0 To 1
    '    LEAD_CAP(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    '    LEAD_CAP(nI).EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    '    LEAD_CAP(nI).EnableTwainEvent = True
    '    'LEAD_LIST(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    '    'LEAD_LIST(nI).EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    '    'LEAD_LIST(nI).EnableTwainEvent = True
    'Next nI
    
    Debug.Print "MDIForm_Load"
    Call InputDataClear

   '���H���}�X�^��񏉊���
    ReDim APNextProcDataSkin(0)
    ReDim APNextProcDataColor(0)

    Call LoadAPSysCfgDataSetting
    'Call LoadAPResDataSetting
    
    '�b�r�n�j�d�s�J�n
    'Call CSTRAN_START
    
    Call MenuUnLock

    '�X�^�b�t���}�X�^��񏉊���
    ReDim APStaffData(0)

    '���������}�X�^��񏉊���
    ReDim APInspData(0)

    '���͎Җ��}�X�^��񏉊���
    ReDim APInpData(0)


    '�X�^�b�t���}�X�^�Ǎ���
    bRet = TRTS0060_Read()
    
    '���������}�X�^�Ǎ���
    bRet = TRTS0062_Read()
    
    '���͎Җ��}�X�^�Ǎ���
    bRet = TRTS0066_Read()
    
End Sub

' @(f)
'
' �@�\      : �c�a�̍ēǂݍ���
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �c�a�̍ēǂݍ��݂��s���B
'
' ���l      :
'
Public Sub ReLoad()
    Dim bRet As Boolean

    Call MsgLog(conProcNum_MAIN, "�c�a�̍ēǂݍ��݂��s���܂��B")   '�K�C�_���X�\��

    Debug.Print "MDIForm_ReLoad"
    Call InputDataClear

    Call MenuUnLock

    '�X�^�b�t���}�X�^��񏉊���
    ReDim APStaffData(0)

    '���������}�X�^��񏉊���
    ReDim APInspData(0)

    '���͎Җ��}�X�^��񏉊���
    ReDim APInpData(0)


    '�X�^�b�t���}�X�^�Ǎ���
    bRet = TRTS0060_Read()
    
    '���������}�X�^�Ǎ���
    bRet = TRTS0062_Read()
    
    '���͎Җ��}�X�^�Ǎ���
    bRet = TRTS0066_Read()
    
End Sub

' @(f)
'
' �@�\      : �X���u���������̓{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u���������̓{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdSkinIn_Click()
    Call mnuSkinIn_Click '�X���u���������̓��j���[
End Sub

' @(f)
'
' �@�\      : �װ���������\���̓{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �װ���������\���̓{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdColorIn_Click()
    Call mnuColorIn_Click '�װ���������\���̓��j���[
End Sub

' @(f)
'
' �@�\      : �ُ�񍐈ꗗ�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ُ�񍐈ꗗ�{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS 2008/09/03
'
Private Sub cmdColorSlbFail_Click()
    Call mnuColorSlbFail_Click '�ُ�񍐈ꗗ���j���[
End Sub

' @(f)
'
' �@�\      : �J���[���ʈꗗ(WEB)�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �J���[���ʈꗗ(WEB)�{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdWEBURL_Color_Result_Click()
    Call mnuWEBURL_Color_Result_Click '�J���[���ʈꗗ(WEB)���j���[
End Sub

' @(f)
'
' �@�\      : �V�X�e���ݒ�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �V�X�e���ݒ�{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdSysCfg_Click()
    Call mnuSysCfg_Click '�V�X�e���ݒ胁�j���[
End Sub

' @(f)
'
' �@�\      : �f�o�b�N���[�h�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �f�o�b�N���[�h�{�^���Ń��j���[���J���B
'
' ���l      : ���W�X�g����nDEBUG_MODE��1�̎��̂݃��j���[���J���B
'          �FCOLORSYS
'
Private Sub mnuDummy0_Click()
    '2002-03-22
    If APSysCfgData.nDEBUG_MODE = 1 Then
        Call mnuDebug_Click
    End If
End Sub


' @(f)
'
' �@�\      : �I���{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �I���{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub ShutButton_Click()
    Call mnuShutDown_Click '�I�����j���[
End Sub

' @(f)
'
' �@�\      : ���j���[�g�p�s�ؑ�
'
' ������    : ARG1 - ���s���j���[��
'
' �Ԃ�l    :
'
' �@�\����  : ���s���j���[�ɉ����đ��̃��j���[
'             ��Ԃ��g�p�s�ɂ���B
'
' ���l      :COLORSYS
'
Private Sub MenuLock(ByVal strMenuName As String)
    
    Select Case strMenuName
        
        Case "mnuSkinIn", "mnuColorIn", "mnuSysCfg", "mnuColorSlbFail", "mnuColorIn_Tok"
            
            mnuSkinIn.Enabled = False
            cmdSkinIn.Enabled = False
            
            mnuColorIn.Enabled = False
            cmdColorIn.Enabled = False
            
            '2008/09/04
            mnuColorSlbFail.Enabled = False
            cmdColorSlbFail.Enabled = False
            
            mnuSysCfg.Enabled = False
            cmdSysCfg.Enabled = False
        
            '2016/04/20 - TAI - S
            mnuColorIn_Tok.Enabled = False
            cmdColorIn_Tok.Enabled = False
            '2016/04/20 - TAI - E
        
    End Select

End Sub

' @(f)
'
' �@�\      : ���j���[�g�p�ؑ�
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���j���[��Ԃ��g�p�ɂ���B
'
' ���l      :COLORSYS
'
Private Sub MenuUnLock()

    mnuSkinIn.Enabled = True
    cmdSkinIn.Enabled = True
    
    mnuColorIn.Enabled = True
    cmdColorIn.Enabled = True
    
    '2008/09/04
    mnuColorSlbFail.Enabled = True
    cmdColorSlbFail.Enabled = True
    
    mnuSysCfg.Enabled = True
    cmdSysCfg.Enabled = True
    
    '2016/04/20 - TAI - S
    mnuColorIn_Tok.Enabled = True
    cmdColorIn_Tok.Enabled = True
    '2016/04/20 - TAI - E
    
End Sub

' @(f)
'
' �@�\      : �X���u���������̓��j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u���������͉�ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuSkinIn_Click()
    Call MenuLock("mnuSkinIn")
        
    '�X���u�������X�g�N���A
    ReDim APSearchListSlbData(0)
    
    ' 20090115 modify by M.Aoyagi �L�[�ύX�����ǉ����
    frmSkinSlbSelWnd.Show vbModeless, Me '�X���u���������͗p�|�X���u�I�����
End Sub

' @(f)
'
' �@�\      : �װ���������\���̓��j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �װ���������\���͉�ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuColorIn_Click()
    Call MenuLock("mnuColorIn")
    
    '2016/04/20 - TAI - S
    '��Əꏊ��"SKY"�ɂ���
    works_sky_tok = WORKS_SKY
    '2016/04/20 - TAI - E

    '�X���u�������X�g�N���A
    ReDim APSearchListSlbData(0)
    
    ' 20090115 modify by M.Aoyagi �L�[�ύX�����ǉ����
    frmColorSlbSelWnd.Show vbModeless, Me '�װ���������\���͗p�|�X���u�I�����
End Sub

' @(f)
'
' �@�\      : �ُ�񍐈ꗗ���j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ُ�񍐈ꗗ��ʂ��J���B
'
' ���l      :COLORSYS 2008/09/03
'
Private Sub mnuColorSlbFail_Click()
    Call MenuLock("mnuColorSlbFail")
    
    '�X���u�������X�g�N���A
    ReDim APSearchListSlbData(0)
    
    frmColorSlbFailWnd.Show vbModeless, Me '�װ���������\���͗p�|�ُ�񍐈ꗗ���
End Sub

' @(f)
'
' �@�\      : �J���[���ʈꗗ(WEB)���j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �J���[���ʈꗗ(WEB)��IE�ŊJ���B
'
' ���l      :COLORSYS 2008/09/03
'
Private Sub mnuWEBURL_Color_Result_Click()
    Dim RetVal
    RetVal = Shell(APSysCfgData.WEBURL_Color_Result, 3)
End Sub

' @(f)
'
' �@�\      : �f�o�b�N���[�h���j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �f�o�b�N���[�h��ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuDebug_Click()
    frmDEBUG.Show vbModeless, Me
End Sub

' @(f)
'
' �@�\      : �V�X�e���ݒ胁�j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �V�X�e���ݒ��ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuSysCfg_Click()
    Dim Result As Boolean
    
    Call MenuLock("mnuSysCfg")

    'Change user
    Result = cUser.ChangeUser
    If Result = False Then
        '���O�I���s��
        Call MsgLog(conProcNum_MAIN, "���O�I���s��")
        Call MenuUnLock
    Else
        '�I�O�I������
        Call MsgLog(conProcNum_MAIN, "���O�I������")
    
        'If Not fMDIWnd Is Nothing Then
        '    Unload fMDIWnd
        '    Set fMDIWnd = Nothing
        'End If
        'Set fMDIWnd = frmSysCfgWnd
        'fMDIWnd.Show
        
        frmSysCfgWnd.SetCallBack Me, CALLBACK_MAIN_RETSYSCFGWND
        frmSysCfgWnd.Show vbModeless, Me '�V�X�e���ݒ���
    End If
End Sub

' @(f)
'
' �@�\      : �I�����j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �I���₢���킹��ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuShutDown_Click()
    Unload fMainWnd '�I���₢���킹
End Sub

' @(f)
'
' �@�\      : �l�c�h�t�H�[���j��������
'
' ������    : ARG1 - �L�����Z���t���O�i�߂�j
'             ARG2 - �j�������[�h�i�߂�j
'
' �Ԃ�l    :
'
' �@�\����  : �l�c�h�t�H�[���j�����̏����B
'
' ���l      : �V�X�e���̏I���₢���킹�O�̏ꍇ�́A
'             �V�X�e���̏I���₢���킹��ʂ�\������B
'             (�R�[���o�b�N�L��)
'
Private Sub MDIForm_QueryUnload(CANCEL As Integer, UnloadMode As Integer)

    If m_shutDownFlag = False Then
        CANCEL = 1
        Dim fmessage As Object
        Set fmessage = New MessageYN
        fmessage.MsgText = "�V�X�e�����I�����܂��B" & vbCrLf & "��낵���ł����H"
        fmessage.AutoDelete = True
        fmessage.SetCallBack Me, CALLBACK_MAIN_SHUTDOWN, True
            Do
                On Error Resume Next
                fmessage.Show vbModeless, Me
                If Err.Number = 0 Then
                    Exit Do
                End If
                DoEvents
            Loop
        Set fmessage = Nothing
    Else
        'frmViewMain.WindowState = 2
        'frmViewMain.Visible = True
        'frmViewMain.ZOrder 0
        fMainWnd.fMDIWnd.WindowState = 2
        fMainWnd.fMDIWnd.Visible = True
        fMainWnd.fMDIWnd.ZOrder 0
        
        Call MsgLog(conProcNum_MAIN, "*********************************************************")
        Call MsgLog(conProcNum_MAIN, "******************** �l�`�h�m���O�I�� ********************")
        Call MsgLog(conProcNum_MAIN, "*********************************************************")
        
        '�l�`�h�m���O�t�@�C���̃N���[�Y
        Close #MainLogFileNumber
        MainLogFileNumber = Empty
        
        If Dir(App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt") <> "" Then
            Call Kill(App.path & "\" & conDefine_LogDirName & "\" & "MAIN_PROCESSING.txt")
        End If
        
        MainModule.UnloadAll
    End If
    
End Sub

' @(f)
'
' �@�\      : �R�[���o�b�N����
'
' ������    : ARG1 - �R�[���o�b�N�ԍ�
'             ARG2 - �R�[���o�b�N�߂�
'
' �Ԃ�l    :
'
' �@�\����  : �R�[���o�b�N�ԍ��Ɩ߂�ɉ����āA���������s���B
'
' ���l      :
'
Public Sub CallBackMessage(ByVal CallNo As Integer, ByVal Result As Integer)
    Dim cnt As Integer
    Dim nI As Integer
    Dim nJ As Integer
    Dim strImageFileName As String
    Dim strMIL_TITLE As String
    Dim strLBLINFO As String
    Dim bRet As Boolean
    Dim strWork As String
    
    Select Case CallNo
    
    '�V�X�e���I��OK
    Case CALLBACK_MAIN_SHUTDOWN
        If Result = CALLBACK_ncResOK Then          'OK
            m_shutDownFlag = True
            On Error Resume Next
            On Error GoTo 0
            Unload Me
        End If
    
    '�X���u���������́|�X���u�I����ʂ���OK
    Case CALLBACK_MAIN_RETSKINSLBSELWND
        If Result = CALLBACK_ncResOK Then          'OK
            frmSkinScanWnd.Show vbModeless, Me '�X���u���������͉�ʂ�
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETSKINSLBSELWND CANCEL"
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If
    
    '�װ���������\���́|�X���u�I����ʂ���OK
    Case CALLBACK_MAIN_RETCOLORSLBSELWND
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND�i���u�{�^�����s�j
            frmDirResWnd.SetCallBack Me, CALLBACK_MAIN_RETDIRRESWND1
            frmDirResWnd.Show vbModeless, Me '���u���e�m�F�^���ʓo�^��ʂֈڍs
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND1
            frmColorScanWnd.Show vbModeless, Me '�װ���������\���͉�ʂ�
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETCOLORSLBSELWND CANCEL"
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If

    '�װ���������\���́|�ُ�񍐈ꗗ��ʂ���OK 2008/09/03
    Case CALLBACK_MAIN_RETCOLORSLBFAILWND
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND�i���u�{�^�����s�j
            frmDirResWnd.SetCallBack Me, CALLBACK_MAIN_RETDIRRESWND2
            frmDirResWnd.Show vbModeless, Me '���u���e�m�F�^���ʓo�^��ʂֈڍs
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND2
            frmColorScanWnd.Show vbModeless, Me '�װ���������\���͉�ʂ�
        Else                                        'CANCEL
            Debug.Print "CALLBACK_MAIN_RETCOLORSLBFAILWND CANCEL"
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
        End If

    '�X���u���������͉�ʂ���OK
    Case CALLBACK_MAIN_RETSKINSCANWND
        If Result = CALLBACK_ncResOK Then          'OK
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdSkinIn_Click '2008/08/04 �X���u���������͊J�n�{�^�����s�i�J�Ԃ��Ή��j
        Else                                        'CANCEL
            ' 20090115 modify by M.Aoyagi �L�[�ύX�����ǉ����
            frmSkinSlbSelWnd.Show vbModeless, Me '�X���u���������́|�X���u�I����ʂ�
        End If
    
    '�װ���������\���͉�ʂ���OK�i1.�X���u�I���V�i���I�j
    Case CALLBACK_MAIN_RETCOLORSCANWND1
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND�i�ۗ��{�^�����s�j
            frmSlbFailScanWnd.SetCallBack Me, CALLBACK_MAIN_RETSLBFAILSCANWND1
            frmSlbFailScanWnd.Show vbModeless, Me '�X���u�ُ�񍐏����͉�ʂֈڍs
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 �װ���������\���̓{�^�����s�i�J�Ԃ��Ή��j
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '���|
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            ' 20090115 modify by M.Aoyagi    �L�[�ύX�����ǉ����
            frmColorSlbSelWnd.Show vbModeless, Me '�װ���������\���́|�X���u�I����ʂ�
        End If
    
    '�װ���������\���͉�ʂ���OK�i2.�ُ�񍐈ꗗ�V�i���I�j
    Case CALLBACK_MAIN_RETCOLORSCANWND2
        If Result = CALLBACK_ncResEXTEND Then          'EXTEND�i�ۗ��{�^�����s�j
            frmSlbFailScanWnd.SetCallBack Me, CALLBACK_MAIN_RETSLBFAILSCANWND2
            frmSlbFailScanWnd.Show vbModeless, Me '�X���u�ُ�񍐏����͉�ʂֈڍs
        ElseIf Result = CALLBACK_ncResOK Then          'OK
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 �ُ�񍐈ꗗ�{�^�����s�i�J�Ԃ��Ή��j
        Else                                        'CANCEL
            frmColorSlbFailWnd.Show vbModeless, Me '�װ���������\���́|�ُ�񍐈ꗗ��ʂ�
        End If
    
    '�X���u�ُ�񍐏����͉�ʂ���OK�i1.�X���u�I���V�i���I�j
    Case CALLBACK_MAIN_RETSLBFAILSCANWND1
        If Result = CALLBACK_ncResOK Then          'OK
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 �װ���������\���̓{�^�����s�i�J�Ԃ��Ή��j
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '���|
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND1
            frmColorScanWnd.Show vbModeless, Me '�װ���������\���͉�ʂ�
        End If
    
    '�X���u�ُ�񍐏����͉�ʂ���OK�i2.�ُ�񍐈ꗗ�V�i���I�j
    Case CALLBACK_MAIN_RETSLBFAILSCANWND2
        If Result = CALLBACK_ncResOK Then          'OK
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 �ُ�񍐈ꗗ�{�^�����s�i�J�Ԃ��Ή��j
        Else                                        'CANCEL
            frmColorScanWnd.SetCallBack Me, CALLBACK_MAIN_RETCOLORSCANWND2
            frmColorScanWnd.Show vbModeless, Me '�װ���������\���͉�ʂ�
        End If
    
    '�V�X�e���ݒ��ʂ���߂�
    Case CALLBACK_MAIN_RETSYSCFGWND
        If Result = CALLBACK_ncResOK Then
            Call MsgLog(conProcNum_MAIN, "�V�X�e���ݒ��ʂ���OK")  '�K�C�_���X�\��
            Call ReLoad
        End If
        
        Call MenuUnLock
        Call RefreshViewMain
        'If Not fMDIWnd Is Nothing Then
        '    Unload fMDIWnd
        '    Set fMDIWnd = Nothing
        'End If
        'Set fMDIWnd = frmViewMain
        'fMDIWnd.Show
    
    '�װ���������\���́|�X���u�I����� -> �X���u�ُ폈�u�w���^���ʓ��͂���OK
    Case CALLBACK_MAIN_RETDIRRESWND1 '�i1.�X���u�I���V�i���I�j
        If Result = CALLBACK_ncResOK Then
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            '2016/04/20 - TAI - S
'            Call cmdColorIn_Click '2008/08/04 �װ���������\���̓{�^�����s�i�J�Ԃ��Ή��j
            If works_sky_tok = WORKS_SKY Then
                Call cmdColorIn_Click               'SKY
            ElseIf works_sky_tok = WORKS_TOK Then
                Call cmdColorIn_Tok_Click           '���|
            End If
            '2016/04/20 - TAI - E
        Else                                        'CANCEL
            frmColorSlbSelWnd.Show vbModeless, Me '�װ���������\���́|�X���u�I����ʂ�
        End If
    
    '�װ���������\���́|�ُ�񍐈ꗗ��� -> �X���u�ُ폈�u�w���^���ʓ��͂���OK
    Case CALLBACK_MAIN_RETDIRRESWND2 '�i2.�ُ�񍐈ꗗ�V�i���I�j
        If Result = CALLBACK_ncResOK Then
            '//���̓f�[�^�N���A���s
            Call InputDataClear
            Call RefreshViewMain
            Call MenuUnLock
            Call cmdColorSlbFail_Click '2008/09/04 �ُ�񍐈ꗗ�{�^�����s�i�J�Ԃ��Ή��j
        Else                                        'CANCEL
            frmColorSlbFailWnd.Show vbModeless, Me '�װ���������\���́|�ُ�񍐈ꗗ��ʂ�
        End If
    
    End Select
End Sub

' @(f)
'
' �@�\      : �f�[�^�N���A����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �f�[�^�N���A�������s���B
'
' ���l      :
'
Private Sub InputDataClear()
    Dim nI As Integer

    Debug.Print "InputDataClear"

    APSlbCont.bProcessing = False
    APSlbCont.nListSelectedIndexP1 = 0
    APSlbCont.nSearchInputModeSelectedIndex = 0
    APSlbCont.strSearchInputSlbNumber = ""

    '�X���u���X�g�f�[�^������
    ReDim APSearchListSlbData(0)

    ReDim APSearchTmpSlbData(0)
    
End Sub

' @(f)
'
' �@�\      : �l�c�h�t�H�[���j������
'
' ������    : ARG1 - �L�����Z���t���O
'
' �Ԃ�l    :
'
' �@�\����  : �l�c�h�t�H�[���j�����s���B
'
' ���l      :
'
Public Sub MDIForm_Unload(CANCEL As Integer)
    'Unload LogoWnd
    Call SaveAPSysCfgDataSetting
    'Call SaveAPResDataSetting
    If Not fMDIWnd Is Nothing Then
        Unload fMDIWnd
        Set fMDIWnd = Nothing
    End If
End Sub

' @(f)
'
' �@�\      : �V�X�e�����Ǎ���
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �V�X�e���������W�X�g������Ǎ��ށB
'
' ���l      :
'
Public Sub LoadAPSysCfgDataSetting()
    Dim nI As Integer
    Dim nCnt As Integer
    
    APSysCfgData.nDEBUG_MODE = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDEBUG_MODE", conDefault_DEBUG_MODE)
    APSysCfgData.nDISP_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDISP_DEBUG", 0)
    APSysCfgData.nFILE_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nFILE_DEBUG", 0)
    APSysCfgData.nTR_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_SKIP", 0)
    APSysCfgData.nDB_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nDB_SKIP", 0)
    APSysCfgData.nSOZAI_DB_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nSOZAI_DB_SKIP", 0)
    APSysCfgData.nSCAN_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nSCAN_SKIP", 0)
    APSysCfgData.nHOSTDATA_DEBUG = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_DEBUG", 0)
    APSysCfgData.nHOSTDATA_SKIP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_SKIP", 0)
    
    '************ COLORSYS
    APSysCfgData.DB_MYUSER_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_DSN", conDefault_DB_MYUSER_DSN)
    APSysCfgData.DB_MYUSER_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_UID", conDefault_DB_MYUSER_UID)
    APSysCfgData.DB_MYUSER_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_PWD", conDefault_DB_MYUSER_PWD)
    APSysCfgData.DB_MYCOMN_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_DSN", conDefault_DB_MYCOMN_DSN)
    APSysCfgData.DB_MYCOMN_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_UID", conDefault_DB_MYCOMN_UID)
    APSysCfgData.DB_MYCOMN_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_PWD", conDefault_DB_MYCOMN_PWD)
    APSysCfgData.DB_SOZAI_DSN = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_DSN", conDefault_DB_SOZAI_DSN)
    APSysCfgData.DB_SOZAI_UID = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_UID", conDefault_DB_SOZAI_UID)
    APSysCfgData.DB_SOZAI_PWD = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_PWD", conDefault_DB_SOZAI_PWD)
    
    APSysCfgData.SHARES_SCNDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_SCNDIR", conDefault_SHARES_SCNDIR)
    APSysCfgData.SHARES_IMGDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_IMGDIR", conDefault_SHARES_IMGDIR)
    ' 20090116 add by M.Aoyagi
    APSysCfgData.SHARES_PDFDIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "SHARES_PDFDIR", conDefault_SHARES_PDFDIR)
    
    APSysCfgData.PHOTOIMG_DIR = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DIR", conDefault_PHOTOIMG_DIR)
    APSysCfgData.PHOTOIMG_DELCHK = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DELCHK", conDefault_PHOTOIMG_DELCHK)
    APSysCfgData.PHOTOIMG_ALLFILES = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_ALLFILES", conDefault_PHOTOIMG_ALLFILES)
    
    '2008/09/01 SystEx. A.K
    APSysCfgData.NowStaffName(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName0", "")
    APSysCfgData.NowStaffName(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName1", "")
    APSysCfgData.NowStaffName(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowStaffName2", "")
    
    APSysCfgData.NowNextProcess(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess0", "")
    APSysCfgData.NowNextProcess(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess1", "")
    APSysCfgData.NowNextProcess(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess2", "")
    
    '2008/09/03 �J���[���ʈꗗ��WEB-URL
    APSysCfgData.WEBURL_Color_Result = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result", conDefault_WEBURL_Color_Result)
    
    '2015/09/15 ���|�J���[���ʈꗗ��WEB-URL
    APSysCfgData.WEBURL_Color_Result_Tok = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result_Tok", conDefault_WEBURL_Color_Result_Tok)
    '************
    
    ' �\�P�b�g�ʐM�Ή�
    APSysCfgData.HOST_IP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "HOST_IP", conDefault_HOST_IP) '�r�W�R��IP
    APSysCfgData.nHOST_PORT = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_PORT", conDefault_nHOST_PORT) '�r�W�R��PORT
    APSysCfgData.nHOST_TOUT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT0", conDefault_nHOST_TOUT0) '�r�W�R���ʐM�^�C���A�E�g�i�S�́j
    APSysCfgData.nHOST_TOUT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT1", conDefault_nHOST_TOUT1) '�r�W�R���ʐM�^�C���A�E�g�i�I�[�v�����j
    APSysCfgData.nHOST_TOUT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT2", conDefault_nHOST_TOUT2) '�r�W�R���ʐM�^�C���A�E�g�i�f�[�^�ʐM�j
    APSysCfgData.nHOST_RETRY = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nHOST_RETRY", conDefault_nHOST_RETRY) '�ʐM���g���C��

    APSysCfgData.TR_IP = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "TR_IP", conDefault_TR_IP) '�e�s�o�ʐM�h�o�A�h���X
    APSysCfgData.nTR_PORT = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_PORT", conDefault_nTR_PORT) '�e�s�o�ʐM�|�[�g�ԍ�
    APSysCfgData.nTR_TOUT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT0", conDefault_nTR_TOUT0) '�e�s�o�ʐM�^�C���A�E�g�i�S�́j
    APSysCfgData.nTR_TOUT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT1", conDefault_nTR_TOUT1) '�e�s�o�ʐM�^�C���A�E�g�i�I�[�v�����j
    APSysCfgData.nTR_TOUT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT2", conDefault_nTR_TOUT2) '�e�s�o�ʐM�^�C���A�E�g�i�f�[�^�ʐM�j
    APSysCfgData.nTR_RETRY = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nTR_RETRY", conDefault_nTR_RETRY) '�ʐM���g���C��
    
    APSysCfgData.nIMAGE_SIZE(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE0", conDefault_nIMAGE_SIZE0)
    APSysCfgData.nIMAGE_SIZE(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE1", conDefault_nIMAGE_SIZE1)
    APSysCfgData.nIMAGE_SIZE(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE2", conDefault_nIMAGE_SIZE2)
    
    APSysCfgData.nIMAGE_ROTATE(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE0", conDefault_nIMAGE_ROTATE0)
    APSysCfgData.nIMAGE_ROTATE(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE1", conDefault_nIMAGE_ROTATE1)
    APSysCfgData.nIMAGE_ROTATE(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE2", conDefault_nIMAGE_ROTATE2)
    
    APSysCfgData.nIMAGE_LEFT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT0", conDefault_nIMAGE_LEFT0)
    APSysCfgData.nIMAGE_TOP(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP0", conDefault_nIMAGE_TOP0)
    APSysCfgData.nIMAGE_WIDTH(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH0", conDefault_nIMAGE_WIDTH0)
    APSysCfgData.nIMAGE_HEIGHT(0) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT0", conDefault_nIMAGE_HEIGHT0)
    
    APSysCfgData.nIMAGE_LEFT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT1", conDefault_nIMAGE_LEFT1)
    APSysCfgData.nIMAGE_TOP(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP1", conDefault_nIMAGE_TOP1)
    APSysCfgData.nIMAGE_WIDTH(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH1", conDefault_nIMAGE_WIDTH1)
    APSysCfgData.nIMAGE_HEIGHT(1) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT1", conDefault_nIMAGE_HEIGHT1)
    
    APSysCfgData.nIMAGE_LEFT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT2", conDefault_nIMAGE_LEFT2)
    APSysCfgData.nIMAGE_TOP(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP2", conDefault_nIMAGE_TOP2)
    APSysCfgData.nIMAGE_WIDTH(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH2", conDefault_nIMAGE_WIDTH2)
    APSysCfgData.nIMAGE_HEIGHT(2) = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT2", conDefault_nIMAGE_HEIGHT2)
    
    If IsDEBUG("SCAN") Then
        APSysCfgData.nIMAGE_LEFT(0) = conDefault_nIMAGE_DEB_LEFT0
        APSysCfgData.nIMAGE_TOP(0) = conDefault_nIMAGE_DEB_TOP0
        APSysCfgData.nIMAGE_WIDTH(0) = conDefault_nIMAGE_DEB_WIDTH0
        APSysCfgData.nIMAGE_HEIGHT(0) = conDefault_nIMAGE_DEB_HEIGHT0
        
        APSysCfgData.nIMAGE_LEFT(1) = conDefault_nIMAGE_DEB_LEFT1
        APSysCfgData.nIMAGE_TOP(1) = conDefault_nIMAGE_DEB_TOP1
        APSysCfgData.nIMAGE_WIDTH(1) = conDefault_nIMAGE_DEB_WIDTH1
        APSysCfgData.nIMAGE_HEIGHT(1) = conDefault_nIMAGE_DEB_HEIGHT1
    
        APSysCfgData.nIMAGE_LEFT(2) = conDefault_nIMAGE_DEB_LEFT2
        APSysCfgData.nIMAGE_TOP(2) = conDefault_nIMAGE_DEB_TOP2
        APSysCfgData.nIMAGE_WIDTH(2) = conDefault_nIMAGE_DEB_WIDTH2
        APSysCfgData.nIMAGE_HEIGHT(2) = conDefault_nIMAGE_DEB_HEIGHT2
    End If
    
    '���H���}�X�^�[�Ǎ�(SKIN)
    nCnt = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", 0)
    ReDim APNextProcDataSkin(0)
    For nI = 1 To nCnt
        APNextProcDataSkin(nI - 1).inp_NextProc = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), "")
        ReDim Preserve APNextProcDataSkin(UBound(APNextProcDataSkin) + 1)
    Next nI

    If UBound(APNextProcDataSkin) = 0 Then
        ReDim APNextProcDataSkin(7)
        APNextProcDataSkin(0).inp_NextProc = ""
        APNextProcDataSkin(1).inp_NextProc = "�������"
        APNextProcDataSkin(2).inp_NextProc = "SLG����"
        APNextProcDataSkin(3).inp_NextProc = "���|����"
        APNextProcDataSkin(4).inp_NextProc = "SKY�ؒf"
        APNextProcDataSkin(5).inp_NextProc = "�\�[�L���O"
        APNextProcDataSkin(6).inp_NextProc = "�w�����ɂĕʓr�w��(�ۗ�����)"
    End If

    '���H���}�X�^�[�Ǎ�(COLOR)
    nCnt = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", 0)
    ReDim APNextProcDataColor(0)
    For nI = 1 To nCnt
        APNextProcDataColor(nI - 1).inp_NextProc = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), "")
        ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
    Next nI

    If UBound(APNextProcDataColor) = 0 Then
        ReDim APNextProcDataColor(6)
        APNextProcDataColor(0).inp_NextProc = ""
        APNextProcDataColor(1).inp_NextProc = "�������"
        APNextProcDataColor(2).inp_NextProc = "�������ăJ���["
        APNextProcDataColor(3).inp_NextProc = "���U�O�d�㌤��"
        APNextProcDataColor(4).inp_NextProc = "SKY�ؒf"
        APNextProcDataColor(5).inp_NextProc = "�w�����ɂĕʓr�w��(�ۗ�����)"
    End If

    ''�ʌ��׃��X�g���i�X���u���j
    ReDim APFaultFaceSkin(13)
    APFaultFaceSkin(0).strCode = ""
    APFaultFaceSkin(0).strName = ""
    APFaultFaceSkin(1).strCode = "����"
    APFaultFaceSkin(1).strName = "�c����"
    APFaultFaceSkin(2).strCode = "ֺ��"
    APFaultFaceSkin(2).strName = "������"
    APFaultFaceSkin(3).strCode = "��Ű"
'    APFaultFaceSkin(3).strName = "�R�[�i�[������"
    APFaultFaceSkin(3).strName = "��Ű������"
    APFaultFaceSkin(4).strCode = "�۶�"
    APFaultFaceSkin(4).strName = "�m���J�~"
    APFaultFaceSkin(5).strCode = "�ï�"
'    APFaultFaceSkin(5).strName = "�X�e�B�b�L���O"
    APFaultFaceSkin(5).strName = "�è��ݸ�"
    APFaultFaceSkin(6).strCode = "��ذ"
'    APFaultFaceSkin(6).strName = "�u���[�f�B���O"
    APFaultFaceSkin(6).strName = "��ذ�ިݸ�"
    APFaultFaceSkin(7).strCode = "����"
'    APFaultFaceSkin(7).strName = "�f�v���b�V����"
    APFaultFaceSkin(7).strName = "����گ���"
    APFaultFaceSkin(8).strCode = "Ƽޭ"
    APFaultFaceSkin(8).strName = "2�d��"
    APFaultFaceSkin(9).strCode = "����"
    APFaultFaceSkin(9).strName = "�i�p"
    APFaultFaceSkin(10).strCode = "��ͺ"
    APFaultFaceSkin(10).strName = "�c����"
    APFaultFaceSkin(11).strCode = "�޸�"
    APFaultFaceSkin(11).strName = "�U�N����"
    APFaultFaceSkin(12).strCode = "˹��"
    APFaultFaceSkin(12).strName = "�Ђ�����"

    ''�������׃��X�g���i�X���u���j
    ReDim APFaultInsideSkin(5)
    APFaultInsideSkin(0).strCode = ""
    APFaultInsideSkin(0).strName = ""
    APFaultInsideSkin(1).strCode = "Ų��"
    APFaultInsideSkin(1).strName = "��������"
    APFaultInsideSkin(2).strCode = "˹��"
    APFaultInsideSkin(2).strName = "�Ђ�����"
    APFaultInsideSkin(3).strCode = "����"
'    APFaultInsideSkin(3).strName = "�Z���^�[�|���V�e�Bor���S�ΐ�"
    APFaultInsideSkin(3).strName = "������ۼè"
    APFaultInsideSkin(4).strCode = "�Ƹ�"
'    APFaultInsideSkin(4).strName = "���j��������"
    APFaultInsideSkin(4).strName = "�Ƃ�������"

    ''�ʌ��׃��X�g���i�J���[�`�F�b�N�j
    ReDim APFaultFaceColor(11)
    APFaultFaceColor(0).strCode = ""
    APFaultFaceColor(0).strName = ""
    APFaultFaceColor(1).strCode = "����"
    APFaultFaceColor(1).strName = "�c����"
    APFaultFaceColor(2).strCode = "ֺ��"
    APFaultFaceColor(2).strName = "������"
    APFaultFaceColor(3).strCode = "��Ű"
    'APFaultFaceColor(3).strName = "�R�[�i�[������"
    APFaultFaceColor(3).strName = "��Ű������"
    APFaultFaceColor(4).strCode = "����"
    APFaultFaceColor(4).strName = "�s���z�[��"
    APFaultFaceColor(5).strCode = "��"
    APFaultFaceColor(5).strName = "�����r"
    APFaultFaceColor(6).strCode = "�۶�"
    APFaultFaceColor(6).strName = "����c��"
    APFaultFaceColor(7).strCode = "ح��"
    APFaultFaceColor(7).strName = "���E�_��"
    APFaultFaceColor(8).strCode = "�޸�"
    APFaultFaceColor(8).strName = "�U�N����"
    APFaultFaceColor(9).strCode = "˹��"
    APFaultFaceColor(9).strName = "�Ђ�����"
    '2016/04/20 - TAI - S
    APFaultFaceColor(10).strCode = "̶��"
    APFaultFaceColor(10).strName = "�[�x��"
    '2016/04/20 - TAI - E

    ''���u��ԃ��X�g
    ReDim APDirRes_Stat(2)
    APDirRes_Stat(0).inp_DirRes_StatCode = ""
    APDirRes_Stat(0).inp_DirRes_Stat = ""
    
    APDirRes_Stat(1).inp_DirRes_StatCode = "1"
    APDirRes_Stat(1).inp_DirRes_Stat = "1:����"


    ''���u���ʃ��X�g
    ReDim APDirRes_Res(2)
    APDirRes_Res(0).inp_DirRes_ResCode = ""
    APDirRes_Res(0).inp_DirRes_Res = ""
    
    APDirRes_Res(1).inp_DirRes_ResCode = "1"
    APDirRes_Res(1).inp_DirRes_Res = "1:�s�K���L��"

End Sub

' @(f)
'
' �@�\      : �V�X�e�����ۑ�
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �V�X�e���������W�X�g���ɕۑ�����B
'
' ���l      :
'
Public Sub SaveAPSysCfgDataSetting()
    Dim nI As Integer
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDEBUG_MODE", APSysCfgData.nDEBUG_MODE
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDISP_DEBUG", APSysCfgData.nDISP_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nFILE_DEBUG", APSysCfgData.nFILE_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_SKIP", APSysCfgData.nTR_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nDB_SKIP", APSysCfgData.nDB_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nSOZAI_DB_SKIP", APSysCfgData.nSOZAI_DB_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nSCAN_SKIP", APSysCfgData.nSCAN_SKIP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_DEBUG", APSysCfgData.nHOSTDATA_DEBUG
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOSTDATA_SKIP", APSysCfgData.nHOSTDATA_SKIP
    
    '************ COLORSYS
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_IP", APSysCfgData.DB_MYUSER_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_UID", APSysCfgData.DB_MYUSER_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYUSER_PWD", APSysCfgData.DB_MYUSER_PWD
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_IP", APSysCfgData.DB_MYCOMN_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_UID", APSysCfgData.DB_MYCOMN_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_MYCOMN_PWD", APSysCfgData.DB_MYCOMN_PWD
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_IP", APSysCfgData.DB_SOZAI_DSN
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_UID", APSysCfgData.DB_SOZAI_UID
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "DB_SOZAI_PWD", APSysCfgData.DB_SOZAI_PWD
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_SCNDIR", APSysCfgData.SHARES_SCNDIR
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_IMGDIR", APSysCfgData.SHARES_IMGDIR
    ' 20090116 add by M.Aoyagi
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "SHARES_PDFDIR", APSysCfgData.SHARES_PDFDIR
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DIR", APSysCfgData.PHOTOIMG_DIR
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_DELCHK", APSysCfgData.PHOTOIMG_DELCHK
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "PHOTOIMG_ALLFILES", APSysCfgData.PHOTOIMG_ALLFILES
    
    '2008/09/01 SystEx. A.K
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName0", APSysCfgData.NowStaffName(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName1", APSysCfgData.NowStaffName(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowStaffName2", APSysCfgData.NowStaffName(2)
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess0", APSysCfgData.NowNextProcess(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess1", APSysCfgData.NowNextProcess(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NowNextProcess2", APSysCfgData.NowNextProcess(2)
    
    '2008/09/03 �J���[���ʈꗗ��WEB-URL
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result", APSysCfgData.WEBURL_Color_Result
    
    '2015/09/15 �J���[���ʈꗗ��WEB-URL
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "WEBURL_Color_Result_Tok", APSysCfgData.WEBURL_Color_Result_Tok
    '************
    
    ' �\�P�b�g�ʐM�Ή�
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "HOST_IP", APSysCfgData.HOST_IP '�r�W�R��IP
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_PORT", APSysCfgData.nHOST_PORT '�r�W�R��PORT
     For nI = 0 To 2
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_TOUT" & CStr(nI), APSysCfgData.nHOST_TOUT(nI) '�r�W�R���ʐM�^�C���A�E�g
    Next nI
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nHOST_RETRY", APSysCfgData.nHOST_RETRY '�ʐM���g���C��
 
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "TR_IP", APSysCfgData.TR_IP '�e�s�o�ʐM�h�o�A�h���X
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_PORT", APSysCfgData.nTR_PORT '�e�s�o�ʐM�|�[�g�ԍ�
    For nI = 0 To 2
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_TOUT" & CStr(nI), APSysCfgData.nTR_TOUT(nI) '�e�s�o�ʐM�^�C���A�E�g
    Next nI
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nTR_RETRY", APSysCfgData.nTR_RETRY '�ʐM���g���C��
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE0", APSysCfgData.nIMAGE_SIZE(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE1", APSysCfgData.nIMAGE_SIZE(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_SIZE2", APSysCfgData.nIMAGE_SIZE(2)
    
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE0", APSysCfgData.nIMAGE_ROTATE(0)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE1", APSysCfgData.nIMAGE_ROTATE(1)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_ROTATE2", APSysCfgData.nIMAGE_ROTATE(2)

    If IsDEBUG("SCAN") = False Then
        For nI = 0 To 2
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_LEFT" & CStr(nI), APSysCfgData.nIMAGE_LEFT(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_TOP" & CStr(nI), APSysCfgData.nIMAGE_TOP(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_WIDTH" & CStr(nI), APSysCfgData.nIMAGE_WIDTH(nI)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nIMAGE_HEIGHT" & CStr(nI), APSysCfgData.nIMAGE_HEIGHT(nI)
        Next nI
    End If

    '���H���}�X�^�[�ۑ�(SKIN)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", UBound(APNextProcDataSkin)
    For nI = 1 To UBound(APNextProcDataSkin)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), APNextProcDataSkin(nI - 1).inp_NextProc
    Next nI

    '���H���}�X�^�[�ۑ�(COLOR)
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
    For nI = 1 To UBound(APNextProcDataColor)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
    Next nI

End Sub

' @(f)
'
' �@�\      : ���C����ʍĕ`��v��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���C����ʂ̍ĕ`���v������B
'
' ���l      :
'
Public Sub ReqRefreshViewMain()
    Call RefreshViewMain
End Sub

' @(f)
'
' �@�\      : ���C����ʍĕ`��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���C����ʂ̍ĕ`���v������B
'
' ���l      : �l�c�h�q�t�H�[�������݂��鎞�̂ݗv������B
'
Private Sub RefreshViewMain()
    If Not fMDIWnd Is Nothing Then
        If fMDIWnd.Name = "frmViewMain" Then
            Call fMainWnd.fMDIWnd.RefreshViewMain
        End If
    End If
End Sub

'2016/04/20 - TAI - S

' @(f)
'
' �@�\      : ���|�J���[���ʈꗗ(WEB)�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���|�J���[���ʈꗗ(WEB)�{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdWEBURL_Color_Result_Tok_Click()
    Call mnuWEBURL_Color_Result_Tok_Click '�J���[���ʈꗗ(WEB)���j���[
End Sub


' @(f)
'
' �@�\      : ���|�J���[���ʈꗗ(WEB)���j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���|�J���[���ʈꗗ(WEB)��IE�ŊJ���B
'
' ���l      :COLORSYS
'
Private Sub mnuWEBURL_Color_Result_Tok_Click()
    Dim RetVal
    RetVal = Shell(APSysCfgData.WEBURL_Color_Result_Tok, 3)
End Sub

' @(f)
'
' �@�\      : ���|�װ���������\���̓{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���|�װ���������\���̓{�^���Ń��j���[���J���B
'
' ���l      :COLORSYS
'
Private Sub cmdColorIn_Tok_Click()
    Call mnuColorIn_Tok_Click '�װ���������\���̓��j���[
End Sub


' @(f)
'
' �@�\      : ���|�װ���������\���̓��j���[
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���|�װ���������\���͉�ʂ��J���B
'
' ���l      :COLORSYS
'
Private Sub mnuColorIn_Tok_Click()
    Call MenuLock("mnuColorIn_Tok")
    
    '2016/04/20 - TAI - S
    '��Əꏊ��"���|"�ɂ���
    works_sky_tok = WORKS_TOK
    '2016/04/20 - TAI - E

    '�X���u�������X�g�N���A
    ReDim APSearchListSlbData(0)
    
    frmColorSlbSelWnd.Show vbModeless, Me '�װ���������\���͗p�|�X���u�I�����
End Sub



'2016/04/20 - TAI - E

