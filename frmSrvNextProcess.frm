VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSrvNextProcess 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "���H���}�X�^"
   ClientHeight    =   4185
   ClientLeft      =   720
   ClientTop       =   900
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2340
      TabIndex        =   3
      Top             =   3420
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelNextProcess 
      Caption         =   "�폜"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5040
      TabIndex        =   2
      Top             =   2700
      Width           =   1200
   End
   Begin VB.ListBox lstNextProcess 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      ItemData        =   "frmSrvNextProcess.frx":0000
      Left            =   1020
      List            =   "frmSrvNextProcess.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3795
   End
   Begin VB.CommandButton cmdAddNextProcess 
      Caption         =   "�ǉ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5040
      TabIndex        =   1
      Top             =   180
      Width           =   1200
   End
   Begin imText6Ctl.imText imtxtNextProcCode 
      Height          =   375
      Left            =   1020
      TabIndex        =   5
      Top             =   180
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   661
      Caption         =   "frmSrvNextProcess.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSrvNextProcess.frx":0072
      Key             =   "frmSrvNextProcess.frx":0090
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   "Aa9�y"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   32
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
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmSrvNextProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSrvNextProcess.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@���H���}�X�^�ǉ��^�폜�t�H�[��
' �@�{���W���[���͎��H���}�X�^�ǉ��^�폜�t�H�[���Ŏg�p����
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
Private Sub cmdAddNextProcess_Click()
    Dim nI As Integer
    Dim bSearch As Boolean
    
    If Trim(imtxtNextProcCode.Text) = "" Then Exit Sub
    
    cmdAddNextProcess.Enabled = False
    
    bSearch = False
    
    For nI = 1 To lstNextProcess.ListCount
        If Trim(imtxtNextProcCode.Text) = lstNextProcess.List(nI - 1) Then
            bSearch = True
            Exit For
        End If
    Next nI
    
    If bSearch = False Then
        lstNextProcess.AddItem Trim(imtxtNextProcCode.Text)
    End If
        
    cmdAddNextProcess.Enabled = True

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
Private Sub cmdDelNextProcess_Click()

    If lstNextProcess.ListIndex > -1 Then
        lstNextProcess.RemoveItem lstNextProcess.ListIndex
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
    
    Dim nI As Integer

'    ReDim APSysCfgData.Group(0)
'    APSysCfgData.nGroupCount = lstGroup.ListCount
'
'    For nI = 1 To APSysCfgData.nGroupCount
'        APSysCfgData.Group(nI - 1) = lstGroup.List(nI - 1)
'        ReDim Preserve APSysCfgData.Group(UBound(APSysCfgData.Group) + 1)
'    Next nI
'
'    '�N���[�Y���Ƀ��W�X�g���֔��f
'    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nGroupCount", APSysCfgData.nGroupCount
'    For nI = 1 To APSysCfgData.nGroupCount
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "Group" & CStr(nI), APSysCfgData.Group(nI - 1)
'    Next nI
'
'
'    ReDim APSysCfgData.StaffNumber(0)
'    ReDim APSysCfgData.StaffName(0)
'    APSysCfgData.nStaffCount = lstStaff.ListCount
'
'    For nI = 1 To APSysCfgData.nStaffCount
'        APSysCfgData.StaffNumber(nI - 1) = Left(lstStaff.List(nI - 1), InStr(lstStaff.List(nI - 1), ":") - 1)
'        APSysCfgData.StaffName(nI - 1) = Mid(lstStaff.List(nI - 1), InStr(lstStaff.List(nI - 1), ":") + 1)
'        ReDim Preserve APSysCfgData.StaffNumber(UBound(APSysCfgData.StaffNumber) + 1)
'        ReDim Preserve APSysCfgData.StaffName(UBound(APSysCfgData.StaffName) + 1)
'    Next nI
'
'    '�N���[�Y���Ƀ��W�X�g���֔��f
'    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nStaffCount", APSysCfgData.nStaffCount
'    For nI = 1 To APSysCfgData.nStaffCount
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "StaffNumber" & CStr(nI), APSysCfgData.StaffNumber(nI - 1)
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "StaffName" & CStr(nI), APSysCfgData.StaffName(nI - 1)
'    Next nI
'
'    Unload Me
'
'    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
'    Set cCallBackObject = Nothing
    
    
    '�ďo���ɂ��A��������
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''�X���u����������
            ReDim APNextProcDataSkin(1)
            '****
                '�󔒂̓V�X�e���ŊǗ��i�K���ǉ��j
                APNextProcDataSkin(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataSkin(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataSkin(UBound(APNextProcDataSkin) + 1)
            Next nI
        
            '���H���}�X�^�[�ۑ�(SKIN)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", UBound(APNextProcDataSkin)
            For nI = 1 To UBound(APNextProcDataSkin)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), APNextProcDataSkin(nI - 1).inp_NextProc
            Next nI
        
        Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
            ReDim APNextProcDataColor(1)
            '****
                '�󔒂̓V�X�e���ŊǗ��i�K���ǉ��j
                APNextProcDataColor(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataColor(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
            Next nI

            '���H���}�X�^�[�ۑ�(COLOR)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
            For nI = 1 To UBound(APNextProcDataColor)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
            Next nI
        
        Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
            '****
                '�󔒂̓V�X�e���ŊǗ��i�K���ǉ��j
                APNextProcDataColor(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataColor(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
            Next nI

            '���H���}�X�^�[�ۑ�(COLOR)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
            For nI = 1 To UBound(APNextProcDataColor)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
            Next nI

    End Select
    
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
End Sub

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
    
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResCANCEL
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
    
    lstNextProcess.Clear
    
    '�ďo���ɂ��A��������
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''�X���u����������
            '�󔒂������ҏW�\
            For nI = 2 To UBound(APNextProcDataSkin)
                lstNextProcess.AddItem APNextProcDataSkin(nI - 1).inp_NextProc
            Next nI
        
        Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
            '�󔒂������ҏW�\
            For nI = 2 To UBound(APNextProcDataColor)
                lstNextProcess.AddItem APNextProcDataColor(nI - 1).inp_NextProc
            Next nI
        
        Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
            '�󔒂������ҏW�\
            For nI = 2 To UBound(APNextProcDataColor)
                lstNextProcess.AddItem APNextProcDataColor(nI - 1).inp_NextProc
            Next nI

    End Select

End Sub

' @(f)
'
' �@�\      : ���H������BOX�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���H������BOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub imtxtNextProcCode_GotFocus()
    imtxtNextProcCode.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' �@�\      : ���H������BOX�t�H�[�J�X����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���H������BOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub imtxtNextProcCode_LostFocus()
    imtxtNextProcCode.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' �@�\      : ���H�����X�gBOX�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���H�����X�gBOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub lstNextProcess_GotFocus()
'    lstNextProcess.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' �@�\      : ���H�����X�gBOX�t�H�[�J�X����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���H�����X�gBOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub lstNextProcess_LostFocus()
'    lstNextProcess.BackColor = conDefine_ColorBKLostFocus
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



