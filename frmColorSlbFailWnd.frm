VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmColorSlbFailWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "�J���[�`�F�b�N�����\���́|�ُ�񍐈ꗗ"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   18690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   18690
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdDirRes 
      Caption         =   "���u"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14100
      TabIndex        =   9
      Top             =   9960
      Width           =   1800
   End
   Begin VB.PictureBox PicSigYellow 
      Height          =   315
      Left            =   3720
      Picture         =   "frmColorSlbFailWnd.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   10140
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicSigRed 
      Height          =   375
      Left            =   5160
      Picture         =   "frmColorSlbFailWnd.frx":0644
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   10020
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox PicSigGreen 
      Height          =   315
      Left            =   3720
      Picture         =   "frmColorSlbFailWnd.frx":0E22
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   10500
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "�߂�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   16620
      TabIndex        =   3
      Top             =   9960
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���яC��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11580
      TabIndex        =   2
      Top             =   9960
      Width           =   1800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�\���X�V"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   16620
      TabIndex        =   0
      Top             =   60
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   9075
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   16007
      _Version        =   393216
      Rows            =   21
      Cols            =   14
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_works 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "SKY"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   27.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbl_nMSFlexGrid1_Selected_Row 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�J���[�`�F�b�N�����\���́|�ُ�񍐈ꗗ"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   0
      Width           =   9795
   End
End
Attribute VB_Name = "frmColorSlbFailWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmColorSlbFailWnd.Frm                ver 1.00 ( '2008.09.03 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N�����\���́|�ُ�񍐈ꗗ�\���t�H�[��
' �@�{���W���[���̓J���[�`�F�b�N�����\���́|�ُ�񍐈ꗗ�\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private nMSFlexGrid1_Selected_Row As Integer ''�O���b�h�P�I���s�ԍ��i�[
Private bMouseControl As Boolean ''�}�E�X�R���g���[���t���O�i�[

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
' ���l      :COLORSYS
'
Private Sub cmdCancel_Click()
    
    cmdCANCEL.Enabled = False ''�A�ŋ֎~�I

    Call SlbSelLock(False)
    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBFAILWND, CALLBACK_ncResCANCEL)
    Unload Me
End Sub

Private Sub cmdDirRes_Click()
    Dim bRet As Boolean
    
    cmdDirRes.Enabled = False '�A�ŋ֎~�I
    
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    
    '�X���u�I���`�F�b�N
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Call WaitMsgBox(Me, "�X���u��I�����Ă��������B")
        Exit Sub
    End If

    Select Case APSlbCont.nSearchInputModeSelectedIndex
        Case 0 '�V�K
        Case 1 '�C��
        Case 2 '�폜
            '�����I��
            Exit Sub
    End Select
    
    bRet = ColorSlbData_Load(True)

    If bRet Then
        Call OKcmdDIR '���u��ʊJ�n(unload me)
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
' ���l      :COLORSYS
'
Private Sub cmdOK_Click()
    Dim bRet As Boolean
    Dim MsgWnd As Message
    Set MsgWnd = New Message

    cmdOK.Enabled = False ''�A�ŋ֎~�I

    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row

    APSlbCont.nSearchInputModeSelectedIndex = 1 '�C���Œ�

    '�X���u�I���`�F�b�N
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '�V�K
                'MsgWnd.MsgText = "���ѓ��͂��s���X���u��I�����Ă��������B"
            Case 1 '�C��
                MsgWnd.MsgText = "���яC�����s���X���u��I�����Ă��������B"
            Case 2 '�폜
                'MsgWnd.MsgText = "���э폜���s���X���u��I�����Ă��������B"
        End Select
        MsgWnd.OK.Visible = True
    '    MsgWnd.AutoDelete = True
        Do
            On Error Resume Next
            MsgWnd.Show vbModal
            If Err.Number = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        Set MsgWnd = Nothing

        cmdOK.Enabled = True '�{�^���L��
        Exit Sub
    End If
    
    '2016/04/20 - TAI - S
    '��Ə�Z�b�g
    works_sky_tok = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_works_sky_tok
    '2016/04/20 - TAI - E

    Set MsgWnd = Nothing

    If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
        '�폜
        'Call ColorDataDel_REQ
    Else
        bRet = ColorSlbData_Load(False)

        cmdOK.Enabled = True '�{�^���L��

        If bRet Then
            Select Case APSlbCont.nSearchInputModeSelectedIndex
                Case 0 '�V�K
                    'Call OKcmdOK '���͊J�n(unload me)
                Case 1 '�C��
                    Call OKcmdOK '���͊J�n(unload me)
            End Select
        End If

    End If

End Sub

Private Function ColorSlbData_Load(ByVal bDirResLoad As Boolean) As Boolean
    Dim bRet As Boolean
    Dim strSource As String
    Dim strDestination As String
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    '�X���u�I���`�F�b�N
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        ColorSlbData_Load = False '�G���[
        Exit Function
    End If
        
    '********************************************************************************************
    'DEBUG POINT �V�K���[�h�Ń��X�g�\���̏ꍇ�A�C���Ώۃ��R�[�h�������ɕ\�������̂ŁA
    '���X�g�I����A�V�K�ł͂Ȃ��A�C�������[�U�[���I�񂾏ꍇ�́A������x���[�h���`�F�b�N���A
    '�V�K�^�C���̐ؑւ����K�v
    '********************************************************************************************
    '�V�K���[�h���H
    If APSlbCont.nSearchInputModeSelectedIndex = 0 Then
        '�I�������X���u�͐V�K���H
        If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).sys_wrt_dte = "" Then
            '�V�K���[�h
        Else
            '�ۑ��ς݂ł���ׁA�C�����[�h�Ɏ����ύX
            APSlbCont.nSearchInputModeSelectedIndex = 1
        End If
    End If
    
    
    '�c�a�������������郂�[�h�@�C���^�폜
    If APSlbCont.nSearchInputModeSelectedIndex <> 0 Then

        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 1 '�C��
                MsgWnd.MsgText = "�f�[�^�x�[�X����X���u����Ǎ��ݒ��ł��B" & vbCrLf & "���΂炭���҂����������B"
            Case 2 '�폜
                MsgWnd.MsgText = "�f�[�^�x�[�X����X���u�����폜���ł��B" & vbCrLf & "���΂炭���҂����������B"
        End Select

        MsgWnd.OK.Visible = False
        MsgWnd.Show vbModeless, Me
        MsgWnd.Refresh
    
    End If
    
    '���я����G���A�փf�[�^�R�s�[
    Call init_APResData
    Select Case APSlbCont.nSearchInputModeSelectedIndex
        Case 0 '�V�K
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no ''�X���u�`���[�WNO
            APResData.slb_chno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno ''�X���u�`���[�WNO
            APResData.slb_aino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino ''�X���u����
            APResData.slb_stat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat ''���
            APResData.slb_col_cnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt ''�J���[��
            APResData.slb_ccno = APSozaiData.slb_ccno           ''�X���uCCNO
            APResData.slb_zkai_dte = APSozaiData.slb_zkai_dte   ''�����
            APResData.slb_ksh = APSozaiData.slb_ksh             ''�|��
            APResData.slb_typ = APSozaiData.slb_typ             ''�^
            APResData.slb_uksk = APSozaiData.slb_uksk           ''����
            APResData.slb_wei = APSozaiData.slb_skin_wei        ''�d�ʁi���ޔ��p�j
            APResData.slb_lngth = APSozaiData.slb_lngth         ''����
            APResData.slb_wdth = APSozaiData.slb_wdth           ''��
            APResData.slb_thkns = APSozaiData.slb_thkns         ''����
            
            '2008/09/01 SystEx. A.K �V�K�̏ꍇ�A�O��f�[�^�i�ێ����f�[�^�j���Z�b�g����B
            APResData.slb_wrt_nme = APSysCfgData.NowStaffName(conDefine_SYSMODE_COLOR) '��������
            APResData.slb_nxt_prcs = APSysCfgData.NowNextProcess(conDefine_SYSMODE_COLOR) '���H��
            
            '�J���[�`�F�b�N
            '�V�K�̏ꍇ�́ASCAN�C���[�W������������B�i���ԃt�@�C���̍폜�j
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG"
            '�C���[�W����
            If Dir(strDestination) <> "" Then
                Kill strDestination
            End If
            
            '�X���u�ُ�
            '�V�K�̏ꍇ�́ASCAN�C���[�W������������B�i���ԃt�@�C���̍폜�j
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
            '�C���[�W����
            If Dir(strDestination) <> "" Then
                Kill strDestination
            End If
            
            ' 20090115 add by M.Aoyagi    �摜�����ǉ��̈�
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1
            
        Case 1 '�C��
            '���уf�[�^�Ǎ���
            bRet = TRTS0014_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, _
                                 APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat, _
                                 APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt)
            If UBound(APResTmpData) = 1 Then
                APResData = APResTmpData(0)
            End If
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
            
            '�J���[�`�F�b�N
            '�o�^�ς�SCAN�C���[�W�����邩�H
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG"
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).bAPScanInput Then
                '�o�^�ς�SCAN�C���[�W��Ǎ��� (conDefine_ImageDirName = TEMP)
                strSource = APSysCfgData.SHARES_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                         "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                         "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
                On Error GoTo ColorSlbData_Load_err:
                Call FileCopy(strSource, strDestination)
                On Error GoTo 0
            Else
                '�C���[�W����
                If Dir(strDestination) <> "" Then
                    Kill strDestination
                End If
            End If
            
            '�X���u�ُ�
            '�o�^�ς�SCAN�C���[�W�����邩�H
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).bAPFailScanInput Then
                '�o�^�ς�SCAN�C���[�W��Ǎ��� (conDefine_ImageDirName = TEMP)
                strSource = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                         "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                         "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
                On Error GoTo ColorSlbData_Load_err:
                Call FileCopy(strSource, strDestination)
                On Error GoTo 0
            Else
                '�C���[�W����
                If Dir(strDestination) <> "" Then
                    Kill strDestination
                End If
            End If
            
            '�X���u�ُ�񍐗p
            APResData.fail_host_send = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_send ''�X���u�ُ�񍐗p�@�r�W�R�����M����
            APResData.fail_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte       ''�X���u�ُ�񍐗p�@�L�^��
            APResData.fail_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_tme       ''�X���u�ُ�񍐗p�@�L�^����
            APResData.fail_sys_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_sys_wrt_dte  ''�X���u�ُ�񍐗p�@�o�^��
            APResData.fail_sys_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_sys_wrt_tme        ''�X���u�ُ�񍐗p�@�o�^����
            
            '���u�w��
            APResData.fail_dir_sys_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_dir_sys_wrt_dte ''���u�w���p�@�L�^���i����L�^���j

            '���u����
            APResData.fail_res_host_send = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_send             ''���u���ʗp�@�r�W�R�����M����
            APResData.fail_res_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_wrt_dte       ''���u���ʗp�@�L�^��
            APResData.fail_res_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_wrt_tme       ''���u���ʗp�@�L�^����

            If bDirResLoad Then
                'DirResLoad
                '���u�w���f�[�^�Ǎ���
                bRet = DBDirResData_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, _
                                     APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat, _
                                     APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt)
                
                ReDim APDirResData(0)
                
                If UBound(APDirResTmpData) <> 0 Then
                    ReDim APDirResData(UBound(APDirResTmpData))
                    APDirResData = APDirResTmpData
                End If
                If bRet = False Then
                    Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                    MsgWnd.OK_Close
                    ColorSlbData_Load = False '�G���[
                    Exit Function
                End If
            End If

            ' 20090115 add by M.Aoyagi    �摜�����ǉ��̈�
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1

        Case 2 '�폜
        
            '*********
            '�폜����
            '*********
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
            APResData.slb_stat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
            APResData.slb_col_cnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
            bRet = TRTS0014_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
        
            bRet = TRTS0052_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
        
            bRet = TRTS0016_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
        
            bRet = TRTS0054_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
        
            bRet = TRTS0022_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                ColorSlbData_Load = False '�G���[
                Exit Function
            End If
        
            MsgWnd.OK_Close
            
            '*********
            '�������ʃ��X�g�ĕ\��
            '*********
            Call WaitMsgBox(Me, "�폜����������I�����܂����B")
            Call cmdSearch_Click
            ColorSlbData_Load = True 'OK
            Exit Function
    End Select
    
    '�c�a�������������郂�[�h�@�C���^�폜�i�Ǎ������b�Z�[�W�\���L��j
    If APSlbCont.nSearchInputModeSelectedIndex <> 0 Then
        MsgWnd.OK_Close
    End If
    
    ColorSlbData_Load = True 'OK
    Exit Function
    
ColorSlbData_Load_err:
    Call MsgLog(conProcNum_MAIN, Err.Source + " " + _
        CStr(Err.Number) + Chr(13) + Err.Description)
    
    Call MsgLog(conProcNum_MAIN, "ColorSlbData_Load FILECOPY SO=" & strSource & " DE=" & strDestination)
    Call WaitMsgBox(Me, "�ۑ��ς݃X�L���i�[�C���[�W�t�@�C���̓Ǎ��G���[���������܂����B" & vbCrLf & vbCrLf & "FILECOPY SO=" & strSource & " DE=" & strDestination)
    
    MsgWnd.OK_Close
    On Error GoTo 0
    ColorSlbData_Load = False '�G���[
    Exit Function
    
End Function

' @(f)
'
' �@�\      : �X���u�I�������n�j�I��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u�I�������n�j�ʒm�B
'
' ���l      : �R�[���o�b�N�ɂĂn�j�ʒm��A�����[�h�B
'
Private Sub OKcmdOK()

    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBFAILWND, CALLBACK_ncResOK)
    Unload Me

End Sub

' @(f)
'
' �@�\      : �X���u�I�������n�j�I���Ə��u��ʃ��N�G�X�g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u�I�������n�j�ʒm�B
'
' ���l      : �R�[���o�b�N�ɂĂn�j�ʒm��A�����[�h�B
'
Private Sub OKcmdDIR()

    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBFAILWND, CALLBACK_ncResEXTEND)
    Unload Me

End Sub

' @(f)
'
' �@�\      : �X���u���\���X�V�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u���̌����ƕ\���X�V���s���B
'
' ���l      : �X���u�������ʕ\���G���A
'
Private Sub cmdSearch_Click()
    Dim nWildCard As Integer
    Dim nI As Integer
    Dim nJ As Integer
    Dim nSEARCH_MAX As Integer
    Dim bRet As Boolean
    Dim strSearchSlbNumber As String '���ۂ̌���������
    Dim strTmpSlbNumber As String '��r�p
    Dim bCmp As Boolean '��r�p
    Dim strChkChar As String
    
    Dim nSlb_Col_Cnt_MAX As Integer
    Dim nFirstDataIndex As Integer
    
    Dim bNoRecord As Boolean '2008/08/30 A.K
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    '�Č������͏�����
    strSearchSlbNumber = ""
    Call InitMSFlexGrid1
    
    nWildCard = 0
    '�n�C�t���f�|�f������Ď��ۂ̌���������փZ�b�g
'    strSearchSlbNumber = ConvSearchSlbNumber(imTextSearchSlbNumber.Text)
    strSearchSlbNumber = ConvSearchSlbNumber("**")
    
'    '���̓��[�h
'    If OptInputMode(0).Value Then '�V�K
'        APSlbCont.nSearchInputModeSelectedIndex = 0 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
'
'        If OptStatus(0).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 0 '����
'        ElseIf OptStatus(1).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 1 '1ht��
'        ElseIf OptStatus(2).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 2 '2ht��
'        ElseIf OptStatus(3).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 3 '3ht��
'        ElseIf OptStatus(4).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 4 '4ht��
'        ElseIf OptStatus(5).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 5 '5ht��
'        ElseIf OptStatus(6).Value Then
'            APSlbCont.nSearchInputStatusSelectedIndex = 6 '6ht��
'        End If
'
'    ElseIf OptInputMode(1).Value Then '�C��
'        APSlbCont.nSearchInputModeSelectedIndex = 1 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
'        APSlbCont.nSearchInputStatusSelectedIndex = 0 '�����i�g�p���Ȃ��j
'    ElseIf OptInputMode(2).Value Then '�폜
'        APSlbCont.nSearchInputModeSelectedIndex = 2 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
'        APSlbCont.nSearchInputStatusSelectedIndex = 0 '�����i�g�p���Ȃ��j
'    End If
    
    
    '****************************
    APSlbCont.nSearchInputModeSelectedIndex = 1 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
    APSlbCont.nSearchInputStatusSelectedIndex = 0 '�����i�g�p���Ȃ��j
    
    
    
    nWildCard = InStr(1, strSearchSlbNumber, "%", vbTextCompare)
    
    'RIAL
    ReDim APSearchListSlbData(0)
    
'    '�V�K���[�h�i�����Ń��C���h�J�[�h�s�j
'    If OptInputMode(0).Value Then
'        '�󔒎w��͕s�B
'        If LenB(imTextSearchSlbNumber.Text) = 0 Then
'            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�s�B
'        If nWildCard <> 0 Then
'            Call WaitMsgBox(Me, "�V�K���[�h�Ń��C���h�J�[�h�̎w��͏o���܂���B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '�擪����T�����܂ł́A0����9�ȊO��s�B
'        For nI = 1 To 5
'            If nI > Len(strSearchSlbNumber) Then Exit For
'            strChkChar = Mid(strSearchSlbNumber, nI, 1)
'            If strChkChar >= "0" And strChkChar <= "9" Then
'                'OK
'            Else
'                'NG
'                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9�ȊO�̎w��͏o���܂���B")
'                imTextSearchSlbNumber.SetFocus
'                Exit Sub
'            End If
'        Next nI
'
'    '�C�����[�h�i�����Ń��C���h�J�[�h�j
'    ElseIf OptInputMode(1).Value Then
'        '�󔒎w��͕s�B
'        If LenB(imTextSearchSlbNumber.Text) = 0 Then
'            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�P�݂͕̂s�B
'        If strSearchSlbNumber = "%" Then
'            Call WaitMsgBox(Me, "���C���h�J�[�h�̎w����@������������܂���B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '�擪����T�����܂ł́A0����9,*�ȊO��s�B
'        For nI = 1 To 5
'            If nI > Len(strSearchSlbNumber) Then Exit For
'            strChkChar = Mid(strSearchSlbNumber, nI, 1)
'            If strChkChar >= "0" And strChkChar <= "9" Then
'                'OK
'            ElseIf strChkChar = "%" Then
'                'OK
'            Else
'                'NG
'                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9,*�ȊO�̎w��͏o���܂���B")
'                imTextSearchSlbNumber.SetFocus
'                Exit Sub
'            End If
'        Next nI
'
'    '�폜���[�h�i�����Ń��C���h�J�[�h�j
'    ElseIf OptInputMode(2).Value Then
'        '�󔒎w��͕s�B
'        If LenB(imTextSearchSlbNumber.Text) = 0 Then
'            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�P�݂͕̂s�B
'        If strSearchSlbNumber = "%" Then
'            Call WaitMsgBox(Me, "���C���h�J�[�h�̎w����@������������܂���B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
'        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
'            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
'            imTextSearchSlbNumber.SetFocus
'            Exit Sub
'        End If
'        '�擪����T�����܂ł́A0����9,*�ȊO��s�B
'        For nI = 1 To 5
'            If nI > Len(strSearchSlbNumber) Then Exit For
'            strChkChar = Mid(strSearchSlbNumber, nI, 1)
'            If strChkChar >= "0" And strChkChar <= "9" Then
'                'OK
'            ElseIf strChkChar = "%" Then
'                'OK
'            Else
'                'NG
'                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9,*�ȊO�̎w��͏o���܂���B")
'                imTextSearchSlbNumber.SetFocus
'                Exit Sub
'            End If
'        Next nI
'
'    End If
    
    MsgWnd.MsgText = "�f�[�^�x�[�X���������ł��B" & vbCrLf & "���΂炭���҂����������B"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
    '��������
'    nSEARCH_MAX = APSysCfgData.nSEARCH_MAX(APSlbCont.nSearchInputModeSelectedIndex)
    'bRet = DBSkinSlbSearchRead(APSlbCont.nSearchInputModeSelectedIndex, nSEARCH_MAX, strSearchSlbNumber)
    
    '�i�����L���͈͂͐�������j
    'bRet = DBSkinSlbSearchRead(APSlbCont.nSearchInputModeSelectedIndex, nSEARCH_MAX, APSysCfgData.nSEARCH_RANGE, strSearchSlbNumber)
    
    '�i�����L���͈͂�9999�������j
    bRet = DBColorSlbSearchRead(1, 0, 9999, strSearchSlbNumber) '1:�ُ�񍐈ꗗ����
        
    '�������ʂ��Z�b�g
    If bRet Then
        
        ReDim APSearchListSlbData(0)
        nJ = 0
        
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '�V�K
                bCmp = False
                nJ = 0
                nSlb_Col_Cnt_MAX = 0
                nFirstDataIndex = 0
                For nI = 0 To UBound(APSearchTmpSlbData) - 1
                    strTmpSlbNumber = APSearchTmpSlbData(nI).slb_no
                    '����No�D���r
                    If strTmpSlbNumber = strSearchSlbNumber Then
                        '��Ԃ��r
                        If CInt(APSearchTmpSlbData(nI).slb_stat) = APSlbCont.nSearchInputStatusSelectedIndex Then
                            bCmp = True '����
                            '*****************************************************
                            'APSlbCont.nSearchInputModeSelectedIndex = 1 '�V�K�ˏC��
                            '*****************************************************
                            'Exit For
                            '�J���[�񐔂̍ő吔���擾
                            If nSlb_Col_Cnt_MAX < CInt(APSearchTmpSlbData(nI).slb_col_cnt) Then
                                nSlb_Col_Cnt_MAX = CInt(APSearchTmpSlbData(nI).slb_col_cnt)
                            End If
                            If CInt(APSearchTmpSlbData(nI).slb_stat) = 1 Then
                                nFirstDataIndex = nI
                            End If
                        End If
                    End If
                Next nI
                
                
                '�V�K�f�[�^�쐬�ǉ�
                If bCmp Then
                    '�ۑ��ς݃f�[�^�L��
                    APSearchListSlbData(nJ).slb_col_cnt = Format(nSlb_Col_Cnt_MAX + 1, "00")
                Else
                    '�ۑ��ς݃f�[�^����
                    APSearchListSlbData(nJ).slb_col_cnt = "01"
                End If
'                If bCmp = False Then
                    
                'APSearchListSlbData(nJ).bAPResEdit = False
                APSearchListSlbData(nJ).bAPScanInput = False
                APSearchListSlbData(nJ).bAPPdfInput = False
                
                APSearchListSlbData(nJ).slb_no = strSearchSlbNumber
                APSearchListSlbData(nJ).slb_chno = Mid(strSearchSlbNumber, 1, 5)
                APSearchListSlbData(nJ).slb_aino = Mid(strSearchSlbNumber, 6)
                
                APSearchListSlbData(nJ).slb_stat = APSlbCont.nSearchInputStatusSelectedIndex
                
                If bCmp Then
                    '�ۑ��ς݃f�[�^�L��
                    '����f�[�^�R�s�[
                    
                    '�\�����X�g�ɃR�s�[
                    '**********************************************************'
                    'nchtaisl
                    'APSozaiTmpData(0).slb_no = "123451234"      ''�X���uNO"
                    APSearchListSlbData(nJ).slb_ksh = APSearchTmpSlbData(nFirstDataIndex).slb_ksh  ''�|��
                    APSearchListSlbData(nJ).slb_uksk = APSearchTmpSlbData(nFirstDataIndex).slb_uksk ''����i�M������j
                    'APSearchListSlbData(nJ).slb_lngth = APSozaiData.slb_lngth ''����
                    'APSearchListSlbData(nJ).slb_color_wei = APSozaiData.slb_color_wei ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
                    APSearchListSlbData(nJ).slb_typ = APSearchTmpSlbData(nFirstDataIndex).slb_typ ''�^
                    'APSearchListSlbData(nJ).slb_skin_wei = APSozaiData.slb_skin_wei ''�d�ʁi���ޔ��p�F����d�ʁj
                    'APSearchListSlbData(nJ).slb_wdth = APSozaiData.slb_wdth ''��
                    'APSearchListSlbData(nJ).slb_thkns = APSozaiData.slb_thkns ''����
                    APSearchListSlbData(nJ).slb_zkai_dte = APSearchTmpSlbData(nFirstDataIndex).slb_zkai_dte ''������i����N�����j
                    '**********************************************************'
                    'skjchjdt�e�[�u��
                    'APSozaiData.slb_chno = "12345"        ''�`���[�WNO
                    'APSearchListSlbData(nJ).slb_ccno = APSozaiData.slb_ccno ''CCNO
                    '**********************************************************'
                    
                    '����:APSozaiData�ɃR�s�[
                    '**********************************************************'
                    'nchtaisl
                    APSozaiData.slb_no = APSearchTmpSlbData(nFirstDataIndex).slb_no      ''�X���uNO"
                    APSozaiData.slb_ksh = APSearchTmpSlbData(nFirstDataIndex).slb_ksh       ''�|��
                    APSozaiData.slb_uksk = APSearchTmpSlbData(nFirstDataIndex).slb_uksk         ''����i�M������j
                    APSozaiData.slb_lngth = APSearchTmpSlbData(nFirstDataIndex).slb_lngth       ''����
                    APSozaiData.slb_color_wei = APSearchTmpSlbData(nFirstDataIndex).slb_wei   ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
                    APSozaiData.slb_typ = APSearchTmpSlbData(nFirstDataIndex).slb_typ           ''�^
'                    APSozaiData.slb_skin_wei = APSearchTmpSlbData(nFirstDataIndex).slb_wei    ''�d�ʁi���ޔ��p�F����d�ʁj
                    APSozaiData.slb_wdth = APSearchTmpSlbData(nFirstDataIndex).slb_wdth         ''��
                    APSozaiData.slb_thkns = APSearchTmpSlbData(nFirstDataIndex).slb_thkns      ''����
                    APSozaiData.slb_zkai_dte = APSearchTmpSlbData(nFirstDataIndex).slb_zkai_dte ''������i����N�����j
                    '**********************************************************'
                    'skjchjdt�e�[�u��
                    APSozaiData.slb_chno = APSearchTmpSlbData(nFirstDataIndex).slb_chno        ''�`���[�WNO
                    APSozaiData.slb_ccno = APSearchTmpSlbData(nFirstDataIndex).slb_ccno        ''CCNO
                    '**********************************************************'
                Else
                    '�ۑ��ς݃f�[�^����
                    '**********************
                    '�f�ޓ�������Ǎ�
                    'bRet = SOZAI_NCHTAISL_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no)
                    
                    bNoRecord = False '2008/08/30 A.K
                    
                    bRet = SOZAI_NCHTAISL_Read(APSearchListSlbData(nJ).slb_no)
                    If UBound(APSozaiTmpData) = 1 Then
                        APSozaiData = APSozaiTmpData(0)
                    Else
                        'NCHTAISL�Y�����R�[�h����
                        bNoRecord = True '2008/08/30 A.K
                    End If
                    If bRet = False Then
                        Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                        MsgWnd.OK_Close
                        Exit Sub
                    End If
                    
                    'bRet = SOZAI_SKJCHJDT_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno)
                    bRet = SOZAI_SKJCHJDT_Read(APSearchListSlbData(nJ).slb_chno)
                    If UBound(APSozaiTmpData) = 1 Then
                        APSozaiData.slb_chno = APSozaiTmpData(0).slb_chno
                        APSozaiData.slb_ccno = APSozaiTmpData(0).slb_ccno
                        
                        If bNoRecord Then '2008/08/30 A.K
                            'NCHTAISL�Y�����R�[�h�����̏ꍇ�̏���
                            'SKJCHJDT����|��,�^�𒊏o
                            APSozaiData.slb_ksh = APSozaiTmpData(0).slb_ksh ''�|��
                            APSozaiData.slb_typ = APSozaiTmpData(0).slb_typ ''�^
                        End If
                        
                    End If
                    If bRet = False Then
                        Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                        MsgWnd.OK_Close
                        Exit Sub
                    End If
                    
                    '���X�g�ɃR�s�[
                    '**********************************************************'
                    'nchtaisl
                    'APSozaiTmpData(0).slb_no = "123451234"      ''�X���uNO"
                    APSearchListSlbData(nJ).slb_ksh = APSozaiData.slb_ksh ''�|��
                    APSearchListSlbData(nJ).slb_uksk = APSozaiData.slb_uksk ''����i�M������j
                    'APSearchListSlbData(nJ).slb_lngth = APSozaiData.slb_lngth ''����
                    'APSearchListSlbData(nJ).slb_color_wei = APSozaiData.slb_color_wei ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
                    APSearchListSlbData(nJ).slb_typ = APSozaiData.slb_typ ''�^
                    'APSearchListSlbData(nJ).slb_skin_wei = APSozaiData.slb_skin_wei ''�d�ʁi���ޔ��p�F����d�ʁj
                    'APSearchListSlbData(nJ).slb_wdth = APSozaiData.slb_wdth ''��
                    'APSearchListSlbData(nJ).slb_thkns = APSozaiData.slb_thkns ''����
                    APSearchListSlbData(nJ).slb_zkai_dte = APSozaiData.slb_zkai_dte ''������i����N�����j
                    '**********************************************************'
                    'skjchjdt�e�[�u��
                    'APSozaiData.slb_chno = "12345"        ''�`���[�WNO
                    'APSearchListSlbData(nJ).slb_ccno = APSozaiData.slb_ccno ''CCNO
                    '**********************************************************'
                    
                    '**********************
                End If
                
                ReDim Preserve APSearchListSlbData(UBound(APSearchListSlbData) + 1)
                nJ = nJ + 1
'                End If
            Case 1 '�C��
            Case 2 '�폜
        End Select
        
        For nI = 0 To UBound(APSearchTmpSlbData) - 1
            APSearchListSlbData(nJ) = APSearchTmpSlbData(nI)
            ReDim Preserve APSearchListSlbData(UBound(APSearchListSlbData) + 1)
            nJ = nJ + 1
        Next nI
    
    End If

    MsgWnd.OK_Close
    
    '�O���b�h�փZ�b�g
    Call SetMSFlexGrid1
    
End Sub

' @(f)
'
' �@�\      : �X���u�I�����b�N�^�A�����b�N
'
' ������    : ARG1 - True=���b�N�^False=�A�����b�N �t���O
'
' �Ԃ�l    :
'
' �@�\����  : �X���u�I����Ԃ̉�ʃ��b�N�^�A�����b�N����B
'
' ���l      :COLORSYS
'
Private Sub SlbSelLock(ByVal sw As Boolean)
    If sw Then
'        cmdOK.Enabled = True
'        MSFlexGrid1.Enabled = False
'        imTextSearchSlbNumber.Enabled = False
'        cmdSearch.Enabled = False
'        OptSearchMode(0).Enabled = False
'        OptSearchMode(1).Enabled = False
'        OptSearchMode(2).Enabled = False
'        OptSearchMode(3).Enabled = False
'
'        lblSearchMAX(2).Enabled = False
'
'        APSlbCont.bProcessing = True '�X���u�I�����b�N�p�������t���O
'        APSlbCont.strSearchInputSlbNumber = imTextSearchSlbNumber.Text '�����X���u�m���D
'        If OptSearchMode(0).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 0 '�����I�v�V�����w��C���f�b�N�X�ԍ�
'        ElseIf OptSearchMode(1).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 1 '�����I�v�V�����w��C���f�b�N�X�ԍ�
'        ElseIf OptSearchMode(2).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 2 '�����I�v�V�����w��C���f�b�N�X�ԍ�
'        ElseIf OptSearchMode(3).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 3 '�����I�v�V�����w��C���f�b�N�X�ԍ�
'        End If
'        '�X���u�I�����ۑ�
'        APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
'        '�q�X���u�͂n�j�{�^�����ɕۑ�
'        'nChildSelectedIndex As Integer '�q�X���u�w��C���f�b�N�X�ԍ� 0�͖��w��
    Else
'        cmdOK.Enabled = False
'        MSFlexGrid1.Enabled = True
'        imTextSearchSlbNumber.Enabled = True
'        cmdSearch.Enabled = True
'        OptSearchMode(0).Enabled = True
'        OptSearchMode(1).Enabled = True
'        OptSearchMode(2).Enabled = True
'        OptSearchMode(3).Enabled = True
'
'        lblSearchMAX(2).Enabled = True
'
'        APSlbCont.bProcessing = False '�X���u�I�����b�N�p�������t���O
    End If
    
    Call MSFlexGrid1_Click

    DoEvents

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
    Dim bRet As Boolean
    
    Select Case CallNo
    
    Case CALLBACK_USEIMGDATA
        '���ɓo�^�f�[�^�����݂���V�i���I
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next
            '�o�^�ς݃X�L���i�[�C���[�W
            
'            Call ImageDataRead
            '�C���[�W�t�@�C���Ǎ���
            'Call ImageLoad
            
            
            'On Error GoTo 0
            'Unload Me
        Else
            
        End If
'        cmdSplitChg.Enabled = True
        
    Case CALLBACK_RES_COLORDATA_DBDEL_REQ
        '�f�[�^�폜�̖⍇�����OK
        If Result = CALLBACK_ncResOK Then          'OK
            bRet = ColorSlbData_Load(False) '�폜�������s
        Else
            
        End If
        
        cmdOK.Enabled = True '�{�^���L��
        
    Case CALLBACK_RES_COLORDATA_HOSTDEL_REQ
        '�f�[�^�폜�̖⍇�����OK�i�r�W�R���֍폜���M�V�i���I�j
        If Result = CALLBACK_ncResOK Then          'OK
            '�r�W�R�����M
            
'           '���n�ɂĒ����i�ʐM�e�X�g���j
            APResData.slb_fault_u_judg = "0"
            APResData.slb_fault_d_judg = "0"
            
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
            APResData.host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_dte
            APResData.host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_tme
            APResData.fail_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte
            APResData.fail_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_tme
            
            frmHostSend.SetCallBack Me, CALLBACK_RES_COLORDATA_HOSTDEL_REQ2
            frmHostSend.Show vbModal, Me '�r�W�R�����M���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
        Else
            '�L�����Z��
            cmdOK.Enabled = True '�{�^���L��
        End If
        
    Case CALLBACK_RES_COLORDATA_HOSTDEL_REQ2
        '�r�W�R���폜�������OK�i�r�W�R���֍폜���M�V�i���I�j
        If Result = CALLBACK_ncResOK Then          'OK
            bRet = ColorSlbData_Load(False) '�폜�������s
        ElseIf Result = CALLBACK_ncResSKIP Then 'SKIP
            bRet = ColorSlbData_Load(False) '�폜�������s
        Else
            '�r�W�R���G���[����
            Call WaitMsgBox(Me, "�r�W�R���ʐM�G���[�����������ׁA�c�a�폜�����͒��f����܂����B")
        End If
        
        cmdOK.Enabled = True '�{�^���L��
        
    End Select

End Sub

' @(f)
'
' �@�\      : �O���b�h�P������
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �O���b�h�P�̏��������s���B
'
' ���l      :
'
Private Sub InitMSFlexGrid1()

    Dim nJ As Integer
    Dim nRow As Integer
    Dim nCol As Integer

    nMSFlexGrid1_Selected_Row = 0
    lbl_nMSFlexGrid1_Selected_Row.Caption = nMSFlexGrid1_Selected_Row
    
    MSFlexGrid1.SelectionMode = flexSelectionByRow
    MSFlexGrid1.FixedCols = 0
    ' 20090115 modify by M.Aoyagi    �摜�����ύX�̈׉��Z
'    MSFlexGrid1.Cols = 16 + 1
    '2016/04/20 - TAI - S
'    MSFlexGrid1.Cols = 18 + 1
    MSFlexGrid1.Cols = 19 + 1
    '2016/04/20 - TAI - E

    MSFlexGrid1.Rows = 1
    
    nRow = 0
    nCol = 0
    MSFlexGrid1.ColWidth(nCol) = 0
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = ""
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1400
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�|��"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�X���uNo."
    
'    '�ُ�ꗗ���X�g�\����p '2008/09/04
'    slb_fault_e_judg As String  ''����E�ʔ���
'    slb_fault_w_judg As String  ''����W�ʔ���
'    slb_fault_s_judg As String  ''����S�ʔ���
'    slb_fault_n_judg As String  ''����N�ʔ���
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "EWSN"
        
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�^"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "����"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "��"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000
    MSFlexGrid1.ColWidth(nCol) = 900                    ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "����"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�Ұ��"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "PDF"
    
    ' 20090115 add by M.Aoyagi    �摜�����ǉ�
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 700
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "����"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�ُ��"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�Ұ��"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "PDF"
    
    ' 20090115 add by M.Aoyagi    �摜�����ǉ�
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 700
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "����"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���u�w��"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���u����"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1000                  ' 20090129 modify by M.Aoyagi    �\���l�߂�׃T�C�Y������
    MSFlexGrid1.ColWidth(nCol) = 700
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���"
    
    '2016/04/20 - TAI - S
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "��Ə�"
    '2016/04/20 - TAI - E
    
    '�^�C�g���s
    For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000012
        MSFlexGrid1.CellBackColor = &H8000000F
    Next nJ

End Sub

' @(f)
'
' �@�\      : �O���b�h�P�f�[�^�Z�b�g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �O���b�h�P�̃f�[�^�Z�b�g���s���B
'
' ���l      :
'
Private Sub SetMSFlexGrid1()
    Dim nJ As Integer
    Dim nCol As Integer
    Dim nRow As Integer
    
    MSFlexGrid1.Rows = 1 + UBound(APSearchListSlbData)
    
    For nRow = 1 To MSFlexGrid1.Rows - 1
    
        nCol = 0
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_ksh '"�|��"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignLeftCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_chno & "-" & APSearchListSlbData(nRow - 1).slb_aino '"�X���uNo."
        
'    '�ُ�ꗗ���X�g�\����p '2008/09/04
'    slb_fault_e_judg As String  ''����E�ʔ���
'    slb_fault_w_judg As String  ''����W�ʔ���
'    slb_fault_s_judg As String  ''����S�ʔ���
'    slb_fault_n_judg As String  ''����N�ʔ���
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_fault_e_judg & _
        APSearchListSlbData(nRow - 1).slb_fault_w_judg & _
        APSearchListSlbData(nRow - 1).slb_fault_s_judg & _
        APSearchListSlbData(nRow - 1).slb_fault_n_judg '"EWSN"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_typ '"�^"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_uksk '"����"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = ConvDpOutStat(conDefine_SYSMODE_COLOR, CInt(APSearchListSlbData(nRow - 1).slb_stat)) '"���"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_col_cnt '"�װ��"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        
        If APSearchListSlbData(nRow - 1).fail_sys_wrt_dte <> "" Then
            '�ُ�񍐂����݂��鎞
            MSFlexGrid1.TextMatrix(nRow, nCol) = "�ۗ�"
            Set MSFlexGrid1.CellPicture = PicSigRed.Picture
            
            If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "1" Then
                '�v�d�a�őS�����̏ꍇ
                If APSearchListSlbData(nRow - 1).fail_res_host_send <> "2" Then
                    '�ۗ������A���u���������A�����M�ł͂Ȃ��ꍇ
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte '"�װ����"�i����o�^���j
                    Set MSFlexGrid1.CellPicture = Nothing
                End If
            End If
        ElseIf APSearchListSlbData(nRow - 1).host_send <> "" Then
            '�r�W�R���ʐM�����M�ς݂̏ꍇ�i�n�j�A�m�f�ɂ�����炸�j
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte '"�װ����"�i����o�^���j
            Set MSFlexGrid1.CellPicture = Nothing
        Else
            If APSearchListSlbData(nRow - 1).sys_wrt_dte <> "" Then
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            Set MSFlexGrid1.CellPicture = Nothing
            Else
                If IsDEBUG("DISP") Then
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte & "?"
                    Set MSFlexGrid1.CellPicture = Nothing
                Else
                    MSFlexGrid1.TextMatrix(nRow, nCol) = ""
                    Set MSFlexGrid1.CellPicture = Nothing
                End If
            End If
        End If
    
        If APSearchListSlbData(nRow - 1).host_send = "0" Then
            '�r�W�R���ʐM���ُ푗�M�̏ꍇ
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "�ʐM�װ"
'            Set MSFlexGrid1.CellPicture = Nothing
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPScanInput, "��", "") '"�װ�Ұ��"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).bAPPdfInput Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "��"
        ElseIf APSearchListSlbData(nRow - 1).sAPPdfInput_ReqDate <> "" Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "��"
        Else
            MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        End If
        'MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPPdfInput, "��", "") '"�װPDF"
    
        ' 20090115 add by M.Aoyagi    �J���[�摜�����\���ǉ�
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).PhotoImgCnt1
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).fail_host_send <> "" Then
            '�r�W�R���ʐM�����M�ς݂̏ꍇ�i�n�j�A�m�f�ɂ�����炸�j
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_sys_wrt_dte '"�ُ��"�i����o�^���j
        Else
            If APSearchListSlbData(nRow - 1).fail_sys_wrt_dte <> "" Then
                If IsDEBUG("DISP") Then
                    '�����M
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_sys_wrt_dte & "?"
                Else
                    MSFlexGrid1.TextMatrix(nRow, nCol) = ""
                End If
            Else
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            End If
        End If
    
        If APSearchListSlbData(nRow - 1).fail_host_send = "0" Then
            '�r�W�R���ʐM���ُ푗�M�̏ꍇ
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "�ʐM�װ"
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPFailScanInput, "��", "") '"�ُ�Ұ��"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).bAPFailPdfInput Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "��"
        ElseIf APSearchListSlbData(nRow - 1).sAPFailPdfInput_ReqDate <> "" Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "��"
        Else
            MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        End If
        'MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPFailPdfInput, "��", "") '"�ُ�PDF"
    
        ' 20090115 add by M.Aoyagi    �摜�����\���ǉ��̈�
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).PhotoImgCnt2
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_dir_sys_wrt_dte '"���u�w��"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "1" Then
            '����
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_res_sys_wrt_dte '"���u����"
        Else
            If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "0" Then
                '������
                MSFlexGrid1.TextMatrix(nRow, nCol) = "��"
            Else
                '�o�^����
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            End If
        End If
    
        '�v�d�a�őS�����̏ꍇ
        If APSearchListSlbData(nRow - 1).fail_res_host_send = "2" Then
            '�r�W�R���ʐM���ُ푗�M�̏ꍇ
            MSFlexGrid1.TextMatrix(nRow, nCol) = "�����M"
        End If
    
        If APSearchListSlbData(nRow - 1).fail_res_host_send = "0" Then
            '�r�W�R���ʐM���ُ푗�M�̏ꍇ
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "�ʐM�װ"
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
        '2008/09/04 �w������ς�
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).fail_dir_prn_out_max = "1" Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "��" '"���"
        Else
            MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        End If
        
        '2016/04/20 - TAI - S
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_works_sky_tok '"��Ə�"
        '2016/04/20 - TAI - E

    Next nRow

    If MSFlexGrid1.Rows > 1 Then
        MSFlexGrid1.Row = 1
    End If

End Sub

Private Sub imTextSearchSlbNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

' @(f)
'
' �@�\      : �X���u�ԍ�����BOX�L�[��
'
' ������    : ARG1 - ASCII�R�[�h
'
' �Ԃ�l    :
'
' �@�\����  : �X���u�ԍ�����BOX�L�[�����̏������s���B
'
' ���l      :
'
Private Sub imTextSearchSlbNumber_KeyPress(KeyAscii As Integer)
    Dim nI As Integer
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = Asc("-") Then
    ElseIf KeyAscii = Asc("*") Then
        'For nI = 1 To LenB(imTextSearchSlbNumber.Text)
        '    If Mid(imTextSearchSlbNumber.Text, nI, 1) = "*" Then
        '        KeyAscii = 0
        '        Beep
        '    End If
        'Next nI
    Else
        If KeyAscii <> 10 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

' @(f)
'
' �@�\      : �O���b�h�P�N���b�N
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �O���b�h�P�N���b�N���̏������s���B
'
' ���l      :
'
Private Sub MSFlexGrid1_Click()
    Dim nJ As Integer
    Dim nNowRow As Integer
    Dim nNowSplitNum As Integer
    Dim nRet As Integer

    bMouseControl = True

    '���݂�Row���ꎞ�ۑ�
    nNowRow = MSFlexGrid1.Row

    '�ȑO�̃Z���N�g�s�𖢃Z���N�g��Ԃɖ߂��B
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H80000008
        MSFlexGrid1.CellBackColor = &H80000005
        Next nJ
    Else
        '�^�C�g���s�̐F��t�������B
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H80000012
        MSFlexGrid1.CellBackColor = &H8000000F
        Next nJ
    End If

    '���݂̃Z���N�g�s�ԍ���ۑ�
    nMSFlexGrid1_Selected_Row = nNowRow
    lbl_nMSFlexGrid1_Selected_Row.Caption = nMSFlexGrid1_Selected_Row
    
    '���݂̍s���Z���N�g�s�ɂ���B
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
            MSFlexGrid1.Col = nJ
            If MSFlexGrid1.Enabled Then
                '�I�𒆂̐F
                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
                    '�폜���[�h�̏ꍇ
                    If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8080FF
                Else
                    If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8000000D
                End If
                
                '�폜���[�h���H
                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
                    '�폜���[�h
                    cmdDirRes.Enabled = False '�֎~�I
                Else
                    If APSearchListSlbData(nMSFlexGrid1_Selected_Row - 1).fail_dir_sys_wrt_dte <> "" Then
                        '�w���L��
                        cmdDirRes.Enabled = True
                        cmdOK.Enabled = False '2008/09/04 ���яC���u�֎~�v�I
                    Else
                        '�w������
                        cmdDirRes.Enabled = False
                        cmdOK.Enabled = True '2008/09/04 ���яC���u���v�I
                    End If
                End If
                
            Else
                '�I�����b�N���̐F
                If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                MSFlexGrid1.CellBackColor = &H808080
            End If
        Next nJ
        If MSFlexGrid1.Enabled Then
            '�I��
        Else
            '�I�����b�N
        End If
    
    Else
    End If

End Sub

' @(f)
'
' �@�\      : �O���b�h�P�t�H�[�J�X�擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �O���b�h�P�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub MSFlexGrid1_GotFocus()
    If nMSFlexGrid1_Selected_Row = 0 Then
        If MSFlexGrid1.Rows > 1 Then
            'MSFlexGrid1.Row = 1
            'Debug.Print MSFlexGrid1.Row
            Call MSFlexGrid1_Click
        End If
    End If
End Sub

' @(f)
'
' �@�\      : �O���b�h�P�Z���ύX
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �O���b�h�P�Z���ύX���̏������s���B
'
' ���l      :
'
Private Sub MSFlexGrid1_SelChange()
    If bMouseControl = False Then
        Call MSFlexGrid1_Click
    End If
    bMouseControl = False
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
    
    Dim nI As Integer
    
    bMouseControl = False
    
'    For nI = 0 To 3
'        lblSearchMAX(nI).Caption = APSysCfgData.nSEARCH_MAX(nI)
'    Next nI
    
    '�I��ԍ��\��
    If IsDEBUG("DISP") Then
        lbl_nMSFlexGrid1_Selected_Row.Visible = True
'        lbl_nMSFlexGrid2_Selected_Row.Visible = True
    End If
    
    '2016/04/20 - TAI - S
    '��Əꏊ�\��
    If works_sky_tok = WORKS_SKY Then
        lbl_works.Caption = "SKY"               'SKY
        lbl_works.ForeColor = &HFF              '��
    ElseIf works_sky_tok = WORKS_TOK Then
        lbl_works.Caption = "���|"              '���|
        lbl_works.ForeColor = &HFF0000          '��
    End If
    '2016/04/20 - TAI - E

'    cmdOK.Enabled = False
    
'    LEAD_SCAN.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
'    LEAD_SCAN.EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
'    LEAD_SCAN.EnableTwainEvent = True
'    LEAD_SCAN.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
'
'    For nI = 0 To 1
'        LEAD1(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
'        LEAD1(nI).EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
'        LEAD1(nI).EnableTwainEvent = True
'        LEAD1(nI).PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
'    Next nI
    
    Call InitMSFlexGrid1

'    If APSlbCont.bProcessing Then '�X���u�I�����b�N�p�������t���O
        '2008.09.03 imTextSearchSlbNumber.Text = APSlbCont.strSearchInputSlbNumber  '�����X���u�m���D
        
'        2008.09.03 OptInputMode(APSlbCont.nSearchInputModeSelectedIndex).Value = True '���̓��[�h�w��C���f�b�N�X�ԍ�
        
        '2008.09.03 bOptInputModeValue(0) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, True, False)
        '2008.09.03 bOptInputModeValue(1) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 1, True, False)
        '2008.09.03 bOptInputModeValue(2) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 2, True, False)
        
        '2008.09.03 OptStatus(APSlbCont.nSearchInputStatusSelectedIndex).Value = True '��ԑI���w��C���f�b�N�X�ԍ�
        
        '�w������
        cmdDirRes.Enabled = False
        cmdOK.Enabled = False '2008/09/04 ���яC���u�֎~�v�I
        
        '�X���u�I�����
        nMSFlexGrid1_Selected_Row = APSlbCont.nListSelectedIndexP1
        Call SetMSFlexGrid1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        Call MSFlexGrid1_Click
        Call SlbSelLock(True)
        
'    End If

    '2008/09/04 ����\��
    Call cmdSearch_Click

End Sub

