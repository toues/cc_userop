VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSkinSlbSelWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "�X���u���������́|�X���u�I��"
   ClientHeight    =   11280
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   16035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11280
   ScaleWidth      =   16035
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdStatChgMode 
      Caption         =   "��ԕύX"
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
      Left            =   14040
      TabIndex        =   24
      Top             =   960
      Width           =   1800
   End
   Begin VB.CommandButton cmdStatChgFix 
      Caption         =   "��Ԍ���"
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
      Left            =   14040
      TabIndex        =   23
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Frame Frame_Status 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�V�K���́|��ԑI��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   7260
      TabIndex        =   15
      Top             =   600
      Width           =   6675
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "6:6ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   4980
         TabIndex        =   22
         Top             =   1140
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "5:5ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   3420
         TabIndex        =   21
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "4:4ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   1860
         TabIndex        =   20
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "2:2ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3420
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1:1ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1860
         TabIndex        =   18
         Top             =   480
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0:����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "3:3ht��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   4980
         TabIndex        =   16
         Top             =   480
         Width           =   1515
      End
   End
   Begin VB.PictureBox PicSigYellow 
      Height          =   315
      Left            =   9120
      Picture         =   "frmSkinSlbSelWnd.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   10320
      Visible         =   0   'False
      Width           =   555
   End
   Begin imText6Ctl.imText imTextSearchSlbNumber 
      Height          =   525
      Left            =   1800
      TabIndex        =   0
      Top             =   1980
      Width           =   3360
      _Version        =   65536
      _ExtentX        =   5927
      _ExtentY        =   926
      Caption         =   "frmSkinSlbSelWnd.frx":0644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSkinSlbSelWnd.frx":06B2
      Key             =   "frmSkinSlbSelWnd.frx":06D0
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
      Format          =   "A9#@"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   10
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
   Begin VB.Frame Frame_InputMode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "���̓��[�h�I��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   6855
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�V�K"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�C��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�폜"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   3900
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicSigRed 
      Height          =   315
      Left            =   9120
      Picture         =   "frmSkinSlbSelWnd.frx":0714
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicSigGreen 
      Height          =   315
      Left            =   9120
      Picture         =   "frmSkinSlbSelWnd.frx":0D56
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   10680
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
      Left            =   12120
      TabIndex        =   7
      Top             =   9900
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   10140
      TabIndex        =   6
      Top             =   9900
      Width           =   1800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "����"
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
      Left            =   5220
      TabIndex        =   1
      Top             =   1980
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7095
      Left            =   180
      TabIndex        =   5
      Top             =   2640
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   21
      Cols            =   10
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      ScrollBars      =   2
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
   Begin VB.Label lbl_nMSFlexGrid1_Selected_Row 
      Height          =   315
      Left            =   1140
      TabIndex        =   12
      Top             =   9900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�X���u���������́|�X���u�I��"
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�X���uNo."
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   9
      Top             =   1980
      Width           =   1635
   End
End
Attribute VB_Name = "frmSkinSlbSelWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSkinSlbSelWnd.Frm                ver 1.00 ( '2008.04.16 SystEx Ayumi Kikuchi )

' @(s)
' �X���u���������́|�X���u�I��\���t�H�[��
' �@�{���W���[���̓X���u���������́|�X���u�I��\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private nMSFlexGrid1_Selected_Row As Integer ''�O���b�h�P�I���s�ԍ��i�[
Private bMouseControl As Boolean ''�}�E�X�R���g���[���t���O�i�[
Private bOptInputModeValue(0 To 2) As Boolean

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
    Call SlbSelLock(False)
    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETSKINSLBSELWND, CALLBACK_ncResCANCEL)
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
' ���l      :COLORSYS
'
Private Sub cmdOK_Click()
    Dim bRet As Boolean
    Dim strSource As String
    Dim strDestination As String
    
    cmdOK.Enabled = False ''�A�ŋ֎~�I
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    
    '�X���u�I���`�F�b�N
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '�V�K
                Call WaitMsgBox(Me, "���ѓ��͂��s���X���u��I�����Ă��������B")
            Case 1 '�C��
                Call WaitMsgBox(Me, "���яC�����s���X���u��I�����Ă��������B")
            Case 2 '�폜
                Call WaitMsgBox(Me, "���э폜���s���X���u��I�����Ă��������B")
        End Select
        
        cmdOK.Enabled = True '�{�^���L��
        
        Exit Sub
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
            'APResData.slb_col_cnt = APSozaiData.       ''�J���[��
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
            APResData.slb_wrt_nme = APSysCfgData.NowStaffName(conDefine_SYSMODE_SKIN) '�X�^�b�t��
            APResData.slb_nxt_prcs = APSysCfgData.NowNextProcess(conDefine_SYSMODE_SKIN) '���H��
            
            '�V�K�̏ꍇ�́ASCAN�C���[�W������������B�i���ԃt�@�C���̍폜�j
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.JPG"
            '�C���[�W����
            If Dir(strDestination) <> "" Then
                Kill strDestination
            End If
            
            ' 20090115 add by M.Aoyagi    �摜�����ǉ��̈�
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1
            
        Case 1 '�C��
            '���уf�[�^�Ǎ���
            bRet = TRTS0012_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, _
                                 APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat)
            If UBound(APResTmpData) = 1 Then
                APResData = APResTmpData(0)
            End If
            If bRet = False Then
                Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
                MsgWnd.OK_Close
                
                cmdOK.Enabled = True '�{�^���L��
                Exit Sub
            End If
            
            '�o�^�ς�SCAN�C���[�W�����邩�H
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.JPG"
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).bAPScanInput Then
                '�o�^�ς�SCAN�C���[�W��Ǎ��� (conDefine_ImageDirName = TEMP)
                strSource = APSysCfgData.SHARES_SCNDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                         "\SKIN" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                         "_" & APResData.slb_stat & "_00.JPG"
                On Error GoTo cmdOK_Click_File_err:
                Call FileCopy(strSource, strDestination)
                On Error GoTo 0
            Else
                '�C���[�W����
                If Dir(strDestination) <> "" Then
                    Kill strDestination
                End If
            End If
            
            ' 20090115 add by M.Aoyagi    �摜�����ǉ��̈�
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1
            
        Case 2 '�폜
        
            MsgWnd.OK_Close
            Call SkinDataDel_REQ
            Exit Sub
    End Select
    
    '�c�a�������������郂�[�h�@�C���^�폜�i�Ǎ������b�Z�[�W�\���L��j
    If APSlbCont.nSearchInputModeSelectedIndex <> 0 Then
        MsgWnd.OK_Close
    End If
    
    cmdOK.Enabled = True '�{�^���L��
    
    Call OKcmdOK '���͊J�n(unload me)
    
    Exit Sub
    
cmdOK_Click_File_err:
    Call MsgLog(conProcNum_MAIN, Err.Source + " " + _
        CStr(Err.Number) + Chr(13) + Err.Description)
    
    Call MsgLog(conProcNum_MAIN, "cmdOK_Click_File FILECOPY SO=" & strSource & " DE=" & strDestination)
    Call WaitMsgBox(Me, "�ۑ��ς݃X�L���i�[�C���[�W�t�@�C���̓Ǎ��G���[���������܂����B" & vbCrLf & vbCrLf & "FILECOPY SO=" & strSource & " DE=" & strDestination)
    
    MsgWnd.OK_Close
    On Error GoTo 0
    
    cmdOK.Enabled = True '�{�^���L��
    Exit Sub
    
End Sub

Private Sub SkinDataDel()
    Dim bRet As Boolean
    
    '*********
    '�폜����
    '*********
    APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
    APResData.slb_stat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
    bRet = TRTS0012_Write(True)
    If bRet = False Then
        Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
        Exit Sub
    End If
    
    bRet = TRTS0050_Write(True)
    If bRet = False Then
        Call WaitMsgBox(Me, "�f�[�^�x�[�X�Ǎ��G���[���������܂����B")
        Exit Sub
    End If
    
    '*********
    '�������ʃ��X�g�ĕ\��
    '*********
    Call WaitMsgBox(Me, "�폜����������I�����܂����B")
    Call cmdSearch_Click

End Sub

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

    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETSKINSLBSELWND, CALLBACK_ncResOK)
    Unload Me

End Sub

' @(f)
'
' �@�\      : �X���u��񌟍��{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u���̌������s���B
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
    
    Dim bNoRecord As Boolean '2008/08/30 A.K
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    '�Č������͏�����
    strSearchSlbNumber = ""
    Call InitMSFlexGrid1
    
    nWildCard = 0
    '�n�C�t���f�|�f������Ď��ۂ̌���������փZ�b�g
    strSearchSlbNumber = ConvSearchSlbNumber(imTextSearchSlbNumber.Text)
    
    '���̓��[�h
    If OptInputMode(0).Value Then '�V�K
        APSlbCont.nSearchInputModeSelectedIndex = 0 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
        
        If OptStatus(0).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 0 '����
        ElseIf OptStatus(1).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 1 '1ht��
        ElseIf OptStatus(2).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 2 '2ht��
        ElseIf OptStatus(3).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 3 '3ht��
        ElseIf OptStatus(4).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 4 '4ht��
        ElseIf OptStatus(5).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 5 '5ht��
        ElseIf OptStatus(6).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 6 '6ht��
        End If
    ElseIf OptInputMode(1).Value Then '�C��
        APSlbCont.nSearchInputModeSelectedIndex = 1 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
        APSlbCont.nSearchInputStatusSelectedIndex = 0 '�����i�g�p���Ȃ��j
    ElseIf OptInputMode(2).Value Then '�폜
        APSlbCont.nSearchInputModeSelectedIndex = 2 '���̓��[�h�I�v�V�����w��C���f�b�N�X�ԍ�
        APSlbCont.nSearchInputStatusSelectedIndex = 0 '�����i�g�p���Ȃ��j
    End If
    
    nWildCard = InStr(1, strSearchSlbNumber, "%", vbTextCompare)
    
    'RIAL
    ReDim APSearchListSlbData(0)
    
    '�V�K���[�h�i�����Ń��C���h�J�[�h�s�j
    If OptInputMode(0).Value Then
        '�󔒎w��͕s�B
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�s�B
        If nWildCard <> 0 Then
            Call WaitMsgBox(Me, "�V�K���[�h�Ń��C���h�J�[�h�̎w��͏o���܂���B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '�擪����T�����܂ł́A0����9�ȊO��s�B
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9�ȊO�̎w��͏o���܂���B")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
    
    
    '�C�����[�h�i�����Ń��C���h�J�[�h�j
    ElseIf OptInputMode(1).Value Then
        '�󔒎w��͕s�B
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�P�݂͕̂s�B
        If strSearchSlbNumber = "%" Then
            Call WaitMsgBox(Me, "���C���h�J�[�h�̎w����@������������܂���B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '�擪����T�����܂ł́A0����9,*�ȊO��s�B
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            ElseIf strChkChar = "%" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9,*�ȊO�̎w��͏o���܂���B")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
    
    '�폜���[�h�i�����Ń��C���h�J�[�h�j
    ElseIf OptInputMode(2).Value Then
        '�󔒎w��͕s�B
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "�X���u�m���D����͂��Ă��������B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�P�݂͕̂s�B
        If strSearchSlbNumber = "%" Then
            Call WaitMsgBox(Me, "���C���h�J�[�h�̎w����@������������܂���B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�X������葽���ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '���C���h�J�[�h�����ŁA�U������菭�Ȃ��ꍇ�͕s�B
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "�X���u�m���D�̌������s���ł��B")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '�擪����T�����܂ł́A0����9,*�ȊO��s�B
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            ElseIf strChkChar = "%" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "�擪����T�����܂ŁA0����9,*�ȊO�̎w��͏o���܂���B")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
        
    End If
    
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
    bRet = DBSkinSlbSearchRead(0, 0, 9999, strSearchSlbNumber)
        
    '�������ʂ��Z�b�g
    If bRet Then
        
        ReDim APSearchListSlbData(0)
        nJ = 0
        
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '�V�K
                bCmp = False
                nJ = 0
                For nI = 0 To UBound(APSearchTmpSlbData) - 1
                    strTmpSlbNumber = APSearchTmpSlbData(nI).slb_no
                    '����No�D���r
                    If strTmpSlbNumber = strSearchSlbNumber Then
                        '��Ԃ��r
                        If CInt(APSearchTmpSlbData(nI).slb_stat) = APSlbCont.nSearchInputStatusSelectedIndex Then
                            bCmp = True '����
                            APSlbCont.nSearchInputModeSelectedIndex = 1 '�V�K�ˏC��
                            Exit For
                        End If
                    End If
                Next nI
            
                '�V�K�f�[�^�쐬�ǉ�
                If bCmp = False Then
                    
                    'APSearchListSlbData(nJ).bAPResEdit = False
                    APSearchListSlbData(nJ).bAPScanInput = False
                    APSearchListSlbData(nJ).bAPPdfInput = False
                    
                    APSearchListSlbData(nJ).slb_no = strSearchSlbNumber
                    APSearchListSlbData(nJ).slb_chno = Mid(strSearchSlbNumber, 1, 5)
                    APSearchListSlbData(nJ).slb_aino = Mid(strSearchSlbNumber, 6)
                    
                    APSearchListSlbData(nJ).slb_stat = APSlbCont.nSearchInputStatusSelectedIndex
                    
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
                    
                    ReDim Preserve APSearchListSlbData(UBound(APSearchListSlbData) + 1)
                    nJ = nJ + 1
                End If
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
    
    ' 20090116 add by M.Aoyagi    �\�����x�Ή�
    MSFlexGrid1.Visible = False
    
    '�O���b�h�փZ�b�g
    Call SetMSFlexGrid1
    
    ' 20090116 add by M.Aoyagi    �\�����x�Ή�
    MSFlexGrid1.Visible = True
    
    ' 20090115 add by M.Aoyagi    ��ԃL�[�ύX�{�^������
    If OptInputMode(0).Value Then
        ' 20090115 add by M.Aoyagi    �V�K���̓L�[�ύX���[�h�{�^���𖳌�
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
    ElseIf OptInputMode(1).Value Then
        ' 20090115 add by M.Aoyagi    �C�����̂݃L�[�ύX���[�h�{�^����L��
        cmdStatChgMode.Enabled = True
'        cmdStatChgFix.Enabled = True
    ElseIf OptInputMode(2).Value Then
        ' 20090115 add by M.Aoyagi    �폜���̓L�[�ύX���[�h�{�^���𖳌�
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
    End If
    
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
    Case CALLBACK_RES_SKINDATA_DBDEL_REQ
        '�f�[�^�폜�̖⍇�����OK
        If Result = CALLBACK_ncResOK Then          'OK
            Call SkinDataDel   '�폜�������s
        Else
            
        End If
        
        cmdOK.Enabled = True '�{�^���L��
    Case CALLBACK_RES_STATECHANGE_SKIN
        '��ԃL�[�ύXOK
        If Result = CALLBACK_ncResOK Then          'OK
            Call cmdStatChgFixExe
        Else
            cmdStatChgFix.Enabled = True            '�{�^���L��
            Call WaitMsgBox(Me, "��ԕύX�����𒆒f���܂����B")
        End If
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
'    MSFlexGrid1.Cols = 9 + 1
    MSFlexGrid1.Cols = 10 + 1
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
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�����"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���ޔ�����"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���ޔ��Ұ��"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���ޔ�PDF"
    
    ' 20090115 add by M.Aoyagi    �摜�����ǉ�
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�摜��"
    
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
        MSFlexGrid1.TextMatrix(nRow, nCol) = ConvDpOutStat(conDefine_SYSMODE_SKIN, CInt(APSearchListSlbData(nRow - 1).slb_stat)) '"���"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_zkai_dte '"�����"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte '"���ޔ�����"�i����o�^���j
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPScanInput, "��", "") '"���ޔ��Ұ��"
    
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
        'MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPPdfInput, "��", "") '"���ޔ�PDF"
    
        ' 20090115 add by M.Aoyagi    �摜�����\���ǉ��̈�
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).PhotoImgCnt1
    
    Next nRow

    If MSFlexGrid1.Rows > 1 Then
        MSFlexGrid1.Row = 1
    End If

End Sub

' @(f)
'
' �@�\      : ��Ԍ���{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �e�[�u���AIMG�t�@�C���APDF�t�@�C���ASCAN�t�@�C���̏�ԃL�[��ύX����B
'
' ���l      :
'
Private Sub cmdStatChgFix_Click()
    Dim bRet      As Boolean
    Dim strSource As String
    Dim strDestination As String
    Dim sSlbno    As String
    Dim sChno     As String
    Dim sAino     As String
    Dim sStat     As String
    Dim sCol      As String
    Dim sStatNew  As String
    Dim sCheckErr As String
    
    Dim fmessage As Object
    Set fmessage = New MessageYN

    cmdStatChgFix.Enabled = False       '�{�^������
    
    ' ��ԕύX�J�n���O������
    Call MsgLog(conProcNum_MAIN, "��ԕύX���J�n���܂��B")

    '�X���u�I���`�F�b�N
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Call WaitMsgBox(Me, "�w��ԁx�ύX���s���X���u��I�����Ă��������B")
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If
    
    ' �K�v�f�[�^�ݒ�
    sSlbno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
    sChno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno
    sAino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino
    sStat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
    sCol = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
    If OptStatus(0).Value Then
        sStatNew = "0" '����
    ElseIf OptStatus(1).Value Then
        sStatNew = "1" '1ht��
    ElseIf OptStatus(2).Value Then
        sStatNew = "2" '2ht��
    ElseIf OptStatus(3).Value Then
        sStatNew = "3" '3ht��
    ElseIf OptStatus(4).Value Then
        sStatNew = "4" '4ht��
    ElseIf OptStatus(5).Value Then
        sStatNew = "5" '5ht��
    ElseIf OptStatus(6).Value Then
        sStatNew = "6" '6ht��
    End If

    ' �ύX��̃f�[�^�����݂��邩�`�F�b�N **********************************************************
    bRet = DBStatChgCheckSKIN(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, sStatNew)
    If bRet = False Then
        sCheckErr = "���Ƀf�[�^�����݂���̂ŕύX�ł��܂���B"
    End If
    
    ' IMG�t�H���_�����݂��邩�`�F�b�N *************************************************************
'    PhotoImgCount("SKIN", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, "00")
    bRet = StatChgFoldCheck("IMG", "SKIN", sChno, sAino, sStatNew, "00")
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "���ɉ摜�t�@�C�������݂���̂ŕύX�ł��܂���B"
        Else
            sCheckErr = sCheckErr & vbCrLf & "���ɉ摜�t�@�C�������݂���̂ŕύX�ł��܂���B"
        End If
    End If
    
    ' PDF�t�H���_�����݂��邩�`�F�b�N *************************************************************
    bRet = StatChgFoldCheck("PDF", "SKIN", sChno, sAino, sStatNew, "00")
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "����PDF�t�@�C�������݂���̂ŕύX�ł��܂���B"
        Else
            sCheckErr = sCheckErr & vbCrLf & "����PDF�t�@�C�������݂���̂ŕύX�ł��܂���B"
        End If
    End If
    
    ' SCAN�t�@�C�������݂��邩�`�F�b�N ************************************************************
    bRet = StatChgFoldCheck("SCAN", "SKIN", sChno, sAino, sStatNew, "00")
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "���ɃX�L�����t�@�C�������݂���̂ŕύX�ł��܂���B"
        Else
            sCheckErr = sCheckErr & vbCrLf & "���ɃX�L�����t�@�C�������݂���̂ŕύX�ł��܂���B"
        End If
    End If

    '�����t�@�C���`�F�b�N�ŏI�m�F *****************************************************************
    If sCheckErr <> "" Then
        Call WaitMsgBox(Me, sCheckErr)
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If

    '��ԕύX���s�m�F
    fmessage.MsgText = "�w��f�[�^�̏�ԃL�[��ύX���܂��B" & vbCrLf & "��낵���ł����H"
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_RES_STATECHANGE_SKIN, False
    fmessage.Show vbModal, Me '���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
    Set fmessage = Nothing
End Sub

' @(f)
'
' �@�\      : ��Ԍ��菈��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �e�[�u���AIMG�t�@�C���APDF�t�@�C���ASCAN�t�@�C���̏�ԃL�[��ύX����B
'
' ���l      :
'
Private Sub cmdStatChgFixExe()
    Dim bRet      As Boolean
    Dim strSource As String
    Dim strDestination As String
    Dim sSlbno    As String
    Dim sChno     As String
    Dim sAino     As String
    Dim sStat     As String
    Dim sCol      As String
    Dim sStatNew  As String
    
    cmdStatChgFix.Enabled = False       '�{�^������
    
    ' ��ԕύX�J�n���O������
'    Call MsgLog(conProcNum_MAIN, "��ԕύX���J�n���܂��B")
    
    ' �K�v�f�[�^�ݒ�
    sSlbno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
    sChno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno
    sAino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino
    sStat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
    sCol = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
    If OptStatus(0).Value Then
        sStatNew = "0" '����
    ElseIf OptStatus(1).Value Then
        sStatNew = "1" '1ht��
    ElseIf OptStatus(2).Value Then
        sStatNew = "2" '2ht��
    ElseIf OptStatus(3).Value Then
        sStatNew = "3" '3ht��
    ElseIf OptStatus(4).Value Then
        sStatNew = "4" '4ht��
    ElseIf OptStatus(5).Value Then
        sStatNew = "5" '5ht��
    ElseIf OptStatus(6).Value Then
        sStatNew = "6" '6ht��
    End If

'    ' �ύX��̃f�[�^�����݂��邩�`�F�b�N **********************************************************
'    bRet = DBStatChgCheckSKIN(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, sStatNew)
'    If bRet = 1 Then
'        Call WaitMsgBox(Me, "���Ƀf�[�^�����݂���̂ŕύX�ł��܂���B")
'        cmdStatChgFix.Enabled = True            '�{�^���L��
'        Exit Sub
'    ElseIf bRet = 2 Then
'        Call WaitMsgBox(Me, "�w���f�[�^�����݂���̂ŕύX�ł��܂���B")
'        cmdStatChgFix.Enabled = True            '�{�^���L��
'        Exit Sub
'    End If
'
'    ' IMG�t�H���_�����݂��邩�`�F�b�N *************************************************************
''    PhotoImgCount("SKIN", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, "00")
'    bRet = StatChgFoldCheck("IMG", "SKIN", sChno, sAino, sStatNew, "00")
'    If bRet = False Then
'        Call WaitMsgBox(Me, "���ɉ摜�t�@�C�������݂���̂ŕύX�ł��܂���B")
'        cmdStatChgFix.Enabled = True            '�{�^���L��
'        Exit Sub
'    End If
'
'    ' PDF�t�H���_�����݂��邩�`�F�b�N *************************************************************
'    bRet = StatChgFoldCheck("PDF", "SKIN", sChno, sAino, sStatNew, "00")
'    If bRet = False Then
'        Call WaitMsgBox(Me, "����PDF�t�@�C�������݂���̂ŕύX�ł��܂���B")
'        cmdStatChgFix.Enabled = True            '�{�^���L��
'        Exit Sub
'    End If
'
'    ' SCAN�t�@�C�������݂��邩�`�F�b�N ************************************************************
'    bRet = StatChgFoldCheck("SCAN", "SKIN", sChno, sAino, sStatNew, "00")
'    If bRet = False Then
'        Call WaitMsgBox(Me, "���ɃX�L�����t�@�C�������݂���̂ŕύX�ł��܂���B")
'        cmdStatChgFix.Enabled = True            '�{�^���L��
'        Exit Sub
'    End If
    
    ' IMG�t�@�C���ύX *****************************************************************************
    bRet = StatChgFoldFix("IMG", "SKIN", sChno, sAino, sStat, sStatNew, "00", "00")
    If bRet = False Then
        Call WaitMsgBox(Me, "�摜�t�@�C���̕ύX�Ɏ��s���܂����B")
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If
    
    ' PDF�t�@�C���ύX *****************************************************************************
    bRet = StatChgFoldFix("PDF", "SKIN", sChno, sAino, sStat, sStatNew, "00", "00")
    If bRet = False Then
        Call WaitMsgBox(Me, "PDF�t�@�C���̕ύX�Ɏ��s���܂����B")
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If
    
    ' SCAN�t�@�C���ύX ****************************************************************************
    bRet = StatChgFoldFix("SCAN", "SKIN", sChno, sAino, sStat, sStatNew, "00", "00")
    If bRet = False Then
        Call WaitMsgBox(Me, "�X�L�����t�@�C���̕ύX�Ɏ��s���܂����B")
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If

    ' �e�[�u���ύX ********************************************************************************
'    bRet = DBStatChgFix(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, sStatNew)
    bRet = DBStatChgFixSKIN(sSlbno, sChno, sAino, sStat, sStatNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "�f�[�^�̕ύX�Ɏ��s���܂����B")
        cmdStatChgFix.Enabled = True            '�{�^���L��
        Exit Sub
    End If

    Call WaitMsgBox(Me, "��ԕύX�͐���ɏI�����܂����B" & vbCrLf & "PDF�̍č쐬�͎蓮�ōs�Ȃ��ĉ������B")
    cmdStatChgFix.Enabled = True        '�{�^���L��

    ' ��ʍX�V
    Call cmdSearch_Click
    
End Sub

' @(f)
'
' �@�\      : ��ԕύX�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ��ԕύX�ׁ̈A��ԃI�v�V�����{�^����L���ɂ���B
'
' ���l      :
'
Private Sub cmdStatChgMode_Click()
    Dim i As Integer

    ' 20090115 add by M.Aoyagi    ��Ԃ�ύX�\�ɂ���
    Frame_Status.Enabled = True '�L��
    For i = 0 To 6
        OptStatus(i).Enabled = True
    Next i

    ' ��Ԍ���{�^����L���ɂ���
    cmdStatChgFix.Enabled = True

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

    ' 20090116 add by M.Aoyagi    �\�����x�Ή�
    MSFlexGrid1.Visible = False

    bMouseControl = True

    '���݂�Row���ꎞ�ۑ�
    nNowRow = MSFlexGrid1.Row

    '�ȑO�̃Z���N�g�s�𖢃Z���N�g��Ԃɖ߂��B
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000008
        MSFlexGrid1.CellBackColor = &H80000005
        Next nJ
    Else
        '�^�C�g���s�̐F��t�������B
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000012
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
                    MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8080FF
                Else
                    MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8000000D
                End If
            Else
                '�I�����b�N���̐F
                MSFlexGrid1.CellForeColor = &H8000000E
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

    ' 20090116 add by M.Aoyagi    �\�����x�Ή�
    MSFlexGrid1.Visible = True

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
        imTextSearchSlbNumber.Text = APSlbCont.strSearchInputSlbNumber  '�����X���u�m���D
        
        OptInputMode(APSlbCont.nSearchInputModeSelectedIndex).Value = True '���̓��[�h�w��C���f�b�N�X�ԍ�
        
        bOptInputModeValue(0) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, True, False)
        bOptInputModeValue(1) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 1, True, False)
        bOptInputModeValue(2) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 2, True, False)
        
        OptStatus(APSlbCont.nSearchInputStatusSelectedIndex).Value = True '��ԑI���w��C���f�b�N�X�ԍ�
        
        ' 20090115 add by M.Aoyagi    �L�[�ύX���[�h�{�^���̏����ݒ�
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
        
        '�X���u�I�����
        nMSFlexGrid1_Selected_Row = APSlbCont.nListSelectedIndexP1
        Call SetMSFlexGrid1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        Call MSFlexGrid1_Click
        Call SlbSelLock(True)
        
'    End If

End Sub

' @(f)
'
' �@�\      : ���̓��[�h�I�v�V�����N���b�N
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���̓��[�h�I�v�V�����N���b�N���̏������s���B
'
' ���l      :conDefine_ColorActive or conDefine_ColorNotActive
'           :COLORSYS
'
Private Sub OptInputMode_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0 '�V�K
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
                
                Frame_Status.Enabled = True '�L��
                For i = 0 To 6
                    OptStatus(i).Enabled = True
                Next i
            End If
            
            ' 20090115 add by M.Aoyagi    �V�K���̓L�[�ύX���[�h�{�^���𖳌�
            cmdStatChgMode.Enabled = False
            cmdStatChgFix.Enabled = False
        Case 1 '�C��
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
        
                Frame_Status.Enabled = False '����
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
        Case 2 '�폜
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
    
            Frame_Status.Enabled = False '����
            For i = 0 To 6
                OptStatus(i).Enabled = False
            Next i
        
            ' 20090115 add by M.Aoyagi    �폜���̓L�[�ύX���[�h�{�^���𖳌�
            cmdStatChgMode.Enabled = False
            cmdStatChgFix.Enabled = False
    End Select

    If bOptInputModeValue(Index) = False Then
        '�ω����������ꍇ
        For i = 0 To 2
            bOptInputModeValue(i) = False
        Next i
        bOptInputModeValue(Index) = True
        
        nMSFlexGrid1_Selected_Row = 0
        APSlbCont.nListSelectedIndexP1 = 0
        
        '�X���u�������X�g�N���A
        ReDim APSearchListSlbData(0)
        '�O���b�h�փZ�b�g
        Call SetMSFlexGrid1
        
    End If

End Sub

' @(f)
'
' �@�\      : ���̓��[�h�I�v�V�����t�H�[�J�X�擾
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���̓��[�h�I�v�V�����t�H�[�J�X�擾���̏������s���B
'
' ���l      :COLORSYS
'
Private Sub OptInputMode_GotFocus(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0 '�V�K
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
                
                Frame_Status.Enabled = True '�L��
                For i = 0 To 6
                    OptStatus(i).Enabled = True
                Next i
            End If
        Case 1 '�C��
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
        
                Frame_Status.Enabled = False '����
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
        Case 2 '�폜
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
    
                Frame_Status.Enabled = False '����
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
    End Select
End Sub

' @(f)
'
' �@�\      : �폜�⍇��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �폜�⍇�������B
'
' ���l      :
'
Private Sub SkinDataDel_REQ()

    Dim bHostSendCmp As Boolean

    Dim fmessage As Object
    Set fmessage = New MessageYN

    fmessage.MsgText = "�I�������f�[�^���폜���܂��B" & vbCrLf & "��낵���ł����H"
'    fmessage.AutoDelete = True
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_RES_SKINDATA_DBDEL_REQ, False

'        Do
'            On Error Resume Next
'            fmessage.Show vbModeless, Me
'            If Err.Number = 0 Then
'                Exit Do
'            End If
'            DoEvents
'        Loop
    fmessage.Show vbModal, Me '���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
    Set fmessage = Nothing
'    End If

End Sub


