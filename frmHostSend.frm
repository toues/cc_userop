VERSION 5.00
Object = "{B6B49C41-8023-4CA6-BDF0-FC5291FC6D71}#18.0#0"; "WCSockControl.ocx"
Begin VB.Form frmHostSend 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "���уf�[�^�r�W�R���o�^"
   ClientHeight    =   3540
   ClientLeft      =   825
   ClientTop       =   1050
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.CommandButton cmdSKIP 
      Caption         =   "�X�L�b�v"
      Height          =   435
      Left            =   7620
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   7620
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer timTimeOut 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin WCSocket.WCSockControl WCSockControl1 
      Height          =   1275
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   2249
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  '����
      Caption         =   "��������������������������"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmHostSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmHostSend.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�r�W�R�����M�\���t�H�[��
' �@�{���W���[���̓r�W�R�����M�\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private cCallBackObject As Object       ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer          ''�R�[���o�b�N�h�c�i�[
Private strResultError As String    ''�ʐM�f�[�^��̃G���[�d��
Private bTimeOutFlag As Boolean         ''�ʐMTimeout�t���O�@True�FTimeout�����@False�FTimeout�������Ȃ�
                
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

' @(f)
'
' �@�\      : �L�����Z���̉���
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �L�����Z������������B
'
' ���l      : �t�H�[�����A�����[�h����B
'
Private Sub cmdCancelClose()
    
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResCANCEL
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' �@�\      : �n�j�̉���
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �n�j����������B
'
' ���l      : �t�H�[�����A�����[�h����B
'
Private Sub cmdOKClose()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' �@�\      : �X�L�b�v�̉���
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L�b�v����������B
'
' ���l      : �t�H�[�����A�����[�h����B
'
Private Sub cmdSKIPClose()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResSKIP
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' �@�\      : �n�j�{�^��/�đ��M�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �n�j�{�^������/�đ��M�{�^������
'
' ���l      : �G���[�������e�m�F�n�j�i�����̓L�����Z���j
'
Private Sub cmdOK_Click()
    If iCallBackID = CALLBACK_HOSTSEND_QUERY Then
        Call HostSend       '�đ��M
    Else
        Call cmdCancelClose '�G���[������
    End If
End Sub

' @(f)
'
' �@�\      : �X�L�b�v�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L�b�v�{�^�������B
'
' ���l      : �G���[�������e�m�F�A�g�n�r�s���M�X�L�b�v�i�����͈ꎞ�n�j�����j
'
Private Sub cmdSKIP_Click()
    Call cmdSKIPClose '�G���[������
End Sub

' @(f)
'
' �@�\      : HOST���M���ʕ���
'
' ������    : ARG1 - ���M���ʃf�[�^
'
' �Ԃ�l    :
'
' �@�\����  : HOST���M���ʂ̏������s���B
'
' ���l      : �\�P�b�g�ʐM�Ή�
'
Private Sub HostSendResult(ByVal strRetData As String)
    Dim nI As Integer
    Dim strResult As String
    Dim strMIL_TITLE As String
    Dim strLBLINFO As String
    Dim nErrNo As Integer
    
    nErrNo = SetResultToAPRegistSlbData(strRetData)

    If nErrNo = 0 Then
        '�G���[�����n�j
        Call cmdOKClose
        Exit Sub
    ElseIf nErrNo > 0 Then
        '�r�W�R������G���[�L��
        lblInfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
        "�r�W�R������G���[���ʒm����܂����B" & vbCrLf & _
        "���e:"
        
        'COLORSYS
'        For nI = 0 To 9
        strResult = strResultError
'        If Trim(strResult) = "" Then Exit For
        lblInfo.Caption = lblInfo.Caption & Trim(strResult) & vbCrLf & "     "
'        Next nI
        
        strLBLINFO = lblInfo.Caption
        
    Else
        
        If nErrNo = -999 Then
            '�n�b�w���������i�^�C���A�E�g�j�̃G���[
            lblInfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�r�W�R���̉���������܂���B�^�C���A�E�g���܂����B" & vbCrLf & "�ʐM�����������ݒ肳��Ă��邩�m�F���Ă��������B"
            
            strLBLINFO = lblInfo.Caption
                
        ElseIf nErrNo = -888 Then
            '�ʐM�G���[�L��
            lblInfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�r�W�R���̉���������܂���B" & vbCrLf & "�ʐM�����������ݒ肳��Ă��邩�m�F���Ă��������B"
            
            strLBLINFO = lblInfo.Caption
        
        ElseIf nErrNo = -777 Then
            '�ʐM�G���[�L��
            lblInfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "��M�f�[�^�擪�ɖ��ߍ��ރf�[�^����" & vbCrLf & "�Ǝ�M�f�[�^�o�C�g���������܂���B"
            
            strLBLINFO = lblInfo.Caption
        
        Else
            '���̑��i��M�f�[�^�t�H�[�}�b�g�j�G���[�L��
            lblInfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�r�W�R������̎�M�f�[�^�ɖ����ȕ���������܂��B" & vbCrLf & "[" & strRetData & "]"
            
            strLBLINFO = lblInfo.Caption
            
        End If
    End If
    
    '2002-03-08 LOG�ɕۑ�
    Call MsgLog(conProcNum_BSCONT, strLBLINFO)
        
    '�{�^������
    If nErrNo > 0 And iCallBackID = CALLBACK_HOSTSEND Then  '���ѓo�^�A�r�W�R������G���[�ʒm������ꍇ
        cmdOK.Visible = True
        cmdSKIP.Visible = False '�X�L�b�v�s��
    Else
        cmdOK.Visible = True
        cmdSKIP.Visible = True
    End If

End Sub

' @(f)
'
' �@�\      : �o�^�����f�[�^�̕���
'
' ������    : ARG1 - �f�[�^
'
' �Ԃ�l    : 0=����I���^��=�ُ�I���i-999�F�ʐM�^�C���A�E�g�@-888�F�ʐM�ł��܂���@-1�F��M�f�[�^�t�H�[�}�b�g�G���[�j
'
' �@�\����  : �o�^�����f�[�^�̕������s���B
'
' ���l      : COLORSYS
'
Private Function SetResultToAPRegistSlbData(ByVal strRetData As String) As Integer
    Dim strBuf As String
    Dim nRet As Integer
    Dim nI As Integer
    Dim sStr As String
    
    sStr = strRetData

    nRet = 0

    'Debug���[�h����SKIP���[�h
    If IsDEBUG("HOSTDATA_SKIP") Then
        Call MsgLog(conProcNum_BSCONT, "��M�f�[�^(���ѓo�^)HOST�X�L�b�v:Err No.[0] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    '�ʐM���[�h
    If strRetData = "" And IsDEBUG("HOSTDATA_DEBUG") = False Then
        If bTimeOutFlag = False Then
            nRet = -888    '�ʐM���[�h�F�ʐM�G���[������
        Else
            nRet = -999    '�ʐM���[�h�F�ʐMTimeout������
        End If
        Call MsgLog(conProcNum_BSCONT, "��M�f�[�^(���ѓo�^):Err No.[" & Format(nRet, "#0") & "] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If

     '�R�[�h�]��
    strRetData = StrConv(strRetData, vbFromUnicode)
    
    '�G���[�R�[�h�擾
    strBuf = Mid(sStr, 24, 2)
    If Trim(strBuf) = "" Then
        nRet = -1                        ' �t�H�[�}�b�g�G���[
    Else
        If IsNumeric(strBuf) = False Then
            nRet = -1                    ' �t�H�[�}�b�g�G���[
        Else
            nRet = CInt(strBuf)
        End If
    End If
    
    'LOG�ɕۑ�
    If IsDEBUG("HOSTDATA_DEBUG") Then
        Call MsgLog(conProcNum_BSCONT, "��M�f�[�^(���ѓo�^)HOST�f�o�b�N:Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
    Else
        Call MsgLog(conProcNum_BSCONT, "��M�f�[�^(���ѓo�^):Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
    End If
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "�G���[�R�[�h[" & strBuf & "]")
    End If
    
    'H1 �f�[�^���`�F�b�N
    strBuf = StrConv(MidB(strRetData, 1, 4), vbUnicode)
    
    If strBuf <> "0000" Then
        nRet = -777                    ' ��M�f�[�^��������Ȃ��ꍇ
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    'H2 �[����
    strBuf = StrConv(MidB(strRetData, 5, 3), vbUnicode)
    
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "�[����[" & strBuf & "]")
    End If
    
    'H3 �g�����U�N�V����
    strBuf = StrConv(MidB(strRetData, 8, 8), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "�g�����U�N�V����[" & strBuf & "]")
    End If

    '��
    strBuf = StrConv(MidB(strRetData, 16, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "��[" & strBuf & "]")
    End If

    '���
    strBuf = StrConv(MidB(strRetData, 21, 3), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "���[" & strBuf & "]")
    End If

    '�G���[�R�[�h
    strBuf = StrConv(MidB(strRetData, 24, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "�G���[�R�[�h[" & strBuf & "]")
    End If

    If IsNumeric(strBuf) Then
        nRet = CInt(strBuf)
    Else
        nRet = -1                        ' �t�H�[�}�b�g�G���[
    End If

    '�G���[���e
    strBuf = StrConv(MidB(strRetData, 26, 50), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "�G���[���e[" & strBuf & "]")
    End If
    
    strResultError = strBuf

    '��
    strBuf = StrConv(MidB(strRetData, 76, 25), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_BSCONT, "��[" & strBuf & "]")
    End If


    SetResultToAPRegistSlbData = nRet
End Function


' @(f)
'
' �@�\      : HOST�ʐM���O�o�̓C�x���g
'
' ������    : ARG1 - �߂�o�̓��O
'
' �Ԃ�l    :
'
' �@�\����  : HOST�ʐM���O�o�̓C�x���g���̏������s���B
'
' ���l      :
'
Private Sub WCSockControl1_ProcessLog(ByVal strBuf As String)
    Call MsgLog(conProcNum_WINSOCKCONT, strBuf)
End Sub

' @(f)
'
' �@�\      : �t�H�[��Active
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �t�H�[��Active���̏������s���B
'
' ���l      :
'
Private Sub Form_Activate()
    
    If iCallBackID = CALLBACK_HOSTSEND_QUERY Then
        Me.Caption = "�X���u���r�W�R���₢���킹"
        cmdOK.Caption = "�đ��M"
    Else
        Me.Caption = "���уf�[�^�r�W�R���o�^"
        cmdOK.Caption = "OK"
   End If
    
   'iHostSendCount = 1
    
    WCSockControl1.RemotePort = APSysCfgData.nHOST_PORT '�r�W�R����PortNo
    WCSockControl1.RemoteHost = APSysCfgData.HOST_IP '�r�W�R����IP
    WCSockControl1.ConnectTimeOut = APSysCfgData.nHOST_TOUT(1) '�I�[�v�����̃^�C���A�E�g �b�Ŏw��
    WCSockControl1.SendTimeOut = APSysCfgData.nHOST_TOUT(2) '�f�[�^�ʐM���̃^�C���A�E�g �b�Ŏw��
    WCSockControl1.RetryTimes = APSysCfgData.nHOST_RETRY '�ʐM�����g���C��
        
    Call HostSend
    
End Sub

' @(f)
'
' �@�\      : �r�W�R�����M����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �r�W�R�����M�������s���B
'
' ���l      :
'
Private Sub HostSend()
    Dim strRet As String
    Dim strSendString As String
    Dim sLocalPath As String
    Dim lLen As Long
    
    '���M
    sLocalPath = App.path & "\"
    
    'If IsDEBUG("DISP") Then
    '    Me.Height = 7635
    '    Me.Width = 14250
    'Else
    '    Me.Height = 2715
    '    Me.Width = 8490
    'End If
    
    strSendString = GetAPResHostData()
    lLen = 150 'COLORSYS
    
    lblInfo.Caption = "���M���ł��B" & vbCrLf & "���΂炭���҂����������B"
    
    strRet = ""
    
    If IsDEBUG("HOSTDATA_DEBUG") Or IsDEBUG("HOSTDATA_SKIP") Then
        WriteHostData sLocalPath & "ReqHostData.txt", strSendString

        If IsDEBUG("HOSTDATA_DEBUG") Then
            strRet = ReadHostData(sLocalPath & "RegistResult.txt")
        End If
        Call HostSendResult(strRet)
        Exit Sub
    End If
    
    bTimeOutFlag = False
    
    timTimeOut.Enabled = False ''�Ď��^�C�}�[�n�e�e
    If APSysCfgData.nHOST_TOUT(0) <> 0 Then
        timTimeOut.Interval = APSysCfgData.nHOST_TOUT(0) * 1000 '�S�̊Ď��̃^�C���A�E�g �b����mS�ɕϊ�
        timTimeOut.Enabled = True '�Ď��^�C�}�[�n�m
    End If
     
    strRet = WCSockControl1.WCSSingleSendRec(strSendString, 1, lLen)
     
    timTimeOut.Enabled = False ''�Ď��^�C�}�[�n�e�e
     
    Call HostSendResult(strRet)
    
End Sub

' @(f)
'
' �@�\      : HOST���M�p�f�[�^�쐬
'
' ������    :
'
' �Ԃ�l    : HOST���M�p�f�[�^
'
' �@�\����  : HOST���M�p�f�[�^�̍쐬���s���B
'
' ���l      :COLORSYS
'
Private Function GetAPResHostData() As String
    Dim strSendString As String
    Dim nI As Integer
    Dim nJ As Integer
    
    strSendString = "XXXX"     '�f�[�^��(1300Bytes) Winsock OCX�Ŗ��ߍ���
    strSendString = strSendString & "A96"     'H2 �[����
    strSendString = strSendString & "EA96"   'H3 �g�����U�N�V������
    strSendString = strSendString & Space(4)    'H3 �]��
    strSendString = strSendString & Space(5)    '1 ��
    strSendString = strSendString & "A96"   '2 ���
  
  
    strSendString = strSendString & _
    Format(Left(APResData.slb_no, 9), "!@@@@@@@@@") '�X���uNo
  
  
    '�ďo���ɂ��A��������
    Select Case cCallBackObject.Name
        '*************************************************************
        Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
            strSendString = strSendString & APResData.host_send_flg
        
            If APResData.host_wrt_dte = "" Then
                APResData.host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
  
        
        '*************************************************************
        Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
            strSendString = strSendString & APResData.host_send_flg
        
            If APResData.fail_host_wrt_dte = "" Then
                APResData.fail_host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.fail_host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
        
  
        '*************************************************************
        Case "frmDirResWnd" ''���u���ʓ���(�����t���O�ȊO�A�X���u�ُ�񍐂Ɠ����j
            strSendString = strSendString & APResData.host_send_flg
            
            If APResData.fail_host_wrt_dte = "" Then
                APResData.fail_host_wrt_dte = Format(Now, "YYYYMMDD")
                APResData.fail_host_wrt_tme = Format(Now, "HHMMSS")
            End If
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
            
            strSendString = strSendString & _
            Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
        
  
        '*************************************************************
        Case "frmColorSlbSelWnd" ''�J���[�`�F�b�N���ה�����폜�����i�J���[�`�F�b�N�|�X���u�I����ʁj
            
            strSendString = strSendString & "9" '����R�[�h
            
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_dte <> "" Then
                '�J���[�`�F�b�N�����\�ő��M�ς�
                strSendString = strSendString & _
                Format(Left(APResData.host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
            
            ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte <> "" Then
                '�X���u�ُ�񍐂ő��M�ς�
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
                
            ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_cmp_flg = "1" Then
                '���u���ʕ񍐂ő��M�ς�
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_dte, 8), "!@@@@@@@@") '��Ɠ��i�N�{���{���jYYYYMMDD
                
                strSendString = strSendString & _
                Format(Left(APResData.fail_host_wrt_tme, 4), "!@@@@") '��Ɠ��i���{���jHHMM
                
            End If

        Case Else
            Call WaitMsgBox(Me, "frmHostSend:GetAPResHostData:�ďo���G���[")
            Call MsgLog(conProcNum_BSCONT, "frmHostSend:GetAPResHostData:�ďo���G���[")
    End Select
  
    If APResData.slb_fault_u_judg = "9" Then
        strSendString = strSendString & "*"
    Else
        strSendString = strSendString & Format(Left(APResData.slb_fault_u_judg & " ", 1), "!@")   '��ʔ���
    End If
  
    If APResData.slb_fault_d_judg = "9" Then
        strSendString = strSendString & "*"
    Else
        strSendString = strSendString & Format(Left(APResData.slb_fault_d_judg & " ", 1), "!@")   '���ʔ���
    End If
  
    strSendString = strSendString & Space(103)    '��
  
  
    GetAPResHostData = strSendString

    Debug.Print "SendData:[" & strSendString & "]"
    Call MsgLog(conProcNum_BSCONT, "���M�f�[�^(���ѓo�^):[" & strSendString & "]")

End Function

' @(f)
'
' �@�\      : HOST��M�p�_�~�[�Ǐo
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : HOST�����ł̃f�o�b�O�p�f�[�^�Ǐo����
'
' ���l      :
'
Private Function ReadHostData(ByVal sFileName As String)
    Dim strReadData As String
    Dim StrTmp As String
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Input As #fp
    Do While Not EOF(fp)
        Line Input #fp, StrTmp
        strReadData = strReadData & StrTmp
    Loop
    Close #fp
    ReadHostData = strReadData
End Function

' @(f)
'
' �@�\      : HOST��M�p�_�~�[����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : HOST�����ł̃f�o�b�O�p�f�[�^��������
'
' ���l      :
'
Private Sub WriteHostData(ByVal sFileName As String, ByVal sSendData As String)
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Output Access Write As #fp
    Print #fp, sSendData
    Close #fp
End Sub

' @(f)
'
' �@�\      : HOST���M�p�^�C���A�E�g�C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : HOST���M�p�^�C���A�E�g�C�x���g���̏������s���B
'
' ���l      : �u�a���n�b�w�^�C���A�E�g�Ď��p
'           :
'

Private Sub timTimeOut_Timer()

    bTimeOutFlag = True
    timTimeOut.Enabled = False '�Ď��^�C�}�[�n�e�e
    WCSockControl1.WCSForceEnd   '�r�W�R���ʐM�����I��
End Sub

