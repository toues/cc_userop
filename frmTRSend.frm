VERSION 5.00
Object = "{B6B49C41-8023-4CA6-BDF0-FC5291FC6D71}#18.0#0"; "WCSockControl.ocx"
Begin VB.Form frmTRSend 
   BackColor       =   &H00C0FFFF&
   Caption         =   "TR�T�[�o�[�ʐM"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   8865
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.Timer timTimeOut 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   7620
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSKIP 
      Caption         =   "�X�L�b�v"
      Height          =   435
      Left            =   7620
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   975
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
Attribute VB_Name = "frmTRSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmTRSend.Frm                ver 1.00

' @(s)
' �J���[�`�F�b�N���тo�b�@�ʐM�T�[�o�[���M�\���t�H�[��
' �@�{���W���[���͒ʐM�T�[�o�[���M�\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private cCallBackObject As Object       ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer          ''�R�[���o�b�N�h�c�i�[
Private sCmdID As String                ''���M�R�}���hID�w�� '2008/09/04
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
' ���l      :2008/09/04 CmdID �ǉ�
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer, ByVal CmdID As String)
    iCallBackID = ObjctID
    sCmdID = CmdID '2008/09/04 CmdID �ǉ�
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
    Call cmdCancelClose '�G���[������
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
' ���l      : �G���[�������e�m�F�A�ʐM�T�[�o�[���M�X�L�b�v�i�����͈ꎞ�n�j�����j
'
Private Sub cmdSKIP_Click()
    Call cmdSKIPClose '�G���[������
End Sub

' @(f)
'
' �@�\      : �ʐM�T�[�o�[���M���ʕ���
'
' ������    : ARG1 - ���M���ʃf�[�^
'
' �Ԃ�l    :
'
' �@�\����  : �ʐM�T�[�o�[���M���ʂ̏������s���B
'
' ���l      : �\�P�b�g�ʐM�Ή�
'
Private Sub TrSendResult(ByVal strRetData As String)
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
        lblinfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
        "�ʐM�T�[�o�[����G���[���ʒm����܂����B" & vbCrLf & _
        "���e:"
        
        'COLORSYS
'        For nI = 0 To 9
        strResult = strResultError
'        If Trim(strResult) = "" Then Exit For
        lblinfo.Caption = lblinfo.Caption & Trim(strResult) & vbCrLf & "     "
'        Next nI
        
        strLBLINFO = lblinfo.Caption
        
    Else
        
        If nErrNo = -999 Then
            '�n�b�w���������i�^�C���A�E�g�j�̃G���[
            lblinfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�ʐM�T�[�o�[�̉���������܂���B�^�C���A�E�g���܂����B" & vbCrLf & "�ʐM�����������ݒ肳��Ă��邩�m�F���Ă��������B"
            
            strLBLINFO = lblinfo.Caption
                
        ElseIf nErrNo = -888 Then
            '�ʐM�G���[�L��
            lblinfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�ʐM�T�[�o�[�̉���������܂���B" & vbCrLf & "�ʐM�����������ݒ肳��Ă��邩�m�F���Ă��������B"
            
            strLBLINFO = lblinfo.Caption
        
        ElseIf nErrNo = -777 Then
            '�ʐM�G���[�L��
            lblinfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "��M�f�[�^�擪�ɖ��ߍ��ރf�[�^����" & vbCrLf & "�Ǝ�M�f�[�^�o�C�g���������܂���B"
            
            strLBLINFO = lblinfo.Caption
        
        Else
            '���̑��i��M�f�[�^�t�H�[�}�b�g�j�G���[�L��
            lblinfo.Caption = "�G���[�ԍ�:" & CStr(nErrNo) & "�@" & _
            "�ʐM�G���[���������܂����B" & vbCrLf & _
            "���e:" & "�ʐM�T�[�o�[����̎�M�f�[�^�ɖ����ȕ���������܂��B" & vbCrLf & "[" & strRetData & "]"
            
            strLBLINFO = lblinfo.Caption
            
        End If
    End If
    
    '2002-03-08 LOG�ɕۑ�
    Call MsgLog(conProcNum_TRCONT, strLBLINFO)
        
    '�{�^������
'    If nErrNo > 0 And iCallBackID = CALLBACK_TRSEND Then  '�o�c�e�쐬�v���A�ʐM�T�[�o�[����G���[�ʒm������ꍇ
        cmdOK.Visible = True
        cmdSKIP.Visible = False '�X�L�b�v�s��
'    Else
'        cmdOK.Visible = True
'        cmdSKIP.Visible = True
'    End If

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
    If IsDEBUG("TR_SKIP") Then
        Call MsgLog(conProcNum_TRCONT, "��M�f�[�^(�o�c�e�쐬�v��)TR�X�L�b�v:Err No.[0] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    '�ʐM���[�h
'    If strRetData = "" And IsDEBUG("HOSTDATA_DEBUG") = False Then
    If strRetData = "" And IsDEBUG("TR_SKIP") = False Then
        If bTimeOutFlag = False Then
            nRet = -888    '�ʐM���[�h�F�ʐM�G���[������
        Else
            nRet = -999    '�ʐM���[�h�F�ʐMTimeout������
        End If
        Call MsgLog(conProcNum_TRCONT, "��M�f�[�^(���ѓo�^):Err No.[" & Format(nRet, "#0") & "] RetData:[]")
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If

     '�R�[�h�]��
    strRetData = StrConv(strRetData, vbFromUnicode)
    
    '�G���[�R�[�h�擾
    strBuf = Mid(sStr, 27, 2)
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
'    If IsDEBUG("HOSTDATA_DEBUG") Then
'        Call MsgLog(conProcNum_TRCONT, "��M�f�[�^(�o�c�e�쐬�v��)HOST�f�o�b�N:Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
'    Else
        Call MsgLog(conProcNum_TRCONT, "��M�f�[�^(�o�c�e�쐬�v��):Err No.[" & Format(nRet, "#0") & "] RetData:[" & sStr & "]")
'    End If
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�G���[�R�[�h[" & strBuf & "]")
    End If
    
    'H1 �f�[�^���`�F�b�N
    strBuf = StrConv(MidB(strRetData, 1, 4), vbUnicode)
    
    If strBuf <> "0000" Then
        nRet = -777                    ' ��M�f�[�^��������Ȃ��ꍇ
        SetResultToAPRegistSlbData = nRet
        Exit Function
    End If
    
    'H2 ���M�N��������
    strBuf = StrConv(MidB(strRetData, 5, 12), vbUnicode)
    
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "���M�N��������[" & strBuf & "]")
    End If
    
    'H3 ���b�Z�[�W�h�c
    strBuf = StrConv(MidB(strRetData, 17, 4), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "���b�Z�[�W�h�c[" & strBuf & "]")
    End If

    'H4 �����h�c
    strBuf = StrConv(MidB(strRetData, 21, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�����h�c[" & strBuf & "]")
    End If

    'H5 �`�����
    strBuf = StrConv(MidB(strRetData, 26, 1), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�`�����[" & strBuf & "]")
    End If

    '�G���[�R�[�h
    strBuf = StrConv(MidB(strRetData, 27, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�G���[�R�[�h[" & strBuf & "]")
    End If

    If IsNumeric(strBuf) Then
        nRet = CInt(strBuf)
    Else
        nRet = -1                        ' �t�H�[�}�b�g�G���[
    End If

    '�G���[���e
    strBuf = StrConv(MidB(strRetData, 29, 144), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�G���[���e[" & strBuf & "]")
    End If
    
    strResultError = strBuf

    '�`���[�W�m�n
    strBuf = StrConv(MidB(strRetData, 173, 5), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�`���[�W�m�n[" & strBuf & "]")
    End If

    '����
    strBuf = StrConv(MidB(strRetData, 178, 4), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "����[" & strBuf & "]")
    End If

    '���
    strBuf = StrConv(MidB(strRetData, 182, 1), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "���[" & strBuf & "]")
    End If

    '�J���[��
    strBuf = StrConv(MidB(strRetData, 183, 2), vbUnicode)
    If IsDEBUG("FILE") Then
        Call MsgLog(conProcNum_TRCONT, "�J���[��[" & strBuf & "]")
    End If

    SetResultToAPRegistSlbData = nRet

End Function


' @(f)
'
' �@�\      : TR�ʐM���O�o�̓C�x���g
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
    
    Select Case sCmdID
        Case "COL01"
            Me.Caption = "�o�c�e�쐬�v�����M"
        Case "COL02"
            Me.Caption = "�w������v�����M"
    End Select
    
    cmdOK.Caption = "OK"
    
   'iHostSendCount = 1
    
    WCSockControl1.RemotePort = APSysCfgData.nTR_PORT '�ʐM�T�[�o�[��PortNo
    WCSockControl1.RemoteHost = APSysCfgData.TR_IP '�r�W�R����IP
    WCSockControl1.ConnectTimeOut = APSysCfgData.nTR_TOUT(1) '�I�[�v�����̃^�C���A�E�g �b�Ŏw��
    WCSockControl1.SendTimeOut = APSysCfgData.nTR_TOUT(2) '�f�[�^�ʐM���̃^�C���A�E�g �b�Ŏw��
    WCSockControl1.RetryTimes = APSysCfgData.nTR_RETRY '�ʐM�����g���C��
        
    Call TrSend
    
End Sub

' @(f)
'
' �@�\      : �ʐM�T�[�o�[���M����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ʐM�T�[�o�[���M�������s���B
'
' ���l      :
'
Private Sub TrSend()
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
    
    strSendString = GetAPResTrData()
    lLen = 38 'COLORSYS
    
    lblinfo.Caption = "���M���ł��B" & vbCrLf & "���΂炭���҂����������B"
    
    strRet = ""
    
'    If IsDEBUG("HOSTDATA_DEBUG") Or IsDEBUG("HOSTDATA_SKIP") Then
'        WriteHostData sLocalPath & "ReqHostData.txt", strSendString
'
'        If IsDEBUG("HOSTDATA_DEBUG") Then
'            strRet = ReadHostData(sLocalPath & "RegistResult.txt")
'        End If
'        Call ApSendResult(strRet)
'        Exit Sub
'    End If
    
    If IsDEBUG("TR_SKIP") Then
        Select Case sCmdID
            Case "COL01"
                WriteTrData sLocalPath & "ReqTrDataCOL01.txt", strSendString
                strRet = "0000YYYYMMDDhhmmPC01COL01000" & Space(144) & "123451234112"
            Case "COL02"
                WriteTrData sLocalPath & "ReqTrDataCOL02.txt", strSendString
                strRet = "0000YYYYMMDDhhmmPC01COL02000" & Space(144) & "123451234112"
        End Select
        
        Call TrSendResult(strRet)
        Exit Sub
    End If
    
    bTimeOutFlag = False
    
    timTimeOut.Enabled = False ''�Ď��^�C�}�[�n�e�e
    If APSysCfgData.nTR_TOUT(0) <> 0 Then
        timTimeOut.Interval = APSysCfgData.nTR_TOUT(0) * 1000 '�S�̊Ď��̃^�C���A�E�g �b����mS�ɕϊ�
        timTimeOut.Enabled = True '�Ď��^�C�}�[�n�m
    End If
     
    strRet = WCSockControl1.WCSSingleSendRec(strSendString, 1, lLen)
     
    timTimeOut.Enabled = False ''�Ď��^�C�}�[�n�e�e
     
    Call TrSendResult(strRet)
    
End Sub

' @(f)
'
' �@�\      : TR���M�p�f�[�^�쐬
'
' ������    :
'
' �Ԃ�l    : TR���M�p�f�[�^
'
' �@�\����  : TR���M�p�f�[�^�̍쐬���s���B
'
' ���l      :COLORSYS
'
Private Function GetAPResTrData() As String
    Dim strSendString As String
    Dim nI As Integer
    Dim nJ As Integer
    
    strSendString = "XXXX"     '�f�[�^��(1300Bytes) Winsock OCX�Ŗ��ߍ���
    strSendString = strSendString & Format(Now, "YYYYMMDDHHMM")    'H2 ���M�N��������
    strSendString = strSendString & "PC01"   'H3 ���b�Z�[�W�h�c
    strSendString = strSendString & sCmdID   'H4 �����h�c
    
    'H5 �`�����
    Select Case sCmdID
        Case "COL01"
            'Me.Caption = "�o�c�e�쐬�v�����M"
            '�ďo���ɂ��A��������
            Select Case cCallBackObject.Name
                '*************************************************************
                Case "frmSkinScanWnd" ''�X���u�������\����
                    strSendString = strSendString & CStr(conDefine_SYSMODE_SKIN)     'H5 �`�����
                '*************************************************************
                Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
                    strSendString = strSendString & CStr(conDefine_SYSMODE_COLOR)     'H5 �`�����
                '*************************************************************
                Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
                    strSendString = strSendString & CStr(conDefine_SYSMODE_SLBFAIL)     'H5 �`�����
                '*************************************************************
                Case Else
                    Call WaitMsgBox(Me, "frmTRSend:GetAPResTrData:�ďo���G���[")
                    Call MsgLog(conProcNum_TRCONT, "frmTRSend:GetAPResTrData:�ďo���G���[")
                    GetAPResTrData = ""
                    Exit Function
            End Select
        Case "COL02"
            'Me.Caption = "�w������v�����M"
            strSendString = strSendString & "0" 'H5 �`����� ���g�p�̂��߁A0:�Œ�
    End Select
    
    
    strSendString = strSendString & _
    Format(Left(APResData.slb_no, 9), "!@@@@@@@@@") '�X���uNo
  
    strSendString = strSendString & _
    Format(Left(APResData.slb_stat, 1), "!@") '���
  
    strSendString = strSendString & Format(CInt(APResData.slb_col_cnt), "00") ''�J���[��
  
    GetAPResTrData = strSendString

    Debug.Print "SendData:[" & strSendString & "]"
    Call MsgLog(conProcNum_TRCONT, "���M�f�[�^(�o�c�e�쐬�v��):[" & strSendString & "]")

End Function


' @(f)
'
' �@�\      : TR���M�p�^�C���A�E�g�C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : TR���M�p�^�C���A�E�g�C�x���g���̏������s���B
'
' ���l      : �u�a���n�b�w�^�C���A�E�g�Ď��p
'           :
'

Private Sub timTimeOut_Timer()

    bTimeOutFlag = True
    timTimeOut.Enabled = False '�Ď��^�C�}�[�n�e�e
    WCSockControl1.WCSForceEnd   '�r�W�R���ʐM�����I��
End Sub

' @(f)
'
' �@�\      : TR��M�p�_�~�[����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : TR�����ł̃f�o�b�O�p�f�[�^��������
'
' ���l      :
'
Private Sub WriteTrData(ByVal sFileName As String, ByVal sSendData As String)
    Dim fp As Integer

    fp = FreeFile
    Open sFileName For Output Access Write As #fp
    Print #fp, sSendData
    Close #fp
End Sub

