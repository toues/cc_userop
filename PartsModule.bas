Attribute VB_Name = "PartsModule"
' @(h) PartsModule.Bas                ver 1.00 ( '01.10.01 SEC Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�֐��p�[�c���W���[��
' �@�{���W���[���͖{�V�X�e���Ŏg�p����֐��p�[�c���W�߂�
' �@���̂ł���B

Option Explicit
    
' @(f)
'
' �@�\      : �\���p��ԕ�����擾
'
' ������    : ARG1 - nMode 0:���ޔ��A1:�װ����
'          : ARG2 - ��Ԕԍ�
'
' �Ԃ�l    : ��ԕ�����
'
' �@�\����  : ��Ԕԍ����R�����g�t���ɕϊ�����B
'
' ���l      :
'
Public Function ConvDpOutStat(ByVal nSysMode As Integer, nStat As Integer) As String

    Select Case nStat
        Case conDefine_SYSMODE_SKIN
            ConvDpOutStat = IIf(nSysMode = conDefine_SYSMODE_SKIN, "0:����", "0:����")
        Case Else
            ConvDpOutStat = CStr(nStat) & ":" & CStr(nStat) & "ht��"
    End Select
    
End Function
    
' @(f)
'
' �@�\      : �X���u�ԍ��ϊ�
'
' ������    : ARG1 - �n�C�t���t���X���u�ԍ�
'
' �Ԃ�l    : �n�C�t�������X���u�ԍ�
'
' �@�\����  : �w�肵���n�C�t���t���X���u�ԍ����f�|�f�n�C�t�������X���u�ԍ��ɕϊ�����B
'
' ���l      : �A�X�^���X�N�f���f������ꍇ�́A�p�[�Z���g�f���f�ɕϊ�����B
'           :COLORSYS
'
Public Function ConvSearchSlbNumber(ByVal strSearchSlbNumber As String) As String
    Dim nI As Integer
    Dim strResSearchSlbNumber As String
    
    '�n�C�t���f�|�f������Ď��ۂ̌���������֕ϊ�
    For nI = 1 To Len(strSearchSlbNumber)
        If Mid(strSearchSlbNumber, nI, 1) <> "-" Then
            If Mid(strSearchSlbNumber, nI, 1) = "*" Then
                strResSearchSlbNumber = strResSearchSlbNumber & "%"
            Else
                strResSearchSlbNumber = strResSearchSlbNumber & Mid(strSearchSlbNumber, nI, 1)
            End If
        End If
    Next nI
    
    ConvSearchSlbNumber = strResSearchSlbNumber
    
End Function
    
' @(f)
'
' �@�\      : �f�o�b�N���[�h����
'
' ������    : ARG1 - �f�o�b�N���[�h������
'
' �Ԃ�l    : True=�f�o�b�N�n�m�^False=�f�o�b�N�n�e�e
'
' �@�\����  : �w�肵���f�o�b�N���[�h�̏�Ԃ𔻕ʂ���B
'
' ���l      :
'
Public Function IsDEBUG(ByVal strDEBUG As String) As Boolean

    IsDEBUG = False
    
    If APSysCfgData.nDEBUG_MODE <> 1 Then Exit Function
    
    Select Case strDEBUG
        Case "DISP"
            If APSysCfgData.nDISP_DEBUG = 1 Then IsDEBUG = True
        Case "FILE"
            If APSysCfgData.nFILE_DEBUG = 1 Then IsDEBUG = True
        Case "TR_SKIP"
            If APSysCfgData.nTR_SKIP = 1 Then IsDEBUG = True
        Case "DB_SKIP"
            If APSysCfgData.nDB_SKIP = 1 Then IsDEBUG = True
        Case "SOZAI_DB_SKIP"
            If APSysCfgData.nSOZAI_DB_SKIP = 1 Then IsDEBUG = True
        Case "SCAN"
            If APSysCfgData.nSCAN_SKIP = 1 Then IsDEBUG = True
        Case "HOSTDATA_DEBUG"
            If APSysCfgData.nHOSTDATA_DEBUG = 1 Then IsDEBUG = True
        Case "HOSTDATA_SKIP"
            If APSysCfgData.nHOSTDATA_SKIP = 1 Then IsDEBUG = True
        'end cho
    End Select
    
End Function
    
'' @(f)
''
'' �@�\      : �X���u�������[�h����
''
'' ������    :
''
'' �Ԃ�l    : True=�����n�m�^False=�����n�e�e
''
'' �@�\����  : �X���u�������[�h�̏�Ԃ𔻕ʂ���B
''
'' ���l      :
''
'Public Function IsAPSplit() As Boolean
'    '�X���u�͑I������Ă��邩�B
'    If APSlbCont.nListSelectedIndexP1 = 0 Then
'        IsAPSplit = False
'    Else
'        '�������[�h�̏ꍇ�B
'        If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).nSplitTotal > 1 Then
'            IsAPSplit = True
'        Else
'            IsAPSplit = False
'        End If
'    End If
'End Function

' @(f)
'
' �@�\      : ���b�Z�[�W���O�̍쐬�A�\���A�ۑ�
'
' ������    : ARG1 - �v���Z�X�ԍ�
'             ARG2 - ���b�Z�[�W
'
' �Ԃ�l    :
'
' �@�\����  : ���b�Z�[�W���O�̍쐬�A�\���A�ۑ����s���B
'
' ���l      : �K�C�_���X�\���B
'
Public Sub MsgLog(ByVal nProcNum As Integer, ByVal strMessage As String)
    Dim strGuidanceMess As String
    
    Select Case nProcNum
        Case conProcNum_MAIN
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
            
            strGuidanceMess = Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision & conDefault_Separator & strMessage
        Case conProcNum_BSCONT
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & "�r�W�R���ʐM�F" & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
              
        Case conProcNum_TRCONT
            If fMainWnd.lstGuidance.ListCount >= conDefine_lGuidanceListMAX Then
                fMainWnd.lstGuidance.RemoveItem 0
            End If
            strGuidanceMess = Now & Space(1) & "�ʐM�T�[�o�[�ʐM�F" & strMessage
            fMainWnd.lstGuidance.AddItem strGuidanceMess
            fMainWnd.lstGuidance.ListIndex = fMainWnd.lstGuidance.ListCount - 1
    
        Case conProcNum_MAINTENANCE
            strGuidanceMess = Now & Space(1) & App.title & " Ver." & App.Major & "." & App.Minor & "." & App.Revision & conDefault_Separator & "�����e�i���X:" & strMessage
        
        Case conProcNum_WINSOCKCONT
            strGuidanceMess = strMessage
            
    End Select
    
    If IsEmpty(MainLogFileNumber) = False Then
        Print #MainLogFileNumber, strGuidanceMess
    End If

End Sub

' @(f)
'
' �@�\      : INPUT MAN�p���̓`�F�b�N
'
' ������    : ARG1 -�@INPUT MAN�@�I�u�W�F�N�g
'
' �Ԃ�l    : TRUE/FALSE
'
' �@�\����  : INPUT MAN�@CAPTION��TEXT�Ɏw�肳��Ă���A�����𖞂��������肵���ʂ�Ԃ��B
'             ��KeepFocus�֐ݒ�\
'
' ���l      : (�����l,����l)(������,...)
'
Public Function LimitCheck(ByVal obj As Object) As Boolean

    Dim obj_str As String
    Dim get_str As String
    Dim text_str, num_str As String
    Dim pos_int, pos_max_int, i, y As Integer
    Dim upper_sing, lower_sing As Single
    Dim array_str() As String
    Dim text_flag_bool As Boolean
    Dim obj_work_str As String

    On Error Resume Next

    '�I�u�W�F�N�g���狖������擾
    obj_str = obj.Caption.Text

    LimitCheck = False
    '���ʂ̗L������
    If (InStr(obj_str, "(") = False) And (InStr(obj_str, ")") = False) Then
        Exit Function
    End If
    obj_work_str = obj
    If obj_work_str = "" Then
        obj_work_str = " "
    End If
    
    '�����͈͎擾
    num_str = Mid(obj_str, 2, (InStr(obj_str, ")") - 2))
    '�������擾
    text_str = Mid(obj_str, InStr(obj_str, ")") + 2, (Len(obj_str) - Len(num_str) - 4))

    If IsNumeric(obj) Then
        '���l�`�F�b�N
        If Len(num_str) <> 0 Then
            '�ŏ��l�擾
            lower_sing = Mid(num_str, 1, (InStr(num_str, ",") - 1))
            '�ő�l�擾
            upper_sing = Mid(num_str, InStr(num_str, ",") + 1, Len(num_str) - 1)
            '�͈̓`�F�b�N
            If (CDbl(obj_work_str) < CDbl(lower_sing)) Or (CDbl(obj_work_str) > CDbl(upper_sing)) Then
                '�G���[���^�[��
                LimitCheck = True
                Exit Function
            Else
                Exit Function
            End If
        Else
            '�G���[���^�[��
            LimitCheck = True
            Exit Function
        End If
    Else
        '�����`�F�b�N
        pos_int = 1
        If Len(text_str) <> 0 Then
            '�������擾
            pos_max_int = Fix(Len(text_str) / 2) + 1
            '�z��̍Ē�`
            ReDim array_str(pos_max_int)
            '�������i�[
            For i = 0 To pos_max_int - 1
                '�����擾
                array_str(i) = Mid(text_str, pos_int, 1)
                '�擾�ʒu�X�V
                pos_int = pos_int + 2
            Next
            '�Ώە����������[�v
            For i = 0 To Len(obj_work_str) - 1
                '�Ώە������P�����擾
                get_str = Mid(obj_work_str, i + 1, 1)
                '�t���O�̏�����
                text_flag_bool = False
                '�����������[�v
                For y = 0 To pos_max_int - 1
                    '�������`�F�b�N
                    If (get_str = array_str(y)) Then
                        text_flag_bool = True
                    End If
                Next
                '����������
                If text_flag_bool = False Then
                    '�G���[���^�[��
                    LimitCheck = True
                    Exit Function
                End If
            Next
        Else
            '�G���[���^�[��
            LimitCheck = True
            Exit Function
        End If
    End If
End Function

' @(f)
'
' �@�\      : �V�X�e���������瑀�Ɠ��t�Z�o
'
' ������    :
'
' �Ԃ�l    : ���Ɠ�("YYYYMMDD")
'
' �@�\����  : �V�X�e���������瑀�Ɠ��t�Z�o
'
' ���l      :
'
Public Function GetSyoGyoDate() As String
    On Error Resume Next                        '�G���[����

    Dim strSysDate              As String       '�V�X�e�����t
    Dim strSysTime              As String       '�V�X�e������

    strSysDate = Format(Date$, "YYYY/MM/DD")
    strSysTime = Format(Time$, "HH:MM:SS")

    ':::: �V�X�e��������0���`7��30���ȍ~�Ȃ�ΑO���̓��t�Ƃ��Čv�Z
    If "00:00:00" <= strSysTime And strSysTime <= "07:29:59" Then
        GetSyoGyoDate = Format(DateAdd("d", -1, strSysDate), "YYYYMMDD")   '���Ɠ��t�Z�b�g
    Else
        GetSyoGyoDate = Format(strSysDate, "YYYYMMDD")                     '���Ɠ��t�Z�b�g
    End If

    Debug.Print GetSyoGyoDate
End Function
' @(f)
'
' �@�\      : INPUT MAN �t�H�[�}�b�g�ݒ菈��
'
' ������    : ARG1 - INPUT MAN �I�t�W�F�N�g
' �@�@�@    �FARG2 - Caption�ݒ�l
' �@�@�@    �FARG3 - Format�ݒ�l
' �@�@�@    �FARG4 - FormatMode�ݒ�l
' �@�@�@    �FARG5 - bAllowSpace�ݒ�l
'
' �Ԃ�l    : ����
'
' �@�\����  : INPUT MAN�̃t�H�[�}�b�g��ݒ肷��B
'
' ���l      :
'
Public Sub SetimTextFormat(ByVal obj As Object, ByVal strCaption As String, ByVal strFormat As String, ByVal iFormatMode As Integer, ByVal bAllowSpace As Boolean)
    obj.Caption.Text = strCaption
    obj.Format = strFormat
    obj.FormatMode = iFormatMode
    obj.AllowSpace = bAllowSpace
    obj.EditMode = 3 '�㏑���i�Œ�j
    obj.HighlightText = True '�e�L�X�g�I��
End Sub


Public Function cnvSplitNum(ByVal strSplitTNum As String) As Integer

    Select Case Trim(strSplitTNum)
        Case "1"
            cnvSplitNum = 1
        Case "2"
            cnvSplitNum = 2
        Case "3"
            cnvSplitNum = 3
        Case "4"
            cnvSplitNum = 4
        Case "5"
            cnvSplitNum = 5
        Case "6"
            cnvSplitNum = 6
        Case "7"
            cnvSplitNum = 7
        
        'Case "X"
        '    cnvSplitNum = 9
        'Case "Y"
        '    cnvSplitNum = 10
        '
        'Case "8"
        '    cnvSplitNum = 8
        'Case "9"
        '    cnvSplitNum = 8
        '
        Case Else
            cnvSplitNum = 0
    End Select

End Function

Public Function cnvSplitTNum(ByVal nSplitNum As Integer) As String

    Select Case nSplitNum
        Case 1
            cnvSplitTNum = "1"
        Case 2
            cnvSplitTNum = "2"
        Case 3
            cnvSplitTNum = "3"
        Case 4
            cnvSplitTNum = "4"
        Case 5
            cnvSplitTNum = "5"
        Case 6
            cnvSplitTNum = "6"
        Case 7
            cnvSplitTNum = "7"
        
        Case 8
            cnvSplitTNum = "9"
        Case 9
            cnvSplitTNum = "X"
        
        Case 10
            cnvSplitTNum = "Y"
        
        Case Else
            cnvSplitTNum = 0
    End Select

End Function

Public Function chkImgFile(ByVal strKEY As String, ByVal nSplitNum As Integer) As Boolean
        
    If Dir(App.path & "\" & conDefine_ImageDirName & "\" & strKEY & Format(nSplitNum, "00") & "(0).jpg") <> "" Then
        chkImgFile = True
    Else
        chkImgFile = False
    End If

End Function

Public Sub clrImgFile(ByVal strKEY As String)
    On Error Resume Next
    Call Kill(App.path & "\" & conDefine_ImageDirName & "\" & strKEY & ".jpg")
    On Error GoTo 0
End Sub

'COLORSYS
Public Sub init_APResData()
    Dim initData As typAPResData
    APResData = initData
End Sub

'
Public Function ReadCSV(ByVal strReadFilePath As String, ByRef strItemName() As String, ByRef strDataField() As String) As Boolean
    Dim READ_FileNumber As Variant ''�t�@�C���ԍ�
    Dim strBuf As String
    Dim strChk As String
    Dim pos, org_pos
    Dim nLine As Integer
    Dim nItemNum As Integer
    Dim strItem As String
    
    If Dir(strReadFilePath) = "" Then
        Call MsgLog(strReadFilePath & "��������܂���B", True)
        ReadCSV = False 'NG
        Exit Function
    End If
    
    ReDim strItemName(0) '�N���A�[
    
    READ_FileNumber = Empty
    READ_FileNumber = FreeFile               ' ���g�p�̃t�@�C���ԍ����擾���܂��B
    Open strReadFilePath For Input As #READ_FileNumber

    '�^�C�g���s�ǎ�ƍ��ڐ��A�s������
    nLine = 0
    Do While Not EOF(READ_FileNumber)            ' �t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
        Line Input #READ_FileNumber, strBuf      ' �s��ϐ��ɓǂݍ��݂܂��B
        Debug.Print strBuf         ' �C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
        'Call MsgLog(strBuf)
        
        pos = 1
        org_pos = 1
        nItemNum = 1
        nLine = nLine + 1
    
        Do While True
             pos = InStr(org_pos, strBuf, ",", vbTextCompare)
             If org_pos <> 1 Then
                 If pos = 0 Then
                     If org_pos > Len(strBuf) + 1 Then
                         '�s�̍Ō�
                         Exit Do
                     Else
                         pos = Len(strBuf) + 1
                     End If
                 End If
             Else
                 '��s�Ȃɂ��Ȃ�
                 If pos = 0 Then
                     Exit Do
                 End If
             End If
             
            If nLine = 1 Then
                '�^�C�g���s
                strItem = Mid(strBuf, org_pos, (pos - org_pos))
                strItemName(UBound(strItemName)) = strItem
                ReDim Preserve strItemName(UBound(strItemName) + 1)
            Else
                '�f�[�^�s
            End If
             
            If pos = Len(strBuf) Then
                Exit Do
            Else
                org_pos = pos + 1
                nItemNum = nItemNum + 1
            End If
        Loop
    Loop
    
    If nLine <> 0 Then
        '�f�[�^�s�ǎ�
        ReDim strDataField(UBound(strItemName), nLine - 1)
        Seek #READ_FileNumber, 1
        nLine = 0
        Do While Not EOF(READ_FileNumber)            ' �t�@�C���̏I�[�܂Ń��[�v���J��Ԃ��܂��B
            Line Input #READ_FileNumber, strBuf      ' �s��ϐ��ɓǂݍ��݂܂��B
            Debug.Print strBuf         ' �C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
            'Call MsgLog(strBuf)
            
            pos = 1
            org_pos = 1
            nItemNum = 1
            nLine = nLine + 1
        
            Do While True
                 pos = InStr(org_pos, strBuf, ",", vbTextCompare)
                 If org_pos <> 1 Then
                     If pos = 0 Then
                         If org_pos > Len(strBuf) + 1 Then
                             '�s�̍Ō�
                             Exit Do
                         Else
                             pos = Len(strBuf) + 1
                         End If
                     End If
                 Else
                     '��s�Ȃɂ��Ȃ�
                     If pos = 0 Then
                         Exit Do
                     End If
                 End If
                 
                If nLine = 1 Then
                    '�^�C�g���s
                Else
                    '�f�[�^�s
                    strItem = Mid(strBuf, org_pos, (pos - org_pos))
                    strDataField(nItemNum - 1, nLine - 2) = strItem
                End If
                 
                If pos = Len(strBuf) Then
                    Exit Do
                Else
                    org_pos = pos + 1
                    nItemNum = nItemNum + 1
                End If
            Loop
        Loop
    End If
    
    If IsEmpty(READ_FileNumber) = False Then
        Close #READ_FileNumber
        READ_FileNumber = Empty
        ReadCSV = True 'OK
    Else
        ReadCSV = False 'NG
    End If

End Function

Public Function getItemDataCSV(ByVal strTGTItemName As String, ByVal nTGTDataNumber As Integer, ByRef strItemName() As String, ByRef strDataField() As String) As String
    Dim nI As Integer
    
    For nI = 0 To UBound(strItemName) - 1
        If Trim(UCase(strItemName(nI))) = Trim(UCase(strTGTItemName)) Then
            getItemDataCSV = strDataField(nI, nTGTDataNumber - 1)
            Exit Function
        End If
    Next nI
    
    getItemDataCSV = ""
    
End Function

' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�Ɏl�̌ܓ����܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�Ɏl�̌ܓ����ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Double, ByVal iDigits As Integer) As Double
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Public Sub WaitMsgBox(ByVal callOwnerObj As Object, ByVal strMessage As String)
    Dim MsgWnd As Message
    Set MsgWnd = New Message
        
    Call MsgLog(conProcNum_MAIN, strMessage)
        
    MsgWnd.MsgText = strMessage
    MsgWnd.OK.Visible = True
'        MsgWnd.AutoDelete = True
    Do
        On Error Resume Next
        MsgWnd.Show vbModal, callOwnerObj
        If Err.Number = 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    Set MsgWnd = Nothing

End Sub

' @(f)
'
' �@�\      : �e��摜�o�^�����J�E���g
'
' ������    : ARG1 - MODE(SKIN COLOR FAIL)�@�X���u���E�J���[�E�ُ�
' �@�@�@    �FARG2 - CHNO   �`���[�W�m�n
' �@�@�@    �FARG3 - AINO   ����
' �@�@�@    �FARG4 - STAT   ���
' �@�@�@    �FARG5 - COLOR  �J���[��
'
' �Ԃ�l    : �J�E���g�������𕶎���Ŗ߂�
'
' �@�\����  : �t�H���_�ɓ����Ă���摜�������J�E���g����
'
' ���l      :
'
Public Function PhotoImgCount(ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStat As String, ByVal sColor As String) As String
    Dim objFso       As Object
    Dim iCnt         As Integer
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    
    On Error GoTo PhotoImgCount_Err
    
    ' �摜�t�H���_�p�X�ݒ�
    sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino

    ' �w��t�H���_�T��
    iCnt = 0
    sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & "_??.JPG"
    sGetFileName = Dir(sPhotoPath & "\" & sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = "jpg" Or Right(sGetFileName, 3) = "JPG" Then
            iCnt = iCnt + 1
        End If

        sGetFileName = Dir()
    Loop
    Debug.Print sChkFileName
    Debug.Print iCnt
    PhotoImgCount = CStr(iCnt)
    
    Exit Function
    
PhotoImgCount_Err:
    PhotoImgCount = CStr(iCnt)
    On Error Resume Next
    
End Function


' @(f)
'
' �@�\      : �A�b�v���[�h�p�t�H���_�m�F
'
' ������    : ARG1 - EXT    IMG�EPDF�ESCAN
' �@�@�@    �FARG2 - MODE(SKIN COLOR FAIL)�@�X���u���E�J���[�E�ُ�
' �@�@�@    �FARG3 - CHNO   �`���[�W�m�n
' �@�@�@    �FARG4 - AINO   ����
' �@�@�@    �FARG5 - STAT   ���
' �@�@�@    �FARG6 - COLOR  �J���[��
'
' �Ԃ�l    : True �f�[�^���^False �f�[�^�L
'
' �@�\����  : �ύX��w��ԁx���g�p�ł��邩�m�F����
'
' ���l      :
'
Public Function StatChgFoldCheck(ByVal sExtDir As String, ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStat As String, ByVal sColor As String) As Boolean
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    Dim sExt         As String
    Dim errNum       As Long
    Dim errDesc      As String
    Dim errSrc       As String
    Dim StrTmp       As String
    
    On Error GoTo StatChgFoldCheck_Err
    
    ' �摜�t�H���_�p�X�ݒ�
    If sExtDir = "IMG" Then
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & "_??.JPG"
    ElseIf sExtDir = "PDF" Then
        sExt = "PDF"
        sPhotoPath = APSysCfgData.SHARES_PDFDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & ".PDF"
    Else
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_SCNDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sMode & "_" & sChno & "_" & sAino & "_" & sStat & "_" & sColor & ".JPG"
    End If

    StatChgFoldCheck = True

    ' �w��t�H���_�T��
    On Error Resume Next
    sGetFileName = Dir(sPhotoPath & "\" & sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = sExt Then
            StatChgFoldCheck = False
            Exit Do
        End If

        sGetFileName = Dir()
    Loop
    Debug.Print sChkFileName
    
    Exit Function
    
StatChgFoldCheck_Err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "StatChgFoldCheck �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    StatChgFoldCheck = False

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : �A�b�v���[�h�p�t�H���_�E�t�@�C���ύX
'
' ������    : ARG1 - MODE       SKIN�X���u���ECOLOR�J���[�EFAIL�ُ�
' �@�@�@    �FARG2 - CHNO       �`���[�W�m�n
' �@�@�@    �FARG3 - AINO       ����
' �@�@�@    �FARG4 - STATOLD    �ύX�O���
' �@�@�@    �FARG4 - STATNEW    �ύX����
' �@�@�@    �FARG5 - COLOROLD      �J���[��
' �@�@�@    �FARG5 - COLORNEW   �J���[��
' �@�@�@    �FARG6 - EXT        IMG�EPDF�ESCAN
'
' �Ԃ�l    : True �f�[�^���^False �f�[�^�L
'
' �@�\����  : �ύX��w��ԁx���g�p�ł��邩�m�F����
'
' ���l      :
'
Public Function StatChgFoldFix(ByVal sExtDir As String, ByVal sMode As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatOld As String, ByVal sStatNew As String, ByVal sColorOld As String, ByVal sColorNew As String) As Boolean
    Dim sPhotoPath   As String
    Dim sGetFileName As String
    Dim sChkFileName As String
    Dim sNewFileName As String    ' �ύX�t�@�C����
    Dim sExt         As String
    Dim sKeepNo      As String
    Dim errNum       As Long
    Dim errDesc      As String
    Dim errSrc       As String
    Dim StrTmp       As String
    
    On Error GoTo StatChgFoldFix_Err
    
    ' �����摜�t�H���_�p�X�ݒ� ********************************************************************
    If sExtDir = "IMG" Then
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_IMGDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & "_??.JPG"
    ElseIf sExtDir = "PDF" Then
        sExt = "PDF"
        sPhotoPath = APSysCfgData.SHARES_PDFDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & ".PDF"
    Else
        sExt = "JPG"
        sPhotoPath = APSysCfgData.SHARES_SCNDIR & "\" & sMode & "\" & sChno & "\" & sAino
        sChkFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatOld & "_" & sColorOld & ".JPG"
    End If

    StatChgFoldFix = True

    ' �w��t�H���_�T�� ****************************************************************************
    sGetFileName = Dir(sChkFileName)
    Do Until sGetFileName = vbNullString
        If Right(sGetFileName, 3) = sExt Then
            ' �ύX���ݒ�
            If sExtDir = "IMG" Then
                sKeepNo = Left(Right(sGetFileName, 6), 2)
                sNewFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColorNew & "_" & sKeepNo & "." & sExt
            Else
                sNewFileName = sPhotoPath & "\" & sMode & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColorNew & "." & sExt
            End If
            Debug.Print "NAME " & sPhotoPath & "\" & sGetFileName & " AS " & sNewFileName
            Call MsgLog(conProcNum_MAIN, "��ԕύX[" & "NAME " & sPhotoPath & "\" & sGetFileName & " AS " & sNewFileName & "]") '�K�C�_���X�\��
            
            ' �t�@�C�����ύX���s
            Name sPhotoPath & "\" & sGetFileName As sNewFileName
        End If

        sGetFileName = Dir()
    Loop
    
    Exit Function
    
StatChgFoldFix_Err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "StatChgFoldFix �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    StatChgFoldFix = False

    On Error GoTo 0
    
End Function

