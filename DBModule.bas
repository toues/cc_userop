Attribute VB_Name = "DBModule"
' @(h) DBModule.Bas                ver 1.00 ( '02.01.10 SEC Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�f�[�^�x�[�X���W���[��
' �@�{���W���[���̓f�[�^�x�[�X�A�N�Z�X��
' �@���߂̂��̂ł���B

'// ODBC Driver Oracle 8.01.66.00
Option Explicit

Const ORAPARM_INPUT As Integer = 1  '���͕ϐ�
Const ORAPARM_OUTPUT As Integer = 2 '�o�͕ϐ�
Const ORAPARM_BOTH As Integer = 3   '���͕ϐ��Əo�͕ϐ��̗���

Const ORATYPE_VARCHAR2 As Integer = 1
Const ORATYPE_NUMBER As Integer = 2
Const ORATYPE_SINT As Integer = 3
Const ORATYPE_FLOAT As Integer = 4
Const ORATYPE_STRING As Integer = 5
Const ORATYPE_DECIMAL As Integer = 7
Const ORATYPE_VARCHAR As Integer = 9
Const ORATYPE_DATE As Integer = 12
Const ORATYPE_REAL As Integer = 21
Const ORATYPE_DOUBLE As Integer = 22
Const ORATYPE_UNSIGNED8 As Integer = 23
Const ORATYPE_UNSIGNED16 As Integer = 25
Const ORATYPE_UNSIGNED32 As Integer = 26
Const ORATYPE_SIGNED8 As Integer = 27
Const ORATYPE_SIGNED16 As Integer = 28
Const ORATYPE_SIGNED32 As Integer = 29
Const ORATYPE_PTR As Integer = 32
Const ORATYPE_OPAQUE As Integer = 58
Const ORATYPE_UINT As Integer = 68
Const ORATYPE_RAW As Integer = 95
Const ORATYPE_CHAR As Integer = 96
Const ORATYPE_CHARZ As Integer = 97
Const ORATYPE_CURSOR As Integer = 102
Const ORATYPE_ROWID As Integer = 104
Const ORATYPE_MLSLABEL As Integer = 105
Const ORATYPE_OBJECT As Integer = 108
Const ORATYPE_REF As Integer = 110
Const ORATYPE_CLOB As Integer = 112
Const ORATYPE_BLOB As Integer = 113
Const ORATYPE_BFILE As Integer = 114
Const ORATYPE_CFILE As Integer = 115
Const ORATYPE_RSLT As Integer = 116
Const ORATYPE_NAMEDCOLLECTION As Integer = 122
Const ORATYPE_COLL As Integer = 122
Const ORATYPE_SYSFIRST As Integer = 228
Const ORATYPE_SYSLAST As Integer = 235
Const ORATYPE_OCTET As Integer = 245
Const ORATYPE_SMALLINT As Integer = 246
Const ORATYPE_VARRAY As Integer = 247
Const ORATYPE_TABLE As Integer = 248
Const ORATYPE_OTMLAST As Integer = 320

Private Const conDef_DB_ProcessName As String = "�J���[�`�F�b�N���тo�b" ''�f�[�^�x�[�X�g�p�v���Z�X��

'Private Const conDef_DB_CHUNK_SIZE As Long = 16384  '�`�����N�T�C�Y

' @(f)
'
' �@�\      : �n�c�a�b�ڑ�������擾
'
' ������    : ARG1 - �ڑ��؂�ւ��t���O
'
' �Ԃ�l    : �ڑ�������
'
' �@�\����  : �n�c�a�b�ڑ���������擾����B
'
' ���l      :COLORSYS
'
Public Function DBConnectStr(ByVal sw As Integer, ByRef host As String, ByRef id As String, ByRef pass As String) As Boolean
    Select Case sw
    
    Case conDefault_DBConnect_MYUSER
        host = APSysCfgData.DB_MYUSER_DSN
        pass = APSysCfgData.DB_MYUSER_UID
        id = APSysCfgData.DB_MYUSER_PWD
        DBConnectStr = True
    
    Case conDefault_DBConnect_MYCOMN
        host = APSysCfgData.DB_MYCOMN_DSN
        pass = APSysCfgData.DB_MYCOMN_UID
        id = APSysCfgData.DB_MYCOMN_PWD
        DBConnectStr = True
    
    Case conDefault_DBConnect_SOZAI
        host = APSysCfgData.DB_SOZAI_DSN
        pass = APSysCfgData.DB_SOZAI_UID
        id = APSysCfgData.DB_SOZAI_PWD
        DBConnectStr = True
    
    Case Else
        DBConnectStr = False
    End Select

End Function

' @(f)
'
' �@�\      : �r�p�k���s����
'
' ������    : ARG1 - �ڑ��؂�ւ��t���O
'             ARG2 - �r�p�k������
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̐ڑ����g�p���Ăr�p�k����������s����
'
' ���l      :
'
Public Function DB_SQL_Execute(ByVal nConnectSw As Integer, ByVal strSQL As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        DB_SQL_Execute = True
        Exit Function
    End If
    
    On Error GoTo DB_SQL_Execute_err
    
    nOpen = 0
    
    ' Oracle�Ƃ̐ڑ����m������
    '-cn.Open DBConnectStr(nConnectSw)
    bRet = DBConnectStr(nConnectSw, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0
    
    Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute ����I��") '�K�C�_���X�\��

    DB_SQL_Execute = True

    On Error GoTo 0
    Exit Function

DB_SQL_Execute_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DB_SQL_Execute �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If
    
    DB_SQL_Execute = False
    
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : �X���u����񌟍�����
'
' ������    : ARG1 - �����I�v�V�����ԍ��i���g�p�j
'             ARG2 - �ő匟������
'             ARG3 - �����X���u�m���D
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      :
'
Public Function DBSkinSlbSearchRead(ByVal nSearchOption As Integer, ByVal nSEARCH_MAX As Integer, ByVal nSERCH_RANGE As Integer, ByVal strSearchSlbNumber As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    Dim lRecCnt As Long
    
    Dim strSERCH_RANGE As String
    
    If nSERCH_RANGE = 9999 Then
        strSERCH_RANGE = ""
    Else
        strSERCH_RANGE = Format(DateAdd("d", -nSERCH_RANGE, Now), "YYYYMMDD")
    End If
    
    ''�c�a�I�t���C���ŋ������͂��s�������Ƃ𔻒f����t���O
'    bAPInputOffline = False
'    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
        
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
'******************************
        'DEMO
        ReDim APSearchTmpSlbData(0)
        
        Call DBSkinSlbSearchReadCSV
        
        Select Case nSearchOption
          Case 0 '�X���u�m���D����
'            For nI = 0 To 19
'                '�X���u�m���D
'                APSearchTmpSlbData(nI).slb_chno = CStr(nI + 10000)
'                APSearchTmpSlbData(nI).slb_aino = CStr(nI + 1000)
'                APSearchTmpSlbData(nI).slb_no = APSearchTmpSlbData(nI).slb_chno & APSearchTmpSlbData(nI).slb_aino
'
'                '���
'                APSearchTmpSlbData(nI).slb_stat = nI Mod 6
'
'                '�|��
'                APSearchTmpSlbData(nI).slb_ksh = "AAAAAA"
'
'                '�^
'                APSearchTmpSlbData(nI).slb_typ = "AAA"
'
'                '����
'                APSearchTmpSlbData(nI).slb_uksk = "AAA"
'
'                '�����
'                APSearchTmpSlbData(nI).slb_zkai_dte = "20080310"
'
'                '���ޔ����сi����L�^���j
'                APSearchTmpSlbData(nI).sys_wrt_dte = "20080310"
'
'                '���ޔ��Ұ��
'                APSearchTmpSlbData(nI).bAPScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                '���ޔ�PDF
'                APSearchTmpSlbData(nI).bAPPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'                ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
'            Next nI
        End Select
'******************************
            
        DBSkinSlbSearchRead = True
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBSkinSlbSearchRead_err
    
    Select Case nSearchOption
    Case 0 '�X���u�m���D����
        strSQL = "SELECT TRTS0012.*,TRTS0040.SLB_PDF_ADDR,TRTS0040.SYS_WRT_DTE AS SYS_WRT_DTE40,TRTS0050.SLB_SCAN_ADDR "
        
        strSQL = strSQL & "FROM (TRTS0012 LEFT JOIN TRTS0040 ON (TRTS0012.SLB_STAT = TRTS0040.SLB_STAT) "
        strSQL = strSQL & "AND (TRTS0012.SLB_NO = TRTS0040.SLB_NO)) "
        strSQL = strSQL & "LEFT JOIN TRTS0050 ON (TRTS0012.SLB_STAT = TRTS0050.SLB_STAT) "
        strSQL = strSQL & "AND (TRTS0012.SLB_NO = TRTS0050.SLB_NO) "
        
        strSQL = strSQL & "WHERE TRTS0012.SLB_NO LIKE '" & strSearchSlbNumber & "' "

        If strSERCH_RANGE <> "" Then
            strSQL = strSQL & _
            "And (TRTS0012.SYS_WRT_DTE >= '" & strSERCH_RANGE & "') "
        End If

        strSQL = strSQL & "ORDER BY TRTS0012.SYS_WRT_DTE DESC ,TRTS0012.SYS_WRT_TME DESC"
    End Select

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    nI = 0
    ReDim APSearchTmpSlbData(nI)
    Do While Not oDS.EOF
        
        APSearchTmpSlbData(nI).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value) ''�X���u�m���D
        APSearchTmpSlbData(nI).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value) ''�X���u�`���[�W�m���D
        APSearchTmpSlbData(nI).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) ''�X���u����
        APSearchTmpSlbData(nI).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value) ''���
        APSearchTmpSlbData(nI).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), "", oDS.Fields("slb_zkai_dte").Value) ''�����
        APSearchTmpSlbData(nI).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), "", oDS.Fields("slb_ksh").Value) ''�|��
        APSearchTmpSlbData(nI).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), "", oDS.Fields("slb_typ").Value) ''�^
        APSearchTmpSlbData(nI).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), "", oDS.Fields("slb_uksk").Value) ''����
        APSearchTmpSlbData(nI).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value) ''�L�^���i����L�^���j
        
        APSearchTmpSlbData(nI).sAPPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE40").Value), "", oDS.Fields("SYS_WRT_DTE40").Value) ''PDF�C���[�W�f�[�^�L�^���i����L�^���j
        
        If IsNull(oDS.Fields("SLB_SCAN_ADDR").Value) = False Then
            APSearchTmpSlbData(nI).bAPScanInput = True ''SCAN�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False ''SCAN�f�[�^�L��t���O
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR").Value) = False Then
            APSearchTmpSlbData(nI).bAPPdfInput = True ''PDF�C���[�W�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False ''PDF�C���[�W�f�[�^�L��t���O
        End If
        
        ' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
        APSearchTmpSlbData(nI).PhotoImgCnt1 = PhotoImgCount("SKIN", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, "00")
        
        ReDim Preserve APSearchTmpSlbData(nI + 1) '�X���u�I����ʌ������X�g
    
        oDS.MoveNext

        nI = nI + 1 '�i�[�p�C���f�b�N�X

        '�ݒ�O�̏ꍇ���������ƂȂ�B
        If nSEARCH_MAX = nI Then
            Exit Do
        '�ő僊�~�b�^�[conDefault_nSEARCH_MAX0 = 9999
        ElseIf nI > conDefault_nSEARCH_MAX0 Then
            Exit Do
        End If
    Loop

    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    DBSkinSlbSearchRead = True

    Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead ����I��") '�K�C�_���X�\��

    On Error GoTo 0

    Exit Function

DBSkinSlbSearchRead_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBSkinSlbSearchRead �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBSkinSlbSearchRead = False

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : �J���[�`�F�b�N��񌟍�����
'
' ������    : ARG1 - �����I�v�V�����ԍ��i0:�ʏ팟��,1:�ُ�񍐈ꗗ����)
'             ARG2 - �ő匟������
'             ARG3 - �����X���u�m���D
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      : 2008/09/03 �ُ�񍐈ꗗ�����ǉ�
'
Public Function DBColorSlbSearchRead(ByVal nSearchOption As Integer, ByVal nSEARCH_MAX As Integer, ByVal nSERCH_RANGE As Integer, ByVal strSearchSlbNumber As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    Dim lRecCnt As Long
    Dim sImageCnt As String     ' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
    
    Dim strRes_Wrt_Dte_Max As String
    Dim strNotCmp_Res_No_MIN As String
    
    Dim strSERCH_RANGE As String
    
    If nSERCH_RANGE = 9999 Then
        strSERCH_RANGE = ""
    Else
        strSERCH_RANGE = Format(DateAdd("d", -nSERCH_RANGE, Now), "YYYYMMDD")
    End If
    
    ''�c�a�I�t���C���ŋ������͂��s�������Ƃ𔻒f����t���O
'    bAPInputOffline = False
'    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
        
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
'******************************
        'DEMO
        ReDim APSearchTmpSlbData(0)
        
        Call DBColorSlbSearchReadCSV
        
        Select Case nSearchOption
          Case 0 '�X���u�m���D����
'            For nI = 0 To 19
'                '�X���u�m���D
'                APSearchTmpSlbData(nI).slb_chno = CStr(nI + 10000)
'                APSearchTmpSlbData(nI).slb_aino = CStr(nI + 1000)
'                APSearchTmpSlbData(nI).slb_no = APSearchTmpSlbData(nI).slb_chno & APSearchTmpSlbData(nI).slb_aino
'
'                '���
'                APSearchTmpSlbData(nI).slb_stat = nI Mod 6
'
'                '���װ��
'                APSearchTmpSlbData(nI).slb_col_cnt = "01"
'
'                '�|��
'                APSearchTmpSlbData(nI).slb_ksh = "AAAAAA"
'
'                '�^
'                APSearchTmpSlbData(nI).slb_typ = "AAA"
'
'                '����
'                APSearchTmpSlbData(nI).slb_uksk = "AAA"
'
'                '�����
'                APSearchTmpSlbData(nI).slb_zkai_dte = "20080310"
'
'                '�װ���сi����L�^���j
'                APSearchTmpSlbData(nI).sys_wrt_dte = "20080310"
'
'                '���r�W�R�����M����
'                APSearchTmpSlbData(nI).host_send = ""
'
'                '�װ�Ұ��
'                APSearchTmpSlbData(nI).bAPScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                '�װPDF
'                APSearchTmpSlbData(nI).bAPPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'***********************************************************************
'                '�ُ�񍐁i����L�^���j
'                APSearchTmpSlbData(nI).fail_sys_wrt_dte = "20080310"
'
'                '�ُ�񍐃r�W�R�����M����
'                APSearchTmpSlbData(nI).fail_host_send = ""
'
'                '�ُ�Ұ��
'                APSearchTmpSlbData(nI).bAPFailScanInput = IIf(nI Mod 2 = 0, False, True)
'
'                '�ُ�PDF
'                APSearchTmpSlbData(nI).bAPFailPdfInput = IIf(nI Mod 2 = 0, True, False)
'
'***********************************************************************
'                'CCNO
'                APSearchTmpSlbData(nI).slb_ccno = "10000"
'
'                '�d�ʁi�װ�����p�FSEG�o���d�ʁj
'                APSearchTmpSlbData(nFirstDataIndex).slb_color_wei
'
'                '����
'                APSearchTmpSlbData(nFirstDataIndex).slb_lngth
'
'                '��
'                APSearchTmpSlbData(nFirstDataIndex).slb_wdth
'
'                '����
'                APSearchTmpSlbData(nFirstDataIndex).slb_thkns
'
'***********************************************************************
'                '���u�w��
'                APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = "20080310"
'
'***********************************************************************
'                '���u����
'                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = "20080310"
'
'                '���u���ʊ����t���O
'                APSearchTmpSlbData(nI).fail_res_cmp_flg = "1"
'
'***********************************************************************
'
'                ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
'            Next nI
        End Select
'******************************
            
        DBColorSlbSearchRead = True
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBColorSlbSearchRead_err
    
    '******************************************************
    'SQL��������
    strSQL = ""
    
    'SQL���O����
    Select Case nSearchOption
        Case 0 '�X���u�m���D����
            '����
        Case 1 '�ُ�񍐈ꗗ
            strSQL = strSQL & "SELECT * FROM ("
    End Select
    
    '******************************************************
    '�X���u�m���D����
    strSQL = strSQL & "SELECT TRTS0014.*, "
    strSQL = strSQL & "TRTS0016.HOST_SEND AS HOST_SEND16, "
    strSQL = strSQL & "TRTS0016.HOST_WRT_DTE AS HOST_WRT_DTE16, "
    strSQL = strSQL & "TRTS0016.HOST_WRT_TME AS HOST_WRT_TME16, "
    strSQL = strSQL & "TRTS0016.SYS_WRT_DTE AS SYS_WRT_DTE16, "
    strSQL = strSQL & "TRTS0016.SYS_WRT_TME AS SYS_WRT_TME16, "
    strSQL = strSQL & "TRTS0042.SLB_PDF_ADDR AS SLB_PDF_ADDR42, "
    strSQL = strSQL & "TRTS0042.SYS_WRT_DTE AS SYS_WRT_DTE42, "
    strSQL = strSQL & "TRTS0044.SLB_PDF_ADDR AS SLB_PDF_ADDR44, "
    strSQL = strSQL & "TRTS0044.SYS_WRT_DTE AS SYS_WRT_DTE44, "
    strSQL = strSQL & "TRTS0052.SLB_SCAN_ADDR AS SLB_SCAN_ADDR52, "
    strSQL = strSQL & "TRTS0054.SLB_SCAN_ADDR AS SLB_SCAN_ADDR54, "
    strSQL = strSQL & "T20.DIR_WRT_DTE_MAX, "
    strSQL = strSQL & "T20.DIR_PRN_OUT_MAX, " '2008/09/04
    strSQL = strSQL & "T22A.NOTCMP_RES_NO_MIN, "
    strSQL = strSQL & "T22B.RES_WRT_DTE_MAX, "
    strSQL = strSQL & "T22C.HOST_SEND AS HOST_SEND22, "
    strSQL = strSQL & "T22C.HOST_WRT_DTE AS HOST_WRT_DTE22, "
    strSQL = strSQL & "T22C.HOST_WRT_TME AS HOST_WRT_TME22 "
    strSQL = strSQL & "FROM ((((((((TRTS0014 "
    
    strSQL = strSQL & "LEFT JOIN TRTS0016 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0016.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0016.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0016.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0042 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0042.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0042.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0042.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0044 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0044.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0044.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0044.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0052 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0052.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0052.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0052.SLB_COL_CNT)) "
    
    strSQL = strSQL & "LEFT JOIN TRTS0054 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = TRTS0054.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = TRTS0054.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = TRTS0054.SLB_COL_CNT)) "
    
    'TRTS0020
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Max(TRTS0020.DIR_NO), TRTS0020.SLB_NO, TRTS0020.SLB_STAT, "
    strSQL = strSQL & "TRTS0020.SLB_COL_CNT, Max(TRTS0020.DIR_PRN_OUT) AS DIR_PRN_OUT_MAX, Max(TRTS0020.DIR_WRT_DTE) AS DIR_WRT_DTE_MAX FROM TRTS0020 "
    strSQL = strSQL & "GROUP BY TRTS0020.SLB_NO, TRTS0020.SLB_STAT, TRTS0020.SLB_COL_CNT "
    strSQL = strSQL & "ORDER BY Max(TRTS0020.DIR_WRT_DTE)) T20 "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T20.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T20.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T20.SLB_COL_CNT)) "
    
    'TRTS0022 A �������̑��ݒ���
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Min(RES_NO) AS NOTCMP_RES_NO_MIN, SLB_NO, SLB_STAT, "
    strSQL = strSQL & "SLB_COL_CNT FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY RES_CMP_FLG, SLB_NO, SLB_STAT, SLB_COL_CNT HAVING (RES_CMP_FLG Is Null)) T22A "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22A.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22A.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22A.SLB_COL_CNT)) "
    
    'TRTS0022 B �����̑��ݒ���
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Max(TRTS0022.RES_NO), TRTS0022.SLB_NO, TRTS0022.SLB_STAT, "
    strSQL = strSQL & "TRTS0022.SLB_COL_CNT, Max(TRTS0022.RES_WRT_DTE) AS RES_WRT_DTE_MAX FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, TRTS0022.RES_CMP_FLG "
    strSQL = strSQL & "HAVING (((TRTS0022.RES_CMP_FLG)='1')) "
    strSQL = strSQL & "ORDER BY Max(TRTS0022.RES_WRT_DTE)) T22B "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22B.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22B.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22B.SLB_COL_CNT)) "
    
    'TRTS0022 C �r�W�R�����M�֌W�擾
    strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT Min(TRTS0022.RES_NO), TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, "
    strSQL = strSQL & "TRTS0022.HOST_SEND, TRTS0022.HOST_WRT_DTE, TRTS0022.HOST_WRT_TME "
    strSQL = strSQL & "FROM TRTS0022 "
    strSQL = strSQL & "GROUP BY TRTS0022.SLB_NO, TRTS0022.SLB_STAT, TRTS0022.SLB_COL_CNT, "
    strSQL = strSQL & "TRTS0022.HOST_SEND, TRTS0022.HOST_WRT_DTE, TRTS0022.HOST_WRT_TME) T22C "
    strSQL = strSQL & "ON (TRTS0014.SLB_NO = T22C.SLB_NO) "
    strSQL = strSQL & "AND (TRTS0014.SLB_STAT = T22C.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0014.SLB_COL_CNT = T22C.SLB_COL_CNT) "
    
    strSQL = strSQL & "WHERE TRTS0014.SLB_NO LIKE '" & strSearchSlbNumber & "' "

    If strSERCH_RANGE <> "" Then
        strSQL = strSQL & _
        "And (TRTS0014.SYS_WRT_DTE >= '" & strSERCH_RANGE & "') "
    End If

    strSQL = strSQL & "ORDER BY TRTS0014.SYS_WRT_DTE DESC ,TRTS0014.SYS_WRT_TME DESC"
    '******************************************************
    'SQL���㏈��
    Select Case nSearchOption
        Case 0 '�X���u�m���D����
            '����
        Case 1 '�ُ�񍐈ꗗ:    (�ُ�񍐗L��) AND ((�����M=2) OR (����) OR (���L AND �����L))
            strSQL = strSQL & ") WHERE (SYS_WRT_DTE16 Is Not Null) AND ((HOST_SEND22 = '2') OR (RES_WRT_DTE_MAX Is Null) OR (RES_WRT_DTE_MAX Is Not Null AND NOTCMP_RES_NO_MIN Is Not Null))"
    End Select

    '******************************************************

    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    nI = 0
    ReDim APSearchTmpSlbData(nI)
    Do While Not oDS.EOF
        
        APSearchTmpSlbData(nI).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value) ''�X���u�m���D
        APSearchTmpSlbData(nI).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value) ''���
        APSearchTmpSlbData(nI).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), "", oDS.Fields("slb_col_cnt").Value) ''�װ��
        APSearchTmpSlbData(nI).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value) ''�X���u�`���[�W�m���D
        APSearchTmpSlbData(nI).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) ''�X���u����
        
        APSearchTmpSlbData(nI).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), "", oDS.Fields("slb_ccno").Value) ''CCNO
        
        APSearchTmpSlbData(nI).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), "", oDS.Fields("slb_zkai_dte").Value) ''�����
        APSearchTmpSlbData(nI).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), "", oDS.Fields("slb_ksh").Value) ''�|��
        APSearchTmpSlbData(nI).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), "", oDS.Fields("slb_typ").Value) ''�^
        APSearchTmpSlbData(nI).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), "", oDS.Fields("slb_uksk").Value) ''����
        
        APSearchTmpSlbData(nI).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), "", oDS.Fields("slb_wei").Value) ''�d��
        APSearchTmpSlbData(nI).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), "", oDS.Fields("slb_lngth").Value) ''����
        APSearchTmpSlbData(nI).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), "", oDS.Fields("slb_wdth").Value) ''��
        APSearchTmpSlbData(nI).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), "", oDS.Fields("slb_thkns").Value) ''����
        
        APSearchTmpSlbData(nI).host_send = IIf(IsNull(oDS.Fields("host_send").Value), "", oDS.Fields("host_send").Value) ''�r�W�R�����M����
        APSearchTmpSlbData(nI).host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte").Value), "", oDS.Fields("host_wrt_dte").Value) ''�r�W�R�����M���i����r�W�R�����M���j
        APSearchTmpSlbData(nI).host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme").Value), "", oDS.Fields("host_wrt_tme").Value) ''�r�W�R�����M�����i����r�W�R�����M�����j
        APSearchTmpSlbData(nI).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value) ''�L�^���i����L�^���j
        APSearchTmpSlbData(nI).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), "", oDS.Fields("sys_wrt_tme").Value) ''�L�^�����i����L�^�����j
        
        '�ُ�ꗗ���X�g�\����p '2008/09/04
        APSearchTmpSlbData(nI).slb_fault_e_judg = IIf(IsNull(oDS.Fields("slb_fault_e_judg").Value), "", oDS.Fields("slb_fault_e_judg").Value)  ''����E�ʔ���
        APSearchTmpSlbData(nI).slb_fault_w_judg = IIf(IsNull(oDS.Fields("slb_fault_w_judg").Value), "", oDS.Fields("slb_fault_w_judg").Value)  ''����W�ʔ���
        APSearchTmpSlbData(nI).slb_fault_s_judg = IIf(IsNull(oDS.Fields("slb_fault_s_judg").Value), "", oDS.Fields("slb_fault_s_judg").Value)  ''����S�ʔ���
        APSearchTmpSlbData(nI).slb_fault_n_judg = IIf(IsNull(oDS.Fields("slb_fault_n_judg").Value), "", oDS.Fields("slb_fault_n_judg").Value)  ''����N�ʔ���
        
        '******************
        '�X���u�ُ�
        APSearchTmpSlbData(nI).fail_host_send = IIf(IsNull(oDS.Fields("host_send16").Value), "", oDS.Fields("host_send16").Value) ''�r�W�R�����M����
        APSearchTmpSlbData(nI).fail_host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte16").Value), "", oDS.Fields("host_wrt_dte16").Value) ''�r�W�R�����M���i����r�W�R�����M���j
        APSearchTmpSlbData(nI).fail_host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme16").Value), "", oDS.Fields("host_wrt_tme16").Value) ''�r�W�R�����M�����i����r�W�R�����M�����j
        APSearchTmpSlbData(nI).fail_sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte16").Value), "", oDS.Fields("sys_wrt_dte16").Value) ''�L�^���i����L�^���j
        APSearchTmpSlbData(nI).fail_sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme16").Value), "", oDS.Fields("sys_wrt_tme16").Value) ''�L�^�����i����L�^�����j
        '******************
        
        APSearchTmpSlbData(nI).sAPPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE42").Value), "", oDS.Fields("SYS_WRT_DTE42").Value) ''PDF�C���[�W�f�[�^�L�^���i����L�^���j
        APSearchTmpSlbData(nI).sAPFailPdfInput_ReqDate = IIf(IsNull(oDS.Fields("SYS_WRT_DTE44").Value), "", oDS.Fields("SYS_WRT_DTE44").Value) ''PDF�C���[�W�f�[�^�L�^���i����L�^���j
        
        If IsNull(oDS.Fields("SLB_SCAN_ADDR52").Value) = False Then
            APSearchTmpSlbData(nI).bAPScanInput = True ''SCAN�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False ''SCAN�f�[�^�L��t���O
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR42").Value) = False Then
            APSearchTmpSlbData(nI).bAPPdfInput = True ''PDF�C���[�W�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False ''PDF�C���[�W�f�[�^�L��t���O
        End If
            
        '******************
        '�X���u�ُ�
        If IsNull(oDS.Fields("SLB_SCAN_ADDR54").Value) = False Then
            APSearchTmpSlbData(nI).bAPFailScanInput = True ''SCAN�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPFailScanInput = False ''SCAN�f�[�^�L��t���O
        End If
        
        If IsNull(oDS.Fields("SLB_PDF_ADDR44").Value) = False Then
            APSearchTmpSlbData(nI).bAPFailPdfInput = True ''PDF�C���[�W�f�[�^�L��t���O
        Else
            APSearchTmpSlbData(nI).bAPFailPdfInput = False ''PDF�C���[�W�f�[�^�L��t���O
        End If
        '******************
            
        '���u�w��
        APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = IIf(IsNull(oDS.Fields("dir_wrt_dte_max").Value), "", oDS.Fields("dir_wrt_dte_max").Value)
        '2008/09/04 ����ς݃t���O
        APSearchTmpSlbData(nI).fail_dir_prn_out_max = IIf(IsNull(oDS.Fields("DIR_PRN_OUT_MAX").Value), "", oDS.Fields("DIR_PRN_OUT_MAX").Value)
            
        '���u����
        strRes_Wrt_Dte_Max = IIf(IsNull(oDS.Fields("res_wrt_dte_max").Value), "", oDS.Fields("res_wrt_dte_max").Value)
        strNotCmp_Res_No_MIN = IIf(IsNull(oDS.Fields("notcmp_res_no_min").Value), "", oDS.Fields("notcmp_res_no_min").Value)
        
        '2016/04/20 - TAI - S
        '��Ə�
        APSearchTmpSlbData(nI).slb_works_sky_tok = IIf(IsNull(oDS.Fields("slb_works_sky_tok").Value), "", oDS.Fields("slb_works_sky_tok").Value) ''��Ə�
        '2016/04/20 - TAI - E
        
        APSearchTmpSlbData(nI).fail_res_host_send = IIf(IsNull(oDS.Fields("host_send22").Value), "", oDS.Fields("host_send22").Value) ''�r�W�R�����M����
        APSearchTmpSlbData(nI).fail_res_host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte22").Value), "", oDS.Fields("host_wrt_dte22").Value) ''�r�W�R�����M���i����r�W�R�����M���j
        APSearchTmpSlbData(nI).fail_res_host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme22").Value), "", oDS.Fields("host_wrt_tme22").Value) ''�r�W�R�����M�����i����r�W�R�����M�����j
        
        APSearchTmpSlbData(nI).fail_res_cmp_flg = "0"
        APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
        
        If strRes_Wrt_Dte_Max <> "" Then
            '�o�^���t�L��
            If strNotCmp_Res_No_MIN <> "" Then
                '���������R�[�h�L��
                APSearchTmpSlbData(nI).fail_res_cmp_flg = "0"
                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
            Else
                '�S�Ċ���
                APSearchTmpSlbData(nI).fail_res_cmp_flg = "1"
                APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = strRes_Wrt_Dte_Max
            End If
        Else
            '�o�^����
            APSearchTmpSlbData(nI).fail_res_cmp_flg = ""
            APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = ""
        End If
        
        ReDim Preserve APSearchTmpSlbData(nI + 1) '�X���u�I����ʌ������X�g
    
        ' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
'        If nSearchOption = 0 Then
            sImageCnt = PhotoImgCount("COLOR", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
            APSearchTmpSlbData(nI).PhotoImgCnt1 = sImageCnt
            sImageCnt = PhotoImgCount("SLBFAIL", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
            APSearchTmpSlbData(nI).PhotoImgCnt2 = sImageCnt
'        Else
'            sImageCnt = PhotoImgCount("SLBFAIL", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, oDS.Fields("slb_col_cnt").Value)
'            APSearchTmpSlbData(nI).PhotoImgCnt1 = sImageCnt
'        End If
        
        oDS.MoveNext

        nI = nI + 1 '�i�[�p�C���f�b�N�X

        '�ݒ�O�̏ꍇ���������ƂȂ�B
        If nSEARCH_MAX = nI Then
            Exit Do
        '�ő僊�~�b�^�[conDefault_nSEARCH_MAX0 = 9999
        ElseIf nI > conDefault_nSEARCH_MAX0 Then
            Exit Do
        End If
    Loop

    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    DBColorSlbSearchRead = True

    Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead ����I��") '�K�C�_���X�\��

    On Error GoTo 0

    Exit Function

DBColorSlbSearchRead_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBColorSlbSearchRead �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBColorSlbSearchRead = False

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : TRTS0012�Ǎ�����
'
' ������    : ARG1 - �X���u�ԍ�
'           : ARG2 - ���
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p����TRTS0012�̃��R�[�h��Ǎ�
'
' ���l      : �X���u�����ѓ��̓f�[�^�Ǎ�
'           :COLORSYS
'
Public Function TRTS0012_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String) As Boolean
'slb_chno          VARCHAR2(5)          /* �X���u�`���[�WNO */
'slb_aino          VARCHAR2(4)          /* �X���u���� */
'slb_stat          VARCHAR2(1)          /* ��� */
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "TRTS0012_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        ReDim APResTmpData(1)
        APResTmpData(0).slb_no = strSlb_No
        APResTmpData(0).slb_chno = Mid(strSlb_No, 1, 5)
        APResTmpData(0).slb_aino = Mid(strSlb_No, 6)
        APResTmpData(0).slb_stat = strSlb_Stat
        TRTS0012_Read = True
        Exit Function
    End If

    On Error GoTo TRTS0012_Read_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT * FROM TRTS0012 WHERE slb_no='" & strSlb_No & "' AND slb_stat='" & strSlb_Stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APResTmpData(0)
    If Not oDS.EOF Then
        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), conDefault_slb_no, oDS.Fields("slb_no").Value)
        APResTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), conDefault_slb_chno, oDS.Fields("slb_chno").Value)
        APResTmpData(0).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), conDefault_slb_aino, oDS.Fields("slb_aino").Value)
        APResTmpData(0).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), conDefault_slb_stat, oDS.Fields("slb_stat").Value)
        
        APResTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), conDefault_slb_ccno, oDS.Fields("slb_ccno").Value)
        APResTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), conDefault_slb_zkai_dte, oDS.Fields("slb_zkai_dte").Value)
        APResTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), conDefault_slb_ksh, oDS.Fields("slb_ksh").Value)
        APResTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), conDefault_slb_typ, oDS.Fields("slb_typ").Value)
        APResTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), conDefault_slb_uksk, oDS.Fields("slb_uksk").Value)
        APResTmpData(0).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), conDefault_slb_wei, oDS.Fields("slb_wei").Value)
        APResTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), conDefault_slb_lngth, oDS.Fields("slb_lngth").Value)
        APResTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), conDefault_slb_wdth, oDS.Fields("slb_wdth").Value)
        APResTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), conDefault_slb_thkns, oDS.Fields("slb_thkns").Value)
        APResTmpData(0).slb_nxt_prcs = IIf(IsNull(oDS.Fields("slb_nxt_prcs").Value), conDefault_slb_nxt_prcs, oDS.Fields("slb_nxt_prcs").Value)
        APResTmpData(0).slb_cmt1 = IIf(IsNull(oDS.Fields("slb_cmt1").Value), conDefault_slb_cmt1, oDS.Fields("slb_cmt1").Value)
        APResTmpData(0).slb_cmt2 = IIf(IsNull(oDS.Fields("slb_cmt2").Value), conDefault_slb_cmt2, oDS.Fields("slb_cmt2").Value)
        
        APResTmpData(0).slb_fault_cd_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s1").Value), conDefault_slb_fault_cd_e_s1, oDS.Fields("slb_fault_cd_e_s1").Value)
        APResTmpData(0).slb_fault_cd_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s2").Value), conDefault_slb_fault_cd_e_s2, oDS.Fields("slb_fault_cd_e_s2").Value)
        APResTmpData(0).slb_fault_cd_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s3").Value), conDefault_slb_fault_cd_e_s3, oDS.Fields("slb_fault_cd_e_s3").Value)
        APResTmpData(0).slb_fault_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_e_s1").Value), conDefault_slb_fault_e_s1, oDS.Fields("slb_fault_e_s1").Value)
        APResTmpData(0).slb_fault_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_e_s2").Value), conDefault_slb_fault_e_s2, oDS.Fields("slb_fault_e_s2").Value)
        APResTmpData(0).slb_fault_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_e_s3").Value), conDefault_slb_fault_e_s3, oDS.Fields("slb_fault_e_s3").Value)
        APResTmpData(0).slb_fault_e_n1 = IIf(IsNull(oDS.Fields("slb_fault_e_n1").Value), conDefault_slb_fault_e_n1, oDS.Fields("slb_fault_e_n1").Value)
        APResTmpData(0).slb_fault_e_n2 = IIf(IsNull(oDS.Fields("slb_fault_e_n2").Value), conDefault_slb_fault_e_n2, oDS.Fields("slb_fault_e_n2").Value)
        APResTmpData(0).slb_fault_e_n3 = IIf(IsNull(oDS.Fields("slb_fault_e_n3").Value), conDefault_slb_fault_e_n3, oDS.Fields("slb_fault_e_n3").Value)
        
        APResTmpData(0).slb_fault_cd_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s1").Value), conDefault_slb_fault_cd_w_s1, oDS.Fields("slb_fault_cd_w_s1").Value)
        APResTmpData(0).slb_fault_cd_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s2").Value), conDefault_slb_fault_cd_w_s2, oDS.Fields("slb_fault_cd_w_s2").Value)
        APResTmpData(0).slb_fault_cd_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s3").Value), conDefault_slb_fault_cd_w_s3, oDS.Fields("slb_fault_cd_w_s3").Value)
        APResTmpData(0).slb_fault_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_w_s1").Value), conDefault_slb_fault_w_s1, oDS.Fields("slb_fault_w_s1").Value)
        APResTmpData(0).slb_fault_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_w_s2").Value), conDefault_slb_fault_w_s2, oDS.Fields("slb_fault_w_s2").Value)
        APResTmpData(0).slb_fault_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_w_s3").Value), conDefault_slb_fault_w_s3, oDS.Fields("slb_fault_w_s3").Value)
        APResTmpData(0).slb_fault_w_n1 = IIf(IsNull(oDS.Fields("slb_fault_w_n1").Value), conDefault_slb_fault_w_n1, oDS.Fields("slb_fault_w_n1").Value)
        APResTmpData(0).slb_fault_w_n2 = IIf(IsNull(oDS.Fields("slb_fault_w_n2").Value), conDefault_slb_fault_w_n2, oDS.Fields("slb_fault_w_n2").Value)
        APResTmpData(0).slb_fault_w_n3 = IIf(IsNull(oDS.Fields("slb_fault_w_n3").Value), conDefault_slb_fault_w_n3, oDS.Fields("slb_fault_w_n3").Value)
        
        APResTmpData(0).slb_fault_cd_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s1").Value), conDefault_slb_fault_cd_s_s1, oDS.Fields("slb_fault_cd_s_s1").Value)
        APResTmpData(0).slb_fault_cd_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s2").Value), conDefault_slb_fault_cd_s_s2, oDS.Fields("slb_fault_cd_s_s2").Value)
        APResTmpData(0).slb_fault_cd_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s3").Value), conDefault_slb_fault_cd_s_s3, oDS.Fields("slb_fault_cd_s_s3").Value)
        APResTmpData(0).slb_fault_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_s_s1").Value), conDefault_slb_fault_s_s1, oDS.Fields("slb_fault_s_s1").Value)
        APResTmpData(0).slb_fault_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_s_s2").Value), conDefault_slb_fault_s_s2, oDS.Fields("slb_fault_s_s2").Value)
        APResTmpData(0).slb_fault_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_s_s3").Value), conDefault_slb_fault_s_s3, oDS.Fields("slb_fault_s_s3").Value)
        APResTmpData(0).slb_fault_s_n1 = IIf(IsNull(oDS.Fields("slb_fault_s_n1").Value), conDefault_slb_fault_s_n1, oDS.Fields("slb_fault_s_n1").Value)
        APResTmpData(0).slb_fault_s_n2 = IIf(IsNull(oDS.Fields("slb_fault_s_n2").Value), conDefault_slb_fault_s_n2, oDS.Fields("slb_fault_s_n2").Value)
        APResTmpData(0).slb_fault_s_n3 = IIf(IsNull(oDS.Fields("slb_fault_s_n3").Value), conDefault_slb_fault_s_n3, oDS.Fields("slb_fault_s_n3").Value)
        
        APResTmpData(0).slb_fault_cd_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s1").Value), conDefault_slb_fault_cd_n_s1, oDS.Fields("slb_fault_cd_n_s1").Value)
        APResTmpData(0).slb_fault_cd_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s2").Value), conDefault_slb_fault_cd_n_s2, oDS.Fields("slb_fault_cd_n_s2").Value)
        APResTmpData(0).slb_fault_cd_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s3").Value), conDefault_slb_fault_cd_n_s3, oDS.Fields("slb_fault_cd_n_s3").Value)
        APResTmpData(0).slb_fault_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_n_s1").Value), conDefault_slb_fault_n_s1, oDS.Fields("slb_fault_n_s1").Value)
        APResTmpData(0).slb_fault_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_n_s2").Value), conDefault_slb_fault_n_s2, oDS.Fields("slb_fault_n_s2").Value)
        APResTmpData(0).slb_fault_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_n_s3").Value), conDefault_slb_fault_n_s3, oDS.Fields("slb_fault_n_s3").Value)
        APResTmpData(0).slb_fault_n_n1 = IIf(IsNull(oDS.Fields("slb_fault_n_n1").Value), conDefault_slb_fault_n_n1, oDS.Fields("slb_fault_n_n1").Value)
        APResTmpData(0).slb_fault_n_n2 = IIf(IsNull(oDS.Fields("slb_fault_n_n2").Value), conDefault_slb_fault_n_n2, oDS.Fields("slb_fault_n_n2").Value)
        APResTmpData(0).slb_fault_n_n3 = IIf(IsNull(oDS.Fields("slb_fault_n_n3").Value), conDefault_slb_fault_n_n3, oDS.Fields("slb_fault_n_n3").Value)
        
        APResTmpData(0).slb_fault_cd_bs_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bs_s").Value), conDefault_slb_fault_cd_bs_s, oDS.Fields("slb_fault_cd_bs_s").Value)
        APResTmpData(0).slb_fault_cd_bm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bm_s").Value), conDefault_slb_fault_cd_bm_s, oDS.Fields("slb_fault_cd_bm_s").Value)
        APResTmpData(0).slb_fault_cd_bn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bn_s").Value), conDefault_slb_fault_cd_bn_s, oDS.Fields("slb_fault_cd_bn_s").Value)
        APResTmpData(0).slb_fault_bs_s = IIf(IsNull(oDS.Fields("slb_fault_bs_s").Value), conDefault_slb_fault_bs_s, oDS.Fields("slb_fault_bs_s").Value)
        APResTmpData(0).slb_fault_bm_s = IIf(IsNull(oDS.Fields("slb_fault_bm_s").Value), conDefault_slb_fault_bm_s, oDS.Fields("slb_fault_bm_s").Value)
        APResTmpData(0).slb_fault_bn_s = IIf(IsNull(oDS.Fields("slb_fault_bn_s").Value), conDefault_slb_fault_bn_s, oDS.Fields("slb_fault_bn_s").Value)
        APResTmpData(0).slb_fault_bs_n = IIf(IsNull(oDS.Fields("slb_fault_bs_n").Value), conDefault_slb_fault_bs_n, oDS.Fields("slb_fault_bs_n").Value)
        APResTmpData(0).slb_fault_bm_n = IIf(IsNull(oDS.Fields("slb_fault_bm_n").Value), conDefault_slb_fault_bm_n, oDS.Fields("slb_fault_bm_n").Value)
        APResTmpData(0).slb_fault_bn_n = IIf(IsNull(oDS.Fields("slb_fault_bn_n").Value), conDefault_slb_fault_bn_n, oDS.Fields("slb_fault_bn_n").Value)
        
        APResTmpData(0).slb_fault_cd_ts_s = IIf(IsNull(oDS.Fields("slb_fault_cd_ts_s").Value), conDefault_slb_fault_cd_ts_s, oDS.Fields("slb_fault_cd_ts_s").Value)
        APResTmpData(0).slb_fault_cd_tm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tm_s").Value), conDefault_slb_fault_cd_tm_s, oDS.Fields("slb_fault_cd_tm_s").Value)
        APResTmpData(0).slb_fault_cd_tn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tn_s").Value), conDefault_slb_fault_cd_tn_s, oDS.Fields("slb_fault_cd_tn_s").Value)
        APResTmpData(0).slb_fault_ts_s = IIf(IsNull(oDS.Fields("slb_fault_ts_s").Value), conDefault_slb_fault_ts_s, oDS.Fields("slb_fault_ts_s").Value)
        APResTmpData(0).slb_fault_tm_s = IIf(IsNull(oDS.Fields("slb_fault_tm_s").Value), conDefault_slb_fault_tm_s, oDS.Fields("slb_fault_tm_s").Value)
        APResTmpData(0).slb_fault_tn_s = IIf(IsNull(oDS.Fields("slb_fault_tn_s").Value), conDefault_slb_fault_tn_s, oDS.Fields("slb_fault_tn_s").Value)
        APResTmpData(0).slb_fault_ts_n = IIf(IsNull(oDS.Fields("slb_fault_ts_n").Value), conDefault_slb_fault_ts_n, oDS.Fields("slb_fault_ts_n").Value)
        APResTmpData(0).slb_fault_tm_n = IIf(IsNull(oDS.Fields("slb_fault_tm_n").Value), conDefault_slb_fault_tm_n, oDS.Fields("slb_fault_tm_n").Value)
        APResTmpData(0).slb_fault_tn_n = IIf(IsNull(oDS.Fields("slb_fault_tn_n").Value), conDefault_slb_fault_tn_n, oDS.Fields("slb_fault_tn_n").Value)
        
        APResTmpData(0).slb_wrt_nme = IIf(IsNull(oDS.Fields("slb_wrt_nme").Value), conDefault_slb_wrt_nme, oDS.Fields("slb_wrt_nme").Value)
        APResTmpData(0).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), conDefault_sys_wrt_dte, oDS.Fields("sys_wrt_dte").Value)
        APResTmpData(0).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), conDefault_sys_wrt_tme, oDS.Fields("sys_wrt_tme").Value)
'        APResTmpData(0).sys_rwrt_dte = IIf(IsNull(oDS.Fields("sys_rwrt_dte").Value), conDefault_sys_rwrt_dte, oDS.Fields("sys_rwrt_dte").Value)
'        APResTmpData(0).sys_rwrt_tme = IIf(IsNull(oDS.Fields("sys_rwrt_tme").Value), conDefault_sys_rwrt_tme, oDS.Fields("sys_rwrt_tme").Value)
'        APResTmpData(0).sys_acs_pros = IIf(IsNull(oDS.Fields("sys_acs_pros").Value), conDefault_sys_acs_pros, oDS.Fields("sys_acs_pros").Value)
'        APResTmpData(0).sys_acs_enum = IIf(IsNull(oDS.Fields("sys_acs_enum").Value), conDefault_sys_acs_enum, oDS.Fields("sys_acs_enum").Value)
       
    End If

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Read ����I��") '�K�C�_���X�\��

    TRTS0012_Read = True

    On Error GoTo 0
    Exit Function

TRTS0012_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0012_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0012��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : ���ѓ��͂̃J�����g�f�[�^������
'
' ���l      : ���ѓ��̓f�[�^��������
'
Public Function TRTS0012_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0012_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0012_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0012_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0012 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0012 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_ccno,"       ''/* �X���uCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* ����� */
        strSQL = strSQL & "slb_ksh,"        ''/* �|�� */
        strSQL = strSQL & "slb_typ,"        ''/* �^ */
        strSQL = strSQL & "slb_uksk,"       ''/* ���� */
        strSQL = strSQL & "slb_wei,"        ''/* �d�� */
        strSQL = strSQL & "slb_lngth,"      ''/* ���� */
        strSQL = strSQL & "slb_wdth,"       ''/* �� */
        strSQL = strSQL & "slb_thkns,"      ''/* ���� */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* ���H�� */
        strSQL = strSQL & "slb_cmt1,"       ''/* �R�����g1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* �R�����g2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* ����E��CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* ����E��CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* ����E��CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* ����E�ʎ��1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* ����E�ʎ��2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* ����E�ʎ��3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* ����E�ʌ�1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* ����E�ʌ�2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* ����W��CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* ����W��CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* ����W��CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* ����W�ʎ��1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* ����W�ʎ��2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* ����W�ʎ��3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* ����W�ʌ�1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* ����W�ʌ�2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* ����S��CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* ����S��CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* ����S��CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* ����S�ʎ��1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* ����S�ʎ��2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* ����S�ʎ��3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* ����S�ʌ�1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* ����S�ʌ�2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* ����N��CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* ����N��CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* ����N��CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* ����N�ʎ��1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* ����N�ʎ��2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* ����N�ʎ��3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* ����N�ʌ�1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* ����N�ʌ�2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* ����N�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* ��������BSCD */
        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* ��������BMCD */
        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* ��������BNCD */
        strSQL = strSQL & "slb_fault_bs_s,"     ''/* ��������BS��� */
        strSQL = strSQL & "slb_fault_bm_s,"     ''/* ��������BM��� */
        strSQL = strSQL & "slb_fault_bn_s,"     ''/* ��������BN��� */
        strSQL = strSQL & "slb_fault_bs_n,"     ''/* ��������BS�� */
        strSQL = strSQL & "slb_fault_bm_n,"     ''/* ��������BM�� */
        strSQL = strSQL & "slb_fault_bn_n,"     ''/* ��������BN�� */
        
        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* ��������TSCD */
        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* ��������TMCD */
        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* ��������TNCD */
        strSQL = strSQL & "slb_fault_ts_s,"     ''/* ��������TS��� */
        strSQL = strSQL & "slb_fault_tm_s,"     ''/* ��������TM��� */
        strSQL = strSQL & "slb_fault_tn_s,"     ''/* ��������TN��� */
        strSQL = strSQL & "slb_fault_ts_n,"     ''/* ��������TS�� */
        strSQL = strSQL & "slb_fault_tm_n,"     ''/* ��������TM�� */
        strSQL = strSQL & "slb_fault_tn_n,"     ''/* ��������TN�� */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* ����E�ʔ��� */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* ����W�ʔ��� */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* ����S�ʔ��� */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* ����N�ʔ��� */
        strSQL = strSQL & "slb_fault_b_judg,"       ''/* ����B�ʔ��� */
        strSQL = strSQL & "slb_fault_t_judg,"       ''/* ����T�ʔ��� */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* �X�^�b�t�� */
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* �X���uCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* ����� */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* �|�� */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* �^ */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* �d�� */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* �� */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* ���H�� */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* �R�����g1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* �R�����g2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* ����E��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* ����E��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* ����E��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* ����E�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* ����E�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* ����E�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* ����E�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* ����E�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* ����W��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* ����W��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* ����W��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* ����W�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* ����W�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* ����W�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* ����W�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* ����W�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* ����S��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* ����S��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* ����S��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* ����S�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* ����S�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* ����S�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* ����S�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* ����S�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* ����N��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* ����N��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* ����N��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* ����N�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* ����N�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* ����N�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* ����N�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* ����N�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* ����N�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* ��������BSCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* ��������BMCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* ��������BNCD */
        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* ��������BS��� */
        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* ��������BM��� */
        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* ��������BN��� */
        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* ��������BS�� */
        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* ��������BM�� */
        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* ��������BN�� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* ��������TSCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* ��������TMCD */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* ��������TNCD */
        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* ��������TS��� */
        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* ��������TM��� */
        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* ��������TN��� */
        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* ��������TS�� */
        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* ��������TM�� */
        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* ��������TN�� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* ����E�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* ����W�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* ����S�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* ����N�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* ����B�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* ����T�ʔ��� */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* �X�^�b�t�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* �o�^�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Write ����I��") '�K�C�_���X�\��

    TRTS0012_Write = True

    On Error GoTo 0
    Exit Function

TRTS0012_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0012_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0012_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0014�Ǎ�����
'
' ������    : ARG1 - �X���u�ԍ�
'           : ARG2 - ���
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p����TRTS0014�̃��R�[�h��Ǎ�
'
' ���l      : �J���[�`�F�b�N���ѓ��̓f�[�^�Ǎ�
'           :COLORSYS
'
Public Function TRTS0014_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String, ByVal strSlb_Col_Cnt As String) As Boolean
'slb_chno          VARCHAR2(5)          /* �X���u�`���[�WNO */
'slb_aino          VARCHAR2(4)          /* �X���u���� */
'slb_stat          VARCHAR2(1)          /* ��� */
'slb_col_cnt       VARCHAR2(2)          /* �װ�� */
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0014_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        ReDim APResTmpData(0)
        
        Call TRTS0014_ReadCSV
'        APResTmpData(0).slb_no = strSlb_No
'        APResTmpData(0).slb_stat = strSlb_Stat
'        APResTmpData(0).slb_col_cnt = Format(CInt(strSlb_Col_Cnt), "00")
        TRTS0014_Read = True
        Exit Function
    End If

    On Error GoTo TRTS0014_Read_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT * FROM TRTS0014 WHERE slb_no='" & strSlb_No & "' AND slb_stat='" & strSlb_Stat & "' AND slb_col_cnt='" & Format(CInt(strSlb_Col_Cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APResTmpData(0)
    If Not oDS.EOF Then
        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), conDefault_slb_no, oDS.Fields("slb_no").Value)
        APResTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), conDefault_slb_chno, oDS.Fields("slb_chno").Value)
        APResTmpData(0).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), conDefault_slb_aino, oDS.Fields("slb_aino").Value)
        APResTmpData(0).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), conDefault_slb_stat, oDS.Fields("slb_stat").Value)
        
        APResTmpData(0).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), conDefault_slb_col_cnt, oDS.Fields("slb_col_cnt").Value)
        
        APResTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("slb_ccno").Value), conDefault_slb_ccno, oDS.Fields("slb_ccno").Value)
        APResTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("slb_zkai_dte").Value), conDefault_slb_zkai_dte, oDS.Fields("slb_zkai_dte").Value)
        APResTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("slb_ksh").Value), conDefault_slb_ksh, oDS.Fields("slb_ksh").Value)
        APResTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("slb_typ").Value), conDefault_slb_typ, oDS.Fields("slb_typ").Value)
        APResTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("slb_uksk").Value), conDefault_slb_uksk, oDS.Fields("slb_uksk").Value)
        APResTmpData(0).slb_wei = IIf(IsNull(oDS.Fields("slb_wei").Value), conDefault_slb_wei, oDS.Fields("slb_wei").Value)
        APResTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("slb_lngth").Value), conDefault_slb_lngth, oDS.Fields("slb_lngth").Value)
        APResTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("slb_wdth").Value), conDefault_slb_wdth, oDS.Fields("slb_wdth").Value)
        APResTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("slb_thkns").Value), conDefault_slb_thkns, oDS.Fields("slb_thkns").Value)
        APResTmpData(0).slb_nxt_prcs = IIf(IsNull(oDS.Fields("slb_nxt_prcs").Value), conDefault_slb_nxt_prcs, oDS.Fields("slb_nxt_prcs").Value)
        APResTmpData(0).slb_cmt1 = IIf(IsNull(oDS.Fields("slb_cmt1").Value), conDefault_slb_cmt1, oDS.Fields("slb_cmt1").Value)
        APResTmpData(0).slb_cmt2 = IIf(IsNull(oDS.Fields("slb_cmt2").Value), conDefault_slb_cmt2, oDS.Fields("slb_cmt2").Value)
        
        APResTmpData(0).slb_fault_cd_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s1").Value), conDefault_slb_fault_cd_e_s1, oDS.Fields("slb_fault_cd_e_s1").Value)
        APResTmpData(0).slb_fault_cd_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s2").Value), conDefault_slb_fault_cd_e_s2, oDS.Fields("slb_fault_cd_e_s2").Value)
        APResTmpData(0).slb_fault_cd_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_e_s3").Value), conDefault_slb_fault_cd_e_s3, oDS.Fields("slb_fault_cd_e_s3").Value)
        APResTmpData(0).slb_fault_e_s1 = IIf(IsNull(oDS.Fields("slb_fault_e_s1").Value), conDefault_slb_fault_e_s1, oDS.Fields("slb_fault_e_s1").Value)
        APResTmpData(0).slb_fault_e_s2 = IIf(IsNull(oDS.Fields("slb_fault_e_s2").Value), conDefault_slb_fault_e_s2, oDS.Fields("slb_fault_e_s2").Value)
        APResTmpData(0).slb_fault_e_s3 = IIf(IsNull(oDS.Fields("slb_fault_e_s3").Value), conDefault_slb_fault_e_s3, oDS.Fields("slb_fault_e_s3").Value)
        APResTmpData(0).slb_fault_e_n1 = IIf(IsNull(oDS.Fields("slb_fault_e_n1").Value), conDefault_slb_fault_e_n1, oDS.Fields("slb_fault_e_n1").Value)
        APResTmpData(0).slb_fault_e_n2 = IIf(IsNull(oDS.Fields("slb_fault_e_n2").Value), conDefault_slb_fault_e_n2, oDS.Fields("slb_fault_e_n2").Value)
        APResTmpData(0).slb_fault_e_n3 = IIf(IsNull(oDS.Fields("slb_fault_e_n3").Value), conDefault_slb_fault_e_n3, oDS.Fields("slb_fault_e_n3").Value)
        
        APResTmpData(0).slb_fault_cd_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s1").Value), conDefault_slb_fault_cd_w_s1, oDS.Fields("slb_fault_cd_w_s1").Value)
        APResTmpData(0).slb_fault_cd_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s2").Value), conDefault_slb_fault_cd_w_s2, oDS.Fields("slb_fault_cd_w_s2").Value)
        APResTmpData(0).slb_fault_cd_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_w_s3").Value), conDefault_slb_fault_cd_w_s3, oDS.Fields("slb_fault_cd_w_s3").Value)
        APResTmpData(0).slb_fault_w_s1 = IIf(IsNull(oDS.Fields("slb_fault_w_s1").Value), conDefault_slb_fault_w_s1, oDS.Fields("slb_fault_w_s1").Value)
        APResTmpData(0).slb_fault_w_s2 = IIf(IsNull(oDS.Fields("slb_fault_w_s2").Value), conDefault_slb_fault_w_s2, oDS.Fields("slb_fault_w_s2").Value)
        APResTmpData(0).slb_fault_w_s3 = IIf(IsNull(oDS.Fields("slb_fault_w_s3").Value), conDefault_slb_fault_w_s3, oDS.Fields("slb_fault_w_s3").Value)
        APResTmpData(0).slb_fault_w_n1 = IIf(IsNull(oDS.Fields("slb_fault_w_n1").Value), conDefault_slb_fault_w_n1, oDS.Fields("slb_fault_w_n1").Value)
        APResTmpData(0).slb_fault_w_n2 = IIf(IsNull(oDS.Fields("slb_fault_w_n2").Value), conDefault_slb_fault_w_n2, oDS.Fields("slb_fault_w_n2").Value)
        APResTmpData(0).slb_fault_w_n3 = IIf(IsNull(oDS.Fields("slb_fault_w_n3").Value), conDefault_slb_fault_w_n3, oDS.Fields("slb_fault_w_n3").Value)
        
        APResTmpData(0).slb_fault_cd_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s1").Value), conDefault_slb_fault_cd_s_s1, oDS.Fields("slb_fault_cd_s_s1").Value)
        APResTmpData(0).slb_fault_cd_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s2").Value), conDefault_slb_fault_cd_s_s2, oDS.Fields("slb_fault_cd_s_s2").Value)
        APResTmpData(0).slb_fault_cd_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_s_s3").Value), conDefault_slb_fault_cd_s_s3, oDS.Fields("slb_fault_cd_s_s3").Value)
        APResTmpData(0).slb_fault_s_s1 = IIf(IsNull(oDS.Fields("slb_fault_s_s1").Value), conDefault_slb_fault_s_s1, oDS.Fields("slb_fault_s_s1").Value)
        APResTmpData(0).slb_fault_s_s2 = IIf(IsNull(oDS.Fields("slb_fault_s_s2").Value), conDefault_slb_fault_s_s2, oDS.Fields("slb_fault_s_s2").Value)
        APResTmpData(0).slb_fault_s_s3 = IIf(IsNull(oDS.Fields("slb_fault_s_s3").Value), conDefault_slb_fault_s_s3, oDS.Fields("slb_fault_s_s3").Value)
        APResTmpData(0).slb_fault_s_n1 = IIf(IsNull(oDS.Fields("slb_fault_s_n1").Value), conDefault_slb_fault_s_n1, oDS.Fields("slb_fault_s_n1").Value)
        APResTmpData(0).slb_fault_s_n2 = IIf(IsNull(oDS.Fields("slb_fault_s_n2").Value), conDefault_slb_fault_s_n2, oDS.Fields("slb_fault_s_n2").Value)
        APResTmpData(0).slb_fault_s_n3 = IIf(IsNull(oDS.Fields("slb_fault_s_n3").Value), conDefault_slb_fault_s_n3, oDS.Fields("slb_fault_s_n3").Value)
        
        APResTmpData(0).slb_fault_cd_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s1").Value), conDefault_slb_fault_cd_n_s1, oDS.Fields("slb_fault_cd_n_s1").Value)
        APResTmpData(0).slb_fault_cd_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s2").Value), conDefault_slb_fault_cd_n_s2, oDS.Fields("slb_fault_cd_n_s2").Value)
        APResTmpData(0).slb_fault_cd_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_cd_n_s3").Value), conDefault_slb_fault_cd_n_s3, oDS.Fields("slb_fault_cd_n_s3").Value)
        APResTmpData(0).slb_fault_n_s1 = IIf(IsNull(oDS.Fields("slb_fault_n_s1").Value), conDefault_slb_fault_n_s1, oDS.Fields("slb_fault_n_s1").Value)
        APResTmpData(0).slb_fault_n_s2 = IIf(IsNull(oDS.Fields("slb_fault_n_s2").Value), conDefault_slb_fault_n_s2, oDS.Fields("slb_fault_n_s2").Value)
        APResTmpData(0).slb_fault_n_s3 = IIf(IsNull(oDS.Fields("slb_fault_n_s3").Value), conDefault_slb_fault_n_s3, oDS.Fields("slb_fault_n_s3").Value)
        APResTmpData(0).slb_fault_n_n1 = IIf(IsNull(oDS.Fields("slb_fault_n_n1").Value), conDefault_slb_fault_n_n1, oDS.Fields("slb_fault_n_n1").Value)
        APResTmpData(0).slb_fault_n_n2 = IIf(IsNull(oDS.Fields("slb_fault_n_n2").Value), conDefault_slb_fault_n_n2, oDS.Fields("slb_fault_n_n2").Value)
        APResTmpData(0).slb_fault_n_n3 = IIf(IsNull(oDS.Fields("slb_fault_n_n3").Value), conDefault_slb_fault_n_n3, oDS.Fields("slb_fault_n_n3").Value)
        
'        APResTmpData(0).slb_fault_cd_bs_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bs_s").Value), conDefault_slb_fault_cd_bs_s, oDS.Fields("slb_fault_cd_bs_s").Value)
'        APResTmpData(0).slb_fault_cd_bm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bm_s").Value), conDefault_slb_fault_cd_bm_s, oDS.Fields("slb_fault_cd_bm_s").Value)
'        APResTmpData(0).slb_fault_cd_bn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_bn_s").Value), conDefault_slb_fault_cd_bn_s, oDS.Fields("slb_fault_cd_bn_s").Value)
'        APResTmpData(0).slb_fault_bs_s = IIf(IsNull(oDS.Fields("slb_fault_bs_s").Value), conDefault_slb_fault_bs_s, oDS.Fields("slb_fault_bs_s").Value)
'        APResTmpData(0).slb_fault_bm_s = IIf(IsNull(oDS.Fields("slb_fault_bm_s").Value), conDefault_slb_fault_bm_s, oDS.Fields("slb_fault_bm_s").Value)
'        APResTmpData(0).slb_fault_bn_s = IIf(IsNull(oDS.Fields("slb_fault_bn_s").Value), conDefault_slb_fault_bn_s, oDS.Fields("slb_fault_bn_s").Value)
'        APResTmpData(0).slb_fault_bs_n = IIf(IsNull(oDS.Fields("slb_fault_bs_n").Value), conDefault_slb_fault_bs_n, oDS.Fields("slb_fault_bs_n").Value)
'        APResTmpData(0).slb_fault_bm_n = IIf(IsNull(oDS.Fields("slb_fault_bm_n").Value), conDefault_slb_fault_bm_n, oDS.Fields("slb_fault_bm_n").Value)
'        APResTmpData(0).slb_fault_bn_n = IIf(IsNull(oDS.Fields("slb_fault_bn_n").Value), conDefault_slb_fault_bn_n, oDS.Fields("slb_fault_bn_n").Value)
'
'        APResTmpData(0).slb_fault_cd_ts_s = IIf(IsNull(oDS.Fields("slb_fault_cd_ts_s").Value), conDefault_slb_fault_cd_ts_s, oDS.Fields("slb_fault_cd_ts_s").Value)
'        APResTmpData(0).slb_fault_cd_tm_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tm_s").Value), conDefault_slb_fault_cd_tm_s, oDS.Fields("slb_fault_cd_tm_s").Value)
'        APResTmpData(0).slb_fault_cd_tn_s = IIf(IsNull(oDS.Fields("slb_fault_cd_tn_s").Value), conDefault_slb_fault_cd_tn_s, oDS.Fields("slb_fault_cd_tn_s").Value)
'        APResTmpData(0).slb_fault_ts_s = IIf(IsNull(oDS.Fields("slb_fault_ts_s").Value), conDefault_slb_fault_ts_s, oDS.Fields("slb_fault_ts_s").Value)
'        APResTmpData(0).slb_fault_tm_s = IIf(IsNull(oDS.Fields("slb_fault_tm_s").Value), conDefault_slb_fault_tm_s, oDS.Fields("slb_fault_tm_s").Value)
'        APResTmpData(0).slb_fault_tn_s = IIf(IsNull(oDS.Fields("slb_fault_tn_s").Value), conDefault_slb_fault_tn_s, oDS.Fields("slb_fault_tn_s").Value)
'        APResTmpData(0).slb_fault_ts_n = IIf(IsNull(oDS.Fields("slb_fault_ts_n").Value), conDefault_slb_fault_ts_n, oDS.Fields("slb_fault_ts_n").Value)
'        APResTmpData(0).slb_fault_tm_n = IIf(IsNull(oDS.Fields("slb_fault_tm_n").Value), conDefault_slb_fault_tm_n, oDS.Fields("slb_fault_tm_n").Value)
'        APResTmpData(0).slb_fault_tn_n = IIf(IsNull(oDS.Fields("slb_fault_tn_n").Value), conDefault_slb_fault_tn_n, oDS.Fields("slb_fault_tn_n").Value)
        
        APResTmpData(0).slb_wrt_nme = IIf(IsNull(oDS.Fields("slb_wrt_nme").Value), conDefault_slb_wrt_nme, oDS.Fields("slb_wrt_nme").Value)
        
        APResTmpData(0).host_send = IIf(IsNull(oDS.Fields("host_send").Value), conDefault_host_send, oDS.Fields("host_send").Value)
        APResTmpData(0).host_wrt_dte = IIf(IsNull(oDS.Fields("host_wrt_dte").Value), conDefault_host_wrt_dte, oDS.Fields("host_wrt_dte").Value)
        APResTmpData(0).host_wrt_tme = IIf(IsNull(oDS.Fields("host_wrt_tme").Value), conDefault_host_wrt_tme, oDS.Fields("host_wrt_tme").Value)
        
        APResTmpData(0).sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), conDefault_sys_wrt_dte, oDS.Fields("sys_wrt_dte").Value)
        APResTmpData(0).sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), conDefault_sys_wrt_tme, oDS.Fields("sys_wrt_tme").Value)
'        APResTmpData(0).sys_rwrt_dte = IIf(IsNull(oDS.Fields("sys_rwrt_dte").Value), conDefault_sys_rwrt_dte, oDS.Fields("sys_rwrt_dte").Value)
'        APResTmpData(0).sys_rwrt_tme = IIf(IsNull(oDS.Fields("sys_rwrt_tme").Value), conDefault_sys_rwrt_tme, oDS.Fields("sys_rwrt_tme").Value)
'        APResTmpData(0).sys_acs_pros = IIf(IsNull(oDS.Fields("sys_acs_pros").Value), conDefault_sys_acs_pros, oDS.Fields("sys_acs_pros").Value)
'        APResTmpData(0).sys_acs_enum = IIf(IsNull(oDS.Fields("sys_acs_enum").Value), conDefault_sys_acs_enum, oDS.Fields("sys_acs_enum").Value)
       
    End If

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Read ����I��") '�K�C�_���X�\��

    TRTS0014_Read = True

    On Error GoTo 0
    Exit Function

TRTS0014_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0014_Read = False

    On Error GoTo 0
End Function

Private Sub TRTS0014_ReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APResTmpData(0)

    bRet = ReadCSV(App.path & "\" & "TRTS0014_Read.csv", strItem(), strData())

    For nI = 0 To UBound(strData, 2) - 1

        ReDim APResTmpData(1)

        APResTmpData(0).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APResTmpData(0).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APResTmpData(0).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_col_cnt = getItemDataCSV("slb_col_cnt", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_ccno = getItemDataCSV("slb_ccno", nI + 1, strItem(), strData())
        APResTmpData(0).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())
        APResTmpData(0).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())
        APResTmpData(0).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())
        APResTmpData(0).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        APResTmpData(0).slb_wei = getItemDataCSV("slb_wei", nI + 1, strItem(), strData())
        APResTmpData(0).slb_lngth = getItemDataCSV("slb_lngth", nI + 1, strItem(), strData())
        APResTmpData(0).slb_wdth = getItemDataCSV("slb_wdth", nI + 1, strItem(), strData())
        APResTmpData(0).slb_thkns = getItemDataCSV("slb_thkns", nI + 1, strItem(), strData())
        APResTmpData(0).slb_nxt_prcs = getItemDataCSV("slb_nxt_prcs", nI + 1, strItem(), strData())
        APResTmpData(0).slb_cmt1 = getItemDataCSV("slb_cmt1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_cmt2 = getItemDataCSV("slb_cmt2", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_e_s1 = getItemDataCSV("slb_fault_cd_e_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_e_s2 = getItemDataCSV("slb_fault_cd_e_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_e_s3 = getItemDataCSV("slb_fault_cd_e_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s1 = getItemDataCSV("slb_fault_e_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s2 = getItemDataCSV("slb_fault_e_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_s3 = getItemDataCSV("slb_fault_e_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n1 = getItemDataCSV("slb_fault_e_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n2 = getItemDataCSV("slb_fault_e_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_e_n3 = getItemDataCSV("slb_fault_e_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_w_s1 = getItemDataCSV("slb_fault_cd_w_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_w_s2 = getItemDataCSV("slb_fault_cd_w_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_w_s3 = getItemDataCSV("slb_fault_cd_w_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s1 = getItemDataCSV("slb_fault_w_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s2 = getItemDataCSV("slb_fault_w_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_s3 = getItemDataCSV("slb_fault_w_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n1 = getItemDataCSV("slb_fault_w_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n2 = getItemDataCSV("slb_fault_w_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_w_n3 = getItemDataCSV("slb_fault_w_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_s_s1 = getItemDataCSV("slb_fault_cd_s_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_s_s2 = getItemDataCSV("slb_fault_cd_s_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_s_s3 = getItemDataCSV("slb_fault_cd_s_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s1 = getItemDataCSV("slb_fault_s_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s2 = getItemDataCSV("slb_fault_s_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_s3 = getItemDataCSV("slb_fault_s_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n1 = getItemDataCSV("slb_fault_s_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n2 = getItemDataCSV("slb_fault_s_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_s_n3 = getItemDataCSV("slb_fault_s_n3", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_fault_cd_n_s1 = getItemDataCSV("slb_fault_cd_n_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_n_s2 = getItemDataCSV("slb_fault_cd_n_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_cd_n_s3 = getItemDataCSV("slb_fault_cd_n_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s1 = getItemDataCSV("slb_fault_n_s1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s2 = getItemDataCSV("slb_fault_n_s2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_s3 = getItemDataCSV("slb_fault_n_s3", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n1 = getItemDataCSV("slb_fault_n_n1", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n2 = getItemDataCSV("slb_fault_n_n2", nI + 1, strItem(), strData())
        APResTmpData(0).slb_fault_n_n3 = getItemDataCSV("slb_fault_n_n3", nI + 1, strItem(), strData())
        
'        APResTmpData(0).slb_fault_cd_bs_s = getItemDataCSV("slb_fault_cd_bs_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_bm_s = getItemDataCSV("slb_fault_cd_bm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_bn_s = getItemDataCSV("slb_fault_cd_bn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bs_s = getItemDataCSV("slb_fault_bs_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bm_s = getItemDataCSV("slb_fault_bm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bn_s = getItemDataCSV("slb_fault_bn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bs_n = getItemDataCSV("slb_fault_bs_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bm_n = getItemDataCSV("slb_fault_bm_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_bn_n = getItemDataCSV("slb_fault_bn_n", nI + 1, strItem(), strData())
'
'        APResTmpData(0).slb_fault_cd_ts_s = getItemDataCSV("slb_fault_cd_ts_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_tm_s = getItemDataCSV("slb_fault_cd_tm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_cd_tn_s = getItemDataCSV("slb_fault_cd_tn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_ts_s = getItemDataCSV("slb_fault_ts_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tm_s = getItemDataCSV("slb_fault_tm_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tn_s = getItemDataCSV("slb_fault_tn_s", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_ts_n = getItemDataCSV("slb_fault_ts_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tm_n = getItemDataCSV("slb_fault_tm_n", nI + 1, strItem(), strData())
'        APResTmpData(0).slb_fault_tn_n = getItemDataCSV("slb_fault_tn_n", nI + 1, strItem(), strData())
        
        APResTmpData(0).slb_wrt_nme = getItemDataCSV("slb_wrt_nme", nI + 1, strItem(), strData())
        
        APResTmpData(0).host_send = getItemDataCSV("host_send", nI + 1, strItem(), strData())
        APResTmpData(0).host_wrt_dte = getItemDataCSV("host_wrt_dte", nI + 1, strItem(), strData())
        APResTmpData(0).host_wrt_tme = getItemDataCSV("host_wrt_tme", nI + 1, strItem(), strData())
        
        APResTmpData(0).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())
        APResTmpData(0).sys_wrt_tme = getItemDataCSV("sys_wrt_tme", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_rwrt_dte = getItemDataCSV("sys_rwrt_dte", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_rwrt_tme = getItemDataCSV("sys_rwrt_tme", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_acs_pros = getItemDataCSV("sys_acs_pros", nI + 1, strItem(), strData())
'        APResTmpData(0).sys_acs_enum = getItemDataCSV("sys_acs_enum", nI + 1, strItem(), strData())
        
        Exit For '�P���̂ݗL��
        

    Next nI

End Sub





' @(f)
'
' �@�\      : TRTS0014��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : ���ѓ��͂̃J�����g�f�[�^������
'
' ���l      : ���ѓ��̓f�[�^��������
'
Public Function TRTS0014_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0014_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0014_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0014_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0014 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0014 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        
        strSQL = strSQL & "slb_col_cnt,"       ''/* �װ�� */
        
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_ccno,"       ''/* �X���uCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* ����� */
        strSQL = strSQL & "slb_ksh,"        ''/* �|�� */
        strSQL = strSQL & "slb_typ,"        ''/* �^ */
        strSQL = strSQL & "slb_uksk,"       ''/* ���� */
        strSQL = strSQL & "slb_wei,"        ''/* �d�� */
        strSQL = strSQL & "slb_lngth,"      ''/* ���� */
        strSQL = strSQL & "slb_wdth,"       ''/* �� */
        strSQL = strSQL & "slb_thkns,"      ''/* ���� */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* ���H�� */
        strSQL = strSQL & "slb_cmt1,"       ''/* �R�����g1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* �R�����g2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* ����E��CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* ����E��CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* ����E��CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* ����E�ʎ��1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* ����E�ʎ��2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* ����E�ʎ��3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* ����E�ʌ�1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* ����E�ʌ�2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* ����W��CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* ����W��CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* ����W��CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* ����W�ʎ��1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* ����W�ʎ��2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* ����W�ʎ��3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* ����W�ʌ�1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* ����W�ʌ�2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* ����S��CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* ����S��CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* ����S��CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* ����S�ʎ��1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* ����S�ʎ��2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* ����S�ʎ��3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* ����S�ʌ�1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* ����S�ʌ�2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* ����N��CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* ����N��CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* ����N��CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* ����N�ʎ��1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* ����N�ʎ��2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* ����N�ʎ��3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* ����N�ʌ�1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* ����N�ʌ�2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* ����N�ʌ�3 */
        
'        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* ��������BSCD */
'        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* ��������BMCD */
'        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* ��������BNCD */
'        strSQL = strSQL & "slb_fault_bs_s,"     ''/* ��������BS��� */
'        strSQL = strSQL & "slb_fault_bm_s,"     ''/* ��������BM��� */
'        strSQL = strSQL & "slb_fault_bn_s,"     ''/* ��������BN��� */
'        strSQL = strSQL & "slb_fault_bs_n,"     ''/* ��������BS�� */
'        strSQL = strSQL & "slb_fault_bm_n,"     ''/* ��������BM�� */
'        strSQL = strSQL & "slb_fault_bn_n,"     ''/* ��������BN�� */
'
'        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* ��������TSCD */
'        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* ��������TMCD */
'        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* ��������TNCD */
'        strSQL = strSQL & "slb_fault_ts_s,"     ''/* ��������TS��� */
'        strSQL = strSQL & "slb_fault_tm_s,"     ''/* ��������TM��� */
'        strSQL = strSQL & "slb_fault_tn_s,"     ''/* ��������TN��� */
'        strSQL = strSQL & "slb_fault_ts_n,"     ''/* ��������TS�� */
'        strSQL = strSQL & "slb_fault_tm_n,"     ''/* ��������TM�� */
'        strSQL = strSQL & "slb_fault_tn_n,"     ''/* ��������TN�� */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* ����E�ʔ��� */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* ����W�ʔ��� */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* ����S�ʔ��� */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* ����N�ʔ��� */
'        strSQL = strSQL & "slb_fault_b_judg,"       ''/* ����B�ʔ��� */
'        strSQL = strSQL & "slb_fault_t_judg,"       ''/* ����T�ʔ��� */
        strSQL = strSQL & "slb_fault_u_judg,"       ''/* ����U�ʔ��� */
        strSQL = strSQL & "slb_fault_d_judg,"       ''/* ����D�ʔ��� */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* �������� */
        
        '2016/04/20 - TAI - S
        strSQL = strSQL & "slb_fault_total_judg,"     ''/* �������� */
        strSQL = strSQL & "slb_works_sky_tok,"        ''/* ��Əꏊ */
        '2016/04/20 - TAI - E
        
        strSQL = strSQL & "host_send,"          ''/* �r�W�R�����M���� */
        strSQL = strSQL & "host_wrt_dte,"       ''/* �r�W�R���o�^�� */
        strSQL = strSQL & "host_wrt_tme,"       ''/* �r�W�R���o�^���� */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","     ''/* �װ�� */
        
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* �X���uCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* ����� */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* �|�� */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* �^ */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* �d�� */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* �� */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* ���H�� */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* �R�����g1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* �R�����g2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* ����E��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* ����E��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* ����E��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* ����E�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* ����E�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* ����E�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* ����E�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* ����E�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* ����W��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* ����W��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* ����W��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* ����W�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* ����W�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* ����W�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* ����W�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* ����W�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* ����S��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* ����S��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* ����S��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* ����S�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* ����S�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* ����S�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* ����S�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* ����S�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* ����N��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* ����N��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* ����N��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* ����N�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* ����N�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* ����N�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* ����N�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* ����N�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* ����N�ʌ�3 */
        
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* ��������BSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* ��������BMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* ��������BNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* ��������BS��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* ��������BM��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* ��������BN��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* ��������BS�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* ��������BM�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* ��������BN�� */
'
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* ��������TSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* ��������TMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* ��������TNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* ��������TS��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* ��������TM��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* ��������TN��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* ��������TS�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* ��������TM�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* ��������TN�� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* ����E�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* ����W�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* ����S�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* ����N�ʔ��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* ����B�ʔ��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* ����T�ʔ��� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_u_judg & "'" & ","      ''/* ����U�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_d_judg & "'" & ","      ''/* ����D�ʔ��� */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* �������� */
        
        '2016/04/20 - TAI - S
        strSQL = strSQL & "'" & APResData.slb_fault_total_judg & "'" & ","        ''/* �������� */
        strSQL = strSQL & "'" & APResData.slb_works_sky_tok & "'" & ","           ''/* ��Əꏊ */
        '2016/04/20 - TAI - E
        
        strSQL = strSQL & "'" & APResData.host_send & "'" & ","             ''/* �r�W�R�����M���� */
        strSQL = strSQL & "'" & APResData.host_wrt_dte & "'" & ","          ''/* �r�W�R���o�^�� */
        strSQL = strSQL & "'" & APResData.host_wrt_tme & "'" & ","          ''/* �r�W�R���o�^���� */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* �o�^�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Write ����I��") '�K�C�_���X�\��

    TRTS0014_Write = True

    On Error GoTo 0
    Exit Function

TRTS0014_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0014_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0014_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0016��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �X���u�ُ�񍐏����͂̃J�����g�f�[�^������
'
' ���l      : �X���u�ُ�񍐏����̓f�[�^��������
'
Public Function TRTS0016_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0016_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0016_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0016_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0016 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0016 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        
        strSQL = strSQL & "slb_col_cnt,"       ''/* �װ�� */
        
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_ccno,"       ''/* �X���uCCNO */
        strSQL = strSQL & "slb_zkai_dte,"       ''/* ����� */
        strSQL = strSQL & "slb_ksh,"        ''/* �|�� */
        strSQL = strSQL & "slb_typ,"        ''/* �^ */
        strSQL = strSQL & "slb_uksk,"       ''/* ���� */
        strSQL = strSQL & "slb_wei,"        ''/* �d�� */
        strSQL = strSQL & "slb_lngth,"      ''/* ���� */
        strSQL = strSQL & "slb_wdth,"       ''/* �� */
        strSQL = strSQL & "slb_thkns,"      ''/* ���� */
        strSQL = strSQL & "slb_nxt_prcs,"       ''/* ���H�� */
        strSQL = strSQL & "slb_cmt1,"       ''/* �R�����g1 */
        strSQL = strSQL & "slb_cmt2,"       ''/* �R�����g2 */
        
        strSQL = strSQL & "slb_fault_cd_e_s1,"      ''/* ����E��CD1 */
        strSQL = strSQL & "slb_fault_cd_e_s2,"      ''/* ����E��CD2 */
        strSQL = strSQL & "slb_fault_cd_e_s3,"      ''/* ����E��CD3 */
        strSQL = strSQL & "slb_fault_e_s1,"     ''/* ����E�ʎ��1 */
        strSQL = strSQL & "slb_fault_e_s2,"     ''/* ����E�ʎ��2 */
        strSQL = strSQL & "slb_fault_e_s3,"     ''/* ����E�ʎ��3 */
        strSQL = strSQL & "slb_fault_e_n1,"     ''/* ����E�ʌ�1 */
        strSQL = strSQL & "slb_fault_e_n2,"     ''/* ����E�ʌ�2 */
        strSQL = strSQL & "slb_fault_e_n3,"     ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_w_s1,"      ''/* ����W��CD1 */
        strSQL = strSQL & "slb_fault_cd_w_s2,"      ''/* ����W��CD2 */
        strSQL = strSQL & "slb_fault_cd_w_s3,"      ''/* ����W��CD3 */
        strSQL = strSQL & "slb_fault_w_s1,"     ''/* ����W�ʎ��1 */
        strSQL = strSQL & "slb_fault_w_s2,"     ''/* ����W�ʎ��2 */
        strSQL = strSQL & "slb_fault_w_s3,"     ''/* ����W�ʎ��3 */
        strSQL = strSQL & "slb_fault_w_n1,"     ''/* ����W�ʌ�1 */
        strSQL = strSQL & "slb_fault_w_n2,"     ''/* ����W�ʌ�2 */
        strSQL = strSQL & "slb_fault_w_n3,"     ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_s_s1,"      ''/* ����S��CD1 */
        strSQL = strSQL & "slb_fault_cd_s_s2,"      ''/* ����S��CD2 */
        strSQL = strSQL & "slb_fault_cd_s_s3,"      ''/* ����S��CD3 */
        strSQL = strSQL & "slb_fault_s_s1,"     ''/* ����S�ʎ��1 */
        strSQL = strSQL & "slb_fault_s_s2,"     ''/* ����S�ʎ��2 */
        strSQL = strSQL & "slb_fault_s_s3,"     ''/* ����S�ʎ��3 */
        strSQL = strSQL & "slb_fault_s_n1,"     ''/* ����S�ʌ�1 */
        strSQL = strSQL & "slb_fault_s_n2,"     ''/* ����S�ʌ�2 */
        strSQL = strSQL & "slb_fault_s_n3,"     ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "slb_fault_cd_n_s1,"      ''/* ����N��CD1 */
        strSQL = strSQL & "slb_fault_cd_n_s2,"      ''/* ����N��CD2 */
        strSQL = strSQL & "slb_fault_cd_n_s3,"      ''/* ����N��CD3 */
        strSQL = strSQL & "slb_fault_n_s1,"     ''/* ����N�ʎ��1 */
        strSQL = strSQL & "slb_fault_n_s2,"     ''/* ����N�ʎ��2 */
        strSQL = strSQL & "slb_fault_n_s3,"     ''/* ����N�ʎ��3 */
        strSQL = strSQL & "slb_fault_n_n1,"     ''/* ����N�ʌ�1 */
        strSQL = strSQL & "slb_fault_n_n2,"     ''/* ����N�ʌ�2 */
        strSQL = strSQL & "slb_fault_n_n3,"     ''/* ����N�ʌ�3 */
        
'        strSQL = strSQL & "slb_fault_cd_bs_s,"      ''/* ��������BSCD */
'        strSQL = strSQL & "slb_fault_cd_bm_s,"      ''/* ��������BMCD */
'        strSQL = strSQL & "slb_fault_cd_bn_s,"      ''/* ��������BNCD */
'        strSQL = strSQL & "slb_fault_bs_s,"     ''/* ��������BS��� */
'        strSQL = strSQL & "slb_fault_bm_s,"     ''/* ��������BM��� */
'        strSQL = strSQL & "slb_fault_bn_s,"     ''/* ��������BN��� */
'        strSQL = strSQL & "slb_fault_bs_n,"     ''/* ��������BS�� */
'        strSQL = strSQL & "slb_fault_bm_n,"     ''/* ��������BM�� */
'        strSQL = strSQL & "slb_fault_bn_n,"     ''/* ��������BN�� */
'
'        strSQL = strSQL & "slb_fault_cd_ts_s,"      ''/* ��������TSCD */
'        strSQL = strSQL & "slb_fault_cd_tm_s,"      ''/* ��������TMCD */
'        strSQL = strSQL & "slb_fault_cd_tn_s,"      ''/* ��������TNCD */
'        strSQL = strSQL & "slb_fault_ts_s,"     ''/* ��������TS��� */
'        strSQL = strSQL & "slb_fault_tm_s,"     ''/* ��������TM��� */
'        strSQL = strSQL & "slb_fault_tn_s,"     ''/* ��������TN��� */
'        strSQL = strSQL & "slb_fault_ts_n,"     ''/* ��������TS�� */
'        strSQL = strSQL & "slb_fault_tm_n,"     ''/* ��������TM�� */
'        strSQL = strSQL & "slb_fault_tn_n,"     ''/* ��������TN�� */
        
        strSQL = strSQL & "slb_fault_e_judg,"       ''/* ����E�ʔ��� */
        strSQL = strSQL & "slb_fault_w_judg,"       ''/* ����W�ʔ��� */
        strSQL = strSQL & "slb_fault_s_judg,"       ''/* ����S�ʔ��� */
        strSQL = strSQL & "slb_fault_n_judg,"       ''/* ����N�ʔ��� */
'        strSQL = strSQL & "slb_fault_b_judg,"       ''/* ����B�ʔ��� */
'        strSQL = strSQL & "slb_fault_t_judg,"       ''/* ����T�ʔ��� */
        strSQL = strSQL & "slb_fault_u_judg,"       ''/* ����U�ʔ��� */
        strSQL = strSQL & "slb_fault_d_judg,"       ''/* ����D�ʔ��� */
        
        strSQL = strSQL & "slb_wrt_nme,"        ''/* �������� */
        
        strSQL = strSQL & "host_send,"          ''/* �r�W�R�����M���� */
        strSQL = strSQL & "host_wrt_dte,"       ''/* �r�W�R���o�^�� */
        strSQL = strSQL & "host_wrt_tme,"       ''/* �r�W�R���o�^���� */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","     ''/* �װ�� */
        
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & APResData.slb_ccno & "'" & ","      ''/* �X���uCCNO */
        strSQL = strSQL & "'" & APResData.slb_zkai_dte & "'" & ","      ''/* ����� */
        strSQL = strSQL & "'" & APResData.slb_ksh & "'" & ","       ''/* �|�� */
        strSQL = strSQL & "'" & APResData.slb_typ & "'" & ","       ''/* �^ */
        strSQL = strSQL & "'" & APResData.slb_uksk & "'" & ","      ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wei & "'" & ","       ''/* �d�� */
        strSQL = strSQL & "'" & APResData.slb_lngth & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_wdth & "'" & ","      ''/* �� */
        strSQL = strSQL & "'" & APResData.slb_thkns & "'" & ","     ''/* ���� */
        strSQL = strSQL & "'" & APResData.slb_nxt_prcs & "'" & ","      ''/* ���H�� */
        strSQL = strSQL & "'" & APResData.slb_cmt1 & "'" & ","      ''/* �R�����g1 */
        strSQL = strSQL & "'" & APResData.slb_cmt2 & "'" & ","      ''/* �R�����g2 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s1 & "'" & ","     ''/* ����E��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s2 & "'" & ","     ''/* ����E��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_e_s3 & "'" & ","     ''/* ����E��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s1 & "'" & ","        ''/* ����E�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s2 & "'" & ","        ''/* ����E�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_s3 & "'" & ","        ''/* ����E�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n1 & "'" & ","        ''/* ����E�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n2 & "'" & ","        ''/* ����E�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_e_n3 & "'" & ","        ''/* ����E�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s1 & "'" & ","     ''/* ����W��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s2 & "'" & ","     ''/* ����W��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_w_s3 & "'" & ","     ''/* ����W��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s1 & "'" & ","        ''/* ����W�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s2 & "'" & ","        ''/* ����W�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_s3 & "'" & ","        ''/* ����W�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n1 & "'" & ","        ''/* ����W�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n2 & "'" & ","        ''/* ����W�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_w_n3 & "'" & ","        ''/* ����W�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s1 & "'" & ","     ''/* ����S��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s2 & "'" & ","     ''/* ����S��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_s_s3 & "'" & ","     ''/* ����S��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s1 & "'" & ","        ''/* ����S�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s2 & "'" & ","        ''/* ����S�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_s3 & "'" & ","        ''/* ����S�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n1 & "'" & ","        ''/* ����S�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n2 & "'" & ","        ''/* ����S�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_s_n3 & "'" & ","        ''/* ����S�ʌ�3 */
        
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s1 & "'" & ","     ''/* ����N��CD1 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s2 & "'" & ","     ''/* ����N��CD2 */
        strSQL = strSQL & "'" & APResData.slb_fault_cd_n_s3 & "'" & ","     ''/* ����N��CD3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s1 & "'" & ","        ''/* ����N�ʎ��1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s2 & "'" & ","        ''/* ����N�ʎ��2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_s3 & "'" & ","        ''/* ����N�ʎ��3 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n1 & "'" & ","        ''/* ����N�ʌ�1 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n2 & "'" & ","        ''/* ����N�ʌ�2 */
        strSQL = strSQL & "'" & APResData.slb_fault_n_n3 & "'" & ","        ''/* ����N�ʌ�3 */
        
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bs_s & "'" & ","     ''/* ��������BSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bm_s & "'" & ","     ''/* ��������BMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_bn_s & "'" & ","     ''/* ��������BNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_s & "'" & ","        ''/* ��������BS��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_s & "'" & ","        ''/* ��������BM��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_s & "'" & ","        ''/* ��������BN��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bs_n & "'" & ","        ''/* ��������BS�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bm_n & "'" & ","        ''/* ��������BM�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_bn_n & "'" & ","        ''/* ��������BN�� */
'
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_ts_s & "'" & ","     ''/* ��������TSCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tm_s & "'" & ","     ''/* ��������TMCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_cd_tn_s & "'" & ","     ''/* ��������TNCD */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_s & "'" & ","        ''/* ��������TS��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_s & "'" & ","        ''/* ��������TM��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_s & "'" & ","        ''/* ��������TN��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_ts_n & "'" & ","        ''/* ��������TS�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tm_n & "'" & ","        ''/* ��������TM�� */
'        strSQL = strSQL & "'" & APResData.slb_fault_tn_n & "'" & ","        ''/* ��������TN�� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_e_judg & "'" & ","      ''/* ����E�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_w_judg & "'" & ","      ''/* ����W�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_s_judg & "'" & ","      ''/* ����S�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_n_judg & "'" & ","      ''/* ����N�ʔ��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_b_judg & "'" & ","      ''/* ����B�ʔ��� */
'        strSQL = strSQL & "'" & APResData.slb_fault_t_judg & "'" & ","      ''/* ����T�ʔ��� */
        
        strSQL = strSQL & "'" & APResData.slb_fault_u_judg & "'" & ","      ''/* ����U�ʔ��� */
        strSQL = strSQL & "'" & APResData.slb_fault_d_judg & "'" & ","      ''/* ����D�ʔ��� */
        
        strSQL = strSQL & "'" & APResData.slb_wrt_nme & "'" & ","           ''/* �������� */
        
        strSQL = strSQL & "'" & APResData.fail_host_send & "'" & ","             ''/* �X���u�ُ�񍐁@�r�W�R�����M���� */
        strSQL = strSQL & "'" & APResData.fail_host_wrt_dte & "'" & ","          ''/* �X���u�ُ�񍐁@�r�W�R���o�^�� */
        strSQL = strSQL & "'" & APResData.fail_host_wrt_tme & "'" & ","          ''/* �X���u�ُ�񍐁@�r�W�R���o�^���� */
        
        strSQL = strSQL & "'" & APResData.fail_sys_wrt_dte & "'" & ","           ''/* �X���u�ُ�񍐁@�o�^�� */
        strSQL = strSQL & "'" & APResData.fail_sys_wrt_tme & "'" & ","           ''/* �X���u�ُ�񍐁@�o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0016_Write ����I��") '�K�C�_���X�\��

    TRTS0016_Write = True

    On Error GoTo 0
    Exit Function

TRTS0016_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0016_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0016_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0022��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : ���u���ʓ��͂̃f�[�^������
'
' ���l      : ���u���ʓ��̓f�[�^��������
'
Public Function TRTS0022_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0022_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0022_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0022_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0022 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        For nI = 0 To UBound(APDirResData) - 1

            If APDirResData(nI).res_sys_wrt_dte = "" Then
                APDirResData(nI).res_sys_wrt_dte = Format(Now, "YYYYMMDD")
                APDirResData(nI).res_sys_wrt_tme = Format(Now, "HHMMSS")
            End If

            '********** ���R�[�h�ǉ� **********
            strSQL = "INSERT INTO TRTS0022 ("
    '        '-------------------------------------
            strSQL = strSQL & "slb_no," ''/* �X���u�m�n */
            strSQL = strSQL & "slb_stat,"   ''/* ��� */
            strSQL = strSQL & "slb_col_cnt,"    ''/* �J���[�� */
            strSQL = strSQL & "res_no," ''/* ���єԍ� */
            strSQL = strSQL & "slb_chno,"   ''/* �X���u�`���[�WNO */
            strSQL = strSQL & "slb_aino,"   ''/* �X���u���� */
            strSQL = strSQL & "res_nme1,"   ''/* ���э���1 */
            strSQL = strSQL & "res_val1,"   ''/* ���ђl1 */
            strSQL = strSQL & "res_uni1,"   ''/* ���ђP��1 */
            strSQL = strSQL & "res_nme2,"   ''/* ���э���2 */
            strSQL = strSQL & "res_val2,"   ''/* ���ђl2 */
            strSQL = strSQL & "res_uni2,"   ''/* ���ђP��2 */
            strSQL = strSQL & "res_cmt1,"   ''/* �R�����g1 */
            strSQL = strSQL & "res_cmt2,"   ''/* �R�����g2 */
            strSQL = strSQL & "res_cmp_flg,"    ''/* ���u�����t���O */
            strSQL = strSQL & "res_aft_stat,"   ''/* ���u���� */
            strSQL = strSQL & "res_wrt_dte,"    ''/* ���͓� */
            strSQL = strSQL & "res_wrt_nme,"    ''/* ���͎Җ� */
            strSQL = strSQL & "host_send,"          ''/* �r�W�R�����M���� */
            strSQL = strSQL & "host_wrt_dte,"       ''/* �r�W�R���o�^�� */
            strSQL = strSQL & "host_wrt_tme,"       ''/* �r�W�R���o�^���� */
            strSQL = strSQL & "sys_wrt_dte,"    ''/* �o�^�� */
            strSQL = strSQL & "sys_wrt_tme,"    ''/* �o�^���� */
            strSQL = strSQL & "sys_rwrt_dte,"   ''/* �X�V�� */
            strSQL = strSQL & "sys_rwrt_tme,"   ''/* �X�V���� */
            strSQL = strSQL & "sys_acs_pros,"   ''/* �A�N�Z�X�v���Z�X�� */
            strSQL = strSQL & "sys_acs_enum"   ''/* �A�N�Z�X�Ј��m�n */
            
            '---
    
            strSQL = strSQL & ") VALUES ("
    
            strSQL = strSQL & "'" & APDirResData(nI).slb_no & "'" & ","        ''/* �X���u�m�n */
            strSQL = strSQL & "'" & APDirResData(nI).slb_stat & "'" & ","      ''/* ��� */
            
            strSQL = strSQL & "'" & Format(CInt(APDirResData(nI).slb_col_cnt), "00") & "'" & ","     ''/* �װ�� */
        
            strSQL = strSQL & "'" & APDirResData(nI).dir_no & "'" & "," ''/* ���єԍ� */
            strSQL = strSQL & "'" & APDirResData(nI).slb_chno & "'" & ","   ''/* �X���u�`���[�WNO */
            strSQL = strSQL & "'" & APDirResData(nI).slb_aino & "'" & ","   ''/* �X���u���� */
            strSQL = strSQL & "'" & APDirResData(nI).dir_nme1 & "'" & ","   ''/* ���э���1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_val1 & "'" & ","   ''/* ���ђl1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_uni1 & "'" & ","   ''/* ���ђP��1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_nme2 & "'" & ","   ''/* ���э���2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_val2 & "'" & ","   ''/* ���ђl2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_uni2 & "'" & ","   ''/* ���ђP��2 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_cmt1 & "'" & ","   ''/* �R�����g1 */
            strSQL = strSQL & "'" & APDirResData(nI).dir_cmt2 & "'" & ","   ''/* �R�����g2 */
            strSQL = strSQL & "'" & APDirResData(nI).res_cmp_flg & "'" & ","    ''/* ���u�����t���O */
            strSQL = strSQL & "'" & APDirResData(nI).res_aft_stat & "'" & ","   ''/* ���u���� */
            strSQL = strSQL & "'" & APDirResData(nI).res_wrt_dte & "'" & ","    ''/* ���͓� */
            strSQL = strSQL & "'" & APDirResData(nI).res_wrt_nme & "'" & ","    ''/* ���͎Җ� */
            strSQL = strSQL & "'" & APResData.fail_res_host_send & "'" & ","        ''/* �r�W�R�����M���� */
            strSQL = strSQL & "'" & APResData.fail_res_host_wrt_dte & "'" & ","     ''/* �r�W�R���o�^�� */
            strSQL = strSQL & "'" & APResData.fail_res_host_wrt_tme & "'" & ","     ''/* �r�W�R���o�^���� */
            strSQL = strSQL & "'" & APDirResData(nI).res_sys_wrt_dte & "'" & ","     ''/* �o�^�� */
            strSQL = strSQL & "'" & APDirResData(nI).res_sys_wrt_tme & "'" & ","    ''/* �o�^���� */
            strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","   ''/* �X�V�� */
            strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","   ''/* �X�V���� */
            strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","   ''/* �A�N�Z�X�v���Z�X�� */
            strSQL = strSQL & "'" & "" & "'" & ")"   ''/* �A�N�Z�X�Ј��m�n */
        
    '        '-------------------------------------
    '
            Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
            '-cn.Execute (strSQL)
            oDB.ExecuteSql (strSQL)
        
        Next nI
    
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0022_Write ����I��") '�K�C�_���X�\��

    TRTS0022_Write = True

    On Error GoTo 0
    Exit Function

TRTS0022_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0022_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0022_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : DBDirResData_Read����
'
' ������    : ARG1 - �X���u�ԍ�
'           : ARG2 - ���
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p����TRTS0020,TRTS0022�̃��R�[�h��Ǎ�
'
' ���l      : �J���[�`�F�b�N�ُ폈�u�w���f�[�^�Ǎ�
'           :COLORSYS
'
Public Function DBDirResData_Read(ByVal strSlb_No As String, ByVal strSlb_Stat As String, ByVal strSlb_Col_Cnt As String) As Boolean
'slb_chno          VARCHAR2(5)          /* �X���u�`���[�WNO */
'slb_aino          VARCHAR2(4)          /* �X���u���� */
'slb_stat          VARCHAR2(1)          /* ��� */
'slb_col_cnt       VARCHAR2(2)          /* �װ�� */
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    'Dim cn As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBDirResData_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        ReDim APDirResTmpData(1)
        APDirResTmpData(0).slb_no = strSlb_No
        APDirResTmpData(0).slb_stat = strSlb_Stat
        APDirResTmpData(0).slb_col_cnt = Format(CInt(strSlb_Col_Cnt), "00")
        APDirResTmpData(0).dir_no = "01"
        APDirResTmpData(0).dir_nme1 = "�w������1"
        APDirResTmpData(0).dir_val1 = "�w���l1"
        APDirResTmpData(0).dir_uni1 = "�w���P��1"
        APDirResTmpData(0).dir_nme2 = "�w������2"
        APDirResTmpData(0).dir_val2 = "�w���l2"
        APDirResTmpData(0).dir_uni2 = "�w���P��2"
        APDirResTmpData(0).dir_cmt1 = "�R�����g1"
        APDirResTmpData(0).dir_cmt2 = "�R�����g2"
        APDirResTmpData(0).dir_wrt_dte = "20080505"
        APDirResTmpData(0).dir_wrt_nme = "�w���Җ�"
        DBDirResData_Read = True
        Exit Function
    End If

    On Error GoTo DBDirResData_Read_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    'ODBC
    'Provider=MSDASQL.1;Password=U3AP;User ID=U3AP;Data Source=ORAM;Extended Properties="DSN=ORAM;UID=U3AP;PWD=U3AP;DBQ=ORAM;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;GDE=F;FRL=F;BAM=IfAllSuccessful;MTS=F;MDI=F;CSR=F;FWC=F;PFC=10;TLO=0;"
    '-cn.Open DBConnectStr(0)

    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    strSQL = "SELECT TRTS0020.*, "
    strSQL = strSQL & "TRTS0022.RES_CMP_FLG, "
    strSQL = strSQL & "TRTS0022.RES_AFT_STAT, "
    strSQL = strSQL & "TRTS0022.RES_WRT_DTE, "
    strSQL = strSQL & "TRTS0022.RES_WRT_NME, "
    strSQL = strSQL & "TRTS0022.SYS_WRT_DTE AS SYS_WRT_DTE22, "
    strSQL = strSQL & "TRTS0022.SYS_WRT_TME AS SYS_WRT_TME22 "
    strSQL = strSQL & "FROM TRTS0020 LEFT JOIN TRTS0022 ON (TRTS0020.DIR_NO = TRTS0022.RES_NO) "
    strSQL = strSQL & "AND (TRTS0020.SLB_COL_CNT = TRTS0022.SLB_COL_CNT) "
    strSQL = strSQL & "AND (TRTS0020.SLB_STAT = TRTS0022.SLB_STAT) "
    strSQL = strSQL & "AND (TRTS0020.SLB_NO = TRTS0022.SLB_NO) "
    strSQL = strSQL & "WHERE TRTS0020.slb_no='" & strSlb_No & "' AND TRTS0020.slb_stat='" & strSlb_Stat & "' AND TRTS0020.slb_col_cnt='" & Format(CInt(strSlb_Col_Cnt), "00") & "' "
    strSQL = strSQL & "ORDER BY TRTS0020.dir_no"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    '-rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    ReDim APDirResTmpData(0)
    Do While Not oDS.EOF
        APDirResTmpData(UBound(APDirResTmpData)).slb_no = IIf(IsNull(oDS.Fields("slb_no").Value), "", oDS.Fields("slb_no").Value)  '' �X���u�m�n
        APDirResTmpData(UBound(APDirResTmpData)).slb_stat = IIf(IsNull(oDS.Fields("slb_stat").Value), "", oDS.Fields("slb_stat").Value)  '' ���
        APDirResTmpData(UBound(APDirResTmpData)).slb_col_cnt = IIf(IsNull(oDS.Fields("slb_col_cnt").Value), "", oDS.Fields("slb_col_cnt").Value) '' �J���[��
        APDirResTmpData(UBound(APDirResTmpData)).dir_no = IIf(IsNull(oDS.Fields("dir_no").Value), "", oDS.Fields("dir_no").Value)  '' �w���ԍ�
        APDirResTmpData(UBound(APDirResTmpData)).slb_chno = IIf(IsNull(oDS.Fields("slb_chno").Value), "", oDS.Fields("slb_chno").Value)  '' �X���u�`���[�WNO
        APDirResTmpData(UBound(APDirResTmpData)).slb_aino = IIf(IsNull(oDS.Fields("slb_aino").Value), "", oDS.Fields("slb_aino").Value) '' �X���u����
        APDirResTmpData(UBound(APDirResTmpData)).dir_nme1 = IIf(IsNull(oDS.Fields("dir_nme1").Value), "", oDS.Fields("dir_nme1").Value)  '' �w������1
        APDirResTmpData(UBound(APDirResTmpData)).dir_val1 = IIf(IsNull(oDS.Fields("dir_val1").Value), "", oDS.Fields("dir_val1").Value)  '' �w���l1
        APDirResTmpData(UBound(APDirResTmpData)).dir_uni1 = IIf(IsNull(oDS.Fields("dir_uni1").Value), "", oDS.Fields("dir_uni1").Value)  '' �w���P��1
        APDirResTmpData(UBound(APDirResTmpData)).dir_nme2 = IIf(IsNull(oDS.Fields("dir_nme2").Value), "", oDS.Fields("dir_nme2").Value)  '' �w������2
        APDirResTmpData(UBound(APDirResTmpData)).dir_val2 = IIf(IsNull(oDS.Fields("dir_val2").Value), "", oDS.Fields("dir_val2").Value)  '' �w���l2
        APDirResTmpData(UBound(APDirResTmpData)).dir_uni2 = IIf(IsNull(oDS.Fields("dir_uni2").Value), "", oDS.Fields("dir_uni2").Value)  '' �w���P��2
        APDirResTmpData(UBound(APDirResTmpData)).dir_cmt1 = IIf(IsNull(oDS.Fields("dir_cmt1").Value), "", oDS.Fields("dir_cmt1").Value)  '' �R�����g1
        APDirResTmpData(UBound(APDirResTmpData)).dir_cmt2 = IIf(IsNull(oDS.Fields("dir_cmt2").Value), "", oDS.Fields("dir_cmt2").Value)  '' �R�����g2
        APDirResTmpData(UBound(APDirResTmpData)).dir_wrt_dte = IIf(IsNull(oDS.Fields("dir_wrt_dte").Value), "", oDS.Fields("dir_wrt_dte").Value) '' �w����
        APDirResTmpData(UBound(APDirResTmpData)).dir_wrt_nme = IIf(IsNull(oDS.Fields("dir_wrt_nme").Value), "", oDS.Fields("dir_wrt_nme").Value) '' �w���Җ�
    APDirResTmpData(UBound(APDirResTmpData)).dir_sys_wrt_dte = IIf(IsNull(oDS.Fields("sys_wrt_dte").Value), "", oDS.Fields("sys_wrt_dte").Value)            ''�o�^��
    APDirResTmpData(UBound(APDirResTmpData)).dir_sys_wrt_tme = IIf(IsNull(oDS.Fields("sys_wrt_tme").Value), "", oDS.Fields("sys_wrt_tme").Value)           ''�o�^����
        
    APDirResTmpData(UBound(APDirResTmpData)).res_cmp_flg = IIf(IsNull(oDS.Fields("res_cmp_flg").Value), "", oDS.Fields("res_cmp_flg").Value)           ''���u�����t���O 1:����
    APDirResTmpData(UBound(APDirResTmpData)).res_aft_stat = IIf(IsNull(oDS.Fields("res_aft_stat").Value), "", oDS.Fields("res_aft_stat").Value)          ''���u���� 1:�s�K���L��i����A�r�L��j
    APDirResTmpData(UBound(APDirResTmpData)).res_wrt_dte = IIf(IsNull(oDS.Fields("res_wrt_dte").Value), "", oDS.Fields("res_wrt_dte").Value)           ''���͓�
    APDirResTmpData(UBound(APDirResTmpData)).res_wrt_nme = IIf(IsNull(oDS.Fields("res_wrt_nme").Value), "", oDS.Fields("res_wrt_nme").Value)           ''���͎Җ�
    APDirResTmpData(UBound(APDirResTmpData)).res_sys_wrt_dte = IIf(IsNull(oDS.Fields("SYS_WRT_DTE22").Value), "", oDS.Fields("SYS_WRT_DTE22").Value)            ''�o�^��
    APDirResTmpData(UBound(APDirResTmpData)).res_sys_wrt_tme = IIf(IsNull(oDS.Fields("SYS_WRT_TME22").Value), "", oDS.Fields("SYS_WRT_TME22").Value)           ''�o�^����
        
        ReDim Preserve APDirResTmpData(UBound(APDirResTmpData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBDirResData_Read ����I��") '�K�C�_���X�\��

    DBDirResData_Read = True

    On Error GoTo 0
    Exit Function

DBDirResData_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "DBDirResData_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBDirResData_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0050��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �X���u���pSCANLOC���̏���
'
' ���l      : �X���u���pSCANLOC��񏑂�����
'
Public Function TRTS0050_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0050_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0050_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0050_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0050 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '�t�@�C�����쐬
        strDestination = conDefault_DEFINE_SCNDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SKIN" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_00.JPG"

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0050 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* �o�^�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0050_Write ����I��") '�K�C�_���X�\��

    TRTS0050_Write = True

    On Error GoTo 0
    Exit Function

TRTS0050_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0050_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0050_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0052��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �J���[�`�F�b�N�pSCANLOC���̏���
'
' ���l      : �J���[�`�F�b�N�pSCANLOC��񏑂�����
'
Public Function TRTS0052_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0052_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0052_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0052_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0052 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '�t�@�C�����쐬
        strDestination = conDefault_DEFINE_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0052 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        strSQL = strSQL & "slb_col_cnt,"       ''/* �J���[�� */
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","      ''/* �J���[�� */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* �o�^�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0052_Write ����I��") '�K�C�_���X�\��

    TRTS0052_Write = True

    On Error GoTo 0
    Exit Function

TRTS0052_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0052_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0052_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0054��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �X���u�ُ�񍐗pSCANLOC���̏���
'
' ���l      : �X���u�ُ�񍐗pSCANLOC��񏑂�����
'
Public Function TRTS0054_Write(ByVal bDeleteOnly As Boolean) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    Dim strDestination As String

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0054_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        TRTS0054_Write = True
        Exit Function
    End If

    On Error GoTo TRTS0054_Write_err

    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0054 WHERE slb_no=" & "'" & APResData.slb_no & "'" & " and " & "slb_stat=" & "'" & APResData.slb_stat & "'" & _
    " and " & "slb_col_cnt=" & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '�t�@�C�����쐬
        strDestination = conDefault_DEFINE_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0054 ("
'        '-------------------------------------
        strSQL = strSQL & "slb_no,"     ''/* �X���u�m�n */
        strSQL = strSQL & "slb_stat,"       ''/* ��� */
        strSQL = strSQL & "slb_col_cnt,"       ''/* �J���[�� */
        strSQL = strSQL & "slb_chno,"       ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "slb_aino,"       ''/* �X���u���� */
        strSQL = strSQL & "slb_scan_addr,"  ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "sys_wrt_dte,"        ''/* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        ''/* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       ''/* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       ''/* �X�V���� */
        strSQL = strSQL & "sys_acs_pros,"       ''/* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "sys_acs_enum"        ''/* �A�N�Z�X�Ј��m�n */
        '---

        strSQL = strSQL & ") VALUES ("

        strSQL = strSQL & "'" & APResData.slb_no & "'" & ","        ''/* �X���u�m�n */
        strSQL = strSQL & "'" & APResData.slb_stat & "'" & ","      ''/* ��� */
        strSQL = strSQL & "'" & Format(CInt(APResData.slb_col_cnt), "00") & "'" & ","      ''/* �J���[�� */
        strSQL = strSQL & "'" & APResData.slb_chno & "'" & ","      ''/* �X���u�`���[�WNO */
        strSQL = strSQL & "'" & APResData.slb_aino & "'" & ","      ''/* �X���u���� */
        strSQL = strSQL & "'" & strDestination & "'" & ","          ''/* SCAN�A�h���X */
        
        strSQL = strSQL & "'" & APResData.sys_wrt_dte & "'" & ","           ''/* �o�^�� */
        strSQL = strSQL & "'" & APResData.sys_wrt_tme & "'" & ","           ''/* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","         ''/* sys_rwrt_dte�X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","           ''/* sys_rwrt_tme�X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ","           ''/* sys_acs_pros�A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & "'" & "" & "'" & ")"                              ''/* sys_acs_enum�A�N�Z�X�Ј��m�n */

'        strSQL = strSQL & "'" & GetSyoGyoDate() & "'" & ","             'VARCHAR2(8)    /* �o�^�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
'        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
'        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
'        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
'        '-------------------------------------
'
        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    '-cn.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0054_Write ����I��") '�K�C�_���X�\��

    TRTS0054_Write = True

    On Error GoTo 0
    Exit Function

TRTS0054_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0054_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0054_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0060�Ǎ�����
'
' ������    :
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : TRTS0060�̃��R�[�h��Ǎ�
'
' ���l      : �X�^�b�t���}�X�^�Ǎ�
'           :COLORSYS
'
Public Function TRTS0060_Read() As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0060_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If UBound(APStaffData) < 1 Then
            ReDim APStaffData(3)
            APStaffData(0).inp_StaffName = "NAME1"
            APStaffData(1).inp_StaffName = "NAME2"
            APStaffData(2).inp_StaffName = "NAME3"
        End If
        TRTS0060_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0060_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0060.* From TRTS0060 ORDER BY TRTS0060.STAFF_NME", cn, adOpenStatic, adLockReadOnly
    
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset("SELECT TRTS0060.* From TRTS0060 ORDER BY TRTS0060.STAFF_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APStaffData(0)
    Do While Not oDS.EOF
        APStaffData(UBound(APStaffData)).inp_StaffName = oDS.Fields("Staff_Nme").Value
        ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Read ����I��") '�K�C�_���X�\��

    TRTS0060_Read = True

    On Error GoTo 0
    
    ''TRTS0060���W�X�g����������
    Call TRTS0060_Reg_Write
    
    Exit Function

TRTS0060_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0060_Read = False
    On Error GoTo 0
    
    ''TRTS0060���W�X�g���Ǎ�����
    Call TRTS0060_Reg_Read

End Function

' @(f)
'
' �@�\      : TRTS0060���W�X�g����������
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɏ���
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɏ���
'           :COLORSYS
'
Public Sub TRTS0060_Reg_Write()
    Dim nI As Integer
    
    ''���W�X�g���ɕۑ�
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPStaffDataCount", UBound(APStaffData)
    For nI = 1 To UBound(APStaffData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APStaffName" & CStr(nI), APStaffData(nI - 1).inp_StaffName
    Next nI

End Sub

' @(f)
'
' �@�\      : TRTS0060���W�X�g���Ǎ�����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'           :COLORSYS
'
Public Sub TRTS0060_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''���W�X�g������ǂݍ���
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPStaffDataCount", 0)
    If nCount = 0 Then
        '�Ј��}�X�^�ǂݍ��݃G���A������
        ReDim APStaffData(1)
        APStaffData(0).inp_StaffName = "guest"
    Else
        '���W�X�g����ǂݍ���
        ReDim APStaffData(0)
        For nI = 1 To nCount
            APStaffData(nI - 1).inp_StaffName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APStaffName" & CStr(nI), "")
            ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' �@�\      : TRTS0060��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'             ARG2 - �X�^�b�t��
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X�^�b�t����������
'
' ���l      : �X�^�b�t���}�X�^����
'           :COLORSYS
'
Public Function TRTS0060_Write(ByVal bDeleteOnly As Boolean, ByVal strStaffName As String) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0060_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If bDeleteOnly = False Then
            APStaffData(UBound(APStaffData)).inp_StaffName = Left(strStaffName, 32)
            ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        End If
        TRTS0060_Write = True
        Exit Function
    End If
    
    ''�f�[�^�ǉ��̏ꍇ��TRTS0060���W�X�g����������
    If bDeleteOnly = False Then
        APStaffData(UBound(APStaffData)).inp_StaffName = Left(strStaffName, 32)
        ReDim Preserve APStaffData(UBound(APStaffData) + 1)
        Call TRTS0060_Reg_Write
    End If
    
    On Error GoTo TRTS0060_Write_err
    
    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0060 WHERE Staff_Nme='" & strStaffName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0060 ("
        strSQL = strSQL & "Staff_Nme,"          'VARCHAR2(32)         /* �X�^�b�t�� */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* �X�V���� */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strStaffName & "'" & ","                'VARCHAR2(32)   /* �X�^�b�t�� */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �o�^�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0060", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Write ����I��") '�K�C�_���X�\��

    TRTS0060_Write = True

    On Error GoTo 0
    Exit Function

TRTS0060_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0060_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0060_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0062�Ǎ�����
'
' ������    :
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : TRTS0062�̃��R�[�h��Ǎ�
'
' ���l      : ���������}�X�^�Ǎ�
'           :COLORSYS
'
Public Function TRTS0062_Read() As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0062_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If UBound(APInspData) < 1 Then
            ReDim APInspData(3)
            APInspData(0).inp_InspName = "NAME1"
            APInspData(1).inp_InspName = "NAME2"
            APInspData(2).inp_InspName = "NAME3"
        End If
        TRTS0062_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0062_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0062.* From TRTS0062 ORDER BY TRTS0062.INSP_NME", cn, adOpenStatic, adLockReadOnly
    
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset("SELECT TRTS0062.* From TRTS0062 ORDER BY TRTS0062.INSP_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APInspData(0)
    Do While Not oDS.EOF
        APInspData(UBound(APInspData)).inp_InspName = oDS.Fields("Insp_Nme").Value
        ReDim Preserve APInspData(UBound(APInspData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Read ����I��") '�K�C�_���X�\��

    TRTS0062_Read = True

    On Error GoTo 0
    
    ''TRTS0062���W�X�g����������
    Call TRTS0062_Reg_Write
    
    Exit Function

TRTS0062_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0062_Read = False
    On Error GoTo 0
    
    ''TRTS0062���W�X�g���Ǎ�����
    Call TRTS0062_Reg_Read

End Function

' @(f)
'
' �@�\      : TRTS0062���W�X�g����������
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɏ���
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɏ���
'           :COLORSYS
'
Public Sub TRTS0062_Reg_Write()
    Dim nI As Integer
    
    ''���W�X�g���ɕۑ�
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPInspDataCount", UBound(APInspData)
    For nI = 1 To UBound(APInspData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APInspName" & CStr(nI), APInspData(nI - 1).inp_InspName
    Next nI

End Sub

' @(f)
'
' �@�\      : TRTS0062���W�X�g���Ǎ�����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'           :COLORSYS
'
Public Sub TRTS0062_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''���W�X�g������ǂݍ���
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPInspDataCount", 0)
    If nCount = 0 Then
        '�Ј��}�X�^�ǂݍ��݃G���A������
        ReDim APInspData(1)
        APInspData(0).inp_InspName = "guest"
    Else
        '���W�X�g����ǂݍ���
        ReDim APInspData(0)
        For nI = 1 To nCount
            APInspData(nI - 1).inp_InspName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APInspName" & CStr(nI), "")
            ReDim Preserve APInspData(UBound(APInspData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' �@�\      : TRTS0062��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'             ARG2 - ��������
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̌���������������
'
' ���l      : ���������}�X�^����
'           :COLORSYS
'
Public Function TRTS0062_Write(ByVal bDeleteOnly As Boolean, ByVal strInspName As String) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0062_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If bDeleteOnly = False Then
            APInspData(UBound(APInspData)).inp_InspName = Left(strInspName, 32)
            ReDim Preserve APInspData(UBound(APInspData) + 1)
        End If
        TRTS0062_Write = True
        Exit Function
    End If
    
    ''�f�[�^�ǉ��̏ꍇ��TRTS0062���W�X�g����������
    If bDeleteOnly = False Then
        APInspData(UBound(APInspData)).inp_InspName = Left(strInspName, 32)
        ReDim Preserve APInspData(UBound(APInspData) + 1)
        Call TRTS0062_Reg_Write
    End If
    
    On Error GoTo TRTS0062_Write_err
    
    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0062 WHERE Insp_Nme='" & strInspName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0062 ("
        strSQL = strSQL & "Insp_Nme,"          'VARCHAR2(32)         /* �X�^�b�t�� */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* �X�V���� */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strInspName & "'" & ","                'VARCHAR2(32)   /* �X�^�b�t�� */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �o�^�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0062", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Write ����I��") '�K�C�_���X�\��

    TRTS0062_Write = True

    On Error GoTo 0
    Exit Function

TRTS0062_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0062_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0062_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : TRTS0066�Ǎ�����
'
' ������    :
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : TRTS0066�̃��R�[�h��Ǎ�
'
' ���l      : ���͎ҏ��}�X�^�Ǎ�
'           :COLORSYS
'
Public Function TRTS0066_Read() As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0066_Read:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If UBound(APInpData) < 1 Then
            ReDim APInpData(3)
            APInpData(0).inp_InpName = "NAME1"
            APInpData(1).inp_InpName = "NAME2"
            APInpData(2).inp_InpName = "NAME3"
        End If
        TRTS0066_Read = True
        Exit Function
    End If
    
    On Error GoTo TRTS0066_Read_err
    
    nOpen = 0
    
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '-rs.Open "SELECT TRTS0066.* From TRTS0066 ORDER BY TRTS0066.INSP_NME", cn, adOpenStatic, adLockReadOnly
    
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset("SELECT TRTS0066.* From TRTS0066 ORDER BY TRTS0066.INP_NME", 0&)
    Debug.Print oDS.RecordCount

    ReDim APInpData(0)
    Do While Not oDS.EOF
        APInpData(UBound(APInpData)).inp_InpName = oDS.Fields("Inp_Nme").Value
        ReDim Preserve APInpData(UBound(APInpData) + 1)
        oDS.MoveNext
    Loop

    '-rs.Close
    '-cn.Close
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Read ����I��") '�K�C�_���X�\��

    TRTS0066_Read = True

    On Error GoTo 0
    
    ''TRTS0066���W�X�g����������
    Call TRTS0066_Reg_Write
    
    Exit Function

TRTS0066_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0066_Read = False
    On Error GoTo 0
    
    ''TRTS0066���W�X�g���Ǎ�����
    Call TRTS0066_Reg_Read

End Function

' @(f)
'
' �@�\      : TRTS0066���W�X�g����������
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɏ���
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɏ���
'           :COLORSYS
'
Public Sub TRTS0066_Reg_Write()
    Dim nI As Integer
    
    ''���W�X�g���ɕۑ�
    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nAPInpDataCount", UBound(APInpData)
    For nI = 1 To UBound(APInpData)
        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "APInpName" & CStr(nI), APInpData(nI - 1).inp_InpName
    Next nI

End Sub

' @(f)
'
' �@�\      : TRTS0066���W�X�g���Ǎ�����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'
' ���l      : �ێ����̃f�[�^�����W�X�g���ɓǍ�
'           :COLORSYS
'
Public Sub TRTS0066_Reg_Read()
    Dim nI As Integer
    Dim nCount As Integer
    
    ''���W�X�g������ǂݍ���
    nCount = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "nAPInpDataCount", 0)
    If nCount = 0 Then
        '�Ј��}�X�^�ǂݍ��݃G���A������
        ReDim APInpData(1)
        APInpData(0).inp_InpName = "guest"
    Else
        '���W�X�g����ǂݍ���
        ReDim APInpData(0)
        For nI = 1 To nCount
            APInpData(nI - 1).inp_InpName = GetSetting(conReg_APPNAME, conReg_APSYSCFG, "APInpName" & CStr(nI), "")
            ReDim Preserve APInpData(UBound(APInpData) + 1)
        Next nI
    End If
    
End Sub

' @(f)
'
' �@�\      : TRTS0066��������
'
' ������    : ARG1 - �폜�̂ݎ��s�t���O
'             ARG2 - ���͎Җ�
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̓��͎Җ���������
'
' ���l      : ���͎Җ��}�X�^����
'           :COLORSYS
'
Public Function TRTS0066_Write(ByVal bDeleteOnly As Boolean, ByVal strInpName As String) As Boolean
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "TRTS0066_Write:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        If bDeleteOnly = False Then
            APInpData(UBound(APInpData)).inp_InpName = Left(strInpName, 32)
            ReDim Preserve APInpData(UBound(APInpData) + 1)
        End If
        TRTS0066_Write = True
        Exit Function
    End If
    
    ''�f�[�^�ǉ��̏ꍇ��TRTS0066���W�X�g����������
    If bDeleteOnly = False Then
        APInpData(UBound(APInpData)).inp_InpName = Left(strInpName, 32)
        ReDim Preserve APInpData(UBound(APInpData) + 1)
        Call TRTS0066_Reg_Write
    End If
    
    On Error GoTo TRTS0066_Write_err
    
    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    '********** ���R�[�h�폜 **********
    strSQL = "DELETE From TRTS0066 WHERE Inp_Nme='" & strInpName & "'"
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '-cn.Execute (strSQL)
    oDB.ExecuteSql (strSQL)

    If bDeleteOnly = False Then

        '********** ���R�[�h�ǉ� **********
        strSQL = "INSERT INTO TRTS0066 ("
        strSQL = strSQL & "Inp_Nme,"          'VARCHAR2(32)         /* �X�^�b�t�� */
        strSQL = strSQL & "sys_wrt_dte,"        'VARCHAR2(8)          /* �o�^�� */
        strSQL = strSQL & "sys_wrt_tme,"        'VARCHAR2(6)          /* �o�^���� */
        strSQL = strSQL & "sys_rwrt_dte,"       'VARCHAR2(8)          /* �X�V�� */
        strSQL = strSQL & "sys_rwrt_tme,"       'VARCHAR2(6)          /* �X�V���� */
        strSQL = strSQL & "sys_acs_pros"        'VARCHAR2(32)         /* �A�N�Z�X�v���Z�X�� */
        strSQL = strSQL & ") VALUES ("
        strSQL = strSQL & "'" & strInpName & "'" & ","                'VARCHAR2(32)   /* �X�^�b�t�� */"
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �o�^�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �o�^���� */
        strSQL = strSQL & "'" & Format(Now, "YYYYMMDD") & "'" & ","     'VARCHAR2(8)    /* �X�V�� */
        strSQL = strSQL & "'" & Format(Now, "HHMMSS") & "'" & ","       'VARCHAR2(6)    /* �X�V���� */
        strSQL = strSQL & "'" & conDef_DB_ProcessName & "'" & ")"       'VARCHAR2(32)   /* �A�N�Z�X�v���Z�X�� */
        '-------------------------------------

        Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
        '-cn.Execute (strSQL)
        oDB.ExecuteSql (strSQL)
    End If

    ''rs.Open "TRTS0066", cn, adOpenStatic, adLockOptimistic, adCmdTable
    ''rs.Close

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    'cn.Close
    ''oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Write ����I��") '�K�C�_���X�\��

    TRTS0066_Write = True

    On Error GoTo 0
    Exit Function

TRTS0066_Write_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next

    Call MsgLog(conProcNum_MAIN, "TRTS0066_Write �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    TRTS0066_Write = False
    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : �f�ޓ����c�a�|NCHTAISL�Ǎ�����
'
' ������    : ARG1 - �X���u�ԍ�
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�ԍ����g�p����NCHTAISL�̃��R�[�h��Ǎ�
'
' ���l      :
'           :COLORSYS
'
Public Function SOZAI_NCHTAISL_Read(ByVal strSlb_No As String) As Boolean
    
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("SOZAI_DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read:�f�ޓ����c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        ReDim APSozaiTmpData(1)
        '**********************************************************'
        'nchtaisl
        APSozaiTmpData(0).slb_no = "123451234"      ''�X���uNO"
        APSozaiTmpData(0).slb_ksh = "ABCDEF"        ''�|��
        APSozaiTmpData(0).slb_uksk = "AB"          ''����i�M������j
        APSozaiTmpData(0).slb_lngth = "12345"       ''����
        APSozaiTmpData(0).slb_color_wei = "12345"   ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
        APSozaiTmpData(0).slb_typ = "ABC"           ''�^
        APSozaiTmpData(0).slb_skin_wei = "12345"    ''�d�ʁi���ޔ��p�F����d�ʁj
        APSozaiTmpData(0).slb_wdth = "1234"         ''��
        APSozaiTmpData(0).slb_thkns = "123.12"      ''����
        APSozaiTmpData(0).slb_zkai_dte = "20080101" ''������i����N�����j
        '**********************************************************'
'        'skjchjdt�e�[�u��
'        APSozaiTmpData(0).slb_chno = "12345"        ''�`���[�WNO
'        APSozaiTmpData(0).slb_ccno = "12345"        ''CCNO
        '**********************************************************'
        SOZAI_NCHTAISL_Read = True
        Exit Function
    End If

    On Error GoTo SOZAI_NCHTAISL_Read_err

    ReDim APSozaiTmpData(0)
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_SOZAI, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '2008/08/30 A.K NCHTAISL�ŐV�f�[�^���o�o�[�W�����i�V�X�e�����t���������f�[�^�������j
    'strSQL = "SELECT * FROM NCHTAISL WHERE slbno='" & strSlb_No & "'"
    
    strSQL = "SELECT * FROM "
    strSQL = strSQL & "(SELECT NCHTAISL.SLBNO,NCHTAISL.�|��,NCHTAISL.�M������,NCHTAISL.����,"
    strSQL = strSQL & "NCHTAISL.SEG�o���d��,NCHTAISL.�^,NCHTAISL.����d��,NCHTAISL.��,"
    strSQL = strSQL & "NCHTAISL.����,NCHTAISL.������t�N,NCHTAISL.������t��,NCHTAISL.������t��,"
    strSQL = strSQL & "((NCHTAISL.������t�N * 10000) + (NCHTAISL.������t�� * 100) + NCHTAISL.������t��) as nYYYYMMDD "
    strSQL = strSQL & "FROM NCHTAISL "
    strSQL = strSQL & "WHERE (((NCHTAISL.SLBNO)='" & strSlb_No & "'))) "
    strSQL = strSQL & "WHERE nYYYYMMDD <= '" & Format(Now, "YYYYMMDD") & "' "
    strSQL = strSQL & "ORDER BY nYYYYMMDD DESC"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    If Not oDS.EOF Then
        ReDim APSozaiTmpData(1)

        '**********************************************************'
        'nchtaisl
        APSozaiTmpData(0).slb_no = IIf(IsNull(oDS.Fields("slbno").Value), "", oDS.Fields("slbno").Value)                      ''�X���uNO"
        APSozaiTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("�|��").Value), "", oDS.Fields("�|��").Value)                         ''�|��
        APSozaiTmpData(0).slb_uksk = IIf(IsNull(oDS.Fields("�M������").Value), "", Left(oDS.Fields("�M������").Value, 2))       ''����i�M������j�ˍ��[����Q���Ɋۂ߂�B
        APSozaiTmpData(0).slb_lngth = IIf(IsNull(oDS.Fields("����").Value), "", oDS.Fields("����").Value)                       ''����
        APSozaiTmpData(0).slb_color_wei = IIf(IsNull(oDS.Fields("SEG�o���d��").Value), "", oDS.Fields("SEG�o���d��").Value)     ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
        APSozaiTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("�^").Value), "", oDS.Fields("�^").Value)                             ''�^
        APSozaiTmpData(0).slb_skin_wei = IIf(IsNull(oDS.Fields("����d��").Value), "", oDS.Fields("����d��").Value)            ''�d�ʁi���ޔ��p�F����d�ʁj
        APSozaiTmpData(0).slb_wdth = IIf(IsNull(oDS.Fields("��").Value), "", oDS.Fields("��").Value)                            ''��
        APSozaiTmpData(0).slb_thkns = IIf(IsNull(oDS.Fields("����").Value), "", oDS.Fields("����").Value)                       ''����
        APSozaiTmpData(0).slb_zkai_dte = IIf(IsNull(oDS.Fields("������t�N").Value), "0000", Format(oDS.Fields("������t�N").Value, "0000")) & _
                                         IIf(IsNull(oDS.Fields("������t��").Value), "00", Format(oDS.Fields("������t��").Value, "00")) & _
                                         IIf(IsNull(oDS.Fields("������t��").Value), "00", Format(oDS.Fields("������t��").Value, "00"))     ''������i����N�����j
        '**********************************************************'

        '���݂�XXX.XX�֊ۂ�
        If APSozaiTmpData(0).slb_thkns <> "" Then
            If IsNumeric(APSozaiTmpData(0).slb_thkns) Then
                APSozaiTmpData(0).slb_thkns = ToHalfAdjust(CDbl(APSozaiTmpData(0).slb_thkns), 2)
            End If
        End If

    End If

    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read ����I��") '�K�C�_���X�\��

    SOZAI_NCHTAISL_Read = True

    On Error GoTo 0
    Exit Function

SOZAI_NCHTAISL_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next
    
    Call MsgLog(conProcNum_MAIN, "SOZAI_NCHTAISL_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    SOZAI_NCHTAISL_Read = False

    On Error GoTo 0
End Function

' @(f)
'
' �@�\      : �f�ޓ����c�a�|SKJCHJDT�Ǎ�����
'
' ������    : ARG1 - �X���u�`���[�W�ԍ�
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �w��̃X���u�`���[�W�ԍ����g�p����SKJCHJDT�̃��R�[�h��Ǎ�
'
' ���l      :
'           :COLORSYS
'
Public Function SOZAI_SKJCHJDT_Read(ByVal strSlb_Chno As String) As Boolean
    
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim Errs1 As Errors
    Dim errLoop As Error
    Dim nI As Integer
    Dim StrTmp As String
    Dim strError As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean

    '�f�o�b�N���[�h��
    If IsDEBUG("SOZAI_DB_SKIP") Then
       Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read:�f�ޓ����c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
        
        ReDim APSozaiTmpData(1)
'        '**********************************************************'
'        'nchtaisl
'        APSozaiTmpData(0).slb_no = "123451234"      ''�X���uNO"
'        APSozaiTmpData(0).slb_ksh = "ABCDEF"        ''�|��
'        APSozaiTmpData(0).slb_uksk = "AB"          ''����i�M������j
'        APSozaiTmpData(0).slb_lngth = "12345"       ''����
'        APSozaiTmpData(0).slb_color_wei = "12345"   ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
'        APSozaiTmpData(0).slb_typ = "ABC"           ''�^
'        APSozaiTmpData(0).slb_skin_wei = "12345"    ''�d�ʁi���ޔ��p�F����d�ʁj
'        APSozaiTmpData(0).slb_wdth = "1234"         ''��
'        APSozaiTmpData(0).slb_thkns = "123.12"      ''����
'        APSozaiTmpData(0).slb_zkai_dte = "20080101" ''������i����N�����j
        '**********************************************************'
        'skjchjdt�e�[�u��
        APSozaiTmpData(0).slb_chno = "12345"        ''�`���[�WNO
        APSozaiTmpData(0).slb_ccno = "12345"        ''CCNO
        '**********************************************************'
        SOZAI_SKJCHJDT_Read = True
        Exit Function
    End If

    On Error GoTo SOZAI_SKJCHJDT_Read_err

    ReDim APSozaiTmpData(0)
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_SOZAI, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '2008/08/30 A.K SKJCHJDT�ŐV�f�[�^���o�o�[�W�����i�V�X�e�����t���������f�[�^�������j
    'strSQL = "SELECT * FROM SKJCHJDT WHERE chno='" & strSlb_Chno & "'"
    
    strSQL = "SELECT * FROM "
    strSQL = strSQL & "(SELECT SKJCHJDT.CHNO,SKJCHJDT.CCNO,"
    strSQL = strSQL & "SKJCHJDT.�|��,SKJCHJDT.�^,"
    strSQL = strSQL & "SKJCHJDT.LS����_1,SKJCHJDT.LS����_2,SKJCHJDT.LS����_3,"
    strSQL = strSQL & "((SKJCHJDT.LS����_1 * 10000) + (SKJCHJDT.LS����_2 * 100) + SKJCHJDT.LS����_3) as nYYYYMMDD "
    strSQL = strSQL & "FROM SKJCHJDT "
    strSQL = strSQL & "WHERE (((SKJCHJDT.CHNO)='" & strSlb_Chno & "'))) "
    strSQL = strSQL & "WHERE nYYYYMMDD <= '" & Format(Now, "YYYYMMDD") & "' "
    strSQL = strSQL & "ORDER BY nYYYYMMDD DESC"
    
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��

    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount

    If Not oDS.EOF Then
        ReDim APSozaiTmpData(1)

        '**********************************************************'
        'skjchjdt�e�[�u��
        APSozaiTmpData(0).slb_chno = IIf(IsNull(oDS.Fields("chno").Value), "", oDS.Fields("chno").Value)        ''�`���[�WNO
        APSozaiTmpData(0).slb_ccno = IIf(IsNull(oDS.Fields("ccno").Value), "", oDS.Fields("ccno").Value)        ''CCNO
        
        '2008/08/30 A.K NCHTAISL�ɊY�����R�[�h���Ȃ��ꍇ�͏�ʉ�ʂō̗p���鍀�ڂ��ꎞ�ۑ�
        APSozaiTmpData(0).slb_ksh = IIf(IsNull(oDS.Fields("�|��").Value), "", oDS.Fields("�|��").Value)        ''�|��
        APSozaiTmpData(0).slb_typ = IIf(IsNull(oDS.Fields("�^").Value), "", oDS.Fields("�^").Value)           ''�^
        '**********************************************************'

    End If

    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read ����I��") '�K�C�_���X�\��

    SOZAI_SKJCHJDT_Read = True

    On Error GoTo 0
    Exit Function

SOZAI_SKJCHJDT_Read_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    nI = 1

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    ' Enumerate Errors collection and display properties of
    ' each Error object.
    'Set Errs1 = oDB.Errors
    'For Each errLoop In Errs1
    '    With errLoop
    '        StrTmp = StrTmp & vbCrLf & "Error #" & nI & ":"
    '        StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
    '        StrTmp = StrTmp & vbCrLf & "   Description  " & .Description
    '        StrTmp = StrTmp & vbCrLf & "   Source       " & .Source
    '        nI = nI + 1
    '    End With
    'Next
    
    Call MsgLog(conProcNum_MAIN, "SOZAI_SKJCHJDT_Read �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��

    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    SOZAI_SKJCHJDT_Read = False

    On Error GoTo 0
End Function

Private Sub DBSkinSlbSearchReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APSearchTmpSlbData(0)

    bRet = ReadCSV(App.path & "\" & "DBSkinSlbSearchRead.csv", strItem(), strData())
    
    For nI = 0 To UBound(strData, 2) - 1
        APSearchTmpSlbData(nI).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())
                
        '���
        APSearchTmpSlbData(nI).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())

        '�|��
        APSearchTmpSlbData(nI).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())

        '�^
        APSearchTmpSlbData(nI).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())

        '����
        APSearchTmpSlbData(nI).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        '�����
        APSearchTmpSlbData(nI).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())

        '���ޔ����сi����L�^���j
        APSearchTmpSlbData(nI).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())

        '���ޔ��Ұ��
        If getItemDataCSV("bAPScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False
        End If

        '���ޔ�PDF
        If getItemDataCSV("bAPPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False
        End If

        ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)
    
    Next nI

End Sub

Private Sub DBColorSlbSearchReadCSV()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim strItem() As String
    Dim strData() As String

    ReDim APSearchTmpSlbData(0)

    bRet = ReadCSV(App.path & "\" & "DBColorSlbSearchRead.csv", strItem(), strData())

    For nI = 0 To UBound(strData, 2) - 1


        APSearchTmpSlbData(nI).slb_chno = getItemDataCSV("slb_chno", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_aino = getItemDataCSV("slb_aino", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).slb_no = getItemDataCSV("slb_no", nI + 1, strItem(), strData())

        '���
        APSearchTmpSlbData(nI).slb_stat = getItemDataCSV("slb_stat", nI + 1, strItem(), strData())

        '���װ��
        APSearchTmpSlbData(nI).slb_col_cnt = getItemDataCSV("slb_col_cnt", nI + 1, strItem(), strData())

        '�|��
        APSearchTmpSlbData(nI).slb_ksh = getItemDataCSV("slb_ksh", nI + 1, strItem(), strData())

        '�^
        APSearchTmpSlbData(nI).slb_typ = getItemDataCSV("slb_typ", nI + 1, strItem(), strData())

        '����
        APSearchTmpSlbData(nI).slb_uksk = getItemDataCSV("slb_uksk", nI + 1, strItem(), strData())

        '�����
        APSearchTmpSlbData(nI).slb_zkai_dte = getItemDataCSV("slb_zkai_dte", nI + 1, strItem(), strData())

        '�װ���сi����L�^���j
        APSearchTmpSlbData(nI).sys_wrt_dte = getItemDataCSV("sys_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).sys_wrt_tme = getItemDataCSV("sys_wrt_tme", nI + 1, strItem(), strData())

        '���r�W�R�����M����
        APSearchTmpSlbData(nI).host_send = getItemDataCSV("host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).host_wrt_dte = getItemDataCSV("host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).host_wrt_tme = getItemDataCSV("host_wrt_tme", nI + 1, strItem(), strData())

        '�װ�Ұ��
        If getItemDataCSV("bAPScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPScanInput = False
        End If

        '�װPDF
        If getItemDataCSV("bAPPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPPdfInput = False
        End If

'***********************************************************************
        '�ُ�񍐁i����L�^���j
        APSearchTmpSlbData(nI).fail_sys_wrt_dte = getItemDataCSV("fail_sys_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_sys_wrt_tme = getItemDataCSV("fail_sys_wrt_tme", nI + 1, strItem(), strData())

        '�ُ�񍐃r�W�R�����M����
        APSearchTmpSlbData(nI).fail_host_send = getItemDataCSV("fail_host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_host_wrt_dte = getItemDataCSV("fail_host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_host_wrt_tme = getItemDataCSV("fail_host_wrt_tme", nI + 1, strItem(), strData())

        '�ُ�Ұ��
        If getItemDataCSV("bAPFailScanInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPFailScanInput = True
        Else
            APSearchTmpSlbData(nI).bAPFailScanInput = False
        End If

        '�ُ�PDF
        If getItemDataCSV("bAPFailPdfInput", nI + 1, strItem(), strData()) = "TRUE" Then
            APSearchTmpSlbData(nI).bAPFailPdfInput = True
        Else
            APSearchTmpSlbData(nI).bAPFailPdfInput = False
        End If

'***********************************************************************
        'CCNO
        APSearchTmpSlbData(nI).slb_ccno = getItemDataCSV("slb_ccno", nI + 1, strItem(), strData())

        '�d�ʁi�װ�����p�FSEG�o���d�� sozai="slb_color_wei"�j
        APSearchTmpSlbData(nI).slb_wei = getItemDataCSV("slb_wei", nI + 1, strItem(), strData())

        '����
        APSearchTmpSlbData(nI).slb_lngth = getItemDataCSV("slb_lngth", nI + 1, strItem(), strData())

        '��
        APSearchTmpSlbData(nI).slb_wdth = getItemDataCSV("slb_wdth", nI + 1, strItem(), strData())

        '����
        APSearchTmpSlbData(nI).slb_thkns = getItemDataCSV("slb_thkns", nI + 1, strItem(), strData())

'***********************************************************************
        '���u�w��
        APSearchTmpSlbData(nI).fail_dir_sys_wrt_dte = getItemDataCSV("fail_dir_sys_wrt_dte", nI + 1, strItem(), strData())

'***********************************************************************
        '���u����
        APSearchTmpSlbData(nI).fail_res_sys_wrt_dte = getItemDataCSV("fail_res_sys_wrt_dte", nI + 1, strItem(), strData())

        '���u���ʊ����t���O
        APSearchTmpSlbData(nI).fail_res_cmp_flg = getItemDataCSV("fail_res_cmp_flg", nI + 1, strItem(), strData())

        '���u���ʃr�W�R�����M����
        APSearchTmpSlbData(nI).fail_res_host_send = getItemDataCSV("fail_res_host_send", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_res_host_wrt_dte = getItemDataCSV("fail_res_host_wrt_dte", nI + 1, strItem(), strData())
        APSearchTmpSlbData(nI).fail_res_host_wrt_tme = getItemDataCSV("fail_res_host_wrt_tme", nI + 1, strItem(), strData())

        ReDim Preserve APSearchTmpSlbData(UBound(APSearchTmpSlbData) + 1)

    Next nI

End Sub


' @(f)
'
' �@�\      : ��ԃL�[�ύX��f�[�^�c�a�m�F
'
' ������    : ARG1 - �����X���u�m���D
'
' �Ԃ�l    : True �f�[�^���^False �f�[�^�L
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      :
'
Public Function DBStatChgCheckSKIN(ByVal sSlbno As String, ByVal sSlbStat As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBStatChgCheckSKIN_err
    
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    ' �ύX��f�[�^�m�F�N�G��
    strSQL = "SELECT TRTS0012.SLB_NO "
    strSQL = strSQL & "FROM TRTS0012 "
    strSQL = strSQL & "WHERE TRTS0012.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0012.SLB_STAT = '" & sSlbStat & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
        ' �f�[�^��
        DBStatChgCheckSKIN = True
    Else
        ' �f�[�^�L
        DBStatChgCheckSKIN = False
    End If
    oDS.Close
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN ����I��") '�K�C�_���X�\��

    On Error GoTo 0

    Exit Function

DBStatChgCheckSKIN_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckSKIN �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBStatChgCheckSKIN = False

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : ��ԃL�[�ύX��f�[�^�c�a�m�F
'
' ������    : ARG1 - �����X���u�m���D
'
' �Ԃ�l    : 0 �f�[�^���^1 TRTS0012�f�[�^�L�^2 TRTS0020�f�[�^�L
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      :
'
Public Function DBStatChgCheckCOLOR(ByVal sSlbno As String, ByVal sSlbStat As String, ByVal sSlbStatNow As String, ByVal sSlbColCnt As String) As Integer
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBStatChgCheckCOLOR_err
    
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    ' �ύX��f�[�^�m�F�N�G��
    strSQL = "SELECT TRTS0014.SLB_NO "
    strSQL = strSQL & "FROM TRTS0014 "
    strSQL = strSQL & "WHERE TRTS0014.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0014.SLB_STAT = '" & sSlbStat & "' "
    strSQL = strSQL & "AND TRTS0014.SLB_COL_CNT = '" & sSlbColCnt & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
        ' �f�[�^��
        DBStatChgCheckCOLOR = 0
    Else
        ' �f�[�^�L
        DBStatChgCheckCOLOR = 1
    End If
    oDS.Close
    
    ' �w���f�[�^�m�F�N�G���E�w���f�[�^�����݂�����ύX���Ȃ�
    strSQL = "SELECT TRTS0020.SLB_NO "
    strSQL = strSQL & "FROM TRTS0020 "
    strSQL = strSQL & "WHERE TRTS0020.SLB_NO = '" & sSlbno & "' "
    strSQL = strSQL & "AND TRTS0020.SLB_STAT = '" & sSlbStatNow & "' "
    strSQL = strSQL & "AND TRTS0020.SLB_COL_CNT = '" & sSlbColCnt & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g�̍쐬
    Set oDS = oDB.CreateDynaset(strSQL, 0&)
    Debug.Print oDS.RecordCount
    If oDS.EOF = True And oDS.BOF = True Then
    Else
        ' �f�[�^�L
        DBStatChgCheckCOLOR = 2
    End If
    oDS.Close
    
    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR ����I��") '�K�C�_���X�\��

    On Error GoTo 0

    Exit Function

DBStatChgCheckCOLOR_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgCheckCOLOR �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBStatChgCheckCOLOR = 1

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : ��ԃL�[�ύX��f�[�^�c�a�m�F
'
' ������    : ARG1 - �����X���u�m���D
'
' �Ԃ�l    : True �f�[�^���^False �f�[�^�L
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      :
'
Public Function DBStatChgFixSKIN(ByVal sSlbno As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatNow As String, ByVal sStatNew As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim sScanAddr As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBStatChgFixSKIN_err
    
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    ' TRTS0012 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0012 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0040 DELETE *****************************************************************************
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    strSQL = "DELETE FROM TRTS0040 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0040 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    sScanAddr = conDefault_DEFINE_PDFDIR & "\SKIN" & "\" & sChno & "\" & sAino & _
                                           "\SKIN" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_00.PDF"
    strSQL = "UPDATE TRTS0040 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0050 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0050 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0050 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    sScanAddr = conDefault_DEFINE_SCNDIR & "\SKIN" & "\" & sChno & "\" & sAino & _
                                           "\SKIN" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_00.JPG"
    strSQL = "UPDATE TRTS0050 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN ����I��") '�K�C�_���X�\��
    DBStatChgFixSKIN = True

    On Error GoTo 0

    Exit Function

DBStatChgFixSKIN_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixSKIN �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        oSess.RollbackTrans
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBStatChgFixSKIN = False

    On Error GoTo 0

End Function

' @(f)
'
' �@�\      : ��ԃL�[�ύX��f�[�^�c�a�m�F
'
' ������    : ARG1 - �����X���u�m���D
'
' �Ԃ�l    : True �f�[�^���^False �f�[�^�L
'
' �@�\����  : �w��̃X���u�ԍ����g�p���ăX���u������������
'
' ���l      :
'
Public Function DBStatChgFixCOLOR(ByVal sSlbno As String, ByVal sChno As String, ByVal sAino As String, ByVal sStatNow As String, ByVal sStatNew As String, ByVal sColCntOld As String, ByVal sColCntNew As String) As Boolean
    ' ADO�̃I�u�W�F�N�g�ϐ���錾����
    Dim oSess As Object     '�I���N���Z�b�V�����I�u�W�F�N�g
    Dim oDB As Object       '�I���N���f�[�^�x�[�X�I�u�W�F�N�g
    Dim oDS As Object       '�I���N���_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sId As String       '���[�U��
    Dim sPass As String     '�p�X���[�h
    Dim sHost As String     '�z�X�g�ڑ�������
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim StrTmp As String
    Dim strSQL As String
    Dim sScanAddr As String
    Dim nOpen As Integer
    Dim bRet As Boolean
    
    '�f�o�b�N���[�h��
    If IsDEBUG("DB_SKIP") Then
        Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR:�c�a�X�L�b�v���[�h�ł��B") '�K�C�_���X�\��
            
        Exit Function
    End If
    
    Call MsgLog(conProcNum_MAIN, "�c�a�I�����C�����[�h�ł��B") '�K�C�_���X�\��
    
    On Error GoTo DBStatChgFixCOLOR_err
    
    nOpen = 0

    ' Oracle�Ƃ̐ڑ����m������
    bRet = DBConnectStr(conDefault_DBConnect_MYUSER, sHost, sId, sPass)

    '�I���N���Z�b�V�����I�u�W�F�N�g�̍쐬
    Set oSess = CreateObject("OracleInProcServer.XOraSession")
    nOpen = 1

    '�I���N���f�[�^�x�[�X�I�u�W�F�N�g�̍쐬
    Set oDB = oSess.OpenDatabase(sHost, sId & "/" & sPass, 0&)
    nOpen = 2

    '�Z�b�V�����g�����U�N�V�����J�n
    oSess.BeginTrans

    ' TRTS0014 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0014 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0016 UPDATE *****************************************************************************
    strSQL = "UPDATE TRTS0016 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0020 DELETE *****************************************************************************
'    strSQL = "DELETE FROM TRTS0020 "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0020 UPDATE *****************************************************************************
'    strSQL = "UPDATE TRTS0020 "
'    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0022 DELETE *****************************************************************************
'    strSQL = "DELETE FROM TRTS0022 "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
'    oDB.ExecuteSql (strSQL)
    
    ' TRTS0022 UPDATE *****************************************************************************
'    strSQL = "UPDATE TRTS0022 "
'    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "' "
'    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
'    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
'    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
'    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
'    oDB.ExecuteSql (strSQL)

    ' TRTS0042 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0042 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0042 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    '$PDFDIR\SKIN\12345\1234\SKIN_12345_1234_0_00.PDF
    sScanAddr = conDefault_DEFINE_PDFDIR & "\COLOR" & "\" & sChno & "\" & sAino & _
                                           "\COLOR" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".PDF"
    strSQL = "UPDATE TRTS0042 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0044 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0044 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0044 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    sScanAddr = conDefault_DEFINE_PDFDIR & "\SLBFAIL" & "\" & sChno & "\" & sAino & _
                                           "\SLBFAIL" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".PDF"
    strSQL = "UPDATE TRTS0044 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_pdf_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0052 DELETE *****************************************************************************
    strSQL = "DELETE FROM TRTS0052 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    ' TRTS0052 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    sScanAddr = conDefault_DEFINE_SCNDIR & "\COLOR" & "\" & sChno & "\" & sAino & _
                                           "\COLOR" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".JPG"
    strSQL = "UPDATE TRTS0052 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0054 DELETE *****************************************************************************
    strSQL = "DELETE TRTS0054 "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNew & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntNew & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)
    
    ' TRTS0054 UPDATE *****************************************************************************
    '�Q�ƃf�B���N�g���E�t�@�C�����쐬
    sScanAddr = conDefault_DEFINE_SCNDIR & "\SLBFAIL" & "\" & sChno & "\" & sAino & _
                                           "\SLBFAIL" & "_" & sChno & "_" & sAino & "_" & sStatNew & "_" & sColCntNew & ".JPG"
    strSQL = "UPDATE TRTS0054 "
    strSQL = strSQL & "SET slb_stat = '" & sStatNew & "', "
    strSQL = strSQL & "slb_col_cnt = '" & sColCntNew & "', "
    strSQL = strSQL & "slb_scan_addr = '" & sScanAddr & "', "
    strSQL = strSQL & "sys_rwrt_dte = '" & Format(Now, "YYYYMMDD") & "', "
    strSQL = strSQL & "sys_rwrt_tme = '" & Format(Now, "HHMMSS") & "', "
    strSQL = strSQL & "sys_acs_pros = '" & conDef_DB_ProcessName & "' "
    strSQL = strSQL & "WHERE slb_no = '" & sSlbno & "' "
    strSQL = strSQL & "AND slb_stat = '" & sStatNow & "' "
    strSQL = strSQL & "AND slb_col_cnt = '" & sColCntOld & "' "
    Call MsgLog(conProcNum_MAIN, "SQL[" & strSQL & "]") '�K�C�_���X�\��
    oDB.ExecuteSql (strSQL)

    '�Z�b�V�����g�����U�N�V�����R�~�b�g
    oSess.CommitTrans

    Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    nOpen = 0

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR ����I��") '�K�C�_���X�\��
    DBStatChgFixCOLOR = True

    On Error GoTo 0

    Exit Function

DBStatChgFixCOLOR_err:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next

    StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(errNum)
    StrTmp = StrTmp & vbCrLf & "   Generated by " & errSrc
    StrTmp = StrTmp & vbCrLf & "   Description  " & errDesc

    Call MsgLog(conProcNum_MAIN, "DBStatChgFixCOLOR �ُ�I��") '�K�C�_���X�\��
    Call MsgLog(conProcNum_MAIN, StrTmp) '�K�C�_���X�\��
    
    If nOpen >= 2 Then
        Set oDB = Nothing    '�f�[�^�x�[�X�I�u�W�F�N�g�����
    End If
    If nOpen >= 1 Then
        Set oSess = Nothing  '�Z�b�V�����I�u�W�F�N�g�����
    End If

    DBStatChgFixCOLOR = False

    On Error GoTo 0

End Function

