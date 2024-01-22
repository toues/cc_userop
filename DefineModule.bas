Attribute VB_Name = "DefineModule"
' @(h) DefineModule.Bas                ver 1.00 ( '08 SEC Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�V�X�e�����C���i�ϐ��^�萔�j��`���W���[��
' �@�{���W���[���̓V�X�e���Ŏg�p����ϐ��^�萔���`����
' �@���߂̂��̂ł���B

Option Explicit

''�v���Z�X�ԍ�
Public Const conProcNum_MAIN As Integer = 1 ''���C������
Public Const conProcNum_BSCONT As Integer = 2 ''�r�W�R���ʐM����
Public Const conProcNum_TRCONT As Integer = 3 ''�\�P�b�g�ʐM����
Public Const conProcNum_MAINTENANCE As Integer = 4 ''�����e�i���X����
'Public Const conProcNum_SELPARSLB As Integer = 5 ''�e�X���u�I������
Public Const conProcNum_WINSOCKCONT As Integer = 6 ''Winsock �C�x���gLOG

''�O���b�g�R���g���[����`
Public Const FlexAlignCenter As Long = 2 ''�\���ʒu�Z���^�[

''�R�[���o�b�N
Public Const CALLBACK_ncResOK = 1 ''�����n�j
Public Const CALLBACK_ncResCANCEL = 2 ''�����L�����Z��
Public Const CALLBACK_ncResSKIP = 3 ''�����X�L�b�v
Public Const CALLBACK_ncResEXTEND = 4 ''�����g��

Public Const CALLBACK_MAIN_SHUTDOWN = 1 ''���C���|�V�X�e���I��
Public Const CALLBACK_MAIN_RETSKINSCANWND = 2 ''���C���|�X���u����������
Public Const CALLBACK_MAIN_RETCOLORSCANWND1 = 3 ''���C���|�װ���������\���� -> �װ���������\���́|�X���u�I���֖߂�
Public Const CALLBACK_MAIN_RETCOLORSCANWND2 = 4 ''���C���|�װ���������\���� -> �װ���������\���́|�ُ�񍐈ꗗ�֖߂�
Public Const CALLBACK_MAIN_RETSLBFAILSCANWND1 = 5 ''���C���|�X���u�ُ�񍐏����� -> �װ���������\���� -> �װ���������\���́|�X���u�I���֖߂�
Public Const CALLBACK_MAIN_RETSLBFAILSCANWND2 = 6 ''���C���|�X���u�ُ�񍐏����� -> �װ���������\���� -> �װ���������\���́|�ُ�񍐈ꗗ�֖߂�

Public Const CALLBACK_MAIN_RETSKINSLBSELWND = 7 ''���C���|�X���u���������́|�X���u�I��
Public Const CALLBACK_MAIN_RETCOLORSLBSELWND = 8 ''���C���|�װ���������\���́|�X���u�I��
Public Const CALLBACK_MAIN_RETCOLORSLBFAILWND = 9 ''���C���|�װ���������\���́|�ُ�񍐈ꗗ

Public Const CALLBACK_MAIN_RETSYSCFGWND = 10 ''���C���|�V�X�e���ݒ�
Public Const CALLBACK_MAIN_RETDIRRESWND1 = 11 ''�X���u�ُ폈�u�w���^���ʓ��� -> �װ���������\���́|�X���u�I���֖߂�
Public Const CALLBACK_MAIN_RETDIRRESWND2 = 12 ''�X���u�ُ폈�u�w���^���ʓ��� -> �װ���������\���́|�ُ�񍐈ꗗ�֖߂�

Public Const CALLBACK_RES_DBSNDDATA_SKIN = 20 ''SKIN�|�c�a�o�^
Public Const CALLBACK_RES_SKINDATA_DBDEL_REQ = 21 ''SKIN�|�f�[�^�폜�⍇��

Public Const CALLBACK_RES_DBSNDDATA_COLOR = 22 ''COLOR�|�c�a�o�^
Public Const CALLBACK_RES_DBSNDDATA_SLBFAIL = 23 ''SLBFAIL�|�c�a�o�^
Public Const CALLBACK_RES_DBSNDDATA_DIRRES = 24 ''DIRRES�|�c�a�o�^
Public Const CALLBACK_RES_HOSTSNDDATA_DIRRES = 25 ''DIRRES�|HOST�o�^
Public Const CALLBACK_RES_COLORDATA_DBDEL_REQ = 26 ''COLOR�|�f�[�^�폜�⍇��
Public Const CALLBACK_RES_COLORDATA_HOSTDEL_REQ = 27 ''COLOR�|�f�[�^�폜�⍇���i�r�W�R���폜�V�i���I�j
Public Const CALLBACK_RES_COLORDATA_HOSTDEL_REQ2 = 28 ''COLOR�|�f�[�^�폜�⍇��(�r�W�R���폜�V�i���I�˂c�a�폜�V�i���I�j

Public Const CALLBACK_RES_DIRPRN_REQ = 30 '�w������⍇�� 2008/09/04
Public Const CALLBACK_RES_DIRPRN_SND = 31 '�w������v������ 2008/09/04

Public Const CALLBACK_RES_STATECHANGE_SKIN = 50  'SKIN�|��ԕύX�⍇�� 2009/01/28
Public Const CALLBACK_RES_STATECHANGE_COLOR = 50 'SKIN�|��ԕύX�⍇�� 2009/01/28

Public Const CALLBACK_OPREGWND = 100 ''������o�^
Public Const CALLBACK_NEXTPROCWND = 101 ''���H���o�^
Public Const CALLBACK_FULLSCANIMAGEWND = 102 ''�t���X�L�����C���[�W�\�����
Public Const CALLBACK_PHOTOIMGUPWND = 103 ''�ʐ^�Y�t���
Public Const CALLBACK_PHOTOIMG_DELETE = 104 '�폜'
Public Const CALLBACK_PHOTOIMG_UPLOAD = 105 '�A�b�v���[�h

Public Const CALLBACK_HOSTSEND = 110 ''�z�X�g���M
'Public Const CALLBACK_HOSTSEND_RESDELETE = 111 ''�z�X�g���M�|���э폜
'Public Const CALLBACK_HOSTSEND_SLBDELETE = 112 ''�z�X�g���M�|�X���u�폜
'Public Const CALLBACK_HOSTSEND_SLBDELETE2 = 113 ''�z�X�g���M�|�X���u�폜�Q
Public Const CALLBACK_HOSTSEND_QUERY = 114 ''�z�X�g���M�|�X���u���₢���킹

Public Const CALLBACK_TRSEND = 115 ''TR���M

Public Const CALLBACK_USEIMGDATA = 200 ''�C���[�W�f�[�^�g�p�i�c�a���݁j
Public Const CALLBACK_GETIMGDATA = 201 ''�C���[�W�f�[�^�擾�i�Ǎ��j

''���W�X�g���p��`
''�A�v���P�[�V������
Public Const conReg_APPNAME As String = "COLORSYS" ''�{�V�X�e����

''�Z�N�V������
Public Const conReg_APSYSCFG As String = "SYSCFG DATA" ''�V�X�e���ݒ���
'Public Const conReg_APSLB As String = "SLB DATA" ''�X���u���
Public Const conReg_APRESULT As String = "RESULT DATA" ''���ѓ��͏��

''���W�X�g�������l

''�f�o�b�O�p
Public Const conDefault_DEBUG_MODE As Integer = 1 ''�f�o�b�O�n�m

''DB
Public Const conDefault_DBConnect_MYUSER As Integer = 0
Public Const conDefault_DBConnect_MYCOMN As Integer = 1
Public Const conDefault_DBConnect_SOZAI As Integer = 2

Public Const conDefault_DB_MYUSER_DSN As String = "ORAM_COL" ''�f�[�^�\�[�X��
Public Const conDefault_DB_MYUSER_UID As String = "UCOL" ''���[�U�[�h�c
Public Const conDefault_DB_MYUSER_PWD As String = "UCOL" ''�p�X���[�h
'Public Const conDefault_DB_MYUSER_UID As String = "UCOLTEST" ''�e�X�g�@���[�U�[�h�c
'Public Const conDefault_DB_MYUSER_PWD As String = "UCOLTEST" ''�e�X�g�@�p�X���[�h

Public Const conDefault_DB_MYCOMN_DSN As String = "ORAM_COL" ''�f�[�^�\�[�X��
Public Const conDefault_DB_MYCOMN_UID As String = "NYKCOMN" ''���[�U�[�h�c
Public Const conDefault_DB_MYCOMN_PWD As String = "NYKCOMN" ''�p�X���[�h
'Public Const conDefault_DB_MYCOMN_UID As String = "NYKCOMNTEST" ''�e�X�g�@���[�U�[�h�c
'Public Const conDefault_DB_MYCOMN_PWD As String = "NYKCOMNTEST" ''�e�X�g�@�p�X���[�h

Public Const conDefault_DB_SOZAI_DSN As String = "ORAM_SOZAI" ''�f�[�^�\�[�X��
Public Const conDefault_DB_SOZAI_UID As String = "JISSEKI1" ''���[�U�[�h�c
Public Const conDefault_DB_SOZAI_PWD As String = "JISSEKI1" ''�p�X���[�h

Public Const conDefault_SHARES_SCNDIR As String = "\\COLDBSRV\shares\SCAN" ''�X�L���i�[�C���[�W�t�@�C���ۑ���p�X��
Public Const conDefault_SHARES_IMGDIR As String = "\\COLDBSRV\shares\IMG" ''�ʐ^�C���[�W�t�@�C���ۑ���p�X��
Public Const conDefault_SHARES_PDFDIR As String = "\\COLDBSRV\shares\PDF" '' 20090124 add by M.Aoyagi    PDF�t�@�C���ۑ���p�X��

Public Const conDefault_DEFINE_SCNDIR As String = "$SCNDIR"
Public Const conDefault_DEFINE_PDFDIR As String = "$PDFDIR"               '' 20090124 add by M.Aoyagi    ��ԕύX���g�p

Public Const conDefault_PHOTOIMG_DIR As String = "c:"       ''�ʐ^���[�J���p�X
Public Const conDefault_PHOTOIMG_DELCHK As Integer = 0      ''�ʐ^�R�s�[���폜�t���O
Public Const conDefault_PHOTOIMG_ALLFILES As Integer = 0    ''�ʐ^�S�Ẵt�@�C���w��t���O

'-------------------------------------
'Public Const conDefault_TRN_MSG_NO As String = "" ''�g�����U�N�V�������b�Z�[�W�ԍ�
'-------------------------------------
'2008-04-28 TRTS0012 Define
Public Const conDefault_slb_no As String = ""
Public Const conDefault_slb_chno As String = ""
Public Const conDefault_slb_aino As String = ""
Public Const conDefault_slb_stat As String = ""

Public Const conDefault_slb_col_cnt As String = ""

Public Const conDefault_slb_ccno As String = ""
Public Const conDefault_slb_zkai_dte As String = ""
Public Const conDefault_slb_ksh As String = ""
Public Const conDefault_slb_typ As String = ""
Public Const conDefault_slb_uksk As String = ""
Public Const conDefault_slb_wei As String = ""
Public Const conDefault_slb_lngth As String = ""
Public Const conDefault_slb_wdth As String = ""
Public Const conDefault_slb_thkns As String = ""
Public Const conDefault_slb_nxt_prcs As String = ""
Public Const conDefault_slb_cmt1 As String = ""
Public Const conDefault_slb_cmt2 As String = ""

Public Const conDefault_slb_fault_cd_e_s1 As String = ""
Public Const conDefault_slb_fault_cd_e_s2 As String = ""
Public Const conDefault_slb_fault_cd_e_s3 As String = ""
Public Const conDefault_slb_fault_e_s1 As String = ""
Public Const conDefault_slb_fault_e_s2 As String = ""
Public Const conDefault_slb_fault_e_s3 As String = ""
Public Const conDefault_slb_fault_e_n1 As String = ""
Public Const conDefault_slb_fault_e_n2 As String = ""
Public Const conDefault_slb_fault_e_n3 As String = ""

Public Const conDefault_slb_fault_cd_w_s1 As String = ""
Public Const conDefault_slb_fault_cd_w_s2 As String = ""
Public Const conDefault_slb_fault_cd_w_s3 As String = ""
Public Const conDefault_slb_fault_w_s1 As String = ""
Public Const conDefault_slb_fault_w_s2 As String = ""
Public Const conDefault_slb_fault_w_s3 As String = ""
Public Const conDefault_slb_fault_w_n1 As String = ""
Public Const conDefault_slb_fault_w_n2 As String = ""
Public Const conDefault_slb_fault_w_n3 As String = ""

Public Const conDefault_slb_fault_cd_s_s1 As String = ""
Public Const conDefault_slb_fault_cd_s_s2 As String = ""
Public Const conDefault_slb_fault_cd_s_s3 As String = ""
Public Const conDefault_slb_fault_s_s1 As String = ""
Public Const conDefault_slb_fault_s_s2 As String = ""
Public Const conDefault_slb_fault_s_s3 As String = ""
Public Const conDefault_slb_fault_s_n1 As String = ""
Public Const conDefault_slb_fault_s_n2 As String = ""
Public Const conDefault_slb_fault_s_n3 As String = ""

Public Const conDefault_slb_fault_cd_n_s1 As String = ""
Public Const conDefault_slb_fault_cd_n_s2 As String = ""
Public Const conDefault_slb_fault_cd_n_s3 As String = ""
Public Const conDefault_slb_fault_n_s1 As String = ""
Public Const conDefault_slb_fault_n_s2 As String = ""
Public Const conDefault_slb_fault_n_s3 As String = ""
Public Const conDefault_slb_fault_n_n1 As String = ""
Public Const conDefault_slb_fault_n_n2 As String = ""
Public Const conDefault_slb_fault_n_n3 As String = ""

Public Const conDefault_slb_fault_cd_bs_s As String = ""
Public Const conDefault_slb_fault_cd_bm_s As String = ""
Public Const conDefault_slb_fault_cd_bn_s As String = ""
Public Const conDefault_slb_fault_bs_s As String = ""
Public Const conDefault_slb_fault_bm_s As String = ""
Public Const conDefault_slb_fault_bn_s As String = ""
Public Const conDefault_slb_fault_bs_n As String = ""
Public Const conDefault_slb_fault_bm_n As String = ""
Public Const conDefault_slb_fault_bn_n As String = ""

Public Const conDefault_slb_fault_cd_ts_s As String = ""
Public Const conDefault_slb_fault_cd_tm_s As String = ""
Public Const conDefault_slb_fault_cd_tn_s As String = ""
Public Const conDefault_slb_fault_ts_s As String = ""
Public Const conDefault_slb_fault_tm_s As String = ""
Public Const conDefault_slb_fault_tn_s As String = ""
Public Const conDefault_slb_fault_ts_n As String = ""
Public Const conDefault_slb_fault_tm_n As String = ""
Public Const conDefault_slb_fault_tn_n As String = ""

Public Const conDefault_slb_wrt_nme As String = ""

Public Const conDefault_host_send As String = ""
Public Const conDefault_host_wrt_dte As String = ""
Public Const conDefault_host_wrt_tme As String = ""

Public Const conDefault_sys_wrt_dte As String = ""
Public Const conDefault_sys_wrt_tme As String = ""
Public Const conDefault_sys_rwrt_dte As String = ""
Public Const conDefault_sys_rwrt_tme As String = ""
Public Const conDefault_sys_acs_pros As String = ""
Public Const conDefault_sys_acs_enum As String = ""

'Public Const conDefault_LINE_NAME As String = "col"
'Public Const conDefault_LINE_NUMBER As String = ""

Public Const conDefault_nSEARCH_MAX0 As Integer = 9999 ''�X���u�m���D
'Public Const conDefault_nSEARCH_MAX1 As Integer = 1 ''����
'Public Const conDefault_nSEARCH_MAX2 As Integer = 10 ''���߉ߋ�
'Public Const conDefault_nSEARCH_MAX3 As Integer = 1 ''��������
'Public Const conDefault_nSEARCH_RANGE As Integer = 90 ''�����L���͈́@�ߋ��H��

'Public Const conDefault_HOST_NAME As String = "QVCB89" ''�z�X�g����
Public Const conDefault_HOST_IP As String = "172.18.192.19" ''�z�X�g IP
Public Const conDefault_nHOST_PORT As String = "15025" ''�z�X�gPort
Public Const conDefault_nHOST_TOUT0 As Long = 35 ''�z�X�g�ʐM�^�C���A�E�g�i�S�́j�i�b�j
Public Const conDefault_nHOST_TOUT1 As Long = 5 ''�z�X�g�ʐM�^�C���A�E�g�i�I�[�v�����j�i�b�j
Public Const conDefault_nHOST_TOUT2 As Long = 10 ''�z�X�g�ʐM�^�C���A�E�g�i�f�[�^�ʐM�j�i�b�j
Public Const conDefault_nHOST_RETRY As Integer = 2 ''�z�X�g�ʐM���g���C��

Public Const conDefault_TR_IP As String = "172.18.128.254" ''�ʐM�T�[�o�[�h�o�A�h���X
Public Const conDefault_nTR_PORT As Integer = 15032 ''�ʐM�T�[�o�[�|�[�g�ԍ�
Public Const conDefault_nTR_TOUT0 As Long = 35 ''�ʐM�T�[�o�[�^�C���A�E�g�i�S�́j�i�b�j
Public Const conDefault_nTR_TOUT1 As Long = 5 ''�ʐM�T�[�o�[�^�C���A�E�g�i�I�[�v�����j�i�b�j
Public Const conDefault_nTR_TOUT2 As Long = 10 ''�ʐM�T�[�o�[�^�C���A�E�g�i�f�[�^�ʐM�j�i�b�j
Public Const conDefault_nTR_RETRY As Integer = 2 ''�ʐM�T�[�o�[���g���C��

Public Const conDefault_nIMAGE_SIZE0 As Integer = 30 ''�C���[�W�\����
Public Const conDefault_nIMAGE_SIZE1 As Integer = 30 ''�C���[�W�\����
Public Const conDefault_nIMAGE_SIZE2 As Integer = 30 ''�C���[�W�\����
Public Const conDefault_nIMAGE_ROTATE0 As Integer = 90 ''�C���[�W��]
Public Const conDefault_nIMAGE_ROTATE1 As Integer = 90 ''�C���[�W��]
Public Const conDefault_nIMAGE_ROTATE2 As Integer = 90 ''�C���[�W��]

'DEMO���ݒ�
Public Const conDefault_nIMAGE_DEB_LEFT0 As Integer = 0 ''�C���[�W0�����W�i�f���j
Public Const conDefault_nIMAGE_DEB_TOP0 As Integer = 0 ''�C���[�W0����W�i�f���j
Public Const conDefault_nIMAGE_DEB_WIDTH0 As Integer = 3467 ''�C���[�W0���i�f���j
Public Const conDefault_nIMAGE_DEB_HEIGHT0 As Integer = 2475 ''�C���[�W0�����i�f���j

Public Const conDefault_nIMAGE_DEB_LEFT1 As Integer = 0 ''�C���[�W1�����W�i�f���j
Public Const conDefault_nIMAGE_DEB_TOP1 As Integer = 0 ''�C���[�W1����W�i�f���j
Public Const conDefault_nIMAGE_DEB_WIDTH1 As Integer = 3467 ''�C���[�W1���i�f���j
Public Const conDefault_nIMAGE_DEB_HEIGHT1 As Integer = 2475 ''�C���[�W1�����i�f���j

Public Const conDefault_nIMAGE_DEB_LEFT2 As Integer = 0 ''�C���[�W2�����W�i�f���j
Public Const conDefault_nIMAGE_DEB_TOP2 As Integer = 0 ''�C���[�W2����W�i�f���j
Public Const conDefault_nIMAGE_DEB_WIDTH2 As Integer = 3467 ''�C���[�W2���i�f���j
Public Const conDefault_nIMAGE_DEB_HEIGHT2 As Integer = 2475 ''�C���[�W2�����i�f���j

'�{�Ԑݒ�
Public Const conDefault_nIMAGE_LEFT0 As Integer = 0 ''�C���[�W0�����W�i�{�ԁj
Public Const conDefault_nIMAGE_TOP0 As Integer = 0 ''�C���[�W0����W�i�{�ԁj
Public Const conDefault_nIMAGE_WIDTH0 As Integer = 3467 ''�C���[�W0���i�{�ԁj
Public Const conDefault_nIMAGE_HEIGHT0 As Integer = 2475 ''�C���[�W0�����i�{�ԁj

Public Const conDefault_nIMAGE_LEFT1 As Integer = 0 ''�C���[�W1�����W�i�{�ԁj
Public Const conDefault_nIMAGE_TOP1 As Integer = 0 ''�C���[�W1����W�i�{�ԁj
Public Const conDefault_nIMAGE_WIDTH1 As Integer = 3467 ''�C���[�W1���i�{�ԁj
Public Const conDefault_nIMAGE_HEIGHT1 As Integer = 2475 ''�C���[�W1�����i�{�ԁj

Public Const conDefault_nIMAGE_LEFT2 As Integer = 0 ''�C���[�W2�����W�i�{�ԁj
Public Const conDefault_nIMAGE_TOP2 As Integer = 0 ''�C���[�W2����W�i�{�ԁj
Public Const conDefault_nIMAGE_WIDTH2 As Integer = 3467 ''�C���[�W2���i�{�ԁj
Public Const conDefault_nIMAGE_HEIGHT2 As Integer = 2475 ''�C���[�W2�����i�{�ԁj

Public Const conAccessLevel_Users As Integer = 0 ''�A�N�Z�X���x���i���[�U�[�j
Public Const conAccessLevel_Administrators As Integer = 1 ''�A�N�Z�X���x���i�Ǘ��j


Public Const conDefault_Separator As String = ":" ''���O�p��؂蕶��

Public Const conDefine_lGuidanceListMAX As Long = 1000 ''�K�C�_���X�\���@���X�g�ő匏��

Public Const conDefine_ImageDirName As String = "TEMP" ''�C���[�W�t�@�C���i�[�t�H���_
Public Const conDefine_LogDirName As String = "LOGS" ''�k�n�f�t�@�C���i�[�t�H���_

Public Const conDefine_SYSMODE_SKIN As Integer = 0
Public Const conDefine_SYSMODE_COLOR As Integer = 1
Public Const conDefine_SYSMODE_SLBFAIL As Integer = 2

Public Const conDefine_ColorActive As Long = &H80000005 ''���[�U�[��`�i�E�C���h�̔w�i�j
Public Const conDefine_ColorNotActive As Long = &H80000013 ''��A�N�e�B�u�^�C�g�������F
Public Const conDefine_ColorBKLostFocus As Long = &H80000005 ''���[�U�[��`�i�E�C���h�̔w�i�j
Public Const conDefine_ColorBKGotFocus As Long = &HFFFF& ''�w�i���F
Public Const conDefine_Color_ForColor_HOST_ERROR As Long = &HFF& ''�ʐM�G���[�@�����F��

Public MainLogFileNumber As Variant ''���O�t�@�C���p �t�@�C���ԍ�

'2008/09/03 �J���[���ʈꗗ��WEB-URL
Public Const conDefault_WEBURL_Color_Result As String = "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://COLDBSRV/CC/jsp/JumpRsltList.jsp?sCall1=sky&sCall2=sky"

'2015/09/15 ���|�J���[���ʈꗗ��WEB-URL
Public Const conDefault_WEBURL_Color_Result_Tok As String = "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://COLDBSRV/CC/jsp/JumpTokRsltList.jsp?sCall1=sky&sCall2=sky"

''�V�X�e�����
Public Type typAPSysCfgData
    nDEBUG_MODE As Integer ''�f�o�b�N���[�h
    nDISP_DEBUG As Integer ''��ʃf�o�b�N�\��
    nFILE_DEBUG As Integer ''LOG�t�@�C���f�o�b�N�\��
    nHOSTDATA_DEBUG As Integer ''���уf�[�^�o�^�z�X�g�ʐM�f�o�b�N���[�h�i�߂�l�𖄂ߍ��݂܂��B�j
    nTR_SKIP As Integer ''�ʐM�T�[�o�[�n�X�L�b�v
    nHOSTDATA_SKIP As Integer ''���уf�[�^�o�^�z�X�g�ʐM�n�X�L�b�v
    nDB_SKIP As Integer ''�c�a�X�L�b�v
    nSOZAI_DB_SKIP As Integer ''�f�ޓ����c�a�X�L�b�v
    nSCAN_SKIP As Integer ''�X�L���i�[�n�X�L�b�v
    
    DB_MYUSER_DSN As String  ''�f�[�^�\�[�X��
    DB_MYUSER_UID As String ''���[�U�[�h�c
    DB_MYUSER_PWD As String ''�p�X���[�h
    DB_MYCOMN_DSN As String  ''�f�[�^�\�[�X��
    DB_MYCOMN_UID As String ''���[�U�[�h�c
    DB_MYCOMN_PWD As String ''�p�X���[�h
    DB_SOZAI_DSN As String  ''�f�[�^�\�[�X��
    DB_SOZAI_UID As String ''���[�U�[�h�c
    DB_SOZAI_PWD As String ''�p�X���[�h
    
    SHARES_SCNDIR As String ''�X�L���i�[�C���[�W�ۑ���p�X
    SHARES_IMGDIR As String ''�ʐ^�C���[�W�ۑ���p�X
    SHARES_PDFDIR As String ''20090116 add by M.Aoyagi PDF�ۑ���p�X
    
    PHOTOIMG_DIR As String          ''�ʐ^���[�J���p�X
    PHOTOIMG_DELCHK As Integer      ''�ʐ^�R�s�[���폜�t���O
    PHOTOIMG_ALLFILES As Integer    ''�ʐ^�S�Ẵt�@�C���w��t���O
    
'    TRN_MSG_NO As String ''2004-12-01 �g�����U�N�V�������b�Z�[�W�ԍ�
    
'    nUSE_OFFICE As Integer ''�������ݒu 0:OFF 1:ON
'    nSEARCH_MAX(0 To 3) As Integer ''���������ݒ�
    
'    nSEARCH_RANGE As Integer ''�����L���͈́@�ߋ��H��
    
   ' �\�P�b�g�ʐM�Ή�
    'HOST_NAME As String ''�z�X�g����
    'nHOST_TOUT(0 To 1) As Integer ''�ʐM�^�C���A�E�g (0)=ALL (1)=IVT
    HOST_IP As String ''�r�W�R���h�o�A�h���X
    nHOST_PORT As Integer ''�r�W�R���|�[�g�ԍ�
    nHOST_TOUT(0 To 2) As Long ''�ʐM�^�C���A�E�g (0)=ALL (1)=OPEN���@(2)=�f�[�^�ʐM��
    nHOST_RETRY As Integer ''�ʐM���g���C��
    
'    nUSE_EMAIL As Integer ''�d�qҰق��g�p���Ĵװ�ʒm 0:OFF 1:ON
    
    TR_IP As String ''�T�[�o�[�h�o�A�h���X
    nTR_PORT As Integer ''�|�[�g�ԍ�
    'nTR_TOUT(0) As Integer ''�ʐM�^�C���A�E�g (0)=ALL
    nTR_TOUT(0 To 2) As Long ''�ʐM�^�C���A�E�g (0)=ALL (1)=OPEN���@(2)=�f�[�^�ʐM��
    nTR_RETRY As Integer ''�ʐM���g���C��
    
'    SMTP As String ''���MҰ� (SMTP) ���ް
'    AP_EMAIL As String ''�װ�ʒm�p�@�d�qҰ� ���ڽ
'    USER_EMAIL(1 To 20) As String ''�װ�ʒm�� �d�qҰ� ���ڽ
    nIMAGE_SIZE(0 To 2) As Integer ''�C���[�W�\���T�C�Y 10,20,30,40,50,60,70,80,90,100
    nIMAGE_ROTATE(0 To 2) As Integer ''�X�L���i�Ǎ�����]�@0,90,180,270
    nIMAGE_LEFT(0 To 2) As Integer ''�؂�o���C���[�W�����W�i�o�����������j
    nIMAGE_TOP(0 To 2) As Integer ''�؂�o���C���[�W����W�i�o�����������j
    nIMAGE_WIDTH(0 To 2) As Integer ''�؂�o���C���[�W���i�o�����������j
    nIMAGE_HEIGHT(0 To 2) As Integer ''�؂�o���C���[�W�����i�o�����������j
    
'    NowLineName As String ''���ݑI�𒆂̃��C����
'    NowLineNumber As String ''���ݑI�𒆂̃��C���ԍ�
'    NowLineType As String ''2002-07-11 ���ݑI�𒆂̃��C���^�C�v
    
'    nLineNumberCount As Integer ''���C���ԍ����X�g�J�E���g
'    LineNumber() As String ''���C���ԍ�
'    LineType() As String ''2002-07-11 ���C���^�C�v
    
'    NowStaffNumber As String ''���ݑI�𒆂̎Ј��ԍ�
    
    NowStaffName(0 To 2) As String ''���ݑI�𒆂̎����i�ێ��p�j
    NowNextProcess(0 To 2) As String ''���ݑI�𒆂̎��H���i�ێ��p�j
    
    WEBURL_Color_Result As String ''�J���[���ʈꗗ��WEB-URL
    WEBURL_Color_Result_Tok As String ''���|�J���[���ʈꗗ��WEB-URL
    
    'NowOperator As String ''���ݑI�𒆂̑����
'    NowGroup As String ''���ݑI�𒆂̔�
    'nOperatorCount As Integer ''����������X�g�J�E���g
    'Operator() As String ''����������X�g
'    nGroupCount As Integer ''������i�ǁj���X�g�J�E���g
'    Group() As String ''������i�ǁj���X�g
    'nStaffCount As Integer ''�Ј����X�g�J�E���g
    'nStaffAccessLevel() As Integer ''�Ј��A�N�Z�X���x��
    'StaffNumber() As String ''�Ј��ԍ�
    'StaffName() As String ''�Ј�����
End Type

'''�V�X�e���R���g���[���f�[�^
'Public Type typAPSysCont
'    bNewEntry As Boolean ''True:�V�K False:�C��
'End Type

''�X���u���R���g���[���f�[�^
Public Type typAPSlbCont
    bProcessing As Boolean ''�X���u�I�����b�N�p�������t���O
    strSearchInputSlbNumber As String ''�����X���u�m���D
    nSearchInputModeSelectedIndex As Integer ''�����I�v�V�����i���̓��[�h�j�w��C���f�b�N�X�ԍ�
    nSearchInputStatusSelectedIndex As Integer ''�����I�v�V�����i��ԓ��́j�w��C���f�b�N�X�ԍ�
    nListSelectedIndexP1 As Integer ''�X���u���X�g�w��C���f�b�N�X+1�ԍ� 0�͖��w��
'    nChildSelectedIndexP1 As Integer ''�q�X���u�w��C���f�b�N�X+1�ԍ� 0�͖��w��
End Type

''�X���u���
Public Type typAPSlbData
''''------------���f�[�^
    '�������X�g
    bWorkSelected As Boolean    ''���[�N�p
    slb_no As String            ''�X���u�m���D
    slb_chno As String          ''�X���u�`���[�W�m���D
    slb_aino As String          ''�X���u����
    slb_stat As String          ''���
    slb_zkai_dte As String      ''�����
    slb_ksh As String           ''�|��
    slb_typ As String           ''�^
    slb_uksk As String          ''����
    sys_wrt_dte As String       ''�L�^���i����L�^���j
    
    '*********************************************
    '�J���[�`�F�b�N
    slb_ccno As String          ''CCNO
    slb_wei As String           ''�d��
    slb_lngth As String         ''����
    slb_wdth As String          ''��
    slb_thkns As String         ''����
    
    slb_col_cnt As String       ''�װ��
    host_send As String         ''�r�W�R�����M ���ʃt���O
    host_wrt_dte As String      ''�r�W�R�����M �L�^��
    host_wrt_tme As String      ''�r�W�R�����M �L�^����
    sys_wrt_tme As String       ''�L�^�����i����L�^�����j
    
    '�ُ�ꗗ���X�g�\����p '2008/09/04
    slb_fault_e_judg As String  ''����E�ʔ���
    slb_fault_w_judg As String  ''����W�ʔ���
    slb_fault_s_judg As String  ''����S�ʔ���
    slb_fault_n_judg As String  ''����N�ʔ���
    
    '*********************************************
    '�X���u�ُ�
    fail_host_send As String    ''�X���u�ُ�p�@�r�W�R�����M���ʃt���O
    fail_host_wrt_dte As String ''�X���u�ُ�p�@�r�W�R�����M �L�^��
    fail_host_wrt_tme As String ''�X���u�ُ�p�@�r�W�R�����M �L�^����
    fail_sys_wrt_dte As String  ''�X���u�ُ�p�@�L�^���i����L�^���j
    fail_sys_wrt_tme As String  ''�X���u�ُ�p�@�L�^�����i����L�^�����j
    '*********************************************
    '���u�w��
    fail_dir_sys_wrt_dte As String  ''���u�w���p�@�L�^���i����L�^���j
    fail_dir_prn_out_max As String  ''�w������ς݃t���O
    '*********************************************
    '���u����
    fail_res_sys_wrt_dte As String  ''���u���ʗp�@�L�^���i����L�^���j
    fail_res_cmp_flg As String      ''���u���ʗp�@�����t���O�i�S�́j
    fail_res_host_send As String    ''���u���ʗp�@�r�W�R�����M���ʃt���O
    fail_res_host_wrt_dte As String ''���u���ʗp�@�r�W�R�����M �L�^��
    fail_res_host_wrt_tme As String ''���u���ʗp�@�r�W�R�����M �L�^����
    '*********************************************
    
    bAPScanInput As Boolean ''SCAN�C���[�W�f�[�^�L��t���O
    bAPFailScanInput As Boolean ''�X���u�ُ�pSCAN�C���[�W�f�[�^�L��t���O
    
    bAPPdfInput As Boolean ''PDF�C���[�W�f�[�^�L��t���O
    sAPPdfInput_ReqDate As String
    bAPFailPdfInput As Boolean ''�X���u�ُ�pPDF�C���[�W�f�[�^�L��t���O
    sAPFailPdfInput_ReqDate As String
    
    PhotoImgCnt1 As String '' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
    PhotoImgCnt2 As String '' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
    
    '2016/04/20 - TAI - S
    slb_works_sky_tok As String         '��Ə�
    '2016/04/20 - TAI - E
End Type

''���уf�[�^�i�X���u���A�J���[�`�F�b�N���p�j
''COLOR
Public Type typAPResData
    slb_no As String            ''�X���uNO
    slb_chno As String          ''�X���u�`���[�WNO
    slb_aino As String          ''�X���u����
    slb_stat As String          ''���
    slb_col_cnt As String       ''�J���[��
    slb_ccno As String          ''�X���uCCNO
    slb_zkai_dte As String      ''�����
    slb_ksh As String           ''�|��
    slb_typ As String           ''�^
    slb_uksk As String          ''����
    slb_wei As String           ''�d��
    slb_lngth As String         ''����
    slb_wdth As String          ''��
    slb_thkns As String         ''����
    slb_nxt_prcs As String      ''���H��
    slb_cmt1 As String          ''�R�����g1
    slb_cmt2 As String          ''�R�����g2
    
    slb_fault_cd_e_s1 As String ''����E��CD1
    slb_fault_cd_e_s2 As String ''����E��CD2
    slb_fault_cd_e_s3 As String ''����E��CD3
    slb_fault_e_s1 As String    ''����E�ʎ��1
    slb_fault_e_s2 As String    ''����E�ʎ��2
    slb_fault_e_s3 As String    ''����E�ʎ��3
    slb_fault_e_n1 As String    ''����E�ʌ�1
    slb_fault_e_n2 As String    ''����E�ʌ�2
    slb_fault_e_n3 As String    ''����E�ʌ�3
    
    slb_fault_cd_w_s1 As String ''����W��CD1
    slb_fault_cd_w_s2 As String ''����W��CD2
    slb_fault_cd_w_s3 As String ''����W��CD3
    slb_fault_w_s1 As String    ''����W�ʎ��1
    slb_fault_w_s2 As String    ''����W�ʎ��2
    slb_fault_w_s3 As String    ''����W�ʎ��3
    slb_fault_w_n1 As String    ''����W�ʌ�1
    slb_fault_w_n2 As String    ''����W�ʌ�2
    slb_fault_w_n3 As String    ''����W�ʌ�3
    
    slb_fault_cd_s_s1 As String ''����S��CD1
    slb_fault_cd_s_s2 As String ''����S��CD2
    slb_fault_cd_s_s3 As String ''����S��CD3
    slb_fault_s_s1 As String    ''����S�ʎ��1
    slb_fault_s_s2 As String    ''����S�ʎ��2
    slb_fault_s_s3 As String    ''����S�ʎ��3
    slb_fault_s_n1 As String    ''����S�ʌ�1
    slb_fault_s_n2 As String    ''����S�ʌ�2
    slb_fault_s_n3 As String    ''����S�ʌ�3
    
    slb_fault_cd_n_s1 As String ''����N��CD1
    slb_fault_cd_n_s2 As String ''����N��CD2
    slb_fault_cd_n_s3 As String ''����N��CD3
    slb_fault_n_s1 As String    ''����N�ʎ��1
    slb_fault_n_s2 As String    ''����N�ʎ��2
    slb_fault_n_s3 As String    ''����N�ʎ��3
    slb_fault_n_n1 As String    ''����N�ʌ�1
    slb_fault_n_n2 As String    ''����N�ʌ�2
    slb_fault_n_n3 As String    ''����N�ʌ�3
    
    slb_fault_cd_bs_s As String ''��������BSCD
    slb_fault_cd_bm_s As String ''��������BMCD
    slb_fault_cd_bn_s As String ''��������BNCD
    slb_fault_bs_s As String    ''��������BS���
    slb_fault_bm_s As String    ''��������BM���
    slb_fault_bn_s As String    ''��������BN���
    slb_fault_bs_n As String    ''��������BS��
    slb_fault_bm_n As String    ''��������BM��
    slb_fault_bn_n As String    ''��������BN��
    
    slb_fault_cd_ts_s As String ''��������TSCD
    slb_fault_cd_tm_s As String ''��������TMCD
    slb_fault_cd_tn_s As String ''��������TNCD
    slb_fault_ts_s As String    ''��������TS���
    slb_fault_tm_s As String    ''��������TM���
    slb_fault_tn_s As String    ''��������TN���
    slb_fault_ts_n As String    ''��������TS��
    slb_fault_tm_n As String    ''��������TM��
    slb_fault_tn_n As String    ''��������TN��
    
    slb_fault_e_judg As String  ''����E�ʔ���
    slb_fault_w_judg As String  ''����W�ʔ���
    slb_fault_s_judg As String  ''����S�ʔ���
    slb_fault_n_judg As String  ''����N�ʔ���
    slb_fault_b_judg As String  ''����B�ʔ���
    slb_fault_t_judg As String  ''����T�ʔ���
    
    slb_fault_u_judg As String  ''����U�ʔ���
    slb_fault_d_judg As String  ''����D�ʔ���
    
    slb_wrt_nme As String       ''��������
    host_send As String         ''�r�W�R�����M����
    host_wrt_dte As String      ''�L�^��
    host_wrt_tme As String      ''�L�^����
    sys_wrt_dte As String       ''�o�^��
    sys_wrt_tme As String       ''�o�^����
    
    fail_host_send As String         ''�X���u�ُ�񍐗p�@�r�W�R�����M����
    fail_host_wrt_dte As String      ''�X���u�ُ�񍐗p�@�L�^��
    fail_host_wrt_tme As String      ''�X���u�ُ�񍐗p�@�L�^����
    fail_sys_wrt_dte As String       ''�X���u�ُ�񍐗p�@�o�^��
    fail_sys_wrt_tme As String       ''�X���u�ُ�񍐗p�@�o�^����
    
    '���u�w��
    fail_dir_sys_wrt_dte As String  ''���u�w���p�@�L�^���i����L�^���j
    
    '���u����
    fail_res_host_send As String         ''���u���ʗp�@�r�W�R�����M����
    fail_res_host_wrt_dte As String      ''���u���ʗp�@�L�^��
    fail_res_host_wrt_tme As String      ''���u���ʗp�@�L�^����
    
    '���ʃt���O
    host_send_flg As String ''�r�W�R�����M�t���O�i�e��ʂő��M�O�ɃZ�b�g�j�폜�n�͖��g�p
    
    PhotoImgCnt As String '' 20090115 add by M.Aoyagi    �摜�o�^�����\���̈גǉ�
    
    '2016/04/20 - TAI - S
    '��������
    slb_fault_total_judg As String
    '��Əꏊ
    slb_works_sky_tok As String
    '2016/04/20 - TAI - E

End Type

''�f�ޓ����f�[�^�i�X���u���A�J���[�`�F�b�N���p�j
''COLOR
Public Type typAPSozaiData
    '**********************************************************'
    'nchtaisl
    slb_no As String            ''�X���uNO
    slb_ksh As String           ''�|��
    slb_uksk As String          ''����i�M������j
    slb_lngth As String         ''����
    slb_color_wei As String     ''�d�ʁi�װ�����p�FSEG�o���d�ʁj
    slb_typ As String           ''�^
    slb_skin_wei As String      ''�d�ʁi���ޔ��p�F����d�ʁj
    slb_wdth As String          ''��
    slb_thkns As String         ''����
    slb_zkai_dte As String      ''������i����N�����j
    '**********************************************************'
    'skjchjdt�e�[�u��
    slb_chno As String          ''�`���[�WNO
    slb_ccno As String          ''CCNO
    '**********************************************************'
End Type

''���u���e�w���m�F�^���ʓo�^�p�f�[�^
''COLOR
Public Type typAPDirResData
    slb_no As String            ''�X���uNO
    slb_chno As String          ''�X���u�`���[�WNO
    slb_aino As String          ''�X���u����
    slb_stat As String          ''���
    slb_col_cnt As String       ''�J���[��
    dir_no As String            ''�w���ԍ�
    
    dir_nme1 As String            ''�w������1
    dir_val1 As String            ''�w���l1
    dir_uni1 As String            ''�w���P��1
    dir_nme2 As String            ''�w������2
    dir_val2 As String            ''�w���l2
    dir_uni2 As String            ''�w���P��2
    dir_cmt1 As String            ''�R�����g1
    dir_cmt2 As String            ''�R�����g2
    dir_wrt_dte  As String            ''�w����
    dir_wrt_nme As String            ''�w���Җ�
    dir_sys_wrt_dte As String            ''�o�^��
    dir_sys_wrt_tme As String            ''�o�^����
    
    res_cmt1 As String            ''�R�����g1�i���g�p�^�\��j
    res_cmt2 As String            ''�R�����g2�i���g�p�^�\��j
    res_cmp_flg As String           ''���u�����t���O 1:����
    res_aft_stat As String          ''���u���� 1:�s�K���L��i����A�r�L��j
    res_wrt_dte  As String          ''���͓�
    res_wrt_nme As String           ''���͎Җ�
    res_sys_wrt_dte As String            ''�o�^��
    res_sys_wrt_tme As String            ''�o�^����
    
End Type

'���׏��
''COLORSYS
Public Type typAPFaultList
    strCode As String
    strName As String
End Type

''�X�^�b�t���
''COLORSYS
Public Type typAPStaffData
    inp_StaffName As String ''�X�^�b�t��
End Type

''���������
''COLORSYS
Public Type typAPInspData
    inp_InspName As String ''��������
End Type

''���͎ҏ��
''COLORSYS
Public Type typAPInpData
    inp_InpName As String ''���͎Җ�
End Type

''���H�����
Public Type typAPNextProcData
    inp_NextProc As String ''���H��
End Type

''���u���
Public Type typAPDirRes_Stat
    inp_DirRes_StatCode As String
    inp_DirRes_Stat As String
End Type

''���u����
Public Type typAPDirRes_Res
    inp_DirRes_ResCode As String
    inp_DirRes_Res As String
End Type

''�V�X�e�����
Public APSysCfgData As typAPSysCfgData ''�V�X�e�����

'''�V�X�e���R���g���[���f�[�^
'Public APSysCont As typAPSysCont ''2001-11-09 �V�X�e���R���g���[���f�[�^

''�X���u���R���g���[���f�[�^
Public APSlbCont As typAPSlbCont ''2001-11-08 �X���u���R���g���[���f�[�^

''�ʌ��׃��X�g���i�X���u���j
Public APFaultFaceSkin() As typAPFaultList
''�������׃��X�g���i�X���u���j
Public APFaultInsideSkin() As typAPFaultList

''�ʌ��׃��X�g���i�J���[�`�F�b�N�j
Public APFaultFaceColor() As typAPFaultList

''���������уf�[�^�i��ʕ\�������W�X�g���ۑ��p�j
Public APResData As typAPResData ''���������уf�[�^�i��ʕ\�������W�X�g���ۑ��p�j
Public APResDataBK As typAPResData ''���������уf�[�^�i�����p�o�b�N�A�b�v�G���A�j

Public APSozaiData As typAPSozaiData ''�f�ޓ����⍇���f�[�^

Public APDirResData() As typAPDirResData ''���u���e�w���m�F�^���ʓo�^�p�f�[�^

''�����X���u���X�g
Public APSearchListSlbData() As typAPSlbData ''�����X���u���X�g

''�X���u�����p�s�l�o
Public APSearchTmpSlbData() As typAPSlbData ''�X���u�����p�s�l�o

''���уf�[�^�Ǎ��ݗp�s�l�o
Public APResTmpData() As typAPResData ''���уf�[�^�Ǎ��ݗp�s�l�o
Public APSozaiTmpData() As typAPSozaiData ''�f�ޓ����⍇���f�[�^
Public APDirResTmpData() As typAPDirResData ''���u���e�w���m�F�^���ʓo�^�p�f�[�^

''�X�^�b�t���}�X�^���
''COLORSYS
Public APStaffData() As typAPStaffData ''�X�^�b�t���}�X�^���

''���������}�X�^���
''COLORSYS
Public APInspData() As typAPInspData ''���������}�X�^���

''���͎Җ��}�X�^���
''COLORSYS
Public APInpData() As typAPInpData ''���͎Җ��}�X�^���

''���H���}�X�^���
Public APNextProcDataSkin() As typAPNextProcData ''���H���}�X�^���
Public APNextProcDataColor() As typAPNextProcData ''���H���}�X�^���

Public APDirRes_Stat() As typAPDirRes_Stat ''���u���
Public APDirRes_Res() As typAPDirRes_Res ''���u����

''�c�a�I�t���C���ŋ������͂��s�������Ƃ𔻒f����t���O
'Public bAPInputOffline As Boolean

'2016/04/20 - TAI - S
'��Əꏊ���
Public works_sky_tok As String
Public Const WORKS_SKY As String = "SKY"       'SKY
Public Const WORKS_TOK As String = "TOK"       '���|
'2016/04/20 - TAI - E

