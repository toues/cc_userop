VERSION 5.00
Object = "{00120101-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "Ltlst12n.ocx"
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "Ltocx12n.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPhotoImgUpView 
   BackColor       =   &H00C0FFFF&
   Caption         =   "�ʐ^�Y�t"
   ClientHeight    =   14625
   ClientLeft      =   855
   ClientTop       =   1125
   ClientWidth     =   17355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   14625
   ScaleWidth      =   17355
   StartUpPosition =   2  '��ʂ̒���
   WindowState     =   2  '�ő剻
   Begin VB.CheckBox chkALLFILES 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2820
      TabIndex        =   40
      Top             =   13620
      Width           =   255
   End
   Begin VB.CheckBox chkDEL 
      Caption         =   "Check1"
      Height          =   255
      Left            =   10500
      TabIndex        =   38
      Top             =   13620
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   17115
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4200
         TabIndex        =   33
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4980
         TabIndex        =   32
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "�^"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4440
         TabIndex        =   30
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4980
         TabIndex        =   29
         Top             =   900
         Width           =   945
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   6120
         TabIndex        =   28
         Top             =   900
         Width           =   705
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6840
         TabIndex        =   27
         Top             =   900
         Width           =   945
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "�|��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   420
         TabIndex        =   26
         Top             =   900
         Width           =   705
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "N304AM"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   25
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "����No"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "47965 - 15"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   23
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   11580
         TabIndex        =   22
         Top             =   900
         Width           =   885
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "20080129"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4980
         TabIndex        =   20
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "CCNo"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4020
         TabIndex        =   19
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   8700
         TabIndex        =   18
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   7920
         TabIndex        =   17
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   8700
         TabIndex        =   16
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   7800
         TabIndex        =   15
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   8700
         TabIndex        =   14
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label lblSlbTitle 
         Alignment       =   1  '�E����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   7860
         TabIndex        =   13
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   12600
         TabIndex        =   12
         Top             =   900
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdFileSel 
      Caption         =   "�Q��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   3
      Top             =   13020
      Width           =   915
   End
   Begin VB.CommandButton cmdUPLOAD 
      Caption         =   "�t�o�k�n�`�c"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      TabIndex        =   2
      Top             =   13020
      Width           =   2595
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   495
      Left            =   14640
      TabIndex        =   1
      Top             =   13860
      Width           =   2595
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�폜"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      TabIndex        =   0
      Top             =   3720
      Width           =   2595
   End
   Begin LEADImgListLibCtl.LEADImgList LEADImgList1 
      Height          =   8655
      Left            =   60
      OleObjectBlob   =   "frmPhotoImgUpView.frx":0000
      TabIndex        =   4
      Top             =   4260
      Width           =   2235
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LEADLib.LEAD LEAD1 
      Height          =   8655
      Left            =   2340
      TabIndex        =   36
      Top             =   4260
      Width           =   14895
      _Version        =   65539
      _ExtentX        =   26273
      _ExtentY        =   15266
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   573
      ScaleWidth      =   989
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin VB.Label lblUPLOADFILE 
      BorderStyle     =   1  '����
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
      Left            =   2880
      TabIndex        =   42
      Top             =   13020
      Width           =   10725
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '����
      Caption         =   "�w��t�H���_���̑S�Ă�JPG�摜��Ώ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   41
      Top             =   13560
      Width           =   6675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "�R�s�[�����폜����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   39
      Top             =   13560
      Width           =   3375
   End
   Begin VB.Label lblDEBUG 
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   37
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSlbTitle 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�J���[��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   35
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label lblSlb 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1980
      TabIndex        =   34
      Top             =   780
      Width           =   2565
   End
   Begin VB.Label lblMainTitle 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�X���u�ُ�񍐏��|�ʐ^�Y�t"
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
      TabIndex        =   10
      Top             =   0
      Width           =   17175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�p�X�^�t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   13020
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�ʐ^�ۑ���t�H���_��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3300
      Width           =   3555
   End
   Begin VB.Label lblShowFolder 
      BorderStyle     =   1  '����
      Caption         =   "\\FILESERVER\���L\IMG\FAULT\12345\1234\"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3300
      Width           =   13485
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '����
      Caption         =   "�\�����̃t�@�C����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3555
   End
   Begin VB.Label lblShowFile 
      BorderStyle     =   1  '����
      Caption         =   "TEST.JPG"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3780
      Width           =   10845
   End
End
Attribute VB_Name = "frmPhotoImgUpView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmPhotoImgUpView.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�ʐ^�摜�Y�t�@�\��ʕ\���t�H�[��
' �@�{���W���[���͎ʐ^�摜�Y�t�@�\��ʕ\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[

Public gnIndex As Long
Public gnMouseX As Long
Public gnMouseY As Long

Private gnSYSMODE As Integer
Private gsSYSFILENAME_MASK As String
Private gnSYSFILECOUNT_MAX As Integer

Private strFileList() As String

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

Private Sub cmdCancel_Click()
    Dim sKind As String
    Dim sColCnt As String
    
    APSysCfgData.PHOTOIMG_DIR = lblUPLOADFILE.Caption
    APSysCfgData.PHOTOIMG_ALLFILES = chkALLFILES.Value
    APSysCfgData.PHOTOIMG_DELCHK = chkDEL.Value
    
    ' 20090115 add by M.Aoyagi
    Select Case gnSYSMODE
        Case conDefine_SYSMODE_SKIN
            sKind = "\SKIN"
            sColCnt = "00"
        Case conDefine_SYSMODE_COLOR
            sKind = "\COLOR"
            sColCnt = APResData.slb_col_cnt
        Case conDefine_SYSMODE_SLBFAIL
            sKind = "\COLOR"
            sColCnt = APResData.slb_col_cnt
    End Select
    APResData.PhotoImgCnt = PhotoImgCount(sKind, APResData.slb_chno, APResData.slb_aino, APResData.slb_stat, sColCnt)
    
    Unload Me
    
    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResCANCEL
    Set cCallBackObject = Nothing
End Sub

Private Sub cmdDelete_Click()
    
    If lblShowFile.Caption = "" Then
            Call WaitMsgBox(Me, "�폜�Ώۂ̉摜������܂���B")
        Exit Sub
    End If
    
    Dim fmessage As Object
    Set fmessage = New MessageYN

    fmessage.MsgText = "�\�����̉摜�t�@�C�����폜���܂��B" & vbCrLf & "��낵���ł����H"
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_PHOTOIMG_DELETE, False
    fmessage.Show vbModal, Me '���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
    Set fmessage = Nothing
    
End Sub

Private Sub cmdUPLOAD_Click()
    
    Dim strMess As String
    Dim nI As Integer
    
    If chkALLFILES.Value = 1 Then
        Call getLocalImgFileList(GetPath(lblUPLOADFILE.Caption))
    Else
        If UCase(Right(lblUPLOADFILE.Caption, 4)) = ".JPG" Then
            If GetFilename(lblUPLOADFILE.Caption) <> "" Then
                ReDim strFileList(1)
                strFileList(0) = GetPath(lblUPLOADFILE.Caption)
                If Right(strFileList(0), 1) = "\" Then
                    strFileList(0) = Left(strFileList(0), Len(strFileList(0)) - 1)
                End If
                strFileList(1) = GetFilename(lblUPLOADFILE.Caption)
            End If
        Else
            '�����t�@�C���w��̏ꍇ
            For nI = 1 To UBound(strFileList)
                If UCase(Right(strFileList(nI), 4)) <> ".JPG" Then
                    '���̑��̃t�@�C���̏ꍇ
                    Call WaitMsgBox(Me, "�w��t�@�C���F" & strFileList(nI) & "�͎g�p�ł��܂���B" & vbCrLf & "������x�A�t�@�C����I���������Ă��������B")
                    Exit Sub
                End If
            Next nI
        End If
    End If
    
    
'    If lblUPLOADFILE.Caption = "" Then
    If UBound(strFileList) < 1 Then
            Call WaitMsgBox(Me, "�w��t�@�C��������܂���B")
        Exit Sub
    End If
    
    Dim fmessage As Object
    Set fmessage = New MessageYN

    strMess = "�w��t�@�C��" & vbCrLf

    For nI = 1 To UBound(strFileList)
        strMess = strMess & strFileList(nI) & vbCrLf
        If nI > 3 Then
            strMess = strMess & "�E�E�E" & vbCrLf
            Exit For
        End If
    Next nI

    strMess = strMess & "���A�b�v���[�h���܂��B" & vbCrLf
    
    If chkDEL.Value = 1 Then
        strMess = strMess & "�A�b�v���[�h��R�s�[���t�@�C���͍폜����܂��B" & vbCrLf
    End If
    
    strMess = strMess & "��낵���ł����H"

    fmessage.MsgText = strMess
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_PHOTOIMG_UPLOAD, False
    fmessage.Show vbModal, Me '���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
    Set fmessage = Nothing
    
End Sub

Private Sub PhotoImgUpLoad()
    Dim nI As Integer
    Dim nSysFileCount As Integer
    Dim strSource As String
    Dim strDestination As String
    
    'LEADImgList1.LoadFromFile txtUPLOADFILE.Text, 0, 0, -1
        '�t�@�C�����쐬
    
    nSysFileCount = gnSYSFILECOUNT_MAX
    
    For nI = 1 To UBound(strFileList)
    
        nSysFileCount = nSysFileCount + 1
        If nSysFileCount > 99 Then nSysFileCount = 1
        
        '��ʂɂ��A����
        Select Case gnSYSMODE
            Case conDefine_SYSMODE_SKIN
                strDestination = APSysCfgData.SHARES_IMGDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                              "\SKIN" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                              "_" & APResData.slb_stat & "_00_" & Format(nSysFileCount, "00") & ".JPG"
            
            Case conDefine_SYSMODE_COLOR
                strDestination = APSysCfgData.SHARES_IMGDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                              "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                              "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & "_" & Format(nSysFileCount, "00") & ".JPG"
        
            Case conDefine_SYSMODE_SLBFAIL
                strDestination = APSysCfgData.SHARES_IMGDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                              "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                              "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & "_" & Format(nSysFileCount, "00") & ".JPG"
            
            Case Else
                Exit Sub
        End Select
        
        strSource = strFileList(0) & "\" & strFileList(nI)
        
        On Error GoTo PhotoImgUpLoad_err:
        Call FileCopy(strSource, strDestination)
        On Error GoTo 0


    Next nI

    gnSYSFILECOUNT_MAX = nSysFileCount

    If chkDEL.Value = 1 Then
        On Error Resume Next
        For nI = 1 To UBound(strFileList)
            strSource = strFileList(0) & "\" & strFileList(nI)
            Call Kill(strSource)
        Next nI
        On Error GoTo 0
        Call WaitMsgBox(Me, "�R�s�[���̍폜���s���܂����B" & vbCrLf & "�A�b�v���[�h���I�����܂����B")
    Else
        Call WaitMsgBox(Me, "�A�b�v���[�h���I�����܂����B")
    End If

    '������
    ReDim strFileList(0)
    strSource = GetPath(lblUPLOADFILE.Caption)
    If Right(strSource, 1) = "\" Then
        strSource = Left(strSource, Len(strSource) - 1)
    End If
    lblUPLOADFILE.Caption = UCase(strSource)

    Call dp_Refresh

    Exit Sub
    
PhotoImgUpLoad_err:

    Call WaitMsgBox(Me, "�A�b�v���[�h�����s���܂����B" & vbCrLf & "�k�`�m����܂��́A�l�b�g���[�N�����m�F���Ă��������B")

    '������
    ReDim strFileList(0)
    strSource = GetPath(lblUPLOADFILE.Caption)
    If Right(strSource, 1) = "\" Then
        strSource = Left(strSource, Len(strSource) - 1)
    End If
    lblUPLOADFILE.Caption = UCase(strSource)

    Call dp_Refresh

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

    Dim bRet As Boolean
    Dim strDestination As String

    gnIndex = -1
    ReDim strFileList(0)

    LEAD1.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD1.EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    LEAD1.EnableTwainEvent = True

    Call GetCurrentAPSlbData

    strDestination = ""

    '�ďo���ɂ��A��������
    Select Case cCallBackObject.Name

        Case "frmSkinScanWnd" ''�X���u����������
            gnSYSMODE = conDefine_SYSMODE_SKIN
'            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
            lblMainTitle.Caption = "�X���u���������́|�ʐ^�Y�t"
            lblSlbTitle(12).Visible = False
            lblSlb(12).Visible = False

            '�t�H���_�쐬
            On Error Resume Next
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SKIN"
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SKIN" & "\" & APResData.slb_chno
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
            Call MkDir(strDestination)
            On Error GoTo 0
    
'            gsSYSFILENAME_MASK = "SKIN_?????_????_?_??_??.JPG"
            gsSYSFILENAME_MASK = "SKIN_" & APResData.slb_chno & "_" & APResData.slb_aino & "_" & APResData.slb_stat & "_" & "00" & "_??.JPG"
    
        Case "frmColorScanWnd" ''�J���[�`�F�b�N�����\����
            gnSYSMODE = conDefine_SYSMODE_COLOR
'            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_COLOR)
            lblMainTitle.Caption = "�J���[�`�F�b�N�����\���́|�ʐ^�Y�t"
            lblSlbTitle(12).Visible = True
            lblSlb(12).Visible = True
            
            '�t�H���_�쐬
            On Error Resume Next
            strDestination = APSysCfgData.SHARES_IMGDIR & "\COLOR"
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\COLOR" & "\" & APResData.slb_chno
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
            Call MkDir(strDestination)
            On Error GoTo 0
            
'            gsSYSFILENAME_MASK = "COLOR_?????_????_?_??_??.JPG"
            gsSYSFILENAME_MASK = "COLOR_" & APResData.slb_chno & "_" & APResData.slb_aino & "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & "_??.JPG"
            
        Case "frmSlbFailScanWnd" ''�X���u�ُ�񍐏�����
            gnSYSMODE = conDefine_SYSMODE_SLBFAIL
'            LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SLBFAIL)
            lblMainTitle.Caption = "�X���u�ُ�񍐏����́|�ʐ^�Y�t"
            lblSlbTitle(12).Visible = True
            lblSlb(12).Visible = True

            '�t�H���_�쐬�i�X���u�ُ�񍐕��j
            On Error Resume Next
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SLBFAIL"
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SLBFAIL" & "\" & APResData.slb_chno
            Call MkDir(strDestination)
            strDestination = APSysCfgData.SHARES_IMGDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
            Call MkDir(strDestination)
            On Error GoTo 0

'            gsSYSFILENAME_MASK = "SLBFAIL_?????_????_?_??_??.JPG"
            gsSYSFILENAME_MASK = "SLBFAIL_" & APResData.slb_chno & "_" & APResData.slb_aino & "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & "_??.JPG"

    End Select
    
    '�t�H���_�Z�b�g
    lblShowFolder.Caption = UCase(strDestination & "\")
    '�t�@�C����������
    lblShowFile.Caption = ""
    
    lblUPLOADFILE.Caption = APSysCfgData.PHOTOIMG_DIR
    chkALLFILES.Value = APSysCfgData.PHOTOIMG_ALLFILES
    chkDEL.Value = APSysCfgData.PHOTOIMG_DELCHK
    
    gnSYSFILECOUNT_MAX = 0 '������
    
    bRet = ImgListLoad(strDestination)
    
    '���X�g���[�h�G���[�̏ꍇ�A�����I����\��
    If bRet = False Then
        Call cmdCancel_Click
    End If
    
End Sub

Private Sub dp_Refresh()
    Dim bRet As Boolean

    LEAD1.Bitmap = 0
    gnIndex = -1
    ReDim strFileList(0)
    
    '�t�@�C����������
    lblShowFile.Caption = ""
    bRet = ImgListLoad(lblShowFolder.Caption)

    '���X�g���[�h�G���[�̏ꍇ�A�����I����\��
    If bRet = False Then
        Call cmdCancel_Click
    End If

End Sub

Private Sub getLocalImgFileList(ByVal strTGT_Folder As String)
   Dim myFileName As String
   Dim strSerchFolder As String
    
    ReDim Preserve strFileList(0)
    
    strSerchFolder = Trim(strTGT_Folder)
    
    If Right(strSerchFolder, 1) = "\" Then
        strFileList(0) = Left(strSerchFolder, Len(strSerchFolder) - 1)
    Else
        strFileList(0) = strSerchFolder
    End If
    
    myFileName = Dir(strFileList(0) & "\*.JPG")

    Do Until myFileName = vbNullString
        ReDim Preserve strFileList(UBound(strFileList) + 1)
        strFileList(UBound(strFileList)) = myFileName
        myFileName = Dir()
    Loop

End Sub

Private Function ImgListLoad(ByVal strTGT_Folder As String) As Boolean
   Dim myFileName As String
    
    LEADImgList1.Clear
    
    myFileName = Dir(strTGT_Folder & "\" & gsSYSFILENAME_MASK)

    Do Until myFileName = vbNullString
        On Error GoTo ImgListLoad_err:
        LEADImgList1.LoadFromFile strTGT_Folder & "\" & myFileName, 0, 0, -1
        On Error GoTo 0
        
        If gnSYSFILECOUNT_MAX < CInt(Left(Right(myFileName, 6), 2)) Then
            gnSYSFILECOUNT_MAX = CInt(Left(Right(myFileName, 6), 2))
        End If
        
        myFileName = Dir()
    Loop

    If IsDEBUG("DISP") Then
        '�\���f�o�b�O���[�h
        lblDebug(0).Visible = True
        lblDebug(0).Caption = gnSYSFILECOUNT_MAX
    Else
        lblDebug(0).Visible = False
    End If
    
    ImgListLoad = True 'OK
    Exit Function

ImgListLoad_err:

    Call WaitMsgBox(Me, "�C���[�W���X�g�Ǎ��ŃG���[���������܂����B:" & strTGT_Folder & "\" & myFileName)
    ImgListLoad = False 'NG

End Function

' @(f)
'
' �@�\      : �J�����g�X���u���擾
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �J�����g�X���u���̎擾���s���B
'
' ���l      :
'
Private Sub GetCurrentAPSlbData()

    lblSlb(0).Caption = APResData.slb_chno & "-" & APResData.slb_aino ''�X���uNo
    lblSlb(1).Caption = ConvDpOutStat(gnSYSMODE, CInt(APResData.slb_stat)) ''���
    lblSlb(2).Caption = APResData.slb_ccno ''CCNo
    lblSlb(3).Caption = APResData.slb_zkai_dte ''�����
    lblSlb(4).Caption = APResData.slb_ksh ''�|��
    lblSlb(5).Caption = APResData.slb_typ ''�^
    lblSlb(6).Caption = APResData.slb_uksk ''����
    lblSlb(7).Caption = APResData.slb_wei ''�d��
    lblSlb(8).Caption = APResData.slb_thkns ''����
    lblSlb(9).Caption = APResData.slb_wdth ''��
    lblSlb(10).Caption = APResData.slb_lngth ''����
'    lblSlb(11).Caption = APResData.sys_wrt_dte ''�L�^��

    If IsNumeric(APResData.slb_col_cnt) Then
        lblSlb(12).Caption = Format(CInt(APResData.slb_col_cnt), "00") ''�J���[��
    Else
        lblSlb(12).Caption = ""
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


Private Sub cmdFileSel_Click()
    
    Dim strFnm As String
    Dim strSource As String
    
    strFnm = lblUPLOADFILE.Caption
    If strFnm <> "" And UCase(Left(strFnm, 4)) <> ".JPG" Then
        strFnm = strFnm & "\"
    End If
    
    If ShowCommonDialog(CommonDialog1, "�J���[�`�F�b�N���d�q���V�X�e���Y�t�摜�t�@�C���̑I��", 0, "�摜�t�@�C��(*.jpg)|*.jpg", strFnm, strFileList) Then
      '** OK �{�^���̏���
        If Trim(strFnm) = "" Then
            Call WaitMsgBox(Me, "�J���[�`�F�b�N����PC �Y�t�摜�t�@�C���͑I������܂���ł����B")
            Exit Sub
        End If
        
        strSource = Trim(strFnm)
        
        If Right(strSource, 1) = "\" Then
            strSource = Left(strSource, Len(strSource) - 1)
        End If
        
        lblUPLOADFILE.Caption = UCase(strSource)
        
    Else
      '** �L�����Z���{�^���̏���
        Call WaitMsgBox(Me, "�J���[�`�F�b�N����PC �Y�t�摜�t�@�C���̑I���̓L�����Z������܂����B")
    End If
    

End Sub

Private Sub lblSlbT_Click(Index As Integer)

End Sub


Private Sub LEADImgList1_ItemSelected(ByVal nIndex As Long)
Call WaitMsgBox(Me, LEADImgList1.Item(nIndex).Text & " ���I������܂����B")

    LEAD1 = LEADImgList1(nIndex)
End Sub

Private Sub LEADImgList1_Click()
    Dim nIndex As Long
    
    nIndex = LEADImgList1.HitTest(CInt(gnMouseX), CInt(gnMouseY))
    
    If (nIndex >= 0) Then
        If gnIndex >= 0 Then
            'LEADImgList1.SelectionColor = RGB(255, 0, 0)
            LEADImgList1.Item(gnIndex).Selected = False
        End If
        'LEADImgList1.SelectionColor = RGB(255, 0, 0)
        LEADImgList1.Item(nIndex).Selected = True
        LEAD1.Bitmap = LEADImgList1(nIndex).Bitmap
        gnIndex = nIndex
        
        lblShowFile.Caption = UCase(GetFilename(LEADImgList1(nIndex).Text))
    End If
End Sub

Private Sub LEADImgList1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    gnMouseX = x
    gnMouseY = y
End Sub


'=======================================================================
'  �R�����_�C�A���O�\��
'=======================================================================
'�y�����z
'  comd    = �R�����_�C�A���O�R���g���[��
'  title   = �^�C�g��
'  mode    = �_�C�A���O���[�h
'              0 = �I�[�v��
'              1 = �ۑ�
'  filt    = �t�B���^
'  fnm     = �y���o�́z�t�@�C���l�[��
'�y�߂�l�z
'  boolean = ��������
'              TRUE  = OK �{�^��
'              FALSE = �L�����Z���{�^��
'�y�����z
'  �E�R�����_�C�A���O��\�����āA�t�@�C�������擾����B
'=======================================================================
Public Function ShowCommonDialog(comd As Variant, title As String, mode As Integer, filt As String, ByRef fnm As String, ByRef Filenames() As String) As Boolean

  Dim dirsv As String
  
   Dim i As Integer
   Dim sFname As String
    Dim intIndex As Integer
  Dim iEndPath As Integer
   Dim iStart As Integer

'** �R�����_�C�A���O�\��
  dirsv = CurDir
  comd.DialogTitle = title
  comd.InitDir = GetPath(fnm)
  comd.FileName = GetFilename(fnm)
  comd.Filter = filt
  comd.CancelError = True
  comd.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
  On Local Error Resume Next
  If mode = 0 Then
    comd.ShowOpen
  Else
    comd.ShowSave
  End If
  If Err = 0 Then
    
  sFname = comd.FileName & vbNullChar
    
   iEndPath = 1
   ' determine if multiple files were selected
   ' null delimiter is not inserted if only 1 file is selected
   If countDelimiters(sFname, vbNullChar) = 1 Then
      Do Until (iEndPath = 0)
         iStart = iEndPath + 1
         iEndPath = InStr(iEndPath + 1, sFname, "\")
      Loop
      ReDim Preserve Filenames(0)
      ' determine if root directory was selected - preserve the "\"
      If countDelimiters(sFname, "\") = 1 Then
         Filenames(0) = Mid(sFname, 1, iStart - 1)
      Else
         Filenames(0) = Mid(sFname, 1, iStart - 2)
      End If
   Else
      iStart = InStr(1, sFname, vbNullChar) + 1
      ReDim Preserve Filenames(0)
      Filenames(0) = Left(sFname, iStart - 2)
   End If

   intIndex = 1
   For i = iStart To Len(sFname)
      If Mid(sFname, i, 1) = vbNullChar Then
        ReDim Preserve Filenames(intIndex)
        Filenames(intIndex) = Mid(sFname, iStart, i - iStart)
        iStart = i + 1
        intIndex = intIndex + 1
      End If
   Next i

   ' display information in proper text box
   For i = 0 To UBound(Filenames)
      If i Then
         Debug.Print Filenames(i)
      Else
         Debug.Print Filenames(i)
      End If
   Next i
    
    fnm = LCase(Trim(comd.FileName))
    ShowCommonDialog = True
  Else
    fnm = ""
    ShowCommonDialog = False
  End If
  On Local Error GoTo 0
  ChDrive Left(dirsv, 2)
  ChDir dirsv

End Function

Private Function countDelimiters(ByVal sFiles As String, _
      ByVal vSearchChar As Variant) As Integer

    Dim iCtr As Integer
    Dim iResult As Integer

    For iCtr = 1 To Len(sFiles)
        If Mid(sFiles, iCtr, 1) = vSearchChar Then iResult = iResult + 1
    Next iCtr

    countDelimiters = iResult

End Function

'=======================================================================
'  �f�B���N�g�����؂�o��
'=======================================================================
Public Function GetPath(path As Variant) As String

  Dim i As Integer
  Dim fnm As String

'** �f�B���N�g�����؂�o��
  fnm = LCase(Trim(path))
  For i = Len(fnm) To 1 Step -1
    If Mid(fnm, i, 1) = "." Then Exit For
  Next
  If i = 0 Then fnm = fnm + "\"
  For i = Len(fnm) To 1 Step -1
    If Mid(fnm, i, 1) = "\" Then Exit For
  Next
  If i > 0 Then
    GetPath = Left(fnm, i)
  Else
    For i = 1 To Len(fnm)
      If Mid(fnm, i, 1) = ":" Then Exit For
    Next
    If i > 0 Then
      GetPath = Left(fnm, i)
    Else
      GetPath = ""
    End If
  End If

End Function

'=======================================================================
'  �t�@�C�����؂�o��
'=======================================================================
Public Function GetFilename(path As Variant) As String

  Dim i As Integer
  Dim fnm As String

'** �t�@�C�����؂�o��
  fnm = LCase(Trim(path))
  If Right(fnm, 1) = "\" Then
    GetFilename = ""
  Else
    For i = Len(fnm) To 1 Step -1
      If Mid(fnm, i, 1) = "\" Then Exit For
    Next
    If i > 0 Then
      GetFilename = Right(fnm, Len(fnm) - i)
    Else
      For i = 1 To Len(fnm)
        If Mid(fnm, i, 1) = ":" Then Exit For
      Next
      If i < Len(fnm) Then
        GetFilename = Right(fnm, Len(fnm) - i)
      Else
        GetFilename = ""
      End If
    End If
  End If

End Function

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
    Dim nI As Integer
    
    Select Case CallNo
    
    Case CALLBACK_PHOTOIMG_DELETE 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            On Error Resume Next
            Kill lblShowFolder.Caption & lblShowFile.Caption
            On Error GoTo 0
        
            Call dp_Refresh
        End If
    
    Case CALLBACK_PHOTOIMG_UPLOAD 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            Call PhotoImgUpLoad
        Else
            'CANCEL
            '������
            ReDim strFileList(0)
            Call WaitMsgBox(Me, "�������L�����Z�����܂����B" & vbCrLf & _
                                "�t�@�C���̎w����s���Ă����ꍇ�́A" & vbCrLf & _
                                "������x�A�t�@�C����I�����Ȃ����Ă��������B")
        End If
    
    End Select

End Sub

