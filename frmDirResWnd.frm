VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDirResWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "���u���e�w���m�F�^���ʓo�^"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   17250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   17250
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdDirPrn 
      Caption         =   "�w�����"
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
      Left            =   60
      TabIndex        =   51
      Top             =   120
      Width           =   1800
   End
   Begin VB.Frame Frame_Status 
      BackColor       =   &H00C0FFFF&
      Caption         =   "���u���ʓ���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   60
      TabIndex        =   40
      Top             =   8520
      Width           =   17115
      Begin VB.CommandButton cmdInput 
         Caption         =   "�K�p"
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
         Left            =   6240
         TabIndex        =   45
         Top             =   900
         Width           =   1800
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         ItemData        =   "frmDirResWnd.frx":0000
         Left            =   2160
         List            =   "frmDirResWnd.frx":0002
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   44
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         ItemData        =   "frmDirResWnd.frx":0004
         Left            =   2160
         List            =   "frmDirResWnd.frx":0006
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   41
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         BackStyle       =   0  '����
         Caption         =   "���u�㌋��"
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
         Index           =   11
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         BackStyle       =   0  '����
         Caption         =   "���u���"
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
         Left            =   120
         TabIndex        =   42
         Top             =   540
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   60
      TabIndex        =   36
      Top             =   7200
      Width           =   17115
      Begin VB.Label lblDirCmt 
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
         Left            =   1860
         TabIndex        =   39
         Top             =   720
         Width           =   15165
      End
      Begin VB.Label lblDirCmt 
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
         Index           =   0
         Left            =   1860
         TabIndex        =   38
         Top             =   300
         Width           =   15165
      End
      Begin VB.Label Label6 
         Caption         =   "�w���R�����g"
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
         Index           =   80
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbRes 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      ItemData        =   "frmDirResWnd.frx":0008
      Left            =   14580
      List            =   "frmDirResWnd.frx":000A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   33
      Top             =   780
      Width           =   2595
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "���͎Җ��F"
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
      Left            =   12720
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   780
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   17115
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
         TabIndex        =   31
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   24
         Left            =   7860
         TabIndex        =   30
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
         Index           =   9
         Left            =   8700
         TabIndex        =   29
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   23
         Left            =   7800
         TabIndex        =   28
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
         Index           =   10
         Left            =   8700
         TabIndex        =   27
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   22
         Left            =   7920
         TabIndex        =   26
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
         Index           =   8
         Left            =   8700
         TabIndex        =   25
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   21
         Left            =   4020
         TabIndex        =   24
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
         Index           =   2
         Left            =   4980
         TabIndex        =   23
         Top             =   360
         Width           =   2805
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
         TabIndex        =   22
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   18
         Left            =   11580
         TabIndex        =   21
         Top             =   900
         Width           =   885
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
         TabIndex        =   20
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   1
         Left            =   60
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
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
         TabIndex        =   18
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   5
         Left            =   420
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label6 
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
         Index           =   8
         Left            =   6120
         TabIndex        =   15
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
         Index           =   5
         Left            =   4980
         TabIndex        =   14
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label6 
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
         Index           =   10
         Left            =   4440
         TabIndex        =   13
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label6 
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
         Index           =   92
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1035
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
         TabIndex        =   11
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label Label6 
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
         Index           =   25
         Left            =   4200
         TabIndex        =   10
         Top             =   1440
         Width           =   705
      End
   End
   Begin VB.PictureBox PicSigYellow 
      Height          =   315
      Left            =   2880
      Picture         =   "frmDirResWnd.frx":000C
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   10920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox PicSigRed 
      Height          =   375
      Left            =   5040
      Picture         =   "frmDirResWnd.frx":0650
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   10800
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox PicSigGreen 
      Height          =   315
      Left            =   3900
      Picture         =   "frmDirResWnd.frx":0E2E
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   10920
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
      Left            =   13020
      TabIndex        =   2
      Top             =   10500
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���M"
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
      Left            =   15300
      TabIndex        =   1
      Top             =   10500
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   3360
      Width           =   17115
      _ExtentX        =   30189
      _ExtentY        =   6694
      _Version        =   393216
      Rows            =   21
      Cols            =   6
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
   Begin VB.Label lblHostSendFlg 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   50
      Top             =   840
      Width           =   885
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��������
      Caption         =   "�޼޺�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4500
      TabIndex        =   49
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblDebug 
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
      Index           =   0
      Left            =   60
      TabIndex        =   48
      Top             =   10380
      Width           =   1275
   End
   Begin VB.Label lblDebug 
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
      Left            =   1380
      TabIndex        =   47
      Top             =   10380
      Width           =   2565
   End
   Begin VB.Label lblDebug 
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
      Left            =   4020
      TabIndex        =   46
      Top             =   10380
      Width           =   2565
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
      Left            =   1920
      TabIndex        =   35
      Top             =   840
      Width           =   2565
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��������
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
      Index           =   7
      Left            =   60
      TabIndex        =   34
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lbl_nMSFlexGrid1_Selected_Row 
      Height          =   315
      Left            =   8340
      TabIndex        =   6
      Top             =   10380
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "���u���e�w���m�F�^���ʓo�^"
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
      TabIndex        =   7
      Top             =   0
      Width           =   17175
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
      TabIndex        =   3
      Top             =   1980
      Width           =   1635
   End
End
Attribute VB_Name = "frmDirResWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmDirResWnd.Frm                ver 1.00 ( '2008.05.15 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N�����\���́|�X���u�I��\���t�H�[��
' �@�{���W���[���̓J���[�`�F�b�N�����\���́|�X���u�I��\���t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private cCallBackObject As Object ''�R�[���o�b�N�I�u�W�F�N�g�i�[
Private iCallBackID As Integer ''�R�[���o�b�N�h�c�i�[

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
    Call SlbSelLock(False)
    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResCANCEL) '2008/09/04 �߂��ύX
    Unload Me
End Sub

' @(f)
'
' �@�\      : �w������{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �w������{�^�������B
'
' ���l      :2008/09/04 �w������@�\
'
Private Sub cmdDirPrn_Click()
    
    Call DirPrnReq

End Sub

' @(f)
'
' �@�\      : �K�p�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �K�p�{�^�������B
'
' ���l      :COLORSYS
'
Private Sub cmdInput_Click()
    
    If nMSFlexGrid1_Selected_Row < 1 Then Exit Sub
    
    ''���u���ʓ��͂����X�g�֓K�p
    APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = APDirRes_Stat(cmbRes(2).ListIndex).inp_DirRes_StatCode
    APDirResData(nMSFlexGrid1_Selected_Row - 1).res_aft_stat = APDirRes_Res(cmbRes(3).ListIndex).inp_DirRes_ResCode
    
    '�����̏ꍇ
    If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = "1" Then
        If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme <> "" Then
        Else
            '���O���󔒂̏ꍇ
            APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme = cmbRes(0).Text '���͎Җ����X�g
        End If
    
        If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte <> "" Then
        Else
            '���t���󔒂̏ꍇ
            APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte = Format(Now, "YYYYMMDD")
        End If
    
    Else
        '�����łȂ��ꍇ
        APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme = "" '���͎Җ����X�g
        APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte = ""
    End If
    
   
    '���X�g�\���X�V
    Call SetMSFlexGrid1

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
    Dim nJ As Integer

    If UBound(APDirResData) < 1 Then Exit Sub

    Call DBSendDataReq_DIRRES

'    Unload Me
'    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) '�����p�� '2008/09/04 �߂��ύX

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

    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) '2008/09/04 �߂��ύX
    Unload Me

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
' ���l      :'2008/09/04 �߂��ύX
'
Public Sub CallBackMessage(ByVal CallNo As Integer, ByVal Result As Integer)
    
    Dim bRet As Boolean
    Dim nI As Integer
    
    Select Case CallNo
    
    Case CALLBACK_OPREGWND 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next

            cmbRes(0).Clear
            For nI = 1 To UBound(APInpData)
                cmbRes(0).AddItem APInpData(nI - 1).inp_InpName
'                cmbRes(0).ListIndex = nI - 1
            Next nI

'            Call InitForm
            'On Error GoTo 0
        End If
    
'    Case CALLBACK_NEXTPROCWND 'COLORSYS
'        If Result = CALLBACK_ncResOK Then          'OK
'            'On Error Resume Next
'
'            cmbRes(1).Clear
'            For nI = 1 To UBound(APNextProcDataColor)
'                cmbRes(1).AddItem APNextProcDataColor(nI - 1).inp_NextProc
'
'            Next nI
'
''            Call InitForm
'            'On Error GoTo 0
'        End If
    
    '�r�W�R�����M�L��A�c�a�o�^�̓o�^�₢���킹OK
    Case CALLBACK_RES_HOSTSNDDATA_DIRRES
            If Result = CALLBACK_ncResOK Then          'OK

'                ''DB�ۑ�����
'                Call SetAPResData(True)
'
'                '�J�����g���ѓ��͏��ꎞ�ۑ�
'                APResDataBK = APResData

                '�������Ł��F�����𑗐M
                APResData.slb_fault_u_judg = "9"
                APResData.slb_fault_d_judg = "9"

                '�r�W�R�����M
                frmHostSend.SetCallBack Me, CALLBACK_HOSTSEND
                frmHostSend.Show vbModal, Me '�r�W�R�����M���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B

            Else
                'DB�o�^�L�����Z��
            End If
    
    '�r�W�R�����M�����A�c�a�o�^�̓o�^�₢���킹OK
    Case CALLBACK_RES_DBSNDDATA_DIRRES
            If Result = CALLBACK_ncResOK Then          'OK

'                ''DB�ۑ�����
'                Call SetAPResData(True)
'
'                '�J�����g���ѓ��͏��ꎞ�ۑ�
'                APResDataBK = APResData

'                '�r�W�R�����M
'                frmHostSend.SetCallBack Me, CALLBACK_HOSTSEND
'                frmHostSend.Show vbModal, Me '�r�W�R�����M���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
'                '/* DB�o�^���s */
                bRet = DB_SAVE_DIRRES(False)
                
'                bRet = TRTS0022_Write(False)

                If bRet Then
                    '����I��
                    Unload Me
                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OK�ŏ����I�� '2008/09/04 �߂��ύX
                End If

            Else
                'DB�o�^�L�����Z��
            End If
    
    '���u���ʓo�^�̃r�W�R���ʐM���OK
    Case CALLBACK_HOSTSEND
            If Result = CALLBACK_ncResOK Then          'OK
                '����I��

'                APResData.fail_host_send = "1" '1:������Z�b�g

'                '/* DB�o�^���s */
                bRet = DB_SAVE_DIRRES(False)
'                bRet = TRTS0022_Write(False)

                Call dpDebug

                If bRet Then
                    '����I��
                    Unload Me
                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OK�ŏ����I�� '2008/09/04 �߂��ύX
                End If
            ElseIf Result = CALLBACK_ncResSKIP Then          'SKIP
                '/* DB�o�^���s */
                bRet = DB_SAVE_DIRRES(True)  '�r�W�R���G���[�L��
                '�����p��
                '�r�W�R���ʐM�K�{�̂��߁A�c�a�ۑ��͍s��Ȃ��B

'                '�r�W�R�����M�X�L�b�v�����i�����O�ɖ߂��B�j
'                APResData.host_send = APResDataBK.host_send
'                APResData.host_wrt_dte = APResDataBK.host_wrt_dte
'                APResData.host_wrt_tme = APResDataBK.host_wrt_tme
'

                Call dpDebug

'                '/* DB�o�^���s */
'                bRet = DB_SAVE_COLOR()
'
'                If bRet Then
'                    '����I��
'                    Unload Me
'                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OK�ŏ����I�� '2008/09/04 �߂��ύX
'                End If
            
            
            Else
                '/* DB�o�^���s */
'                bRet = DB_SAVE_DIRRES(True) '���f
                '�L�����Z���i�G���[�����ɂāAOK�{�^�����������ꍇ�A�ďo����ʂɖ߂�B�j
                '�r�W�R���ʐM�K�{�̂��߁A�c�a�ۑ��͍s��Ȃ��B

'                '�r�W�R�����M�X�L�b�v�����i�����O�ɖ߂��B�j
'                APResData.host_send = "0" '0:�ُ���Z�b�g
'                APResData.host_wrt_dte = APResDataBK.host_wrt_dte
'                APResData.host_wrt_tme = APResDataBK.host_wrt_tme
                
                Call WaitMsgBox(Me, "���M�^�c�a�ۑ������𒆒f���܂����B")

                Call dpDebug

            End If

        '�w������⍇�� 2008/09/04
        Case CALLBACK_RES_DIRPRN_REQ
            If Result = CALLBACK_ncResOK Then          'OK
                frmTRSend.SetCallBack Me, CALLBACK_RES_DIRPRN_SND, "COL02"
                frmTRSend.Show vbModal, Me
            Else
            End If

        '�w������v�����M���� 2008/09/04
        Case CALLBACK_RES_DIRPRN_SND
            If Result = CALLBACK_ncResOK Then          'OK
                Call WaitMsgBox(Me, "�w������v���͐���I�����܂����B")
            Else
                Call WaitMsgBox(Me, "�w������v���͎��s���܂����B")
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
    MSFlexGrid1.Cols = 5 + 1
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
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���u���"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 10000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "�w�����e"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���u�㌋��"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���͓�"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "���͎�"
    
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
    
    Dim strDir As String
    
    MSFlexGrid1.Rows = 1 + UBound(APDirResData)
    
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
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APDirResData(nRow - 1).res_cmp_flg = "1", "����", "")  '"���u���"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignLeftCenter
        '�w�����e
        strDir = APDirResData(nRow - 1).dir_nme1 & " " & APDirResData(nRow - 1).dir_val1 & " " & APDirResData(nRow - 1).dir_uni1 & " " & _
        APDirResData(nRow - 1).dir_nme2 & " " & APDirResData(nRow - 1).dir_val2 & " " & APDirResData(nRow - 1).dir_uni2
        MSFlexGrid1.TextMatrix(nRow, nCol) = strDir
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APDirResData(nRow - 1).res_aft_stat = "1", "�s�K���L��", "") '"���u�㌋��"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APDirResData(nRow - 1).res_wrt_dte '"���͓�"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APDirResData(nRow - 1).res_wrt_nme '"���͎�"
        
    
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
' �@�\      : ���͎Җ��o�^�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���͎Җ��o�^�{�^�������B
'
' ���l      :
'           :COLORSYS
'
Private Sub cmdUser_Click()
    frmOpRegWnd.SetCallBack Me, CALLBACK_OPREGWND
    frmOpRegWnd.Show vbModal, Me '�T�[�o�[�f�[�^�ǉ��^�폜���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
End Sub

Private Sub lblHostSendFlg_DblClick()
    If APResData.host_send_flg = "1" Then
        '�����̏ꍇ
        APResData.host_send_flg = "0" '�V�K�ɕύX
    Else
        APResData.host_send_flg = "1" '�����ɕύX
    End If

    lblHostSendFlg.Caption = IIf(APResData.host_send_flg = "0", "0:�V�K", "1:����")

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
'                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
'                    '�폜���[�h�̏ꍇ
'                    MSFlexGrid1.CellForeColor = &H8000000E
'                    MSFlexGrid1.CellBackColor = &H8080FF
'                Else
                    MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8000000D
'                End If

                ''���u���ʓ��̓G���A���X�V
                cmbRes(2).ListIndex = IIf(APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = "1", 1, 0)
                cmbRes(3).ListIndex = IIf(APDirResData(nMSFlexGrid1_Selected_Row - 1).res_aft_stat = "1", 1, 0)

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
    
    Call GetCurrentAPSlbData
    
    Call InitMSFlexGrid1

'    If APSlbCont.bProcessing Then '�X���u�I�����b�N�p�������t���O
'        imTextSearchSlbNumber.Text = APSlbCont.strSearchInputSlbNumber  '�����X���u�m���D
        
'        OptInputMode(APSlbCont.nSearchInputModeSelectedIndex).Value = True '���̓��[�h�w��C���f�b�N�X�ԍ�
'        OptStatus(APSlbCont.nSearchInputStatusSelectedIndex).Value = True '��ԑI���w��C���f�b�N�X�ԍ�
        
        '�X���u�I�����
'        nMSFlexGrid1_Selected_Row = APSlbCont.nListSelectedIndexP1

nMSFlexGrid1_Selected_Row = 0

        Call SetMSFlexGrid1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        Call MSFlexGrid1_Click
        Call SlbSelLock(True)
        
'    End If

End Sub

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

    Dim nI As Integer
    Dim nJ As Integer

'    lblInputMode.Caption = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, "�y�V�K�z", "�y�C���z")

    If APResData.fail_host_wrt_dte <> "" Then
        APResData.host_send_flg = "1" '����
    Else
        APResData.host_send_flg = "0" '�V�K
    End If

    lblHostSendFlg.Caption = IIf(APResData.host_send_flg = "0", "0:�V�K", "1:����")

'    '�ُ�񍐂����݂���ꍇ�́A�u���M�v�{�^���𖳌��ɂ���B
'    If APResData.fail_sys_wrt_dte <> "" Then
'        cmdOK.Enabled = False
'    Else
'        cmdOK.Enabled = True
'    End If

    '�J�����g�X���u��񃍁[�h

    Call dpDebug

    lblSlb(0).Caption = APResData.slb_chno & "-" & APResData.slb_aino ''�X���uNo
    lblSlb(1).Caption = ConvDpOutStat(conDefine_SYSMODE_COLOR, CInt(APResData.slb_stat)) ''���
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

    lblSlb(12).Caption = Format(CInt(APResData.slb_col_cnt), "00") ''�J���[��

    '���͎Җ����X�gBOX�ݒ�
    cmbRes(0).Clear
    For nJ = 1 To UBound(APInpData)
        cmbRes(0).AddItem APInpData(nJ - 1).inp_InpName
'        If APDirResData.slb_wrt_nme = APInpData(nJ - 1).inp_InpName Then
'            cmbRes(0).ListIndex = nJ - 1
'        End If
    Next nJ

'    '���H�����X�gBOX�ݒ�
'    cmbRes(1).Clear
'    For nJ = 1 To UBound(APNextProcDataColor)
'        cmbRes(1).AddItem APNextProcDataColor(nJ - 1).inp_NextProc
'        If APResData.slb_nxt_prcs = APNextProcDataColor(nJ - 1).inp_NextProc Then
'            cmbRes(1).ListIndex = nJ - 1
'        End If
'    Next nJ

'    '�R�����g��񃍁[�h
    lblDirCmt(0).Caption = APDirResData(0).dir_cmt1
    lblDirCmt(1).Caption = APDirResData(0).dir_cmt2

    ''���u��ԃ��X�gBOX�ݒ�
    cmbRes(2).Clear
    For nJ = 1 To UBound(APDirRes_Stat)
        cmbRes(2).AddItem APDirRes_Stat(nJ - 1).inp_DirRes_Stat
    Next nJ
    
    ''���u���ʃ��X�gBOX�ݒ�
    cmbRes(3).Clear
    For nJ = 1 To UBound(APDirRes_Res)
        cmbRes(3).AddItem APDirRes_Res(nJ - 1).inp_DirRes_Res
    Next nJ
    
End Sub

Private Sub dpDebug()

    Dim nI As Integer

    If IsDEBUG("DISP") Then
        '�\���f�o�b�O���[�h
        For nI = 0 To 2
            lblDEBUG(nI).Visible = True
        Next nI
        
        lblDEBUG(0).Caption = APResData.fail_host_send
        lblDEBUG(1).Caption = APResData.fail_host_wrt_dte
        lblDEBUG(2).Caption = APResData.fail_host_wrt_tme
        
    Else
        For nI = 0 To 2
            lblDEBUG(nI).Visible = False
        Next nI
    End If

End Sub

' @(f)
'
' �@�\      : ���уf�[�^�o�^�₢���킹����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : ���уf�[�^�o�^�₢���킹��ʂ��J���B
'
' ���l      : �R�[���o�b�N�L��B
'
Private Sub DBSendDataReq_DIRRES()
    Dim nI As Integer
    Dim bAllCmp As Boolean
    Dim fmessage As Object
    Set fmessage = New MessageYN

    '�S�Ċ������̃`�F�b�N
    bAllCmp = True '����
    For nI = 0 To UBound(APDirResData) - 1
        If APDirResData(nI).res_cmp_flg <> "1" Then
            bAllCmp = False '�������L��
            Exit For
        End If
    Next nI
    
    If bAllCmp Then
        '�����̏ꍇ
        fmessage.MsgText = "�S�Ċ����ƂȂ�܂��̂ŁA�r�W�R���֊����𑗐M��A�c�a�֓o�^���܂��B" & vbCrLf & "��낵���ł����H"
    '    fmessage.AutoDelete = True
        fmessage.AutoDelete = False
        fmessage.SetCallBack Me, CALLBACK_RES_HOSTSNDDATA_DIRRES, False
    Else
        '�������L��
        fmessage.MsgText = "���͌��ʂ��c�a�֓o�^���܂��B" & vbCrLf & "��낵���ł����H"
    '    fmessage.AutoDelete = True
        fmessage.AutoDelete = False
        fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_DIRRES, False
        
    End If
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

' @(f)
'
' �@�\      : �c�a�ۑ�����
'
' ������    :
'
' �Ԃ�l    : True ����I���^False �ُ�I��
'
' �@�\����  : �c�a�ۑ��������s���B
'
' ���l      :
'
Private Function DB_SAVE_DIRRES(ByVal bHostSendError As Boolean) As Boolean
    Dim bNOErrorFlg As Boolean
    Dim bRet As Boolean
    Dim MsgWnd As Message
    Set MsgWnd = New Message

    MsgWnd.MsgText = "�f�[�^�x�[�X�T�[�o�[�ɕۑ����ł��B" & vbCrLf & "���΂炭���҂����������B"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
    bNOErrorFlg = True '�G���[����

    '�r�W�R���ʐM�G���[������
    If bHostSendError Then
        
'        MsgWnd.OK_Close
'        MsgWnd.MsgText = "�r�W�R���ʐM�����폈���o���Ȃ��ׁA" & vbCrLf & "�c�a�ۑ��𒆒f���܂����B"
'        MsgWnd.OK.Visible = True
'    '    MsgWnd.AutoDelete = True
'        Do
'            On Error Resume Next
'            MsgWnd.Show vbModal
'            If Err.Number = 0 Then
'                Exit Do
'            End If
'            DoEvents
'        Loop
'        Set MsgWnd = Nothing
'
'        bNOErrorFlg = False '�G���[�L��
'        DB_SAVE_DIRRES = bNOErrorFlg
'        Exit Function
        
        APResData.fail_res_host_send = "0"       ''/* �r�W�R�����M���� */
        APResData.fail_res_host_wrt_dte = APResData.fail_host_wrt_dte    ''/* �r�W�R���o�^�� */
        APResData.fail_res_host_wrt_tme = APResData.fail_host_wrt_tme     ''/* �r�W�R���o�^���� */
    
    Else
        APResData.fail_res_host_send = "1"       ''/* �r�W�R�����M���� */
        APResData.fail_res_host_wrt_dte = APResData.fail_host_wrt_dte    ''/* �r�W�R���o�^�� */
        APResData.fail_res_host_wrt_tme = APResData.fail_host_wrt_tme     ''/* �r�W�R���o�^���� */
    End If

    'TRTS0022 �o�^
    bRet = TRTS0022_Write(False)
    
    If bRet = False Then
        bNOErrorFlg = False '�G���[�L��
        MsgWnd.OK_Close
        MsgWnd.MsgText = "�c�a�ۑ��G���[���������܂����B" & vbCrLf & "�����𒆒f���܂����B"
        MsgWnd.OK.Visible = True
'        MsgWnd.AutoDelete = True
        Do
            On Error Resume Next
            MsgWnd.Show vbModal
            If Err.Number = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        Set MsgWnd = Nothing
        DB_SAVE_DIRRES = bNOErrorFlg
    
    Else
        MsgWnd.OK_Close
        MsgWnd.MsgText = "�c�a�ۑ�������I�����܂����B"
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
        DB_SAVE_DIRRES = bNOErrorFlg
    
    End If

End Function

' @(f)
'
' �@�\      : �w������₢���킹����
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �w������₢���킹��ʂ��J���B
'
' ���l      : �R�[���o�b�N�L��B2008/09/04
'
Private Sub DirPrnReq()
    Dim fmessage As Object
    Set fmessage = New MessageYN

    fmessage.MsgText = "�w�����[�̈�����s���܂��B" & vbCrLf & "��낵���ł����H"
'    fmessage.AutoDelete = True
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_RES_DIRPRN_REQ, False
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
' ���l      :2008/09/04
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer)
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
End Sub

