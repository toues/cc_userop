VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "Ltocx12n.ocx"
Begin VB.Form frmSkinScanWnd 
   BackColor       =   &H00C0FFC0&
   Caption         =   "�J���[�`�F�b�N���ѓ��́��X�L���i�[�Ǎ���"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   1  '��Ű ̫�т̒���
   WindowState     =   2  '�ő剻
   Begin VB.CommandButton cmdOK 
      Caption         =   "���M"
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
      Left            =   17280
      TabIndex        =   56
      Top             =   13980
      Width           =   1755
   End
   Begin VB.CommandButton cmdCANCEL 
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
      Height          =   435
      Left            =   60
      TabIndex        =   73
      Top             =   60
      Width           =   1935
   End
   Begin VB.CommandButton cmdFullImage 
      Caption         =   "�S�̕\��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   17280
      TabIndex        =   72
      Top             =   5340
      Width           =   1755
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "�X�L���i�[�Ǎ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   17280
      TabIndex        =   71
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Height          =   2955
      Left            =   120
      TabIndex        =   16
      Top             =   1020
      Width           =   17115
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
         Index           =   1
         ItemData        =   "frmSkinScanWnd.frx":0000
         Left            =   12600
         List            =   "frmSkinScanWnd.frx":0002
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   11
         Top             =   1380
         Width           =   4455
      End
      Begin VB.CommandButton cmdNextProc 
         Caption         =   "���H��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11520
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1095
      End
      Begin imText6Ctl.imText imText 
         Height          =   465
         Index           =   0
         Left            =   2160
         TabIndex        =   12
         Top             =   1920
         Width           =   14895
         _Version        =   65536
         _ExtentX        =   26273
         _ExtentY        =   820
         Caption         =   "frmSkinScanWnd.frx":0004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0072
         Key             =   "frmSkinScanWnd.frx":0090
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
         AlignVertical   =   2
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "�y"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   50
         LengthAsByte    =   0
         Text            =   "�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   1
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   465
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   2400
         Width           =   14895
         _Version        =   65536
         _ExtentX        =   26273
         _ExtentY        =   820
         Caption         =   "frmSkinScanWnd.frx":00C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0132
         Key             =   "frmSkinScanWnd.frx":0150
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
         AlignVertical   =   2
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "�y"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   50
         LengthAsByte    =   0
         Text            =   "�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O�P�Q�R�S�T�U�V�W�X�O"
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   1
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   300
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0184
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":01F2
         Key             =   "frmSkinScanWnd.frx":0210
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   8
         LengthAsByte    =   0
         Text            =   "20080829"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0254
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":02C2
         Key             =   "frmSkinScanWnd.frx":02E0
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "A9#"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   6
         LengthAsByte    =   0
         Text            =   "N304AM"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   2
         Left            =   4980
         TabIndex        =   4
         Top             =   300
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0324
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0392
         Key             =   "frmSkinScanWnd.frx":03B0
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "12345"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   3
         Left            =   4980
         TabIndex        =   5
         Top             =   840
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":03F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0462
         Key             =   "frmSkinScanWnd.frx":0480
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "A9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "ABC"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   4
         Left            =   6840
         TabIndex        =   6
         Top             =   840
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":04C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0532
         Key             =   "frmSkinScanWnd.frx":0550
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "A9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "ABC"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   5
         Left            =   4980
         TabIndex        =   7
         Top             =   1380
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0594
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0602
         Key             =   "frmSkinScanWnd.frx":0620
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "12345"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   6
         Left            =   8700
         TabIndex        =   8
         Top             =   300
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":06D2
         Key             =   "frmSkinScanWnd.frx":06F0
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "#"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   6
         LengthAsByte    =   0
         Text            =   "123.12"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   7
         Left            =   8700
         TabIndex        =   9
         Top             =   840
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0734
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":07A2
         Key             =   "frmSkinScanWnd.frx":07C0
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
         LengthAsByte    =   0
         Text            =   "1234"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imSozai 
         Height          =   405
         Index           =   8
         Left            =   8700
         TabIndex        =   10
         Top             =   1380
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0872
         Key             =   "frmSkinScanWnd.frx":0890
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "12345"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
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
         TabIndex        =   93
         Top             =   1440
         Width           =   705
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
         TabIndex        =   70
         Top             =   360
         Width           =   1035
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
         TabIndex        =   69
         Top             =   900
         Width           =   435
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
         TabIndex        =   68
         Top             =   900
         Width           =   705
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
         TabIndex        =   67
         Top             =   900
         Width           =   705
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   1440
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
         TabIndex        =   64
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "�R�����g"
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
         Left            =   180
         TabIndex        =   63
         Top             =   1980
         Width           =   2715
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
         TabIndex        =   62
         Top             =   360
         Width           =   885
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
         TabIndex        =   61
         Top             =   360
         Width           =   705
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
         TabIndex        =   60
         Top             =   1440
         Width           =   825
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
         TabIndex        =   59
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
         Index           =   2
         Left            =   12600
         TabIndex        =   58
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�L�^��"
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
         Index           =   12
         Left            =   11460
         TabIndex        =   57
         Top             =   420
         Width           =   1125
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
         TabIndex        =   27
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label Label6 
         Alignment       =   2  '��������
         Caption         =   "�i�����������L���j"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   60
         TabIndex        =   18
         Top             =   2520
         Width           =   2055
      End
   End
   Begin VB.Timer timOpening 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   540
   End
   Begin VB.CommandButton cmdPhotoImgUp 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ʐ^�Y�t"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   17280
      Style           =   1  '���̨���
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "�X�^�b�t���F"
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
      Left            =   12780
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   540
      Width           =   1815
   End
   Begin VB.ComboBox cmbRes 
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
      Index           =   0
      ItemData        =   "frmSkinScanWnd.frx":08D4
      Left            =   14640
      List            =   "frmSkinScanWnd.frx":08D6
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin LEADLib.LEAD LEAD_SCAN 
      Height          =   315
      Left            =   18000
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   7620
      Visible         =   0   'False
      Width           =   315
      _Version        =   65539
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   17
      ScaleWidth      =   17
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin LEADLib.LEAD LEAD1 
      Height          =   7755
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   17115
      _Version        =   65539
      _ExtentX        =   30189
      _ExtentY        =   13679
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   513
      ScaleWidth      =   1137
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   75
      Top             =   12180
      Width           =   17115
      Begin VB.ComboBox cmbRes 
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
         Index           =   27
         ItemData        =   "frmSkinScanWnd.frx":08D8
         Left            =   13500
         List            =   "frmSkinScanWnd.frx":08DA
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   54
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   26
         ItemData        =   "frmSkinScanWnd.frx":08DC
         Left            =   13500
         List            =   "frmSkinScanWnd.frx":08DE
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   52
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   25
         ItemData        =   "frmSkinScanWnd.frx":08E0
         Left            =   13500
         List            =   "frmSkinScanWnd.frx":08E2
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   50
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   24
         ItemData        =   "frmSkinScanWnd.frx":08E4
         Left            =   10860
         List            =   "frmSkinScanWnd.frx":08E6
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   48
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   23
         ItemData        =   "frmSkinScanWnd.frx":08E8
         Left            =   10860
         List            =   "frmSkinScanWnd.frx":08EA
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   46
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   22
         ItemData        =   "frmSkinScanWnd.frx":08EC
         Left            =   10860
         List            =   "frmSkinScanWnd.frx":08EE
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   44
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   21
         ItemData        =   "frmSkinScanWnd.frx":08F0
         Left            =   8160
         List            =   "frmSkinScanWnd.frx":08F2
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   42
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   20
         ItemData        =   "frmSkinScanWnd.frx":08F4
         Left            =   8160
         List            =   "frmSkinScanWnd.frx":08F6
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   40
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   19
         ItemData        =   "frmSkinScanWnd.frx":08F8
         Left            =   8160
         List            =   "frmSkinScanWnd.frx":08FA
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   38
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   18
         ItemData        =   "frmSkinScanWnd.frx":08FC
         Left            =   5520
         List            =   "frmSkinScanWnd.frx":08FE
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   36
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   17
         ItemData        =   "frmSkinScanWnd.frx":0900
         Left            =   5520
         List            =   "frmSkinScanWnd.frx":0902
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   34
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   16
         ItemData        =   "frmSkinScanWnd.frx":0904
         Left            =   5520
         List            =   "frmSkinScanWnd.frx":0906
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   32
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   15
         ItemData        =   "frmSkinScanWnd.frx":0908
         Left            =   2880
         List            =   "frmSkinScanWnd.frx":090A
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   30
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   14
         ItemData        =   "frmSkinScanWnd.frx":090C
         Left            =   2880
         List            =   "frmSkinScanWnd.frx":090E
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   28
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   13
         ItemData        =   "frmSkinScanWnd.frx":0910
         Left            =   2880
         List            =   "frmSkinScanWnd.frx":0912
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   25
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   12
         ItemData        =   "frmSkinScanWnd.frx":0914
         Left            =   240
         List            =   "frmSkinScanWnd.frx":0916
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   23
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   11
         ItemData        =   "frmSkinScanWnd.frx":0918
         Left            =   240
         List            =   "frmSkinScanWnd.frx":091A
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   21
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
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
         Index           =   10
         ItemData        =   "frmSkinScanWnd.frx":091C
         Left            =   240
         List            =   "frmSkinScanWnd.frx":091E
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   19
         Top             =   780
         Width           =   1755
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   11
         Left            =   1980
         TabIndex        =   22
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":098E
         Key             =   "frmSkinScanWnd.frx":09AC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   12
         Left            =   1980
         TabIndex        =   24
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":09E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0A4E
         Key             =   "frmSkinScanWnd.frx":0A6C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   13
         Left            =   4620
         TabIndex        =   26
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0AA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0B0E
         Key             =   "frmSkinScanWnd.frx":0B2C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   14
         Left            =   4620
         TabIndex        =   29
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0B60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0BCE
         Key             =   "frmSkinScanWnd.frx":0BEC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   15
         Left            =   4620
         TabIndex        =   31
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0C20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0C8E
         Key             =   "frmSkinScanWnd.frx":0CAC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   16
         Left            =   7260
         TabIndex        =   33
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0CE0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0D4E
         Key             =   "frmSkinScanWnd.frx":0D6C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   17
         Left            =   7260
         TabIndex        =   35
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0DA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0E0E
         Key             =   "frmSkinScanWnd.frx":0E2C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   18
         Left            =   7260
         TabIndex        =   37
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0E60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0ECE
         Key             =   "frmSkinScanWnd.frx":0EEC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   19
         Left            =   9960
         TabIndex        =   39
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0F20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":0F8E
         Key             =   "frmSkinScanWnd.frx":0FAC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   20
         Left            =   9960
         TabIndex        =   41
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":0FE0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":104E
         Key             =   "frmSkinScanWnd.frx":106C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   21
         Left            =   9960
         TabIndex        =   43
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":10A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":110E
         Key             =   "frmSkinScanWnd.frx":112C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   22
         Left            =   12600
         TabIndex        =   45
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":1160
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":11CE
         Key             =   "frmSkinScanWnd.frx":11EC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   23
         Left            =   12600
         TabIndex        =   47
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":1220
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":128E
         Key             =   "frmSkinScanWnd.frx":12AC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   24
         Left            =   12600
         TabIndex        =   49
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":12E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":134E
         Key             =   "frmSkinScanWnd.frx":136C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   25
         Left            =   15240
         TabIndex        =   51
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":13A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":140E
         Key             =   "frmSkinScanWnd.frx":142C
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   26
         Left            =   15240
         TabIndex        =   53
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":1460
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":14CE
         Key             =   "frmSkinScanWnd.frx":14EC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   27
         Left            =   15240
         TabIndex        =   55
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":1520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":158E
         Key             =   "frmSkinScanWnd.frx":15AC
         BackColor       =   -2147483643
         EditMode        =   0
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   10
         Left            =   1980
         TabIndex        =   20
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSkinScanWnd.frx":15E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSkinScanWnd.frx":164E
         Key             =   "frmSkinScanWnd.frx":166C
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "99"
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   2
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "��������"
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
         Index           =   13
         Left            =   10320
         TabIndex        =   89
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�s�n�o"
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
         Left            =   13380
         TabIndex        =   88
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�a�n�s"
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
         Left            =   10680
         TabIndex        =   87
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�m��"
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
         Left            =   8040
         TabIndex        =   86
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�r��"
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
         Left            =   5340
         TabIndex        =   85
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "�v��"
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
         Index           =   3
         Left            =   2700
         TabIndex        =   84
         Top             =   480
         Width           =   825
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
         Index           =   2
         Left            =   60
         TabIndex        =   83
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "���ׁi��ށE���j"
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
         TabIndex        =   82
         Top             =   180
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "S"
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
         Index           =   14
         Left            =   10620
         TabIndex        =   81
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "S"
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
         Index           =   15
         Left            =   13260
         TabIndex        =   80
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "M"
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
         Index           =   16
         Left            =   10620
         TabIndex        =   79
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "M"
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
         Index           =   17
         Left            =   13260
         TabIndex        =   78
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "N"
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
         Index           =   19
         Left            =   10620
         TabIndex        =   77
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  '�E����
         Caption         =   "N"
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
         Index           =   20
         Left            =   13260
         TabIndex        =   76
         Top             =   1800
         Width           =   225
      End
   End
   Begin VB.Label lblPhotoCnt 
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
      Left            =   18240
      TabIndex        =   95
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lblPhotoCntTitle 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
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
      Index           =   26
      Left            =   17400
      TabIndex        =   94
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "�������X�L�����C���[�W"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   92
      Top             =   4020
      Width           =   4275
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      BackStyle       =   0  '����
      Caption         =   "�X���u����������"
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
      Index           =   0
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   17235
   End
   Begin VB.Label lblInputMode 
      Caption         =   "�y�V�K�z"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   90
      Top             =   540
      Width           =   1395
   End
End
Attribute VB_Name = "frmSkinScanWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSkinScanWnd.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' �J���[�`�F�b�N���тo�b�@�X�L���i�[�Ǎ��݃t�H�[��
' �@�{���W���[���̓X�L���i�[�Ǎ��݃t�H�[���Ŏg�p����
' �@���߂̂��̂ł���B

Option Explicit

Private nPreBkColor As Long         ''���O�̔w�i�F

Private bUpdateImageFlg As Boolean ''�C���[�W�ω��L�薳���t���O

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
    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETSKINSCANWND, CALLBACK_ncResCANCEL)
End Sub

Private Sub cmdNextProc_Click()
    frmSrvNextProcess.SetCallBack Me, CALLBACK_NEXTPROCWND
    frmSrvNextProcess.Show vbModal, Me '�T�[�o�[�f�[�^�ǉ��^�폜���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
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

    Call DBSendDataReq_SKIN

'    Unload Me
'    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETSKINSCANWND, CALLBACK_ncResOK) '�����p��

End Sub

' @(f)
'
' �@�\      : �S�̕\���{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �S�̕\���{�^�������B
'
' ���l      :
'
Private Sub cmdFullImage_Click()
    
    frmFullImage.SetCallBack Me, CALLBACK_FULLSCANIMAGEWND
    frmFullImage.LEAD1.Bitmap = LEAD1.Bitmap
    frmFullImage.LEAD1.PaintSizeMode = PAINTSIZEMODE_FIT '�����`�̑傫�����ő�ɂȂ�悤�ɁA�N���C�A���g�̈�̕��������̂����ꂩ�ɍ��킹�A�c��̃T�C�Y�̓A�X�y�N�g����ێ�����悤�ɒ��߂��܂��B
    frmFullImage.Show vbModal, Me '�T�[�o�[�f�[�^�ǉ��^�폜���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B

End Sub

' @(f)
'
' �@�\      :�ʐ^�Y�t�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �ʐ^�Y�t�{�^�������B
'
' ���l      :
'
Private Sub cmdPhotoImgUp_Click()
    frmPhotoImgUpView.SetCallBack Me, CALLBACK_PHOTOIMGUPWND
    On Error Resume Next '�����I���̏ꍇ�̉��
    frmPhotoImgUpView.Show vbModal, Me '�T�[�o�[�f�[�^�ǉ��^�폜���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
    On Error GoTo 0
End Sub

' @(f)
'
' �@�\      : �X�^�b�t���o�^�{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�^�b�t���o�^�{�^�������B
'
' ���l      :
'           :COLORSYS
'
Private Sub cmdUser_Click()
    frmOpRegWnd.SetCallBack Me, CALLBACK_OPREGWND
    frmOpRegWnd.Show vbModal, Me '�T�[�o�[�f�[�^�ǉ��^�폜���́A���̏�����s�Ƃ���ׁAvbModal�Ƃ���B
End Sub

' @(f)
'
' �@�\      : �X�L���i�[�ǂݎ�芮��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L���i�[�ǂݎ�芮�����̏����B
'
' ���l      :
'
Private Sub LEAD_SCAN_TwainPage()
    Dim lBitMapDC As Long
    Dim nJ As Integer
    
    If IsDEBUG("SCAN") Then
    Else
        '��ʂɃC���[�W�����݂��邩�B
        If LEAD_SCAN.Bitmap <> 0 Then
            If APSysCfgData.nIMAGE_ROTATE(conDefine_SYSMODE_SKIN) <> 0 Then
                LEAD_SCAN.FastRotate APSysCfgData.nIMAGE_ROTATE(conDefine_SYSMODE_SKIN)
            End If
        End If
    End If
    
    On Error Resume Next
    
    lBitMapDC = LEAD_SCAN.GetBitmapDC
    
    On Error GoTo 0
    
'    For nJ = 0 To 1
        LEAD1.Capture lBitMapDC, APSysCfgData.nIMAGE_LEFT(conDefine_SYSMODE_SKIN), APSysCfgData.nIMAGE_TOP(conDefine_SYSMODE_SKIN), _
                                                APSysCfgData.nIMAGE_WIDTH(conDefine_SYSMODE_SKIN), APSysCfgData.nIMAGE_HEIGHT(conDefine_SYSMODE_SKIN)
'    Next nJ
    
    LEAD_SCAN.ReleaseBitmapDC
    
    '�ǂݎ�肪�����̊m�F�͕K�v�Ȃ��B
    'Dim MsgWnd As Message
    'Set MsgWnd = New Message
    
    'MsgWnd.MsgText = "�X�L���i�[�ǂݎ�肪�������܂����B" & vbCrLf
    'MsgWnd.OK.Visible = False
    
    '�ǂݎ�肪�����̊m�F�͕K�v�Ȃ��B
    Call MsgLog(conProcNum_MAIN, "�X�L���i�[�ǂݎ�肪�������܂����B" & vbCrLf) '�K�C�_���X�\��
    'MsgWnd.Show vbModeless, Me
    'MsgWnd.Refresh
    'DoEvents
    'MsgWnd.OK.Visible = True
    
    '
    'Call LEAD1.Save(App.Path & "\" & conDefine_ImageDirName & "\" & "SCAN" & Format(nNowSplitNum, "00") & "(0).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    Call LEAD1.Save(App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.JPG", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
'    Call LEAD1(1).Save(App.Path & "\" & conDefine_ImageDirName & "\" & "SCAN" & Format(nNowSplitNum, "00") & "(1).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    
    Call ButtonEnable(True)
    
    bUpdateImageFlg = True '�C���[�W�ω��L��B
    
End Sub

' @(f)
'
' �@�\      : �\�����C���[�W�̉�]
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �\�����C���[�W�̉�]���s���B
'
' ���l      : �i���g�p�j
'
Private Sub cmdRotate_Click()
    '��ʂɃC���[�W�����݂��邩�B
    If LEAD1.Bitmap <> 0 Then
        LEAD1.FastRotate 90
    End If
End Sub

' @(f)
'
' �@�\      : �s�h�e�t�@�C���ۑ�
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �s�h�e�t�@�C���֕ۑ����s���B
'
' ���l      : �i���g�p�j
'
Private Sub cmdSaveTIF_Click()
    Debug.Print LEAD_SCAN.Save("d:\TEST.jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    Debug.Print LEAD1.Save("d:\TEST(0).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
'    Debug.Print LEAD1(1).Save("d:\TEST(1).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
End Sub

' @(f)
'
' �@�\      : �X�L���i�[�Ǎ��{�^��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L���i�[�Ǎ��{�^�������B
'
' ���l      :
'
Private Sub cmdScan_Click()
        Dim fmessage As Object
        Set fmessage = New MessageYN
        fmessage.MsgText = "�X�L���i�[�Ǎ����J�n���܂��B" & vbCrLf & "�����͂�낵���ł����H"
        fmessage.AutoDelete = True
        fmessage.SetCallBack Me, CALLBACK_GETIMGDATA, True
            Do
                On Error Resume Next
                fmessage.Show vbModeless, Me
                If Err.Number = 0 Then
                    Exit Do
                End If
                DoEvents
            Loop
        Set fmessage = Nothing

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
    Dim nI As Integer
    
    Select Case CallNo
    
    Case CALLBACK_GETIMGDATA
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next
            DoEvents
            Call ImageScan
            'On Error GoTo 0
        End If
    
    Case CALLBACK_OPREGWND 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next

            cmbRes(0).Clear
            For nI = 1 To UBound(APStaffData)
                cmbRes(0).AddItem APStaffData(nI - 1).inp_StaffName
'                cmbRes(0).ListIndex = nI - 1
            Next nI

'            Call InitForm
            'On Error GoTo 0
        End If
    
    Case CALLBACK_NEXTPROCWND 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next

            cmbRes(1).Clear
            For nI = 1 To UBound(APNextProcDataSkin)
                cmbRes(1).AddItem APNextProcDataSkin(nI - 1).inp_NextProc
                
            Next nI
 
'            Call InitForm
            'On Error GoTo 0
        End If
    
    'SKIN���уf�[�^�̓o�^�₢���킹OK
    Case CALLBACK_RES_DBSNDDATA_SKIN
            If Result = CALLBACK_ncResOK Then          'OK
                
                ''DB�ۑ�����
                Call SetAPResData
                
                '�J�����g���ѓ��͏��ꎞ�ۑ�
                APResDataBK = APResData
                
                '/* DB�o�^���s */
                bRet = DB_SAVE_SKIN()
                
                If bRet Then
                    '�c�a�ۑ�����I���̏ꍇ
                    '�o�c�e�쐬�v���ʒm
                    APResData.slb_col_cnt = "00"
                    frmTRSend.SetCallBack Me, CALLBACK_TRSEND, "COL01"
                    frmTRSend.Show vbModal, Me
                Else
                    Call WaitMsgBox(Me, "���M�^�c�a�ۑ������𒆒f���܂����B")
                End If
                
            Else
                'DB�o�^�L�����Z��
            End If
    
    Case CALLBACK_TRSEND
            If Result = CALLBACK_ncResOK Then          'OK
                Call WaitMsgBox(Me, "�c�a�ۑ�������I�����܂����B")
            Else
                Call WaitMsgBox(Me, "�o�c�e�쐬�v���͎��s���܂������A" & vbCrLf & "�c�a�ۑ��͐���I�����܂����B")
            End If
    
            '�쐬�v���̂n�j�^�m�f�ɂ�����炸����I��
            Unload Me
            Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETSKINSCANWND, CALLBACK_ncResOK) 'OK�ŏ����I��
    
    Case CALLBACK_PHOTOIMGUPWND
            If Result = CALLBACK_ncResOK Then          'OK
            Else
                ' 20090115 add by M.Aoyagi
                lblPhotoCnt.Caption = APResData.PhotoImgCnt
            End If
    
    End Select

End Sub

' @(f)
'
' �@�\      : �{�^���R���g���[��
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �{�^���R���g���[�������B
'
' ���l      :
'
Private Sub ButtonEnable(ByVal bEnable As Boolean)
    cmdScan.Enabled = bEnable
    cmdFullImage.Enabled = bEnable
    cmdOK.Enabled = bEnable
    cmdCANCEL.Enabled = bEnable
End Sub

' @(f)
'
' �@�\      : �X�L���i�[�ǎ�J�n
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L���i�[�ǎ�J�n�����B
'
' ���l      :
'
Private Sub ImageScan()
    Dim nRet As Integer
    Dim Msg As String

    On Error Resume Next
    Call ButtonEnable(False)
    
    If IsDEBUG("SCAN") Then
        
        Dim MsgWnd As Message
        Set MsgWnd = New Message
        
        MsgWnd.MsgText = "�X�L���i�[�ǂݍ��ݒ��ł��B" & vbCrLf & "���΂炭���҂����������B"
        MsgWnd.OK.Visible = False
        MsgWnd.Show vbModeless, Me
        MsgWnd.Refresh
        DoEvents
        
        nRet = LEAD_SCAN.Load(App.path & "\TEST_SKIN.jpg", 0, 0, 1)
        
        MsgWnd.OK_Close
        
        Call LEAD_SCAN_TwainPage
    Else
        'nRet = LEAD_SCAN_TWAIN_ACQUIRE()
        nRet = LEAD_SCAN.TwainAcquire(Me.hWnd)
    End If
    On Error GoTo 0
    
    If nRet <> 0 Then
        Msg = "�װ '" & CStr(nRet) & ", " & DecodeError(nRet) & ""
        Call WaitMsgBox(Me, Msg)
        Call ButtonEnable(True)
    End If
End Sub

' @(f)
'
' �@�\      : �X�L���i�[�ǎ�
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X�L���i�[�ǎ揈���B
'
' ���l      :
'
Private Function LEAD_SCAN_TWAIN_ACQUIRE() As Integer
Dim nRet As Integer

Dim MsgWnd As Message
Set MsgWnd = New Message

MsgWnd.MsgText = "�X�L���i�[�ǂݍ��ݒ��ł��B" & vbCrLf & "���΂炭���҂����������B"
MsgWnd.OK.Visible = False
MsgWnd.Show vbModeless, Me
MsgWnd.Refresh
DoEvents

On Error GoTo ERRORHANDLER
'�C���[�W�̎擾���ɁA�\�������`��������`���܂��B
LEAD_SCAN.AutoSetRects = True
'�����ĕ`��𖳌��ɂ��܂��B
LEAD_SCAN.AutoRepaint = False
'TWAIN�\�[�X�}�l�[�W����I�����܂��B

Screen.MousePointer = 11 '�}�E�X�|�C���^�������v��
LEAD_SCAN.TwainEnumSources (hWnd)
Screen.MousePointer = 0 '�}�E�X�|�C���^��W����

LEAD_SCAN.TwainSourceName = LEAD_SCAN.TwainSourceList(0)
Debug.Print LEAD_SCAN.TwainSourceName

'�J�X�^��TWAIN�l��ݒ肵�܂��B
LEAD_SCAN.TwainMaxPages = -1               '�f�t�H���g
LEAD_SCAN.TwainAppAuthor = ""              '�f�t�H���g

LEAD_SCAN.TwainAppFamily = ""              '�f�t�H���g
LEAD_SCAN.TwainFrameLeft = -1              '�f�t�H���g
LEAD_SCAN.TwainFrameTop = -1               '�f�t�H���g
'LEAD_SCAN.TwainFrameWidth = 10080          '7 �C���`
'LEAD_SCAN.TwainFrameHeight = 12960         '9 �C���`
LEAD_SCAN.TwainFrameWidth = -1          '7 �C���`
LEAD_SCAN.TwainFrameHeight = -1         '9 �C���`
LEAD_SCAN.TwainBits = 1                    '1 bit/plane

LEAD_SCAN.TwainPixelType = TWAIN_PIX_HALF  '�����C���[�W

'LEAD_SCAN.TwainPixelType = TWAIN_PIX_GRAY
'LEAD_SCAN.TwainRes = -1                    '�f�t�H���g�𑜓x
LEAD_SCAN.TwainRes = 600                    '�f�t�H���g�𑜓x
LEAD_SCAN.TwainContrast = TWAIN_DEFAULT_CONTRAST        '�f�t�H���g

LEAD_SCAN.TwainIntensity = TWAIN_DEFAULT_INTENSITY      '�f�t�H���g
LEAD_SCAN.EnableTwainFeeder = TWAIN_FEEDER_DEFAULT      '�f�t�H���g
LEAD_SCAN.EnableTwainAutoFeed = TWAIN_AUTOFEED_DEFAULT  '�f�t�H���g
'TwainRealize���\�b�h�����s���A
'�ݒ���e���m���ɔ��f���ꂽ���m�F���܂��B
Screen.MousePointer = 11 '�}�E�X�|�C���^�������v��
LEAD_SCAN.TwainRealize (hWnd)
Screen.MousePointer = 0 '�}�E�X�|�C���^��W����
'TWAIN�C���^�[�t�F�[�X���\���ɂ��A�C���[�W���擾���܂��B
LEAD_SCAN.TwainFlags = 0

nRet = LEAD_SCAN.TwainAcquire(hWnd)

LEAD_SCAN_TWAIN_ACQUIRE = nRet

MsgWnd.OK_Close

Exit Function
ERRORHANDLER:
Call MsgLog(conProcNum_MAIN, Err.Source + " " + _
    CStr(Err.Number) + Chr(13) + Err.Description)

LEAD_SCAN_TWAIN_ACQUIRE = Err.Number

MsgWnd.OK_Close

Call ButtonEnable(True)

End Function

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
    
''    Call clrImgFile("SCAN")
    
    bUpdateImageFlg = False '�C���[�W�ω��������Z�b�g�B

    LEAD1.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD1.EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    LEAD1.EnableTwainEvent = True
    LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)

    LEAD_SCAN.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD_SCAN.EnableMethodErrors = False 'False   �V�X�e���G���[�C�x���g�𔭐������Ȃ�
    LEAD_SCAN.EnableTwainEvent = True
    LEAD_SCAN.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)

    Call GetCurrentAPSlbData
    
    timOpening.Interval = 500
    timOpening.Enabled = True

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
    Dim nJ As Integer
    Dim nRet As Integer
    
    Dim strDestination As String

    '�Ǎ��ݍς݃C���[�W�f�[�^������ꍇ�\������ 'nBitmapListIndexP1 �O�F������ �|�P�F�X�L�b�v
    strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.JPG"
    If Dir(strDestination) <> "" Then
        nRet = LEAD1.Load(App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.jpg", 0, 0, 1)
    End If

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

    lblInputMode.Caption = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, "�y�V�K�z", "�y�C���z")
    lblInputMode.Refresh

    '�J�����g�X���u��񃍁[�h
    lblSlb(0).Caption = APResData.slb_chno & "-" & APResData.slb_aino ''�X���uNo
    lblSlb(1).Caption = ConvDpOutStat(conDefine_SYSMODE_SKIN, CInt(APResData.slb_stat)) ''���
    lblSlb(2).Caption = APResData.sys_wrt_dte ''�L�^��

    '2008/09/01 SystEx. A.K
    imSozai(0).Text = APResData.slb_zkai_dte ''�����
    imSozai(1).Text = APResData.slb_ksh ''�|��
    imSozai(2).Text = APResData.slb_ccno ''CCNo
    imSozai(3).Text = APResData.slb_typ ''�^
    imSozai(4).Text = APResData.slb_uksk ''����
    imSozai(5).Text = APResData.slb_wei ''�d��
    imSozai(6).Text = APResData.slb_thkns ''����
    imSozai(7).Text = APResData.slb_wdth ''��
    imSozai(8).Text = APResData.slb_lngth ''����

    '�X�^�b�t�����X�gBOX�ݒ�
    cmbRes(0).Clear
    For nJ = 1 To UBound(APStaffData)
        cmbRes(0).AddItem APStaffData(nJ - 1).inp_StaffName
        If APResData.slb_wrt_nme = APStaffData(nJ - 1).inp_StaffName Then
            cmbRes(0).ListIndex = nJ - 1
        End If
    Next nJ

    '���H�����X�gBOX�ݒ�
    cmbRes(1).Clear
    For nJ = 1 To UBound(APNextProcDataSkin)
        cmbRes(1).AddItem APNextProcDataSkin(nJ - 1).inp_NextProc
        If APResData.slb_nxt_prcs = APNextProcDataSkin(nJ - 1).inp_NextProc Then
            cmbRes(1).ListIndex = nJ - 1
        End If
    Next nJ

    '�R�����g��񃍁[�h
    imText(0).Text = APResData.slb_cmt1
    imText(1).Text = APResData.slb_cmt2

    '�ʌ��׃��X�gBOX�ݒ�
    nI = 10
    'E-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_e_s1 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_e_n1
    nI = nI + 1

    'E-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_e_s2 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_e_n2
    nI = nI + 1

    'E-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_e_s3 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_e_n3
    nI = nI + 1

    'W-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_w_s1 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_w_n1
    nI = nI + 1

    'W-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_w_s2 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_w_n2
    nI = nI + 1

    'W-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_w_s3 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_w_n3
    nI = nI + 1

    'S-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_s_s1 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_s_n1
    nI = nI + 1

    'S-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_s_s2 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_s_n2
    nI = nI + 1

    'S-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_s_s3 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_s_n3
    nI = nI + 1

    'N-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_n_s1 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_n_n1
    nI = nI + 1

    'N-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_n_s2 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_n_n2
    nI = nI + 1

    'N-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceSkin)
        cmbRes(nI).AddItem APFaultFaceSkin(nJ - 1).strName
        If APResData.slb_fault_n_s3 = APFaultFaceSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_n_n3
    nI = nI + 1

    '�������׃��X�gBOX�ݒ�
    nI = 22
    'B-S
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_bs_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_bs_n
    nI = nI + 1

    'B-M
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_bm_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_bm_n
    nI = nI + 1

    'B-N
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_bn_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_bn_n
    nI = nI + 1

    'T-S
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_ts_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_ts_n
    nI = nI + 1

    'T-M
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_tm_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_tm_n
    nI = nI + 1

    'T-N
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultInsideSkin)
        cmbRes(nI).AddItem APFaultInsideSkin(nJ - 1).strName
        If APResData.slb_fault_tn_s = APFaultInsideSkin(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    
    imText(nI).Text = APResData.slb_fault_tn_n
    nI = nI + 1

    ' 20090115 add by M.Aoyagi    �摜�����ǉ�
    lblPhotoCnt.Caption = APResData.PhotoImgCnt

End Sub

' @(f)
'
' �@�\      : �\������p�^�C�}�[�C�x���g
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �\������p�^�C�}�[�C�x���g���̏������s���B
'
' ���l      :
'
Private Sub timOpening_Timer()
    timOpening.Enabled = False
    Call InitForm
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
Private Sub DBSendDataReq_SKIN()
    Dim fmessage As Object
    Set fmessage = New MessageYN

    '�o�^�ɕK�v�ȃC���[�W�Ǝ��ѓ��̓f�[�^�����݂��邩�B
'    If CheckAPInputComplete() Then
    fmessage.MsgText = "�X���u���������͂̎��уf�[�^��o�^���܂��B" & vbCrLf & "��낵���ł����H"
'    fmessage.AutoDelete = True
    fmessage.AutoDelete = False
'    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SKIN, True
    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SKIN, False
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
' �@�\      : ���͏��ݒ�
'
' ������    :
'
' �Ԃ�l    :
'
' �@�\����  : �X���u���̐ݒ���s���B
'
' ���l      :
'
Private Sub SetAPResData()
    
    Dim nI As Integer
    Dim nJ As Integer
    Dim bRet As Boolean
    Dim bFault As Boolean
    
    ''DB�ۑ�����
    
    '�X�^�b�t��
    APResData.slb_wrt_nme = cmbRes(0).Text
    
    '���H��
    APResData.slb_nxt_prcs = cmbRes(1).Text
    
    '2008/09/01 SystEx. A.K ���݃f�[�^��ێ�����B
    APSysCfgData.NowStaffName(conDefine_SYSMODE_SKIN) = APResData.slb_wrt_nme '�X�^�b�t��
    APSysCfgData.NowNextProcess(conDefine_SYSMODE_SKIN) = APResData.slb_nxt_prcs '���H��
    
    '2008/09/01 SystEx. A.K
    APResData.slb_zkai_dte = imSozai(0).Text ''�����
    APResData.slb_ksh = imSozai(1).Text ''�|��
    APResData.slb_ccno = imSozai(2).Text ''CCNo
    APResData.slb_typ = imSozai(3).Text ''�^
    APResData.slb_uksk = imSozai(4).Text ''����
    APResData.slb_wei = imSozai(5).Text ''�d��
    APResData.slb_thkns = imSozai(6).Text ''����
    APResData.slb_wdth = imSozai(7).Text ''��
    APResData.slb_lngth = imSozai(8).Text ''����
    
    '�R�����g�P
    APResData.slb_cmt1 = imText(0).Text
    
    '�R�����g�Q
    APResData.slb_cmt2 = imText(1).Text
    
    ''����o�^���t��ݒ�
    If APResData.sys_wrt_dte = "" Then
        APResData.sys_wrt_dte = Format(Now, "YYYYMMDD")
        APResData.sys_wrt_tme = Format(Now, "HHMMSS")
    End If
    
    '���ׂ�ݒ�
    nI = 10
    'E1��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s1 = ""
        APResData.slb_fault_cd_e_s1 = ""
    Else
        APResData.slb_fault_e_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n1 = ""
    End If
    nI = nI + 1
    
    'E2��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s2 = ""
        APResData.slb_fault_cd_e_s2 = ""
    Else
        APResData.slb_fault_e_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n2 = ""
    End If
    nI = nI + 1
    
    'E3��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s3 = ""
        APResData.slb_fault_cd_e_s3 = ""
    Else
        APResData.slb_fault_e_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n3 = ""
    End If
    nI = nI + 1
        
    'W1��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s1 = ""
        APResData.slb_fault_cd_w_s1 = ""
    Else
        APResData.slb_fault_w_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n1 = ""
    End If
    nI = nI + 1
    
    'W2��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s2 = ""
        APResData.slb_fault_cd_w_s2 = ""
    Else
        APResData.slb_fault_w_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n2 = ""
    End If
    nI = nI + 1
    
    'W3��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s3 = ""
        APResData.slb_fault_cd_w_s3 = ""
    Else
        APResData.slb_fault_w_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n3 = ""
    End If
    nI = nI + 1
    
    'S1��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s1 = ""
        APResData.slb_fault_cd_s_s1 = ""
    Else
        APResData.slb_fault_s_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n1 = ""
    End If
    nI = nI + 1
    
    'S2��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s2 = ""
        APResData.slb_fault_cd_s_s2 = ""
    Else
        APResData.slb_fault_s_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n2 = ""
    End If
    nI = nI + 1
    
    'S3��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s3 = ""
        APResData.slb_fault_cd_s_s3 = ""
    Else
        APResData.slb_fault_s_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n3 = ""
    End If
    nI = nI + 1
        
    'N1��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s1 = ""
        APResData.slb_fault_cd_n_s1 = ""
    Else
        APResData.slb_fault_n_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s1 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n1 = ""
    End If
    nI = nI + 1
    
    'N2��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s2 = ""
        APResData.slb_fault_cd_n_s2 = ""
    Else
        APResData.slb_fault_n_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s2 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n2 = ""
    End If
    nI = nI + 1
    
    'N3��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s3 = ""
        APResData.slb_fault_cd_n_s3 = ""
    Else
        APResData.slb_fault_n_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s3 = APFaultFaceSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n3 = ""
    End If
    nI = nI + 1
    
    '�������׃��X�gBOX�擾
    'BS��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_bs_s = ""
        APResData.slb_fault_cd_bs_s = ""
    Else
        APResData.slb_fault_bs_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_bs_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_bs_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_bs_n = ""
    End If
    nI = nI + 1
    
    'BM��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_bm_s = ""
        APResData.slb_fault_cd_bm_s = ""
    Else
        APResData.slb_fault_bm_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_bm_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_bm_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_bm_n = ""
    End If
    nI = nI + 1
    
    'BN��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_bn_s = ""
        APResData.slb_fault_cd_bn_s = ""
    Else
        APResData.slb_fault_bn_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_bn_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_bn_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_bn_n = ""
    End If
    nI = nI + 1
    
    'TS��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_ts_s = ""
        APResData.slb_fault_cd_ts_s = ""
    Else
        APResData.slb_fault_ts_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_ts_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_ts_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_ts_n = ""
    End If
    nI = nI + 1
    
    'TM��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_tm_s = ""
        APResData.slb_fault_cd_tm_s = ""
    Else
        APResData.slb_fault_tm_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_tm_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_tm_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_tm_n = ""
    End If
    nI = nI + 1
    
    'TN��
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_tn_s = ""
        APResData.slb_fault_cd_tn_s = ""
    Else
        APResData.slb_fault_tn_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_tn_s = APFaultInsideSkin(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_tn_n = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_tn_n = ""
    End If
    nI = nI + 1
    
    
    '���ה����ݒ�
    'E����
    bFault = False '���ז���
    Do While True
        'E1
        If IsNumeric(APResData.slb_fault_e_n1) Then
            If APResData.slb_fault_e_s1 <> "" And CInt(APResData.slb_fault_e_n1) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'E2
        If IsNumeric(APResData.slb_fault_e_n2) Then
            If APResData.slb_fault_e_s2 <> "" And CInt(APResData.slb_fault_e_n2) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'E3
        If IsNumeric(APResData.slb_fault_e_n3) Then
            If APResData.slb_fault_e_s3 <> "" And CInt(APResData.slb_fault_e_n3) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_e_judg = IIf(bFault, "1", "0")
    
    'W����
    bFault = False '���ז���
    Do While True
        'W1
        If IsNumeric(APResData.slb_fault_w_n1) Then
            If APResData.slb_fault_w_s1 <> "" And CInt(APResData.slb_fault_w_n1) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'W2
        If IsNumeric(APResData.slb_fault_w_n2) Then
            If APResData.slb_fault_w_s2 <> "" And CInt(APResData.slb_fault_w_n2) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'W3
        If IsNumeric(APResData.slb_fault_w_n3) Then
            If APResData.slb_fault_w_s3 <> "" And CInt(APResData.slb_fault_w_n3) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_w_judg = IIf(bFault, "1", "0")
    
    'N����
    bFault = False '���ז���
    Do While True
        'N1
        If IsNumeric(APResData.slb_fault_n_n1) Then
            If APResData.slb_fault_n_s1 <> "" And CInt(APResData.slb_fault_n_n1) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'N2
        If IsNumeric(APResData.slb_fault_n_n2) Then
            If APResData.slb_fault_n_s2 <> "" And CInt(APResData.slb_fault_n_n2) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'N3
        If IsNumeric(APResData.slb_fault_n_n3) Then
            If APResData.slb_fault_n_s3 <> "" And CInt(APResData.slb_fault_n_n3) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_n_judg = IIf(bFault, "1", "0")
    
    'S����
    bFault = False '���ז���
    Do While True
        'S1
        If IsNumeric(APResData.slb_fault_s_n1) Then
            If APResData.slb_fault_s_s1 <> "" And CInt(APResData.slb_fault_s_n1) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'S2
        If IsNumeric(APResData.slb_fault_s_n2) Then
            If APResData.slb_fault_s_s2 <> "" And CInt(APResData.slb_fault_s_n2) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'S3
        If IsNumeric(APResData.slb_fault_s_n3) Then
            If APResData.slb_fault_s_s3 <> "" And CInt(APResData.slb_fault_s_n3) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_s_judg = IIf(bFault, "1", "0")
    
    'B����
    bFault = False '���ז���
    Do While True
        'B1
        If IsNumeric(APResData.slb_fault_bs_n) Then
            If APResData.slb_fault_bs_s <> "" And CInt(APResData.slb_fault_bs_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'B2
        If IsNumeric(APResData.slb_fault_bm_n) Then
            If APResData.slb_fault_bm_s <> "" And CInt(APResData.slb_fault_bm_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'B3
        If IsNumeric(APResData.slb_fault_bn_n) Then
            If APResData.slb_fault_bn_s <> "" And CInt(APResData.slb_fault_bn_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_b_judg = IIf(bFault, "1", "0")
    
    'T����
    bFault = False '���ז���
    Do While True
        'T1
        If IsNumeric(APResData.slb_fault_ts_n) Then
            If APResData.slb_fault_ts_s <> "" And CInt(APResData.slb_fault_ts_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'T2
        If IsNumeric(APResData.slb_fault_tm_n) Then
            If APResData.slb_fault_tm_s <> "" And CInt(APResData.slb_fault_tm_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        'T3
        If IsNumeric(APResData.slb_fault_tn_n) Then
            If APResData.slb_fault_tn_s <> "" And CInt(APResData.slb_fault_tn_n) > 0 Then
                bFault = True
                Exit Do
            End If
        End If
        Exit Do
    Loop
    APResData.slb_fault_t_judg = IIf(bFault, "1", "0")

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
Private Function DB_SAVE_SKIN() As Boolean
    Dim bNOErrorFlg As Boolean
'    Dim APResDataBK As typAPResData
    Dim nI As Integer
    Dim nJ As Integer
    Dim strImageFileName As String
    Dim bRet As Boolean
    Dim bFault As Boolean
    Dim strSource As String
    Dim strDestination As String
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message

    MsgWnd.MsgText = "�f�[�^�x�[�X�T�[�o�[�ɕۑ����ł��B" & vbCrLf & "���΂炭���҂����������B"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
'    ''�c�a�I�t���C���ŋ������͂��s�������Ƃ𔻒f����t���O
'    If bAPInputOffline Then
'        MsgWnd.OK_Close
'        bNOErrorFlg = True '�G���[����
'        DB_SAVE_SKIN = bNOErrorFlg
'        Exit Function
'    End If
    
'    '�J�����g���ѓ��͏��ꎞ�ۑ�
'    APResDataBK = APResData
    

    bNOErrorFlg = True '�G���[����

    'TRTS0012 �o�^
    bRet = TRTS0012_Write(False)
    If bRet = False Then
        bNOErrorFlg = False '�G���[�L��
        MsgWnd.OK_Close
        Exit Function
    End If

'    '�����܂ŁA�G���[�����̏ꍇ
'    If bNOErrorFlg Then
'        '�g�����U�N�V�����ʒm����
'        'Call CSTRAN_DB_SAVE_START
'    End If
'
'    '//�o�^���s
'    '//�o�^���s
'
    ''�X�L�����C���[�W��ۑ�
    '�X�L���������C���[�W�����邩�H
    'strDestination
    strSource = App.path & "\" & conDefine_ImageDirName & "\" & "SKIN.JPG"
    If Dir(strSource) <> "" Then
    
        Call MsgLog(conProcNum_MAIN, "�X�L�����C���[�W�t�@�C���i�L��j�F" & strSource) '�K�C�_���X�\��
    
        '�t�H���_�쐬
        On Error Resume Next
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SKIN"
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SKIN" & "\" & APResData.slb_chno
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
        Call MkDir(strDestination)
        On Error GoTo 0
        
        '�t�@�C�����쐬
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SKIN" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SKIN" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_00.JPG"
        
        On Error GoTo DB_SAVE_SKIN_File_err:
        Call FileCopy(strSource, strDestination)
        On Error GoTo 0
    
        'TRTS0050 �o�^
        bRet = TRTS0050_Write(False)
        If bRet = False Then
            bNOErrorFlg = False '�G���[�L��
        End If
    
    Else
        
        Call MsgLog(conProcNum_MAIN, "�X�L�����C���[�W�t�@�C���i�����j�F" & strSource) '�K�C�_���X�\��
        
        '�C���[�W����
        If Dir(strDestination) <> "" Then
            'Kill strDestination
        End If
    
        'TRTS0050 �o�^
        bRet = TRTS0050_Write(True)
        If bRet = False Then
            bNOErrorFlg = False '�G���[�L��
        End If
    
    End If
    
    MsgWnd.OK_Close
    
    DB_SAVE_SKIN = bNOErrorFlg

    Exit Function

DB_SAVE_SKIN_File_err:
    On Error GoTo 0
    
    MsgWnd.OK_Close
    
    Call MsgLog(conProcNum_MAIN, strDestination & ":DB_SAVE_SKIN_File_err:�C���[�W�t�@�C���̕ۑ��Ɏ��s���܂����B") '�K�C�_���X�\��
    
    bNOErrorFlg = False '�G���[�L��
    
    DB_SAVE_SKIN = bNOErrorFlg

End Function


' @(f)
'
' �@�\      : ���ѓ���BOX�t�H�[�J�X�擾
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ���BOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub imText_GotFocus(Index As Integer)
    nPreBkColor = imText(Index).BackColor
    imText(Index).BackColor = conDefine_ColorBKGotFocus '�w�i���F
End Sub

''---
'Private Sub cmbRes_GotFocus(Index As Integer)
'    nPreBkColor = cmbRes(Index).BackColor
'    cmbRes(Index).BackColor = conDefine_ColorBKGotFocus '�w�i���F
'End Sub



' @(f)
'
' �@�\      : ���ѓ���BOX�t�H�[�J�X����
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ���BOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub imText_LostFocus(Index As Integer)
    imText(Index).BackColor = nPreBkColor
End Sub

''---
'Private Sub cmbRes_LostFocus(Index As Integer)
'    cmbRes(Index).BackColor = nPreBkColor
'End Sub

' @(f)
'
' �@�\      : ���ѓ���BOX�ύX
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ���BOX�ύX���̏������s���B
'
' ���l      :
'
'---
Private Sub imText_Change(Index As Integer)
    If Len(imText(Index).Text) = imText(Index).MaxLength Then
        SendKeys "{TAB}", True
    End If
End Sub

Private Sub imText_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

Private Sub cmbRes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

'###########################################################
' @(f)
'
' �@�\      : ���ѓ��́i�f�ށjBOX�t�H�[�J�X�擾
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ��́i�f�ށjBOX�t�H�[�J�X�擾���̏������s���B
'
' ���l      :
'
Private Sub imSozai_GotFocus(Index As Integer)
    nPreBkColor = imSozai(Index).BackColor
    imSozai(Index).BackColor = conDefine_ColorBKGotFocus '�w�i���F
End Sub

' @(f)
'
' �@�\      : ���ѓ��́i�f�ށjBOX�t�H�[�J�X����
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ��́i�f�ށjBOX�t�H�[�J�X���Ŏ��̏������s���B
'
' ���l      :
'
Private Sub imSozai_LostFocus(Index As Integer)
    imSozai(Index).BackColor = nPreBkColor
End Sub

' @(f)
'
' �@�\      : ���ѓ��́i�f�ށjBOX�ύX
'
' ������    : ARG1 - �C���f�b�N�X�ԍ�
'
' �Ԃ�l    :
'
' �@�\����  : ���ѓ��́i�f�ށjBOX�ύX���̏������s���B
'
' ���l      :
'
'---
Private Sub imSozai_Change(Index As Integer)
    If Len(imSozai(Index).Text) = imSozai(Index).MaxLength Then
        SendKeys "{TAB}", True
    End If
End Sub

Private Sub imSozai_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

Private Sub imSozai_KeyPress(Index As Integer, KeyAscii As Integer)
    '���ږ��A����`�F�b�N
    Select Case Index
        Case 1 '�|�� 20080909
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
                'OK
            ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
                'OK
            ElseIf KeyAscii = Asc("-") Then
                'OK
            Else
                'NG
                KeyAscii = 0
            End If
            
        Case 6 '���@XXX.XX
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
                'OK
            ElseIf KeyAscii = Asc(".") Then
                'OK
            Else
                'NG
                KeyAscii = 0
            End If
    End Select
    
End Sub


' @(f)
'
' �@�\      : �e�L�X�g�{�b�N�X�`�F�b�N
'
' ������    : ARG1 - ���ڂ̃C���f�b�N�X
'             ARG2 - �L�����Z���t���O
'
' �Ԃ�l    :
'
' �@�\����  : �e�L�X�g�{�b�N�X�`�F�b�N�p
'
' ���l      :
'
Private Sub imSozai_Validate(Index As Integer, CANCEL As Boolean)
    Dim dAns As Double
    '���ږ��A����`�F�b�N
    Select Case Index
        Case 6 '���@XXX.XX
            If IsNumeric(imSozai(Index).Text) Then
                '���l
                dAns = CDbl(imSozai(Index).Text)
                If dAns > 999.99 Then dAns = 999.99
                If dAns < 0 Then dAns = 0
                imSozai(Index).Text = Format(dAns, "0.00")
            ElseIf Trim(imSozai(Index).Text) = "" Then
                imSozai(Index).Text = ""
            Else
                'NG
                imSozai(Index).Text = ""
                CANCEL = True
            End If
    End Select

End Sub

'###########################################################

