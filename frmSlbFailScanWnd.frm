VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{00120003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "Ltocx12n.ocx"
Begin VB.Form frmSlbFailScanWnd 
   BackColor       =   &H00C0FFC0&
   Caption         =   "カラーチェック実績入力＆スキャナー読込み"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   WindowState     =   2  '最大化
   Begin VB.CommandButton cmdOK 
      Caption         =   "送信"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17280
      TabIndex        =   44
      Top             =   13560
      Width           =   1755
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "戻る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   61
      Top             =   60
      Width           =   1935
   End
   Begin VB.CommandButton cmdFullImage 
      Caption         =   "全体表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   17280
      TabIndex        =   60
      Top             =   5340
      Width           =   1755
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "スキャナー読込"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   17280
      TabIndex        =   59
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
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         ItemData        =   "frmSlbFailScanWnd.frx":0000
         Left            =   12600
         List            =   "frmSlbFailScanWnd.frx":0002
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   11
         Top             =   1380
         Width           =   4455
      End
      Begin VB.CommandButton cmdNextProc 
         Caption         =   "次工程"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "frmSlbFailScanWnd.frx":0004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0072
         Key             =   "frmSlbFailScanWnd.frx":0090
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
         Format          =   "Ｚ"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   50
         LengthAsByte    =   0
         Text            =   "１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０"
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
         Caption         =   "frmSlbFailScanWnd.frx":00C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0132
         Key             =   "frmSlbFailScanWnd.frx":0150
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
         Format          =   "Ｚ"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   50
         LengthAsByte    =   0
         Text            =   "１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０１２３４５６７８９０"
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
         Caption         =   "frmSlbFailScanWnd.frx":0184
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":01F2
         Key             =   "frmSlbFailScanWnd.frx":0210
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
         Caption         =   "frmSlbFailScanWnd.frx":0254
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":02C2
         Key             =   "frmSlbFailScanWnd.frx":02E0
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
         Caption         =   "frmSlbFailScanWnd.frx":0324
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0392
         Key             =   "frmSlbFailScanWnd.frx":03B0
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
         Caption         =   "frmSlbFailScanWnd.frx":03F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0462
         Key             =   "frmSlbFailScanWnd.frx":0480
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
         Caption         =   "frmSlbFailScanWnd.frx":04C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0532
         Key             =   "frmSlbFailScanWnd.frx":0550
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
         Caption         =   "frmSlbFailScanWnd.frx":0594
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0602
         Key             =   "frmSlbFailScanWnd.frx":0620
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
         Caption         =   "frmSlbFailScanWnd.frx":0664
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":06D2
         Key             =   "frmSlbFailScanWnd.frx":06F0
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
         Caption         =   "frmSlbFailScanWnd.frx":0734
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":07A2
         Key             =   "frmSlbFailScanWnd.frx":07C0
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
         Caption         =   "frmSlbFailScanWnd.frx":0804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0872
         Key             =   "frmSlbFailScanWnd.frx":0890
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
         Alignment       =   1  '右揃え
         Caption         =   "重量"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   72
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "造塊日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   58
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "型"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   57
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "向先"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   56
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "鋼種"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   55
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "ｽﾗﾌﾞNo"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   54
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "47965 - 15"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   53
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "状態"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   52
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "コメント"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   51
         Top             =   1980
         Width           =   2715
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "CCNo"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   50
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "厚"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   49
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "長"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   48
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "幅"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   47
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   12600
         TabIndex        =   46
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "記録日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   45
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Alignment       =   2  '中央揃え
         Caption         =   "（製造条件等記入）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Width           =   2115
      End
   End
   Begin VB.Timer timOpening 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   540
   End
   Begin VB.CommandButton cmdPhotoImgUp 
      BackColor       =   &H00FFFF80&
      Caption         =   "写真添付"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   17280
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1755
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "検査員名："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      ItemData        =   "frmSlbFailScanWnd.frx":08D4
      Left            =   14640
      List            =   "frmSlbFailScanWnd.frx":08D6
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin LEADLib.LEAD LEAD_SCAN 
      Height          =   315
      Left            =   18000
      TabIndex        =   62
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
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   17115
      _Version        =   65539
      _ExtentX        =   30189
      _ExtentY        =   12938
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      ScaleHeight     =   485
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
      TabIndex        =   63
      Top             =   11760
      Width           =   17115
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   21
         ItemData        =   "frmSlbFailScanWnd.frx":08D8
         Left            =   18780
         List            =   "frmSlbFailScanWnd.frx":08DA
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   42
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   20
         ItemData        =   "frmSlbFailScanWnd.frx":08DC
         Left            =   18780
         List            =   "frmSlbFailScanWnd.frx":08DE
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   40
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   19
         ItemData        =   "frmSlbFailScanWnd.frx":08E0
         Left            =   18780
         List            =   "frmSlbFailScanWnd.frx":08E2
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   38
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   18
         ItemData        =   "frmSlbFailScanWnd.frx":08E4
         Left            =   16140
         List            =   "frmSlbFailScanWnd.frx":08E6
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   36
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   17
         ItemData        =   "frmSlbFailScanWnd.frx":08E8
         Left            =   16140
         List            =   "frmSlbFailScanWnd.frx":08EA
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   34
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   16
         ItemData        =   "frmSlbFailScanWnd.frx":08EC
         Left            =   16140
         List            =   "frmSlbFailScanWnd.frx":08EE
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   32
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         ItemData        =   "frmSlbFailScanWnd.frx":08F0
         Left            =   13500
         List            =   "frmSlbFailScanWnd.frx":08F2
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   30
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         ItemData        =   "frmSlbFailScanWnd.frx":08F4
         Left            =   13500
         List            =   "frmSlbFailScanWnd.frx":08F6
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   28
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         ItemData        =   "frmSlbFailScanWnd.frx":08F8
         Left            =   13500
         List            =   "frmSlbFailScanWnd.frx":08FA
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   25
         Top             =   780
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         ItemData        =   "frmSlbFailScanWnd.frx":08FC
         Left            =   10860
         List            =   "frmSlbFailScanWnd.frx":08FE
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   23
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         ItemData        =   "frmSlbFailScanWnd.frx":0900
         Left            =   10860
         List            =   "frmSlbFailScanWnd.frx":0902
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   21
         Top             =   1260
         Width           =   1755
      End
      Begin VB.ComboBox cmbRes 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         ItemData        =   "frmSlbFailScanWnd.frx":0904
         Left            =   10860
         List            =   "frmSlbFailScanWnd.frx":0906
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   19
         Top             =   780
         Width           =   1755
      End
      Begin imText6Ctl.imText imText 
         Height          =   405
         Index           =   11
         Left            =   12600
         TabIndex        =   22
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0908
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0976
         Key             =   "frmSlbFailScanWnd.frx":0994
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
         Left            =   12600
         TabIndex        =   24
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":09C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0A36
         Key             =   "frmSlbFailScanWnd.frx":0A54
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
         Left            =   15240
         TabIndex        =   26
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0A88
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0AF6
         Key             =   "frmSlbFailScanWnd.frx":0B14
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
         Left            =   15240
         TabIndex        =   29
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0B48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0BB6
         Key             =   "frmSlbFailScanWnd.frx":0BD4
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
         Left            =   15240
         TabIndex        =   31
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0C08
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0C76
         Key             =   "frmSlbFailScanWnd.frx":0C94
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
         Left            =   17880
         TabIndex        =   33
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0CC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0D36
         Key             =   "frmSlbFailScanWnd.frx":0D54
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
         Left            =   17880
         TabIndex        =   35
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0D88
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0DF6
         Key             =   "frmSlbFailScanWnd.frx":0E14
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
         Left            =   17880
         TabIndex        =   37
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0E48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0EB6
         Key             =   "frmSlbFailScanWnd.frx":0ED4
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
         Left            =   20580
         TabIndex        =   39
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0F08
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":0F76
         Key             =   "frmSlbFailScanWnd.frx":0F94
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
         Left            =   20580
         TabIndex        =   41
         Top             =   1260
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":0FC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":1036
         Key             =   "frmSlbFailScanWnd.frx":1054
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
         Left            =   20580
         TabIndex        =   43
         Top             =   1740
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":1088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":10F6
         Key             =   "frmSlbFailScanWnd.frx":1114
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
         Left            =   12600
         TabIndex        =   20
         Top             =   780
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "frmSlbFailScanWnd.frx":1148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSlbFailScanWnd.frx":11B6
         Key             =   "frmSlbFailScanWnd.frx":11D4
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
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   21
         Left            =   9960
         TabIndex        =   102
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   21
         Left            =   8220
         TabIndex        =   101
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   20
         Left            =   9960
         TabIndex        =   100
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   20
         Left            =   8220
         TabIndex        =   99
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   19
         Left            =   9960
         TabIndex        =   98
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   19
         Left            =   8220
         TabIndex        =   97
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   18
         Left            =   7320
         TabIndex        =   96
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   18
         Left            =   5580
         TabIndex        =   95
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   17
         Left            =   7320
         TabIndex        =   94
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   17
         Left            =   5580
         TabIndex        =   93
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   16
         Left            =   7320
         TabIndex        =   92
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   16
         Left            =   5580
         TabIndex        =   91
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   4680
         TabIndex        =   90
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   2940
         TabIndex        =   89
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   4680
         TabIndex        =   88
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   2940
         TabIndex        =   87
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   4680
         TabIndex        =   86
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   2940
         TabIndex        =   85
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         Left            =   2040
         TabIndex        =   84
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         Left            =   300
         TabIndex        =   83
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   2040
         TabIndex        =   82
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   300
         TabIndex        =   81
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblimText 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   2040
         TabIndex        =   80
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblcmbRes 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   300
         TabIndex        =   79
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "Ｎ面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   68
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "Ｓ面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   67
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "Ｗ面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   66
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "Ｅ面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   65
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         Caption         =   "欠陥（種類・個数）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   64
         Top             =   180
         Width           =   2745
      End
   End
   Begin VB.Label lblPhotoCntTitle 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "枚数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      TabIndex        =   106
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label lblPhotoCnt 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18240
      TabIndex        =   105
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lblHostSendFlg 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18120
      TabIndex        =   104
      Top             =   120
      Width           =   885
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      Caption         =   "ﾋﾞｼﾞｺﾝ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   17280
      TabIndex        =   103
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblOK 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10620
      TabIndex        =   78
      Top             =   14100
      Width           =   8415
   End
   Begin VB.Label lblDebug 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   17280
      TabIndex        =   77
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label lblDebug 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   17280
      TabIndex        =   76
      Top             =   3180
      Width           =   1725
   End
   Begin VB.Label lblDebug 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   17280
      TabIndex        =   75
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      Caption         =   "カラー回数"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   12780
      TabIndex        =   74
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblSlb 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   14640
      TabIndex        =   73
      Top             =   120
      Width           =   2565
   End
   Begin VB.Label Label2 
      Caption         =   "スラブ異常報告書スキャンイメージ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      TabIndex        =   71
      Top             =   4020
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "スラブ異常報告書入力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      TabIndex        =   70
      Top             =   0
      Width           =   17235
   End
   Begin VB.Label lblInputMode 
      Caption         =   "【新規】"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   69
      Top             =   540
      Width           =   1395
   End
End
Attribute VB_Name = "frmSlbFailScanWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSlbFailScanWnd.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　スラブ異常報告書スキャナー読込みフォーム
' 　本モジュールはスキャナー読込みフォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

Private nPreBkColor As Long         ''直前の背景色

Private bUpdateImageFlg As Boolean ''イメージ変化有り無しフラグ


' @(f)
'
' 機能      : 戻るボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 戻るボタン処理。
'
' 備考      :
'
Private Sub cmdCancel_Click()
    
    Call SetAPResData(False)
    
    Unload Me
    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResCANCEL) '2008/09/04 戻り先変更
End Sub

Private Sub cmdNextProc_Click()
    frmSrvNextProcess.SetCallBack Me, CALLBACK_NEXTPROCWND
    frmSrvNextProcess.Show vbModal, Me 'サーバーデータ追加／削除中は、他の処理を不可とする為、vbModalとする。
End Sub

' @(f)
'
' 機能      : ＯＫボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＯＫボタン処理。
'
' 備考      :
'
Private Sub cmdOK_Click()

    Dim nI As Integer
    Dim nJ As Integer

    Call DBSendDataReq_SLBFAIL

'    Unload Me
'    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) '処理継続 '2008/09/04 戻り先変更

End Sub

' @(f)
'
' 機能      : 全体表示ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 全体表示ボタン処理。
'
' 備考      :
'
Private Sub cmdFullImage_Click()
    
    frmFullImage.SetCallBack Me, CALLBACK_FULLSCANIMAGEWND
    frmFullImage.LEAD1.Bitmap = LEAD1.Bitmap
    frmFullImage.LEAD1.PaintSizeMode = PAINTSIZEMODE_FIT '長方形の大きさが最大になるように、クライアント領域の幅か高さのいずれかに合わせ、残りのサイズはアスペクト比を維持するように調節します。
    frmFullImage.Show vbModal, Me 'サーバーデータ追加／削除中は、他の処理を不可とする為、vbModalとする。

End Sub

' @(f)
'
' 機能      :写真添付ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 写真添付ボタン処理。
'
' 備考      :
'
Private Sub cmdPhotoImgUp_Click()
    frmPhotoImgUpView.SetCallBack Me, CALLBACK_PHOTOIMGUPWND
    On Error Resume Next '強制終了の場合の回避
    frmPhotoImgUpView.Show vbModal, Me 'サーバーデータ追加／削除中は、他の処理を不可とする為、vbModalとする。
    On Error GoTo 0
End Sub

' @(f)
'
' 機能      : スタッフ名登録ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スタッフ名登録ボタン処理。
'
' 備考      :
'           :COLORSYS
'
Private Sub cmdUser_Click()
    frmOpRegWnd.SetCallBack Me, CALLBACK_OPREGWND
    frmOpRegWnd.Show vbModal, Me 'サーバーデータ追加／削除中は、他の処理を不可とする為、vbModalとする。
End Sub

Private Sub lblHostSendFlg_DblClick()
    If APResData.host_send_flg = "1" Then
        '訂正の場合
        APResData.host_send_flg = "0" '新規に変更
    Else
        APResData.host_send_flg = "1" '訂正に変更
    End If

    lblHostSendFlg.Caption = IIf(APResData.host_send_flg = "0", "0:新規", "1:訂正")

End Sub

' @(f)
'
' 機能      : スキャナー読み取り完了
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキャナー読み取り完了時の処理。
'
' 備考      :
'
Private Sub LEAD_SCAN_TwainPage()
    Dim lBitMapDC As Long
    Dim nJ As Integer
    
    If IsDEBUG("SCAN") Then
    Else
        '画面にイメージが存在するか。
        If LEAD_SCAN.Bitmap <> 0 Then
            If APSysCfgData.nIMAGE_ROTATE(conDefine_SYSMODE_SLBFAIL) <> 0 Then
                LEAD_SCAN.FastRotate APSysCfgData.nIMAGE_ROTATE(conDefine_SYSMODE_SLBFAIL)
            End If
        End If
    End If
    
    On Error Resume Next
    
    lBitMapDC = LEAD_SCAN.GetBitmapDC
    
    On Error GoTo 0
    
'    For nJ = 0 To 1
        LEAD1.Capture lBitMapDC, APSysCfgData.nIMAGE_LEFT(conDefine_SYSMODE_SLBFAIL), APSysCfgData.nIMAGE_TOP(conDefine_SYSMODE_SLBFAIL), _
                                                APSysCfgData.nIMAGE_WIDTH(conDefine_SYSMODE_SLBFAIL), APSysCfgData.nIMAGE_HEIGHT(conDefine_SYSMODE_SLBFAIL)
'    Next nJ
    
    LEAD_SCAN.ReleaseBitmapDC
    
    '読み取りが完了の確認は必要なし。
    'Dim MsgWnd As Message
    'Set MsgWnd = New Message
    
    'MsgWnd.MsgText = "スキャナー読み取りが完了しました。" & vbCrLf
    'MsgWnd.OK.Visible = False
    
    '読み取りが完了の確認は必要なし。
    Call MsgLog(conProcNum_MAIN, "スキャナー読み取りが完了しました。" & vbCrLf) 'ガイダンス表示
    'MsgWnd.Show vbModeless, Me
    'MsgWnd.Refresh
    'DoEvents
    'MsgWnd.OK.Visible = True
    
    '
    'Call LEAD1.Save(App.Path & "\" & conDefine_ImageDirName & "\" & "SCAN" & Format(nNowSplitNum, "00") & "(0).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    'Call LEAD1.Save(App.Path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    Call LEAD1.Save(App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
'    Call LEAD1(1).Save(App.Path & "\" & conDefine_ImageDirName & "\" & "SCAN" & Format(nNowSplitNum, "00") & "(1).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    
    Call ButtonEnable(True)
    
    bUpdateImageFlg = True 'イメージ変化有り。
    
End Sub

' @(f)
'
' 機能      : 表示中イメージの回転
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 表示中イメージの回転を行う。
'
' 備考      : （未使用）
'
Private Sub cmdRotate_Click()
    '画面にイメージが存在するか。
    If LEAD1.Bitmap <> 0 Then
        LEAD1.FastRotate 90
    End If
End Sub

' @(f)
'
' 機能      : ＴＩＦファイル保存
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ＴＩＦファイルへ保存を行う。
'
' 備考      : （未使用）
'
Private Sub cmdSaveTIF_Click()
    Debug.Print LEAD_SCAN.Save("d:\TEST.jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
    Debug.Print LEAD1.Save("d:\TEST(0).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
'    Debug.Print LEAD1(1).Save("d:\TEST(1).jpg", FILE_JFIF, 8, 255, SAVE_OVERWRITE)
End Sub

' @(f)
'
' 機能      : スキャナー読込ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキャナー読込ボタン処理。
'
' 備考      :
'
Private Sub cmdScan_Click()
        Dim fmessage As Object
        Set fmessage = New MessageYN
        fmessage.MsgText = "スキャナー読込を開始します。" & vbCrLf & "準備はよろしいですか？"
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
' 機能      : コールバック処理
'
' 引き数    : ARG1 - コールバック番号
'             ARG2 - コールバック戻り
'
' 返り値    :
'
' 機能説明  : コールバック番号と戻りに応じて、次処理を行う。
'
' 備考      :
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
            For nI = 1 To UBound(APInspData)
                cmbRes(0).AddItem APInspData(nI - 1).inp_InspName
'                cmbRes(0).ListIndex = nI - 1
            Next nI

'            Call InitForm
            'On Error GoTo 0
        End If
    
    Case CALLBACK_NEXTPROCWND 'COLORSYS
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next

            cmbRes(1).Clear
            For nI = 1 To UBound(APNextProcDataColor)
                cmbRes(1).AddItem APNextProcDataColor(nI - 1).inp_NextProc
                
            Next nI
 
'            Call InitForm
            'On Error GoTo 0
        End If
    
    'COLOR実績データの登録問い合わせOK
    Case CALLBACK_RES_DBSNDDATA_SLBFAIL
            If Result = CALLBACK_ncResOK Then          'OK
                
                ''DB保存準備
                Call SetAPResData(True)
                
                'カレント実績入力情報一時保存
                APResDataBK = APResData
                
                'ビジコン送信
                frmHostSend.SetCallBack Me, CALLBACK_HOSTSEND
                frmHostSend.Show vbModal, Me 'ビジコン送信中は、他の処理を不可とする為、vbModalとする。
            
            Else
                'DB登録キャンセル
            End If
    
    'COLORビジコン通信
    Case CALLBACK_HOSTSEND
            If Result = CALLBACK_ncResOK Then          'OK
                '正常終了
                
                APResData.fail_host_send = "1" '0:正常をセット
                
                '/* DB登録実行 */
                bRet = DB_SAVE_SLBFAIL()
                
                Call dpDebug
                
                If bRet Then
                    'ＤＢ保存正常終了の場合
                    'ＰＤＦ作成要求通知
                    frmTRSend.SetCallBack Me, CALLBACK_TRSEND, "COL01"
                    frmTRSend.Show vbModal, Me
                Else
                    Call WaitMsgBox(Me, "送信／ＤＢ保存処理を中断しました。")
                End If
                
            ElseIf Result = CALLBACK_ncResSKIP Then          'SKIP
                '処理継続
                
                'ビジコン送信スキップ処理（処理前に戻す。）
                APResData.fail_host_send = "0" 'フラグのみエラー扱い
                APResData.fail_host_wrt_dte = APResDataBK.fail_host_wrt_dte
                APResData.fail_host_wrt_tme = APResDataBK.fail_host_wrt_tme
                
                '/* DB登録実行 */
                bRet = DB_SAVE_SLBFAIL()
                
                Call dpDebug
                
                If bRet Then
                    'ＤＢ保存正常終了の場合
                    'ＰＤＦ作成要求通知
                    frmTRSend.SetCallBack Me, CALLBACK_TRSEND, "COL01"
                    frmTRSend.Show vbModal, Me
                Else
                    Call WaitMsgBox(Me, "送信／ＤＢ保存処理を中断しました。")
                End If
            Else
                'キャンセル（エラー発生にて、OKボタンを押した場合、呼出元画面に戻る。）
                
                'ビジコン送信スキップ処理（処理前に戻す。）
                APResData.fail_host_send = "0" '0:異常をセット
                APResData.fail_host_wrt_dte = APResDataBK.fail_host_wrt_dte
                APResData.fail_host_wrt_tme = APResDataBK.fail_host_wrt_tme
                
                Call WaitMsgBox(Me, "送信／ＤＢ保存処理を中断しました。")
                
                Call dpDebug
                
            End If
            
    Case CALLBACK_TRSEND
            If Result = CALLBACK_ncResOK Then          'OK
                Call WaitMsgBox(Me, "ＤＢ保存が正常終了しました。")
            Else
                Call WaitMsgBox(Me, "ＰＤＦ作成要求は失敗しましたが、" & vbCrLf & "ＤＢ保存は正常終了しました。")
            End If
    
            '作成要求のＯＫ／ＮＧにかかわらず正常終了
            '正常終了
            Unload Me
            Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OKで処理終了 '2008/09/04 戻り先変更
    
    Case CALLBACK_PHOTOIMGUPWND
            If Result = CALLBACK_ncResOK Then          'OK
            Else
                ' 20090115 add by M.Aoyagi
'                lblPhotoCnt.Caption = APResData.PhotoImgCnt
                lblPhotoCnt.Caption = PhotoImgCount("SLBFAIL", APResData.slb_chno, APResData.slb_aino, APResData.slb_stat, APResData.slb_col_cnt)
            End If
    
    End Select

End Sub

' @(f)
'
' 機能      : ボタンコントロール
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ボタンコントロール処理。
'
' 備考      :
'
Private Sub ButtonEnable(ByVal bEnable As Boolean)
    cmdScan.Enabled = bEnable
    cmdFullImage.Enabled = bEnable
    cmdOK.Enabled = bEnable
    cmdCANCEL.Enabled = bEnable

    If bEnable Then
        'スラブ異常報告ビジコン正常送信済みで、処置指示が存在する場合は、「送信」ボタンを無効にする。
        If APResData.fail_host_send = "1" And APResData.fail_dir_sys_wrt_dte <> "" Then
            cmdOK.Enabled = False
            lblOK.Caption = "※ビジコン正常送信済みで、処置指示が存在する為、この画面からの「送信」は出来ません。"
        Else
            cmdOK.Enabled = True
        End If
    End If
End Sub

' @(f)
'
' 機能      : スキャナー読取開始
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキャナー読取開始処理。
'
' 備考      :
'
Private Sub ImageScan()
    Dim nRet As Integer
    Dim Msg As String

    On Error Resume Next
    Call ButtonEnable(False)
    
    If IsDEBUG("SCAN") Then
        
        Dim MsgWnd As Message
        Set MsgWnd = New Message
        
        MsgWnd.MsgText = "スキャナー読み込み中です。" & vbCrLf & "しばらくお待ちください。"
        MsgWnd.OK.Visible = False
        MsgWnd.Show vbModeless, Me
        MsgWnd.Refresh
        DoEvents
        
        nRet = LEAD_SCAN.Load(App.path & "\TEST_SLBFAIL.jpg", 0, 0, 1)
        
        MsgWnd.OK_Close
        
        Call LEAD_SCAN_TwainPage
    Else
        'nRet = LEAD_SCAN_TWAIN_ACQUIRE()
        nRet = LEAD_SCAN.TwainAcquire(Me.hWnd)
    End If
    On Error GoTo 0
    
    If nRet <> 0 Then
        Msg = "ｴﾗｰ '" & CStr(nRet) & ", " & DecodeError(nRet) & ""
        Call WaitMsgBox(Me, Msg)
        Call ButtonEnable(True)
    End If
End Sub

' @(f)
'
' 機能      : スキャナー読取
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スキャナー読取処理。
'
' 備考      :
'
Private Function LEAD_SCAN_TWAIN_ACQUIRE() As Integer
Dim nRet As Integer

Dim MsgWnd As Message
Set MsgWnd = New Message

MsgWnd.MsgText = "スキャナー読み込み中です。" & vbCrLf & "しばらくお待ちください。"
MsgWnd.OK.Visible = False
MsgWnd.Show vbModeless, Me
MsgWnd.Refresh
DoEvents

On Error GoTo ERRORHANDLER
'イメージの取得時に、表示長方形を自動定義します。
LEAD_SCAN.AutoSetRects = True
'自動再描画を無効にします。
LEAD_SCAN.AutoRepaint = False
'TWAINソースマネージャを選択します。

Screen.MousePointer = 11 'マウスポインタを砂時計化
LEAD_SCAN.TwainEnumSources (hWnd)
Screen.MousePointer = 0 'マウスポインタを標準化

LEAD_SCAN.TwainSourceName = LEAD_SCAN.TwainSourceList(0)
Debug.Print LEAD_SCAN.TwainSourceName

'カスタムTWAIN値を設定します。
LEAD_SCAN.TwainMaxPages = -1               'デフォルト
LEAD_SCAN.TwainAppAuthor = ""              'デフォルト

LEAD_SCAN.TwainAppFamily = ""              'デフォルト
LEAD_SCAN.TwainFrameLeft = -1              'デフォルト
LEAD_SCAN.TwainFrameTop = -1               'デフォルト
'LEAD_SCAN.TwainFrameWidth = 10080          '7 インチ
'LEAD_SCAN.TwainFrameHeight = 12960         '9 インチ
LEAD_SCAN.TwainFrameWidth = -1          '7 インチ
LEAD_SCAN.TwainFrameHeight = -1         '9 インチ
LEAD_SCAN.TwainBits = 1                    '1 bit/plane

LEAD_SCAN.TwainPixelType = TWAIN_PIX_HALF  '白黒イメージ

'LEAD_SCAN.TwainPixelType = TWAIN_PIX_GRAY
'LEAD_SCAN.TwainRes = -1                    'デフォルト解像度
LEAD_SCAN.TwainRes = 600                    'デフォルト解像度
LEAD_SCAN.TwainContrast = TWAIN_DEFAULT_CONTRAST        'デフォルト

LEAD_SCAN.TwainIntensity = TWAIN_DEFAULT_INTENSITY      'デフォルト
LEAD_SCAN.EnableTwainFeeder = TWAIN_FEEDER_DEFAULT      'デフォルト
LEAD_SCAN.EnableTwainAutoFeed = TWAIN_AUTOFEED_DEFAULT  'デフォルト
'TwainRealizeメソッドを実行し、
'設定内容が確実に反映されたか確認します。
Screen.MousePointer = 11 'マウスポインタを砂時計化
LEAD_SCAN.TwainRealize (hWnd)
Screen.MousePointer = 0 'マウスポインタを標準化
'TWAINインターフェースを非表示にし、イメージを取得します。
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
' 機能      : フォームロード
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームロード時の処理を行う。
'
' 備考      :
'
Private Sub Form_Load()
    
''    Call clrImgFile("SCAN")
    
    bUpdateImageFlg = False 'イメージ変化無しをセット。

    LEAD1.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD1.EnableMethodErrors = False 'False   システムエラーイベントを発生させない
    LEAD1.EnableTwainEvent = True
    LEAD1.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SLBFAIL)

    LEAD_SCAN.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
    LEAD_SCAN.EnableMethodErrors = False 'False   システムエラーイベントを発生させない
    LEAD_SCAN.EnableTwainEvent = True
    LEAD_SCAN.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SLBFAIL)

    Call GetCurrentAPSlbData
    
    timOpening.Interval = 500
    timOpening.Enabled = True



End Sub

' @(f)
'
' 機能      : フォームの初期化
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : フォームの初期化処理。
'
' 備考      :
'
Private Sub InitForm()
    Dim nI As Integer
    Dim nJ As Integer
    Dim nRet As Integer
    
    Dim strDestination As String

    '読込み済みイメージデータがある場合表示する｡ 'nBitmapListIndexP1 ０：未入力 －１：スキップ
    strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
    If Dir(strDestination) <> "" Then
        nRet = LEAD1.Load(App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.jpg", 0, 0, 1)
    End If

End Sub

' @(f)
'
' 機能      : カレントスラブ情報取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : カレントスラブ情報の取得を行う。
'
' 備考      :
'
Private Sub GetCurrentAPSlbData()

    Dim nI As Integer
    Dim nJ As Integer

    lblInputMode.Caption = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, "【新規】", "【修正】")
    lblInputMode.Refresh

    If APResData.fail_host_wrt_dte <> "" Then
        APResData.host_send_flg = "1" '訂正
    Else
        APResData.host_send_flg = "0" '新規
    End If

    lblHostSendFlg.Caption = IIf(APResData.host_send_flg = "0", "0:新規", "1:訂正")

'    'スラブ異常報告ビジコン正常送信済みで、処置指示が存在する場合は、「送信」ボタンを無効にする。
'    If APResData.fail_host_send = "1" And APResData.fail_dir_sys_wrt_dte <> "" Then
'        cmdOK.Enabled = False
'        lblOK.Caption = "※ビジコン正常送信済みで、処置指示が存在する為、この画面からの「送信」は出来ません。"
'    Else
'        cmdOK.Enabled = True
'    End If

    Call ButtonEnable(True)

    'カレントスラブ情報ロード
    Call dpDebug

    lblSlb(0).Caption = APResData.slb_chno & "-" & APResData.slb_aino ''スラブNo
    lblSlb(1).Caption = ConvDpOutStat(conDefine_SYSMODE_SLBFAIL, CInt(APResData.slb_stat)) ''状態
    lblSlb(2).Caption = Format(CInt(APResData.slb_col_cnt), "00") ''カラー回数
    lblSlb(3).Caption = APResData.sys_wrt_dte ''記録日
    
    '2008/09/01 SystEx. A.K
    imSozai(0).Text = APResData.slb_zkai_dte ''造塊日
    imSozai(1).Text = APResData.slb_ksh ''鋼種
    imSozai(2).Text = APResData.slb_ccno ''CCNo
    imSozai(3).Text = APResData.slb_typ ''型
    imSozai(4).Text = APResData.slb_uksk ''向先
    imSozai(5).Text = APResData.slb_wei ''重量
    imSozai(6).Text = APResData.slb_thkns ''厚み
    imSozai(7).Text = APResData.slb_wdth ''幅
    imSozai(8).Text = APResData.slb_lngth ''長さ

    '検査員名リストBOX設定
    cmbRes(0).Clear
    For nJ = 1 To UBound(APInspData)
        cmbRes(0).AddItem APInspData(nJ - 1).inp_InspName
        If APResData.slb_wrt_nme = APInspData(nJ - 1).inp_InspName Then
            cmbRes(0).ListIndex = nJ - 1
        End If
    Next nJ

    '次工程リストBOX設定
    cmbRes(1).Clear
    For nJ = 1 To UBound(APNextProcDataColor)
        cmbRes(1).AddItem APNextProcDataColor(nJ - 1).inp_NextProc
        If APResData.slb_nxt_prcs = APNextProcDataColor(nJ - 1).inp_NextProc Then
            cmbRes(1).ListIndex = nJ - 1
        End If
    Next nJ

    'コメント情報ロード
    imText(0).Text = APResData.slb_cmt1
    imText(1).Text = APResData.slb_cmt2

    '面欠陥リストBOX設定
    nI = 10
    'E-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_e_s1 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_e_n1
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'E-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_e_s2 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_e_n2
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'E-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_e_s3 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_e_n3
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'W-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_w_s1 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_w_n1
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'W-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_w_s2 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_w_n2
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'W-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_w_s3 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_w_n3
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'S-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_s_s1 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_s_n1
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'S-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_s_s2 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_s_n2
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'S-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_s_s3 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_s_n3
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'N-1
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_n_s1 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_n_n1
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'N-2
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_n_s2 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_n_n2
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

    'N-3
    cmbRes(nI).Clear
    For nJ = 1 To UBound(APFaultFaceColor)
        cmbRes(nI).AddItem APFaultFaceColor(nJ - 1).strName
        If APResData.slb_fault_n_s3 = APFaultFaceColor(nJ - 1).strName Then
            cmbRes(nI).ListIndex = nJ - 1
        End If
    Next nJ
    lblcmbRes(nI).Caption = cmbRes(nI).Text
    
    imText(nI).Text = APResData.slb_fault_n_n3
    lblimText(nI).Caption = imText(nI).Text
    nI = nI + 1

'    '内部欠陥リストBOX設定
'    nI = 22
'    'B-S
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_bs_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_bs_n
'    nI = nI + 1
'
'    'B-M
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_bm_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_bm_n
'    nI = nI + 1
'
'    'B-N
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_bn_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_bn_n
'    nI = nI + 1
'
'    'T-S
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_ts_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_ts_n
'    nI = nI + 1
'
'    'T-M
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_tm_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_tm_n
'    nI = nI + 1
'
'    'T-N
'    cmbRes(nI).Clear
'    For nJ = 1 To UBound(APFaultInsideCOLOR)
'        cmbRes(nI).AddItem APFaultInsideCOLOR(nJ - 1).strName
'        If APResData.slb_fault_tn_s = APFaultInsideCOLOR(nJ - 1).strName Then
'            cmbRes(nI).ListIndex = nJ - 1
'        End If
'    Next nJ
'
'    imText(nI).Text = APResData.slb_fault_tn_n
'    nI = nI + 1

    '欠陥リストは使用不可とする。
    For nI = 10 To 21
        cmbRes(nI).Enabled = False
        imText(nI).Enabled = False
    
        cmbRes(nI).Visible = False
        imText(nI).Visible = False
    Next nI

    ' 20090115 add by M.Aoyagi    画像枚数追加
    lblPhotoCnt.Caption = APResData.PhotoImgCnt
    lblPhotoCnt.Caption = PhotoImgCount("SLBFAIL", APResData.slb_chno, APResData.slb_aino, APResData.slb_stat, APResData.slb_col_cnt)

End Sub

' @(f)
'
' 機能      : 表示動作用タイマーイベント
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 表示動作用タイマーイベント時の処理を行う。
'
' 備考      :
'
Private Sub timOpening_Timer()
    timOpening.Enabled = False
    Call InitForm
End Sub

' @(f)
'
' 機能      : 実績データ登録問い合わせ処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 実績データ登録問い合わせ画面を開く。
'
' 備考      : コールバック有り。
'
Private Sub DBSendDataReq_SLBFAIL()
    Dim fmessage As Object
    Set fmessage = New MessageYN

    '登録に必要なイメージと実績入力データが存在するか。
'    If CheckAPInputComplete() Then
    fmessage.MsgText = "スラブ異常報告書入力の実績データを登録します。" & vbCrLf & "よろしいですか？"
'    fmessage.AutoDelete = True
    fmessage.AutoDelete = False
'    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SLBFAIL, True
    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SLBFAIL, False
'        Do
'            On Error Resume Next
'            fmessage.Show vbModeless, Me
'            If Err.Number = 0 Then
'                Exit Do
'            End If
'            DoEvents
'        Loop
    fmessage.Show vbModal, Me '他の処理を不可とする為、vbModalとする。
    Set fmessage = Nothing
'    End If

End Sub

' @(f)
'
' 機能      : 入力情報設定
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ情報の設定を行う。
'
' 備考      :
'
Private Sub SetAPResData(ByVal bDateTimeSet As Boolean)

    Dim nI As Integer
    Dim bFault As Boolean

    '検査員名
    APResData.slb_wrt_nme = cmbRes(0).Text
    
    '次工程
    APResData.slb_nxt_prcs = cmbRes(1).Text
    
    '2008/09/01 SystEx. A.K 現在データを保持する。
    APSysCfgData.NowStaffName(conDefine_SYSMODE_SLBFAIL) = APResData.slb_wrt_nme '検査員名
    APSysCfgData.NowNextProcess(conDefine_SYSMODE_SLBFAIL) = APResData.slb_nxt_prcs '次工程
    '2008/09/01 SystEx. A.K カラー側にも保持する。
    APSysCfgData.NowStaffName(conDefine_SYSMODE_COLOR) = APResData.slb_wrt_nme '検査員名
    APSysCfgData.NowNextProcess(conDefine_SYSMODE_COLOR) = APResData.slb_nxt_prcs '次工程
    
    '2008/09/01 SystEx. A.K
    APResData.slb_zkai_dte = imSozai(0).Text ''造塊日
    APResData.slb_ksh = imSozai(1).Text ''鋼種
    APResData.slb_ccno = imSozai(2).Text ''CCNo
    APResData.slb_typ = imSozai(3).Text ''型
    APResData.slb_uksk = imSozai(4).Text ''向先
    APResData.slb_wei = imSozai(5).Text ''重量
    APResData.slb_thkns = imSozai(6).Text ''厚み
    APResData.slb_wdth = imSozai(7).Text ''幅
    APResData.slb_lngth = imSozai(8).Text ''長さ
    
    'コメント１
    APResData.slb_cmt1 = imText(0).Text
    
    'コメント２
    APResData.slb_cmt2 = imText(1).Text
    
    If bDateTimeSet Then
        ''初回登録日付を設定
        If APResData.fail_sys_wrt_dte = "" Then
            APResData.fail_sys_wrt_dte = Format(Now, "YYYYMMDD")
            APResData.fail_sys_wrt_tme = Format(Now, "HHMMSS")
        End If
    End If
    
    '欠陥を設定
    nI = 10
    'E1面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s1 = ""
        APResData.slb_fault_cd_e_s1 = ""
    Else
        APResData.slb_fault_e_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n1 = ""
    End If
    nI = nI + 1
    
    'E2面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s2 = ""
        APResData.slb_fault_cd_e_s2 = ""
    Else
        APResData.slb_fault_e_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n2 = ""
    End If
    nI = nI + 1
    
    'E3面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_e_s3 = ""
        APResData.slb_fault_cd_e_s3 = ""
    Else
        APResData.slb_fault_e_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_e_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_e_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_e_n3 = ""
    End If
    nI = nI + 1
        
    'W1面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s1 = ""
        APResData.slb_fault_cd_w_s1 = ""
    Else
        APResData.slb_fault_w_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n1 = ""
    End If
    nI = nI + 1
    
    'W2面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s2 = ""
        APResData.slb_fault_cd_w_s2 = ""
    Else
        APResData.slb_fault_w_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n2 = ""
    End If
    nI = nI + 1
    
    'W3面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_w_s3 = ""
        APResData.slb_fault_cd_w_s3 = ""
    Else
        APResData.slb_fault_w_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_w_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_w_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_w_n3 = ""
    End If
    nI = nI + 1
    
    'S1面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s1 = ""
        APResData.slb_fault_cd_s_s1 = ""
    Else
        APResData.slb_fault_s_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n1 = ""
    End If
    nI = nI + 1
    
    'S2面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s2 = ""
        APResData.slb_fault_cd_s_s2 = ""
    Else
        APResData.slb_fault_s_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n2 = ""
    End If
    nI = nI + 1
    
    'S3面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_s_s3 = ""
        APResData.slb_fault_cd_s_s3 = ""
    Else
        APResData.slb_fault_s_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_s_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_s_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_s_n3 = ""
    End If
    nI = nI + 1
        
    'N1面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s1 = ""
        APResData.slb_fault_cd_n_s1 = ""
    Else
        APResData.slb_fault_n_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s1 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n1 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n1 = ""
    End If
    nI = nI + 1
    
    'N2面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s2 = ""
        APResData.slb_fault_cd_n_s2 = ""
    Else
        APResData.slb_fault_n_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s2 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n2 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n2 = ""
    End If
    nI = nI + 1
    
    'N3面
    If cmbRes(nI).ListIndex <= 0 Then
        APResData.slb_fault_n_s3 = ""
        APResData.slb_fault_cd_n_s3 = ""
    Else
        APResData.slb_fault_n_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strName
        APResData.slb_fault_cd_n_s3 = APFaultFaceColor(cmbRes(nI).ListIndex).strCode
    End If
    If IsNumeric(imText(nI).Text) Then
        APResData.slb_fault_n_n3 = Format(CInt(imText(nI).Text), "00")
    Else
        APResData.slb_fault_n_n3 = ""
    End If
    nI = nI + 1
    
'    '内部欠陥リストBOX取得
'    'BS面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_bs_s = ""
'        APResData.slb_fault_cd_bs_s = ""
'    Else
'        APResData.slb_fault_bs_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_bs_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_bs_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_bs_n = ""
'    End If
'    nI = nI + 1
'
'    'BM面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_bm_s = ""
'        APResData.slb_fault_cd_bm_s = ""
'    Else
'        APResData.slb_fault_bm_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_bm_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_bm_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_bm_n = ""
'    End If
'    nI = nI + 1
'
'    'BN面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_bn_s = ""
'        APResData.slb_fault_cd_bn_s = ""
'    Else
'        APResData.slb_fault_bn_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_bn_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_bn_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_bn_n = ""
'    End If
'    nI = nI + 1
'
'    'TS面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_ts_s = ""
'        APResData.slb_fault_cd_ts_s = ""
'    Else
'        APResData.slb_fault_ts_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_ts_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_ts_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_ts_n = ""
'    End If
'    nI = nI + 1
'
'    'TM面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_tm_s = ""
'        APResData.slb_fault_cd_tm_s = ""
'    Else
'        APResData.slb_fault_tm_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_tm_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_tm_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_tm_n = ""
'    End If
'    nI = nI + 1
'
'    'TN面
'    If cmbRes(nI).ListIndex <= 0 Then
'        APResData.slb_fault_tn_s = ""
'        APResData.slb_fault_cd_tn_s = ""
'    Else
'        APResData.slb_fault_tn_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strName
'        APResData.slb_fault_cd_tn_s = APFaultInsideCOLOR(cmbRes(nI).ListIndex).strCode
'    End If
'    If IsNumeric(imText(nI).Text) Then
'        APResData.slb_fault_tn_n = Format(CInt(imText(nI).Text), "00")
'    Else
'        APResData.slb_fault_tn_n = ""
'    End If
'    nI = nI + 1
    
    
    '欠陥判定を設定
    'E判定
    bFault = False '欠陥無し
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
    
    'W判定
    bFault = False '欠陥無し
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
    
    'N判定
    bFault = False '欠陥無し
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
    
    'S判定
    bFault = False '欠陥無し
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
    
'    'B判定
'    bFault = False '欠陥無し
'    Do While True
'        'B1
'        If IsNumeric(APResData.slb_fault_bs_n) Then
'            If APResData.slb_fault_bs_s <> "" And CInt(APResData.slb_fault_bs_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        'B2
'        If IsNumeric(APResData.slb_fault_bm_n) Then
'            If APResData.slb_fault_bm_s <> "" And CInt(APResData.slb_fault_bm_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        'B3
'        If IsNumeric(APResData.slb_fault_bn_n) Then
'            If APResData.slb_fault_bn_s <> "" And CInt(APResData.slb_fault_bn_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        Exit Do
'    Loop
'    APResData.slb_fault_b_judg = IIf(bFault, "1", "0")
'
'    'T判定
'    bFault = False '欠陥無し
'    Do While True
'        'T1
'        If IsNumeric(APResData.slb_fault_ts_n) Then
'            If APResData.slb_fault_ts_s <> "" And CInt(APResData.slb_fault_ts_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        'T2
'        If IsNumeric(APResData.slb_fault_tm_n) Then
'            If APResData.slb_fault_tm_s <> "" And CInt(APResData.slb_fault_tm_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        'T3
'        If IsNumeric(APResData.slb_fault_tn_n) Then
'            If APResData.slb_fault_tn_s <> "" And CInt(APResData.slb_fault_tn_n) > 0 Then
'                bFault = True
'                Exit Do
'            End If
'        End If
'        Exit Do
'    Loop
'    APResData.slb_fault_t_judg = IIf(bFault, "1", "0")
    
    If IsNumeric(APResData.slb_ccno) Then
        If CLng(APResData.slb_ccno) >= 10000 And CLng(APResData.slb_ccno) <= 19999 Then
            '1万番台
            'U=E and S
            'D=W and N
            '0:欠陥無し
            '1:欠陥有り
            '******** U判定 ********
            If CInt(APResData.slb_fault_e_judg) = 1 Or CInt(APResData.slb_fault_s_judg) = 1 Then
                APResData.slb_fault_u_judg = "1"
            Else
                APResData.slb_fault_u_judg = "0"
                'カラー２回目以降の変換
                If CInt(APResData.slb_col_cnt) > 1 Then
                    APResData.slb_fault_u_judg = "9"
                End If
            End If
            '******** D判定 ********
            If CInt(APResData.slb_fault_w_judg) = 1 Or CInt(APResData.slb_fault_n_judg) = 1 Then
                APResData.slb_fault_d_judg = "1"
            Else
                APResData.slb_fault_d_judg = "0"
                'カラー２回目以降の変換
                If CInt(APResData.slb_col_cnt) > 1 Then
                    APResData.slb_fault_d_judg = "9"
                End If
            End If

        ElseIf CLng(APResData.slb_ccno) >= 60000 And CLng(APResData.slb_ccno) <= 69999 Then
            '6万番台
            'U=W and S
            'D=E and N
            '0:欠陥無し
            '1:欠陥有り
            '******** U判定 ********
            If CInt(APResData.slb_fault_w_judg) = 1 Or CInt(APResData.slb_fault_s_judg) = 1 Then
                APResData.slb_fault_u_judg = "1"
            Else
                APResData.slb_fault_u_judg = "0"
                'カラー２回目以降の変換
                If CInt(APResData.slb_col_cnt) > 1 Then
                    APResData.slb_fault_u_judg = "9"
                End If
            End If
            '******** D判定 ********
            If CInt(APResData.slb_fault_e_judg) = 1 Or CInt(APResData.slb_fault_n_judg) = 1 Then
                APResData.slb_fault_d_judg = "1"
            Else
                APResData.slb_fault_d_judg = "0"
                'カラー２回目以降の変換
                If CInt(APResData.slb_col_cnt) > 1 Then
                    APResData.slb_fault_d_judg = "9"
                End If
            End If

        Else
            'CCNOが判定範囲外です。
            Call MsgLog(conProcNum_MAIN, "DB_SAVE_SLBFAIL:CCNOが判定範囲外です。:" & APResData.slb_ccno) 'ガイダンス表示
        End If
    
    Else
        'CCNOがありませんでした。
        Call MsgLog(conProcNum_MAIN, "DB_SAVE_SLBFAIL:CCNOがありませんでした。:" & APResData.slb_ccno) 'ガイダンス表示
    End If

    'UD判定後変換
    '既に異常報告が作成されている場合
    If APResData.fail_sys_wrt_dte <> "" Then
        If APResData.slb_fault_e_judg = "1" Then APResData.slb_fault_e_judg = "2"
        If APResData.slb_fault_w_judg = "1" Then APResData.slb_fault_w_judg = "2"
        If APResData.slb_fault_s_judg = "1" Then APResData.slb_fault_s_judg = "2"
        If APResData.slb_fault_n_judg = "1" Then APResData.slb_fault_n_judg = "2"
        If APResData.slb_fault_u_judg = "1" Then APResData.slb_fault_u_judg = "2"
        If APResData.slb_fault_d_judg = "1" Then APResData.slb_fault_d_judg = "2"
    End If

End Sub


' @(f)
'
' 機能      : ＤＢ保存処理
'
' 引き数    :
'
' 返り値    : True 正常終了／False 異常終了
'
' 機能説明  : ＤＢ保存処理を行う。
'
' 備考      :
'
Private Function DB_SAVE_SLBFAIL() As Boolean
    Dim bNOErrorFlg As Boolean
'    Dim APResDataBK As typAPResData
    Dim nI As Integer
    Dim nJ As Integer
    Dim strImageFileName As String
    Dim bRet As Boolean
    Dim nRet As Integer
    Dim nBlock As Integer
    Dim bFault As Boolean
    Dim strSource As String
    Dim strDestination As String
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message

    MsgWnd.MsgText = "データベースサーバーに保存中です。" & vbCrLf & "しばらくお待ちください。"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
'    ''ＤＢオフラインで強制入力を行ったことを判断するフラグ
'    If bAPInputOffline Then
'        MsgWnd.OK_Close
'        bNOErrorFlg = True 'エラー無し
'        DB_SAVE_SLBFAIL = bNOErrorFlg
'        Exit Function
'    End If
    
'    'カレント実績入力情報一時保存
'    APResDataBK = APResData
    

    bNOErrorFlg = True 'エラー無し

    '*** カラーチェック検査表 ***
    ''初回登録日付を設定
    If APResData.sys_wrt_dte = "" Then
        APResData.sys_wrt_dte = Format(Now, "YYYYMMDD")
        APResData.sys_wrt_tme = Format(Now, "HHMMSS")
    End If
    'TRTS0014 登録
    bRet = TRTS0014_Write(False)
    If bRet = False Then
        bNOErrorFlg = False 'エラー有り
        MsgWnd.OK_Close
        DB_SAVE_SLBFAIL = bNOErrorFlg
        Exit Function
    End If
    
    ''スキャンイメージを保存
    'スキャンしたイメージがあるか？
    'strDestination
    strSource = App.path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG"
    If Dir(strSource) <> "" Then
    
        'フォルダ作成（カラーチェック分）
        On Error Resume Next
        strDestination = APSysCfgData.SHARES_SCNDIR & "\COLOR"
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\COLOR" & "\" & APResData.slb_chno
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
        Call MkDir(strDestination)
        On Error GoTo 0
        
        'ファイル名作成
        strDestination = APSysCfgData.SHARES_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
        
        On Error GoTo DB_SAVE_SLBFAIL_File_err:
        Call FileCopy(strSource, strDestination)
        On Error GoTo 0
    
        'TRTS0052 登録(COLOR_SCANLOC)
        bRet = TRTS0052_Write(False)
        If bRet = False Then
            bNOErrorFlg = False 'エラー有り
        End If
    
    Else
        'イメージ無し
        If Dir(strDestination) <> "" Then
            'Kill strDestination
        End If
    
        'TRTS0052 登録(COLOR_SCANLOC)
        bRet = TRTS0052_Write(True)
        If bRet = False Then
            bNOErrorFlg = False 'エラー有り
        End If
    
    End If
    '******

    '*** スラブ異常報告書 ***
    'TRTS0016 登録
    bRet = TRTS0016_Write(False)
    If bRet = False Then
        bNOErrorFlg = False 'エラー有り
    End If
    '******

'    'ここまで、エラー無しの場合
'    If bNOErrorFlg Then
'        'トランザクション通知処理
'        'Call CSTRAN_DB_SAVE_START
'    End If
'
'    '//登録実行
'    '//登録実行
'
    ''スキャンイメージを保存
    'スキャンしたイメージがあるか？
    'strDestination
    strSource = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
    If Dir(strSource) <> "" Then
    
        'フォルダ作成（スラブ異常報告分）
        On Error Resume Next
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL"
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno
        Call MkDir(strDestination)
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino
        Call MkDir(strDestination)
        On Error GoTo 0
        
        'ファイル名作成
        strDestination = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                      "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                      "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
        
        On Error GoTo DB_SAVE_SLBFAIL_File_err:
        Call FileCopy(strSource, strDestination)
        On Error GoTo 0
    
        'TRTS0054 登録(SLBFAIL_SCANLOC)
        bRet = TRTS0054_Write(False)
        If bRet = False Then
            bNOErrorFlg = False 'エラー有り
        End If
    
    Else
        'イメージ無し
        If Dir(strDestination) <> "" Then
            'Kill strDestination
        End If
    
        'TRTS0054 登録(SLBFAIL_SCANLOC)
        bRet = TRTS0054_Write(True)
        If bRet = False Then
            bNOErrorFlg = False 'エラー有り
        End If
    
    End If
    
    MsgWnd.OK_Close

    DB_SAVE_SLBFAIL = bNOErrorFlg

    Exit Function

DB_SAVE_SLBFAIL_File_err:
    On Error GoTo 0
    
    MsgWnd.OK_Close
    
    Call MsgLog(conProcNum_MAIN, strDestination & ":DB_SAVE_SLBFAIL_File_err:イメージファイルの保存に失敗しました。") 'ガイダンス表示
    
    bNOErrorFlg = False 'エラー有り
    
    DB_SAVE_SLBFAIL = bNOErrorFlg

End Function


' @(f)
'
' 機能      : 実績入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imText_GotFocus(Index As Integer)
    nPreBkColor = imText(Index).BackColor
    imText(Index).BackColor = conDefine_ColorBKGotFocus '背景黄色
End Sub

''---
'Private Sub cmbRes_GotFocus(Index As Integer)
'    nPreBkColor = cmbRes(Index).BackColor
'    cmbRes(Index).BackColor = conDefine_ColorBKGotFocus '背景黄色
'End Sub



' @(f)
'
' 機能      : 実績入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
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
' 機能      : 実績入力BOX変更
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力BOX変更時の処理を行う。
'
' 備考      :
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
' 機能      : 実績入力（素材）BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力（素材）BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imSozai_GotFocus(Index As Integer)
    nPreBkColor = imSozai(Index).BackColor
    imSozai(Index).BackColor = conDefine_ColorBKGotFocus '背景黄色
End Sub

' @(f)
'
' 機能      : 実績入力（素材）BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力（素材）BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imSozai_LostFocus(Index As Integer)
    imSozai(Index).BackColor = nPreBkColor
End Sub

' @(f)
'
' 機能      : 実績入力（素材）BOX変更
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 実績入力（素材）BOX変更時の処理を行う。
'
' 備考      :
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
    '項目毎、特殊チェック
    Select Case Index
        Case 1 '鋼種 20080909
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
            
        Case 6 '厚　XXX.XX
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
' 機能      : テキストボックスチェック
'
' 引き数    : ARG1 - 項目のインデックス
'             ARG2 - キャンセルフラグ
'
' 返り値    :
'
' 機能説明  : テキストボックスチェック用
'
' 備考      :
'
Private Sub imSozai_Validate(Index As Integer, CANCEL As Boolean)
    Dim dAns As Double
    '項目毎、特殊チェック
    Select Case Index
        Case 6 '厚　XXX.XX
            If IsNumeric(imSozai(Index).Text) Then
                '数値
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
Private Sub dpDebug()

    Dim nI As Integer

    If IsDEBUG("DISP") Then
        '表示デバッグモード
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
' 機能      : コールバック設定
'
' 引き数    : ARG1 - コールバックオブジェクト
'             ARG2 - コールバックＩＤ
'
' 返り値    :
'
' 機能説明  : 戻り先コールバック情報を設定する。
'
' 備考      :2008/09/04
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer)
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
End Sub


