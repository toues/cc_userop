VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSysCfgWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "システム設定"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   12330
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Frame Frame2 
      Caption         =   "スラブ異常報告書入力－スキャナ読込設定"
      Height          =   4335
      Left            =   8220
      TabIndex        =   82
      Top             =   180
      Width           =   3915
      Begin VB.ComboBox cmbRotate 
         Height          =   300
         Index           =   2
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   13
         Top             =   1140
         Width           =   795
      End
      Begin VB.ComboBox cmbImageSize 
         Height          =   300
         Index           =   2
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   12
         Top             =   660
         Width           =   795
      End
      Begin imText6Ctl.imText imtxtImageLeft 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   14
         Top             =   2280
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":006E
         Key             =   "frmSysCfgWnd.frx":008C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageTop 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         Top             =   2700
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":00D0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":013E
         Key             =   "frmSysCfgWnd.frx":015C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageWidth 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   3120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":01A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":020E
         Key             =   "frmSysCfgWnd.frx":022C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageHeight 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Top             =   3540
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0270
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":02DE
         Key             =   "frmSysCfgWnd.frx":02FC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.Label Label21 
         Caption         =   "読込時イメージ回転"
         Height          =   315
         Left            =   240
         TabIndex        =   96
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label20 
         Caption         =   "°"
         Height          =   315
         Left            =   2820
         TabIndex        =   95
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "イメージ表示サイズ"
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "％"
         Height          =   315
         Left            =   2820
         TabIndex        =   93
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "左座標"
         Height          =   255
         Index           =   11
         Left            =   660
         TabIndex        =   92
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "上座標"
         Height          =   255
         Index           =   10
         Left            =   660
         TabIndex        =   91
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "幅"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   90
         Top             =   3180
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "高さ"
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   89
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label17 
         Caption         =   "切り出し設定"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   11
         Left            =   2220
         TabIndex        =   87
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "イメージ"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   86
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   10
         Left            =   2220
         TabIndex        =   85
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   9
         Left            =   2220
         TabIndex        =   84
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   8
         Left            =   2220
         TabIndex        =   83
         Top             =   3540
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "カラーチェック検査表入力－スキャナ読込設定"
      Height          =   4335
      Left            =   4200
      TabIndex        =   67
      Top             =   180
      Width           =   3915
      Begin VB.ComboBox cmbRotate 
         Height          =   300
         Index           =   1
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   7
         Top             =   1140
         Width           =   795
      End
      Begin VB.ComboBox cmbImageSize 
         Height          =   300
         Index           =   1
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   6
         Top             =   660
         Width           =   795
      End
      Begin imText6Ctl.imText imtxtImageLeft 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   2280
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":03AE
         Key             =   "frmSysCfgWnd.frx":03CC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageTop 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   2700
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0410
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":047E
         Key             =   "frmSysCfgWnd.frx":049C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageWidth 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   3120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":04E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":054E
         Key             =   "frmSysCfgWnd.frx":056C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageHeight 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   3540
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":05B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":061E
         Key             =   "frmSysCfgWnd.frx":063C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.Label Label16 
         Caption         =   "読込時イメージ回転"
         Height          =   315
         Left            =   240
         TabIndex        =   81
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label15 
         Caption         =   "°"
         Height          =   315
         Left            =   2820
         TabIndex        =   80
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "イメージ表示サイズ"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "％"
         Height          =   315
         Left            =   2820
         TabIndex        =   78
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "左座標"
         Height          =   255
         Index           =   7
         Left            =   660
         TabIndex        =   77
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "上座標"
         Height          =   255
         Index           =   6
         Left            =   660
         TabIndex        =   76
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "幅"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   75
         Top             =   3180
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "高さ"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   74
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "切り出し設定"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   7
         Left            =   2220
         TabIndex        =   72
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "イメージ"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   71
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   6
         Left            =   2220
         TabIndex        =   70
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   5
         Left            =   2220
         TabIndex        =   69
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   4
         Left            =   2220
         TabIndex        =   68
         Top             =   3540
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "通信設定　（通信サーバー）"
      Height          =   2055
      Left            =   180
      TabIndex        =   57
      Top             =   6900
      Width           =   7935
      Begin imText6Ctl.imText imtxtTR_TOUT 
         Height          =   315
         Index           =   0
         Left            =   2820
         TabIndex        =   25
         Top             =   720
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":06EE
         Key             =   "frmSysCfgWnd.frx":070C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtTR_PORT 
         Height          =   315
         Left            =   5895
         TabIndex        =   28
         Top             =   240
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0750
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":07BE
         Key             =   "frmSysCfgWnd.frx":07DC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtTR_IP 
         Height          =   315
         Left            =   2340
         TabIndex        =   24
         Top             =   240
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":088E
         Key             =   "frmSysCfgWnd.frx":08AC
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
         AllowSpace      =   0
         Format          =   "9#"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   15
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtTR_TOUT 
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   26
         Top             =   1140
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":08F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":095E
         Key             =   "frmSysCfgWnd.frx":097C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtTR_TOUT 
         Height          =   315
         Index           =   2
         Left            =   2820
         TabIndex        =   27
         Top             =   1560
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":09C0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0A2E
         Key             =   "frmSysCfgWnd.frx":0A4C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtTR_RETRY 
         Height          =   315
         Left            =   5895
         TabIndex        =   29
         Top             =   720
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0A90
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0AFE
         Key             =   "frmSysCfgWnd.frx":0B1C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.Label Label1 
         Caption         =   "通信サーバー IPアドレス "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "ポート番号"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   65
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "通信タイムアウト(全体監視)"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   64
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   63
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "通信タイムアウト（オープン時）"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   62
         Top             =   1200
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   61
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "通信タイムアウト（データ通信）"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   60
         Top             =   1620
         Width           =   2475
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   59
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "通信リトライ回数"
         Height          =   195
         Index           =   5
         Left            =   4500
         TabIndex        =   58
         Top             =   780
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "通信設定　（ビジコン）"
      Height          =   2055
      Left            =   180
      TabIndex        =   47
      Top             =   4740
      Width           =   7935
      Begin imText6Ctl.imText imtxtHOST_IP 
         Height          =   315
         Left            =   2340
         TabIndex        =   18
         Top             =   240
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0B60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0BCE
         Key             =   "frmSysCfgWnd.frx":0BEC
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
         Format          =   "Ｚ"
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   256
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtHOST_TOUT 
         Height          =   315
         Index           =   0
         Left            =   2820
         TabIndex        =   19
         Top             =   720
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0C30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0C9E
         Key             =   "frmSysCfgWnd.frx":0CBC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtHOST_TOUT 
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   20
         Top             =   1140
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0D00
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0D6E
         Key             =   "frmSysCfgWnd.frx":0D8C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtHOST_PORT 
         Height          =   315
         Left            =   5895
         TabIndex        =   22
         Top             =   240
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0DD0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0E3E
         Key             =   "frmSysCfgWnd.frx":0E5C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtHOST_TOUT 
         Height          =   315
         Index           =   2
         Left            =   2820
         TabIndex        =   21
         Top             =   1560
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0EA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0F0E
         Key             =   "frmSysCfgWnd.frx":0F2C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   3
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtHOST_RETRY 
         Height          =   315
         Left            =   5895
         TabIndex        =   23
         Top             =   720
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":0F70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":0FDE
         Key             =   "frmSysCfgWnd.frx":0FFC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.Label Label2 
         Caption         =   "ビジコン IPアドレス "
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "通信タイムアウト(全体監視)"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   55
         Top             =   780
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   54
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   53
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "通信タイムアウト（オープン時）"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   52
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label Label1 
         Caption         =   "ポート番号"
         Height          =   195
         Index           =   3
         Left            =   4500
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "通信タイムアウト（データ通信）"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   50
         Top             =   1620
         Width           =   2595
      End
      Begin VB.Label Label3 
         Caption         =   "秒"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   49
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "通信リトライ回数"
         Height          =   195
         Index           =   4
         Left            =   4500
         TabIndex        =   48
         Top             =   780
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "スラブ肌調査入力－スキャナ読込設定"
      Height          =   4335
      Left            =   180
      TabIndex        =   32
      Top             =   180
      Width           =   3915
      Begin imText6Ctl.imText imtxtImageLeft 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   2280
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":1040
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":10AE
         Key             =   "frmSysCfgWnd.frx":10CC
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.ComboBox cmbImageSize 
         Height          =   300
         Index           =   0
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   0
         Top             =   660
         Width           =   795
      End
      Begin VB.ComboBox cmbRotate 
         Height          =   300
         Index           =   0
         Left            =   1980
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   1
         Top             =   1140
         Width           =   795
      End
      Begin imText6Ctl.imText imtxtImageTop 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   2700
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":1110
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":117E
         Key             =   "frmSysCfgWnd.frx":119C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageWidth 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   3120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":11E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":124E
         Key             =   "frmSysCfgWnd.frx":126C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin imText6Ctl.imText imtxtImageHeight 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   3540
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmSysCfgWnd.frx":12B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSysCfgWnd.frx":131E
         Key             =   "frmSysCfgWnd.frx":133C
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
         AlignHorizontal =   1
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   3
         Left            =   2220
         TabIndex        =   46
         Top             =   3540
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   2
         Left            =   2220
         TabIndex        =   45
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   44
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "イメージ"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   37
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Pixels"
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   43
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "切り出し設定"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "高さ"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   41
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label10 
         Caption         =   "幅"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   40
         Top             =   3180
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "上座標"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   39
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "左座標"
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   38
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "％"
         Height          =   315
         Left            =   2820
         TabIndex        =   36
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "イメージ表示サイズ"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "°"
         Height          =   315
         Left            =   2820
         TabIndex        =   34
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "読込時イメージ回転"
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "キャンセル"
      Height          =   555
      Left            =   10500
      TabIndex        =   31
      Top             =   6540
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   8940
      TabIndex        =   30
      Top             =   6540
      Width           =   1095
   End
End
Attribute VB_Name = "frmSysCfgWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSysCfgWnd.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　システム設定表示フォーム
' 　本モジュールはシステム設定表示フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

' @(f)
'
' 機能      : イメージ表示率入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ表示率入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub cmbImageSize_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim nI As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            For nI = 10 To 100 Step 10
                If APSysCfgData.nIMAGE_SIZE(Index) = nI Then
                    cmbImageSize(Index).ListIndex = cmbImageSize(Index).ListCount - 1
                End If
            Next nI
    End Select
End Sub

' @(f)
'
' 機能      : イメージ回転入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ回転入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub cmbRotate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim nI As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            For nI = 0 To 270 Step 90
                If APSysCfgData.nIMAGE_ROTATE(Index) = nI Then
                    cmbRotate(Index).ListIndex = cmbRotate(Index).ListCount - 1
                End If
            Next nI
    End Select
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
    
    ' ソケット通信対応
    APSysCfgData.HOST_IP = imtxtHOST_IP.Text
    APSysCfgData.nHOST_PORT = imtxtHOST_PORT.Text
    For nI = 0 To 2
        APSysCfgData.nHOST_TOUT(nI) = imtxtHOST_TOUT(nI).Text
    Next nI
    APSysCfgData.nHOST_RETRY = imtxtHOST_RETRY.Text
    
    APSysCfgData.TR_IP = imtxtTR_IP.Text
    APSysCfgData.nTR_PORT = imtxtTR_PORT.Text
    
    For nI = 0 To 2
        APSysCfgData.nTR_TOUT(nI) = imtxtTR_TOUT(nI).Text
    Next nI
    APSysCfgData.nTR_RETRY = imtxtTR_RETRY.Text
    
    For nI = 0 To 2
        APSysCfgData.nIMAGE_SIZE(nI) = CInt(cmbImageSize(nI).Text)
        APSysCfgData.nIMAGE_ROTATE(nI) = CInt(cmbRotate(nI).Text)
        APSysCfgData.nIMAGE_LEFT(nI) = CInt(imtxtImageLeft(nI).Text)
        APSysCfgData.nIMAGE_TOP(nI) = CInt(imtxtImageTop(nI).Text)
        APSysCfgData.nIMAGE_WIDTH(nI) = CInt(imtxtImageWidth(nI).Text)
        APSysCfgData.nIMAGE_HEIGHT(nI) = CInt(imtxtImageHeight(nI).Text)
    Next nI
    
    Unload Me
    
    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
    Set cCallBackObject = Nothing
End Sub

' @(f)
'
' 機能      : キャンセルボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : キャンセルボタン処理。
'
' 備考      :
'
Private Sub cmdCancel_Click()
    Unload Me

    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResCANCEL
    Set cCallBackObject = Nothing
End Sub

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
    Call InitForm
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
    Dim Index As Integer
    
    ' ソケット通信対応
    imtxtHOST_IP.Text = APSysCfgData.HOST_IP
    imtxtHOST_PORT.Text = APSysCfgData.nHOST_PORT
    imtxtHOST_RETRY.Text = APSysCfgData.nHOST_RETRY
    For nI = 0 To 2
        imtxtHOST_TOUT(nI).Text = APSysCfgData.nHOST_TOUT(nI)
    Next nI

    imtxtTR_IP.Text = APSysCfgData.TR_IP
    imtxtTR_PORT.Text = APSysCfgData.nTR_PORT
    imtxtTR_RETRY.Text = APSysCfgData.nTR_RETRY
    For nI = 0 To 2
        imtxtTR_TOUT(nI).Text = APSysCfgData.nTR_TOUT(nI)
    Next nI
    
    For Index = 0 To 2
        For nI = 10 To 100 Step 10
            cmbImageSize(Index).AddItem CStr(nI)
            If APSysCfgData.nIMAGE_SIZE(Index) = nI Then
                cmbImageSize(Index).ListIndex = cmbImageSize(Index).ListCount - 1
            End If
        Next nI
        
        For nI = 0 To 270 Step 90
            cmbRotate(Index).AddItem CStr(nI)
            If APSysCfgData.nIMAGE_ROTATE(Index) = nI Then
                cmbRotate(Index).ListIndex = cmbRotate(Index).ListCount - 1
            End If
        Next nI
    Next Index
    
    For nI = 0 To 2
        imtxtImageLeft(nI).Text = CStr(APSysCfgData.nIMAGE_LEFT(nI))
        imtxtImageTop(nI).Text = CStr(APSysCfgData.nIMAGE_TOP(nI))
        imtxtImageWidth(nI).Text = CStr(APSysCfgData.nIMAGE_WIDTH(nI))
        imtxtImageHeight(nI).Text = CStr(APSysCfgData.nIMAGE_HEIGHT(nI))
    Next nI
    
End Sub

' @(f)
'
' 機能      : FTP通信IPアドレス入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信IPアドレス入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_IP_GotFocus()
    imtxtTR_IP.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : FTP通信IPアドレス入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : FTP通信IPアドレス入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_IP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtTR_IP.Text = APSysCfgData.TR_IP
    End Select
End Sub

' @(f)
'
' 機能      : FTP通信IPアドレス入力BOXキー押
'
' 引き数    : ARG1 - ASCIIコード
'
' 返り値    :
'
' 機能説明  : FTP通信IPアドレス入力BOXキー押時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_IP_KeyPress(KeyAscii As Integer)
    Dim nI As Integer
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = Asc(".") Then
    Else
        If KeyAscii <> 10 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If

End Sub

' @(f)
'
' 機能      : FTP通信IPアドレス入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信IPアドレス入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_IP_LostFocus()
    imtxtTR_IP.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : FTP通信ポート番号入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信ポート番号入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_PORT_GotFocus()
    imtxtTR_PORT.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : FTP通信ポート番号入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : FTP通信ポート番号入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_PORT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtTR_PORT.Text = APSysCfgData.nTR_PORT
    End Select
End Sub

' @(f)
'
' 機能      : FTP通信ポート番号入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信ポート番号入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_PORT_LostFocus()
    imtxtTR_PORT.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : FTP通信リトライ回数入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信リトライ回数入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
Private Sub imtxtTR_RETRY_GotFocus()
    ' ソケット通信対応
    imtxtTR_RETRY.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : FTP通信リトライ回数入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : FTP通信リトライ回数入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_RETRY_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ソケット通信対応
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtTR_RETRY.Text = APSysCfgData.nTR_RETRY
    End Select
End Sub

' @(f)
'
' 機能      : FTP通信リトライ回数入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信リトライ回数入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_RETRY_LostFocus()
    ' ソケット通信対応
    imtxtTR_RETRY.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : FTP通信タイムアウト入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : FTP通信タイムアウト入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_TOUT_GotFocus(Index As Integer)
    imtxtTR_TOUT(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : FTP通信タイムアウト入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : FTP通信タイムアウト入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_TOUT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtTR_TOUT(Index).Text = APSysCfgData.nTR_TOUT(Index)
    End Select
End Sub

' @(f)
'
' 機能      : FTP通信タイムアウト入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : FTP通信タイムアウト入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtTR_TOUT_LostFocus(Index As Integer)
    imtxtTR_TOUT(Index).BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : HOST通信ビジコンIP入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : HOST通信ビジコンIP入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_IP_GotFocus()
    ' ソケット通信対応
    imtxtHOST_IP.BackColor = conDefine_ColorBKGotFocus
End Sub


' @(f)
'
' 機能      : ビジコン通信IPアドレス入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : ビジコン通信IPアドレス入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_IP_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ソケット通信対応
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtHOST_IP.Text = APSysCfgData.HOST_IP
    End Select

End Sub

' @(f)
'
' 機能      : ビジコン通信IPアドレス入力BOXキー押
'
' 引き数    : ARG1 - ASCIIコード
'
' 返り値    :
'
' 機能説明  : ビジコン通信IPアドレス入力BOXキー押時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_IP_KeyPress(KeyAscii As Integer)
    ' ソケット通信対応
    Dim nI As Integer
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = Asc(".") Then
    Else
        If KeyAscii <> 10 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If

End Sub

' @(f)
'
' 機能      : HOST通信ビジコンIPアドレス入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : HOST通信ビジコンIPアドレス入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_IP_LostFocus()
    ' ソケット通信対応
    imtxtHOST_IP.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : ビジコン通信ポート番号入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ビジコン通信ポート番号入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_PORT_GotFocus()
    ' ソケット通信対応
    imtxtHOST_PORT.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : ビジコン通信ポート番号入力BOXキー押下
'
' 引き数    : ARG1 - キーコード
'             ARG2 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : ビジコン通信ポート番号入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_PORT_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ソケット通信対応
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtHOST_PORT.Text = APSysCfgData.nHOST_PORT
    End Select
End Sub

' @(f)
'
' 機能      : ビジコン通信ボート入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ビジコン通信ボート入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_PORT_LostFocus()
    ' ソケット通信対応
    imtxtHOST_PORT.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : ビジコン通信リトライ回数入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ビジコン通信リトライ回数入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_RETRY_GotFocus()
    ' ソケット通信対応
    imtxtHOST_RETRY.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : ビジコン通信リトライ回数入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : ビジコン通信リトライ回数入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_RETRY_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ソケット通信対応
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtHOST_RETRY.Text = APSysCfgData.nHOST_RETRY
    End Select
End Sub

' @(f)
'
' 機能      : ビジコン通信リトライ回数入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : ビジコン通信リトライ回数入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
Private Sub imtxtHOST_RETRY_LostFocus()
    ' ソケット通信対応
    imtxtHOST_RETRY.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : HOST通信タイムアウト入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : HOST通信タイムアウト入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_TOUT_GotFocus(Index As Integer)
    imtxtHOST_TOUT(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : HOST通信タイムアウト入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : HOST通信タイムアウト入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_TOUT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtHOST_TOUT(Index).Text = APSysCfgData.nHOST_TOUT(Index)
    End Select
End Sub

' @(f)
'
' 機能      : HOST通信タイムアウト入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : HOST通信タイムアウト入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtHOST_TOUT_LostFocus(Index As Integer)
    imtxtHOST_TOUT(Index).BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : イメージ高さ入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ高さ入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageHeight_GotFocus(Index As Integer)
    imtxtImageHeight(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : イメージ高さ入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ高さ入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageHeight_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtImageHeight(Index).Text = APSysCfgData.nIMAGE_HEIGHT(Index)
    End Select
End Sub

' @(f)
'
' 機能      : イメージ高さ入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ高さ入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageHeight_LostFocus(Index As Integer)
    imtxtImageHeight(Index).BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : イメージ左座標入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ左座標入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageLeft_GotFocus(Index As Integer)
    imtxtImageLeft(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : イメージ左座標入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ左座標入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageLeft_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtImageLeft(Index).Text = APSysCfgData.nIMAGE_LEFT(Index)
    End Select
End Sub

' @(f)
'
' 機能      : イメージ左座標入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ左座標入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageLeft_LostFocus(Index As Integer)
    imtxtImageLeft(Index).BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : イメージ上座標入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ上座標入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageTop_GotFocus(Index As Integer)
    imtxtImageTop(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : イメージ上座標入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ上座標入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageTop_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtImageTop(Index).Text = APSysCfgData.nIMAGE_TOP(Index)
    End Select
End Sub

' @(f)
'
' 機能      : イメージ上座標入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ上座標入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageTop_LostFocus(Index As Integer)
    imtxtImageTop(Index).BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : イメージ幅入力BOXフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ幅入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageWidth_GotFocus(Index As Integer)
    imtxtImageWidth(Index).BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : イメージ幅入力BOXキー押下
'
' 引き数    : ARG1 - インデックス番号
'             ARG2 - キーコード
'             ARG3 - シフトフラグ
'
' 返り値    :
'
' 機能説明  : イメージ幅入力BOXキー押下時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageWidth_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
        Case vbKeyEscape
            imtxtImageWidth(Index).Text = APSysCfgData.nIMAGE_WIDTH(Index)
    End Select
End Sub

' @(f)
'
' 機能      : イメージ幅入力BOXフォーカス消滅
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : イメージ幅入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtImageWidth_LostFocus(Index As Integer)
    imtxtImageWidth(Index).BackColor = conDefine_ColorBKLostFocus
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
' 備考      :
'
Public Sub SetCallBack(ByVal callBackObj As Object, ByVal ObjctID As Integer)
    iCallBackID = ObjctID
    Set cCallBackObject = callBackObj
End Sub


