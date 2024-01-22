VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmColorSlbSelWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "カラーチェック検査表入力－スラブ選択"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   18690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   18690
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame_StatusChg 
      BackColor       =   &H00C0FFFF&
      Caption         =   "状態変更"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   14040
      TabIndex        =   24
      Top             =   600
      Width           =   4395
      Begin VB.CommandButton cmdStatChgFix 
         Caption         =   "状態決定"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2280
         TabIndex        =   27
         Top             =   480
         Width           =   1800
      End
      Begin VB.CommandButton cmdStatChgMode 
         Caption         =   "状態変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1800
      End
      Begin imText6Ctl.imText imTextColCntNew 
         Height          =   525
         Left            =   2280
         TabIndex        =   28
         Top             =   1080
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   926
         Caption         =   "frmColorSlbSelWnd.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmColorSlbSelWnd.frx":006E
         Key             =   "frmColorSlbSelWnd.frx":008C
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
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   1
         MaxLength       =   2
         LengthAsByte    =   0
         Text            =   ""
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
      Begin VB.Label lblColCntNew 
         BackStyle       =   0  '透明
         Caption         =   "カラー回数"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdDirRes 
      Caption         =   "処置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14100
      TabIndex        =   23
      Top             =   9960
      Width           =   1800
   End
   Begin VB.Frame Frame_Status 
      BackColor       =   &H00C0FFFF&
      Caption         =   "新規入力－状態選択"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   7260
      TabIndex        =   15
      Top             =   600
      Width           =   6675
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "6:6ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   4980
         TabIndex        =   22
         Top             =   1140
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "5:5ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   3420
         TabIndex        =   21
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "4:4ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   1860
         TabIndex        =   20
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "2:2ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3420
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1:1ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1860
         TabIndex        =   18
         Top             =   480
         Width           =   1515
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0:白皮"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "3:3ht後"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   4980
         TabIndex        =   16
         Top             =   480
         Width           =   1515
      End
   End
   Begin VB.PictureBox PicSigYellow 
      Height          =   315
      Left            =   3720
      Picture         =   "frmColorSlbSelWnd.frx":00D0
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   10140
      Visible         =   0   'False
      Width           =   555
   End
   Begin imText6Ctl.imText imTextSearchSlbNumber 
      Height          =   525
      Left            =   1800
      TabIndex        =   0
      Top             =   1980
      Width           =   3360
      _Version        =   65536
      _ExtentX        =   5927
      _ExtentY        =   926
      Caption         =   "frmColorSlbSelWnd.frx":0714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmColorSlbSelWnd.frx":0782
      Key             =   "frmColorSlbSelWnd.frx":07A0
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
      Format          =   "A9#@"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   10
      LengthAsByte    =   0
      Text            =   ""
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
   Begin VB.Frame Frame_InputMode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "入力モード選択"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   6855
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "新規"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "修正"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton OptInputMode 
         BackColor       =   &H00C0FFFF&
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   3900
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicSigRed 
      Height          =   375
      Left            =   5160
      Picture         =   "frmColorSlbSelWnd.frx":07E4
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   10020
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox PicSigGreen 
      Height          =   315
      Left            =   3720
      Picture         =   "frmColorSlbSelWnd.frx":0FC2
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   10500
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "戻る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   16620
      TabIndex        =   7
      Top             =   9960
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11580
      TabIndex        =   6
      Top             =   9960
      Width           =   1800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "検索"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5220
      TabIndex        =   1
      Top             =   1980
      Width           =   1800
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7095
      Left            =   180
      TabIndex        =   5
      Top             =   2640
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   21
      Cols            =   15
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_works 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "SKY"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   27.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3360
      TabIndex        =   29
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label lbl_nMSFlexGrid1_Selected_Row 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "カラーチェック検査表入力－スラブ選択"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Width           =   9435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "スラブNo."
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   9
      Top             =   1980
      Width           =   1635
   End
End
Attribute VB_Name = "frmColorSlbSelWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmColorSlbSelWnd.Frm                ver 1.00 ( '2008.05.15 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック検査表入力－スラブ選択表示フォーム
' 　本モジュールはカラーチェック検査表入力－スラブ選択表示フォームで使用する
' 　ためのものである。

Option Explicit

Private nMSFlexGrid1_Selected_Row As Integer ''グリッド１選択行番号格納
Private bMouseControl As Boolean ''マウスコントロールフラグ格納
Private bOptInputModeValue(0 To 2) As Boolean

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
' 備考      :COLORSYS
'
Private Sub cmdCancel_Click()
    
    cmdCANCEL.Enabled = False ''連打禁止！

    Call SlbSelLock(False)
    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBSELWND, CALLBACK_ncResCANCEL)
    Unload Me
End Sub

Private Sub cmdDirRes_Click()
    Dim bRet As Boolean
    
    cmdDirRes.Enabled = False '連打禁止！
    
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    
    'スラブ選択チェック
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Call WaitMsgBox(Me, "スラブを選択してください。")
        Exit Sub
    End If

    Select Case APSlbCont.nSearchInputModeSelectedIndex
        Case 0 '新規
        Case 1 '修正
        Case 2 '削除
            '処理終了
            Exit Sub
    End Select
    
    '2016/04/20 - TAI - S
    '作業場チェック
    '違う場所のスラブの操作は禁止
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    If works_sky_tok <> APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_works_sky_tok Then
        If works_sky_tok = WORKS_SKY Then
            MsgWnd.MsgText = "特鋼のスラブは処理できません。"
        ElseIf works_sky_tok = WORKS_TOK Then
            MsgWnd.MsgText = "SKYのスラブは処理できません。"
        End If
        
        MsgWnd.OK.Visible = True
        Do
            On Error Resume Next
            MsgWnd.Show vbModal
            If Err.Number = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        Set MsgWnd = Nothing
        
        cmdOK.Enabled = True 'ボタン有効
        Exit Sub
    End If
    '2016/04/20 - TAI - E
    
    bRet = ColorSlbData_Load(True)

    If bRet Then
        Call OKcmdDIR '処置画面開始(unload me)
    End If
    
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
' 備考      :COLORSYS
'
Private Sub cmdOK_Click()
    Dim bRet As Boolean
    Dim MsgWnd As Message
    Set MsgWnd = New Message
 
    cmdOK.Enabled = False ''連打禁止！
 
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    
    'スラブ選択チェック
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '新規
                MsgWnd.MsgText = "実績入力を行うスラブを選択してください。"
            Case 1 '修正
                MsgWnd.MsgText = "実績修正を行うスラブを選択してください。"
            Case 2 '削除
                MsgWnd.MsgText = "実績削除を行うスラブを選択してください。"
        End Select
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
        
        cmdOK.Enabled = True 'ボタン有効
        Exit Sub
    End If
 
    '2016/04/20 - TAI - S
    '作業場チェック
    '違う場所のスラブの操作は禁止
    If works_sky_tok <> APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_works_sky_tok Then
        If works_sky_tok = WORKS_SKY Then
            MsgWnd.MsgText = "特鋼のスラブは処理できません。"
        ElseIf works_sky_tok = WORKS_TOK Then
            MsgWnd.MsgText = "SKYのスラブは処理できません。"
        End If
        
        MsgWnd.OK.Visible = True
        Do
            On Error Resume Next
            MsgWnd.Show vbModal
            If Err.Number = 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        Set MsgWnd = Nothing
        
        cmdOK.Enabled = True 'ボタン有効
        Exit Sub
    End If
    '2016/04/20 - TAI - E
 
    Set MsgWnd = Nothing
 
    If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
        '削除
        Call ColorDataDel_REQ
    Else
        bRet = ColorSlbData_Load(False)
    
        cmdOK.Enabled = True 'ボタン有効
    
        If bRet Then
            Select Case APSlbCont.nSearchInputModeSelectedIndex
                Case 0 '新規
                    Call OKcmdOK '入力開始(unload me)
                Case 1 '修正
                    Call OKcmdOK '入力開始(unload me)
            End Select
        End If
 
    End If
 
End Sub

Private Function ColorSlbData_Load(ByVal bDirResLoad As Boolean) As Boolean
    Dim bRet As Boolean
    Dim strSource As String
    Dim strDestination As String
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    'スラブ選択チェック
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        ColorSlbData_Load = False 'エラー
        Exit Function
    End If
        
    '********************************************************************************************
    'DEBUG POINT 新規モードでリスト表示の場合、修正対象レコードも同時に表示されるので、
    'リスト選択後、新規ではなく、修正をユーザーが選んだ場合は、もう一度モードをチェックし、
    '新規／修正の切替えが必要
    '********************************************************************************************
    '新規モードか？
    If APSlbCont.nSearchInputModeSelectedIndex = 0 Then
        '選択したスラブは新規か？
        If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).sys_wrt_dte = "" Then
            '新規モード
        Else
            '保存済みである為、修正モードに自動変更
            APSlbCont.nSearchInputModeSelectedIndex = 1
        End If
    End If
    
    
    'ＤＢ処理が発生するモード　修正／削除
    If APSlbCont.nSearchInputModeSelectedIndex <> 0 Then

        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 1 '修正
                MsgWnd.MsgText = "データベースからスラブ情報を読込み中です。" & vbCrLf & "しばらくお待ちください。"
            Case 2 '削除
                MsgWnd.MsgText = "データベースからスラブ情報を削除中です。" & vbCrLf & "しばらくお待ちください。"
        End Select

        MsgWnd.OK.Visible = False
        MsgWnd.Show vbModeless, Me
        MsgWnd.Refresh
    
    End If
    
    '実績処理エリアへデータコピー
    Call init_APResData
    Select Case APSlbCont.nSearchInputModeSelectedIndex
        Case 0 '新規
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no ''スラブチャージNO
            APResData.slb_chno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno ''スラブチャージNO
            APResData.slb_aino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino ''スラブ合番
            APResData.slb_stat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat ''状態
            APResData.slb_col_cnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt ''カラー回数
            APResData.slb_ccno = APSozaiData.slb_ccno           ''スラブCCNO
            APResData.slb_zkai_dte = APSozaiData.slb_zkai_dte   ''造塊日
            APResData.slb_ksh = APSozaiData.slb_ksh             ''鋼種
            APResData.slb_typ = APSozaiData.slb_typ             ''型
            APResData.slb_uksk = APSozaiData.slb_uksk           ''向先
            APResData.slb_wei = APSozaiData.slb_skin_wei        ''重量（ｽﾗﾌﾞ肌用）
            APResData.slb_lngth = APSozaiData.slb_lngth         ''長さ
            APResData.slb_wdth = APSozaiData.slb_wdth           ''幅
            APResData.slb_thkns = APSozaiData.slb_thkns         ''厚み
            
            '2008/09/01 SystEx. A.K 新規の場合、前回データ（保持中データ）をセットする。
            APResData.slb_wrt_nme = APSysCfgData.NowStaffName(conDefine_SYSMODE_COLOR) '検査員名
            APResData.slb_nxt_prcs = APSysCfgData.NowNextProcess(conDefine_SYSMODE_COLOR) '次工程
            
            'カラーチェック
            '新規の場合は、SCANイメージを初期化する。（中間ファイルの削除）
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG"
            'イメージ無し
            If Dir(strDestination) <> "" Then
                Kill strDestination
            End If
            
            'スラブ異常
            '新規の場合は、SCANイメージを初期化する。（中間ファイルの削除）
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
            'イメージ無し
            If Dir(strDestination) <> "" Then
                Kill strDestination
            End If
            
            ' 20090115 add by M.Aoyagi    画像枚数追加の為
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1
            
        Case 1 '修正
            '実績データ読込み
            bRet = TRTS0014_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, _
                                 APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat, _
                                 APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt)
            If UBound(APResTmpData) = 1 Then
                APResData = APResTmpData(0)
            End If
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
            
            'カラーチェック
            '登録済みSCANイメージがあるか？
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "COLOR.JPG"
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).bAPScanInput Then
                '登録済みSCANイメージを読込み (conDefine_ImageDirName = TEMP)
                strSource = APSysCfgData.SHARES_SCNDIR & "\COLOR" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                         "\COLOR" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                         "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
                On Error GoTo ColorSlbData_Load_err:
                Call FileCopy(strSource, strDestination)
                On Error GoTo 0
            Else
                'イメージ無し
                If Dir(strDestination) <> "" Then
                    Kill strDestination
                End If
            End If
            
            'スラブ異常
            '登録済みSCANイメージがあるか？
            strDestination = App.path & "\" & conDefine_ImageDirName & "\" & "SLBFAIL.JPG"
            If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).bAPFailScanInput Then
                '登録済みSCANイメージを読込み (conDefine_ImageDirName = TEMP)
                strSource = APSysCfgData.SHARES_SCNDIR & "\SLBFAIL" & "\" & APResData.slb_chno & "\" & APResData.slb_aino & _
                                                         "\SLBFAIL" & "_" & APResData.slb_chno & "_" & APResData.slb_aino & _
                                                         "_" & APResData.slb_stat & "_" & Format(CInt(APResData.slb_col_cnt), "00") & ".JPG"
                On Error GoTo ColorSlbData_Load_err:
                Call FileCopy(strSource, strDestination)
                On Error GoTo 0
            Else
                'イメージ無し
                If Dir(strDestination) <> "" Then
                    Kill strDestination
                End If
            End If
            
            'スラブ異常報告用
            APResData.fail_host_send = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_send ''スラブ異常報告用　ビジコン送信結果
            APResData.fail_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte       ''スラブ異常報告用　記録日
            APResData.fail_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_tme       ''スラブ異常報告用　記録時刻
            APResData.fail_sys_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_sys_wrt_dte  ''スラブ異常報告用　登録日
            APResData.fail_sys_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_sys_wrt_tme        ''スラブ異常報告用　登録時刻
            
            '処置指示
            APResData.fail_dir_sys_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_dir_sys_wrt_dte ''処置指示用　記録日（初回記録日）

            '処置結果
            APResData.fail_res_host_send = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_send             ''処置結果用　ビジコン送信結果
            APResData.fail_res_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_wrt_dte       ''処置結果用　記録日
            APResData.fail_res_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_host_wrt_tme       ''処置結果用　記録時刻

            If bDirResLoad Then
                'DirResLoad
                '処置指示データ読込み
                bRet = DBDirResData_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, _
                                     APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat, _
                                     APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt)
                
                ReDim APDirResData(0)
                
                If UBound(APDirResTmpData) <> 0 Then
                    ReDim APDirResData(UBound(APDirResTmpData))
                    APDirResData = APDirResTmpData
                End If
                If bRet = False Then
                    Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                    MsgWnd.OK_Close
                    ColorSlbData_Load = False 'エラー
                    Exit Function
                End If
            End If

            ' 20090115 add by M.Aoyagi    画像枚数追加の為
            APResData.PhotoImgCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).PhotoImgCnt1

        Case 2 '削除
        
            '*********
            '削除処理
            '*********
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
            APResData.slb_stat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
            APResData.slb_col_cnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
            bRet = TRTS0014_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
        
            bRet = TRTS0052_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
        
            bRet = TRTS0016_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
        
            bRet = TRTS0054_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
        
            bRet = TRTS0022_Write(True)
            If bRet = False Then
                Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                MsgWnd.OK_Close
                ColorSlbData_Load = False 'エラー
                Exit Function
            End If
        
            MsgWnd.OK_Close
            
            '*********
            '検索結果リスト再表示
            '*********
            Call WaitMsgBox(Me, "削除処理が正常終了しました。")
            Call cmdSearch_Click
            ColorSlbData_Load = True 'OK
            Exit Function
    End Select
    
    'ＤＢ処理が発生するモード　修正／削除（読込中メッセージ表示有り）
    If APSlbCont.nSearchInputModeSelectedIndex <> 0 Then
        MsgWnd.OK_Close
    End If
    
    ColorSlbData_Load = True 'OK
    Exit Function
    
ColorSlbData_Load_err:
    Call MsgLog(conProcNum_MAIN, Err.Source + " " + _
        CStr(Err.Number) + Chr(13) + Err.Description)
    
    Call MsgLog(conProcNum_MAIN, "ColorSlbData_Load FILECOPY SO=" & strSource & " DE=" & strDestination)
    Call WaitMsgBox(Me, "保存済みスキャナーイメージファイルの読込エラーが発生しました。" & vbCrLf & vbCrLf & "FILECOPY SO=" & strSource & " DE=" & strDestination)
    
    MsgWnd.OK_Close
    On Error GoTo 0
    ColorSlbData_Load = False 'エラー
    Exit Function
    
End Function

' @(f)
'
' 機能      : スラブ選択処理ＯＫ終了
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ選択処理ＯＫ通知。
'
' 備考      : コールバックにてＯＫ通知後アンロード。
'
Private Sub OKcmdOK()

    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBSELWND, CALLBACK_ncResOK)
    Unload Me

End Sub

' @(f)
'
' 機能      : スラブ選択処理ＯＫ終了と処置画面リクエスト
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ選択処理ＯＫ通知。
'
' 備考      : コールバックにてＯＫ通知後アンロード。
'
Private Sub OKcmdDIR()

    Call fMainWnd.CallBackMessage(CALLBACK_MAIN_RETCOLORSLBSELWND, CALLBACK_ncResEXTEND)
    Unload Me

End Sub

' @(f)
'
' 機能      : スラブ情報検索ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : スラブ情報の検索を行う。
'
' 備考      : スラブ検索結果表示エリア
'
Private Sub cmdSearch_Click()
    Dim nWildCard As Integer
    Dim nI As Integer
    Dim nJ As Integer
    Dim nSEARCH_MAX As Integer
    Dim bRet As Boolean
    Dim strSearchSlbNumber As String '実際の検索文字列
    Dim strTmpSlbNumber As String '比較用
    Dim bCmp As Boolean '比較用
    Dim strChkChar As String
    
    Dim nSlb_Col_Cnt_MAX As Integer
    Dim nFirstDataIndex As Integer
    
    Dim bNoRecord As Boolean '2008/08/30 A.K
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    '再検索時は初期化
    strSearchSlbNumber = ""
    Call InitMSFlexGrid1
    
    nWildCard = 0
    'ハイフン’－’を取って実際の検索文字列へセット
    strSearchSlbNumber = ConvSearchSlbNumber(imTextSearchSlbNumber.Text)
    
    '入力モード
    If OptInputMode(0).Value Then '新規
        APSlbCont.nSearchInputModeSelectedIndex = 0 '入力モードオプション指定インデックス番号
        
        If OptStatus(0).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 0 '白皮
        ElseIf OptStatus(1).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 1 '1ht後
        ElseIf OptStatus(2).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 2 '2ht後
        ElseIf OptStatus(3).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 3 '3ht後
        ElseIf OptStatus(4).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 4 '4ht後
        ElseIf OptStatus(5).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 5 '5ht後
        ElseIf OptStatus(6).Value Then
            APSlbCont.nSearchInputStatusSelectedIndex = 6 '6ht後
        End If
    ElseIf OptInputMode(1).Value Then '修正
        APSlbCont.nSearchInputModeSelectedIndex = 1 '入力モードオプション指定インデックス番号
        APSlbCont.nSearchInputStatusSelectedIndex = 0 '無効（使用しない）
    ElseIf OptInputMode(2).Value Then '削除
        APSlbCont.nSearchInputModeSelectedIndex = 2 '入力モードオプション指定インデックス番号
        APSlbCont.nSearchInputStatusSelectedIndex = 0 '無効（使用しない）
    End If
    
    nWildCard = InStr(1, strSearchSlbNumber, "%", vbTextCompare)
    
    'RIAL
    ReDim APSearchListSlbData(0)
    
    '新規モード（検索でワイルドカード不可）
    If OptInputMode(0).Value Then
        '空白指定は不可。
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "スラブＮｏ．を入力してください。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード不可。
        If nWildCard <> 0 Then
            Call WaitMsgBox(Me, "新規モードでワイルドカードの指定は出来ません。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、９文字より多い場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、６文字より少ない場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '先頭から５文字までは、0から9以外を不可。
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "先頭から５文字まで、0から9以外の指定は出来ません。")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
    
    '修正モード（検索でワイルドカード可）
    ElseIf OptInputMode(1).Value Then
        '空白指定は不可。
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "スラブＮｏ．を入力してください。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード１つのみは不可。
        If strSearchSlbNumber = "%" Then
            Call WaitMsgBox(Me, "ワイルドカードの指定方法が正しくありません。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、９文字より多い場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、６文字より少ない場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '先頭から５文字までは、0から9,*以外を不可。
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            ElseIf strChkChar = "%" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "先頭から５文字まで、0から9,*以外の指定は出来ません。")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
    
    '削除モード（検索でワイルドカード可）
    ElseIf OptInputMode(2).Value Then
        '空白指定は不可。
        If LenB(imTextSearchSlbNumber.Text) = 0 Then
            Call WaitMsgBox(Me, "スラブＮｏ．を入力してください。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード１つのみは不可。
        If strSearchSlbNumber = "%" Then
            Call WaitMsgBox(Me, "ワイルドカードの指定方法が正しくありません。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、９文字より多い場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) > 9) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        'ワイルドカード無しで、６文字より少ない場合は不可。
        If (nWildCard = 0) And (Len(strSearchSlbNumber) < 6) Then
            Call WaitMsgBox(Me, "スラブＮｏ．の桁数が不正です。")
            imTextSearchSlbNumber.SetFocus
            Exit Sub
        End If
        '先頭から５文字までは、0から9,*以外を不可。
        For nI = 1 To 5
            If nI > Len(strSearchSlbNumber) Then Exit For
            strChkChar = Mid(strSearchSlbNumber, nI, 1)
            If strChkChar >= "0" And strChkChar <= "9" Then
                'OK
            ElseIf strChkChar = "%" Then
                'OK
            Else
                'NG
                Call WaitMsgBox(Me, "先頭から５文字まで、0から9,*以外の指定は出来ません。")
                imTextSearchSlbNumber.SetFocus
                Exit Sub
            End If
        Next nI
        
    End If
    
    MsgWnd.MsgText = "データベースを検索中です。" & vbCrLf & "しばらくお待ちください。"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
    '検索件数
'    nSEARCH_MAX = APSysCfgData.nSEARCH_MAX(APSlbCont.nSearchInputModeSelectedIndex)
    'bRet = DBSkinSlbSearchRead(APSlbCont.nSearchInputModeSelectedIndex, nSEARCH_MAX, strSearchSlbNumber)
    
    '（検索有効範囲は制限あり）
    'bRet = DBSkinSlbSearchRead(APSlbCont.nSearchInputModeSelectedIndex, nSEARCH_MAX, APSysCfgData.nSEARCH_RANGE, strSearchSlbNumber)
    
    '（検索有効範囲は9999無制限）
    bRet = DBColorSlbSearchRead(0, 0, 9999, strSearchSlbNumber)
        
    '検索結果をセット
    If bRet Then
        
        ReDim APSearchListSlbData(0)
        nJ = 0
        
        Select Case APSlbCont.nSearchInputModeSelectedIndex
            Case 0 '新規
                bCmp = False
                nJ = 0
                nSlb_Col_Cnt_MAX = 0
                nFirstDataIndex = 0
                For nI = 0 To UBound(APSearchTmpSlbData) - 1
                    strTmpSlbNumber = APSearchTmpSlbData(nI).slb_no
                    'ｽﾗﾌﾞNo．を比較
                    If strTmpSlbNumber = strSearchSlbNumber Then
                        '状態を比較
                        If CInt(APSearchTmpSlbData(nI).slb_stat) = APSlbCont.nSearchInputStatusSelectedIndex Then
                            bCmp = True '存在
                            '*****************************************************
                            'APSlbCont.nSearchInputModeSelectedIndex = 1 '新規⇒修正
                            '*****************************************************
                            'Exit For
                            'カラー回数の最大数を取得
                            If nSlb_Col_Cnt_MAX < CInt(APSearchTmpSlbData(nI).slb_col_cnt) Then
                                nSlb_Col_Cnt_MAX = CInt(APSearchTmpSlbData(nI).slb_col_cnt)
                            End If
                            If CInt(APSearchTmpSlbData(nI).slb_stat) = 1 Then
                                nFirstDataIndex = nI
                            End If
                        End If
                    End If
                Next nI
                
                
                '新規データ作成追加
                If bCmp Then
                    '保存済みデータ有り
                    APSearchListSlbData(nJ).slb_col_cnt = Format(nSlb_Col_Cnt_MAX + 1, "00")
                Else
                    '保存済みデータ無し
                    APSearchListSlbData(nJ).slb_col_cnt = "01"
                End If
'                If bCmp = False Then
                    
                'APSearchListSlbData(nJ).bAPResEdit = False
                APSearchListSlbData(nJ).bAPScanInput = False
                APSearchListSlbData(nJ).bAPPdfInput = False
                
                APSearchListSlbData(nJ).slb_no = strSearchSlbNumber
                APSearchListSlbData(nJ).slb_chno = Mid(strSearchSlbNumber, 1, 5)
                APSearchListSlbData(nJ).slb_aino = Mid(strSearchSlbNumber, 6)
                
                APSearchListSlbData(nJ).slb_stat = APSlbCont.nSearchInputStatusSelectedIndex
                
                If bCmp Then
                    '保存済みデータ有り
                    '初回データコピー
                    
                    '表示リストにコピー
                    '**********************************************************'
                    'nchtaisl
                    'APSozaiTmpData(0).slb_no = "123451234"      ''スラブNO"
                    APSearchListSlbData(nJ).slb_ksh = APSearchTmpSlbData(nFirstDataIndex).slb_ksh  ''鋼種
                    APSearchListSlbData(nJ).slb_uksk = APSearchTmpSlbData(nFirstDataIndex).slb_uksk ''向先（熱延向先）
                    'APSearchListSlbData(nJ).slb_lngth = APSozaiData.slb_lngth ''長さ
                    'APSearchListSlbData(nJ).slb_color_wei = APSozaiData.slb_color_wei ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
                    APSearchListSlbData(nJ).slb_typ = APSearchTmpSlbData(nFirstDataIndex).slb_typ ''型
                    'APSearchListSlbData(nJ).slb_skin_wei = APSozaiData.slb_skin_wei ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
                    'APSearchListSlbData(nJ).slb_wdth = APSozaiData.slb_wdth ''幅
                    'APSearchListSlbData(nJ).slb_thkns = APSozaiData.slb_thkns ''厚み
                    APSearchListSlbData(nJ).slb_zkai_dte = APSearchTmpSlbData(nFirstDataIndex).slb_zkai_dte ''造塊日（造塊年月日）
                    '**********************************************************'
                    'skjchjdtテーブル
                    'APSozaiData.slb_chno = "12345"        ''チャージNO
                    'APSearchListSlbData(nJ).slb_ccno = APSozaiData.slb_ccno ''CCNO
                    '**********************************************************'
                    
                    '内部:APSozaiDataにコピー
                    '**********************************************************'
                    'nchtaisl
                    APSozaiData.slb_no = APSearchTmpSlbData(nFirstDataIndex).slb_no      ''スラブNO"
                    APSozaiData.slb_ksh = APSearchTmpSlbData(nFirstDataIndex).slb_ksh       ''鋼種
                    APSozaiData.slb_uksk = APSearchTmpSlbData(nFirstDataIndex).slb_uksk         ''向先（熱延向先）
                    APSozaiData.slb_lngth = APSearchTmpSlbData(nFirstDataIndex).slb_lngth       ''長さ
                    APSozaiData.slb_color_wei = APSearchTmpSlbData(nFirstDataIndex).slb_wei   ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
                    APSozaiData.slb_typ = APSearchTmpSlbData(nFirstDataIndex).slb_typ           ''型
'                    APSozaiData.slb_skin_wei = APSearchTmpSlbData(nFirstDataIndex).slb_wei    ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
                    APSozaiData.slb_wdth = APSearchTmpSlbData(nFirstDataIndex).slb_wdth         ''幅
                    APSozaiData.slb_thkns = APSearchTmpSlbData(nFirstDataIndex).slb_thkns      ''厚み
                    APSozaiData.slb_zkai_dte = APSearchTmpSlbData(nFirstDataIndex).slb_zkai_dte ''造塊日（造塊年月日）
                    '**********************************************************'
                    'skjchjdtテーブル
                    APSozaiData.slb_chno = APSearchTmpSlbData(nFirstDataIndex).slb_chno        ''チャージNO
                    APSozaiData.slb_ccno = APSearchTmpSlbData(nFirstDataIndex).slb_ccno        ''CCNO
                    '**********************************************************'
                Else
                    '保存済みデータ無し
                    '**********************
                    '素材統括から読込
                    'bRet = SOZAI_NCHTAISL_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no)
                    
                    bNoRecord = False '2008/08/30 A.K
                    
                    bRet = SOZAI_NCHTAISL_Read(APSearchListSlbData(nJ).slb_no)
                    If UBound(APSozaiTmpData) = 1 Then
                        APSozaiData = APSozaiTmpData(0)
                    Else
                        'NCHTAISL該当レコード無し
                        bNoRecord = True '2008/08/30 A.K
                    End If
                    If bRet = False Then
                        Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                        MsgWnd.OK_Close
                        Exit Sub
                    End If
                    
                    'bRet = SOZAI_SKJCHJDT_Read(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno)
                    bRet = SOZAI_SKJCHJDT_Read(APSearchListSlbData(nJ).slb_chno)
                    If UBound(APSozaiTmpData) = 1 Then
                        APSozaiData.slb_chno = APSozaiTmpData(0).slb_chno
                        APSozaiData.slb_ccno = APSozaiTmpData(0).slb_ccno
                        
                        If bNoRecord Then '2008/08/30 A.K
                            'NCHTAISL該当レコード無しの場合の処理
                            'SKJCHJDTから鋼種,型を抽出
                            APSozaiData.slb_ksh = APSozaiTmpData(0).slb_ksh ''鋼種
                            APSozaiData.slb_typ = APSozaiTmpData(0).slb_typ ''型
                        End If
                        
                    End If
                    If bRet = False Then
                        Call WaitMsgBox(Me, "データベース読込エラーが発生しました。")
                        MsgWnd.OK_Close
                        Exit Sub
                    End If
                    
                    'リストにコピー
                    '**********************************************************'
                    'nchtaisl
                    'APSozaiTmpData(0).slb_no = "123451234"      ''スラブNO"
                    APSearchListSlbData(nJ).slb_ksh = APSozaiData.slb_ksh ''鋼種
                    APSearchListSlbData(nJ).slb_uksk = APSozaiData.slb_uksk ''向先（熱延向先）
                    'APSearchListSlbData(nJ).slb_lngth = APSozaiData.slb_lngth ''長さ
                    'APSearchListSlbData(nJ).slb_color_wei = APSozaiData.slb_color_wei ''重量（ｶﾗｰﾁｪｯｸ用：SEG出側重量）
                    APSearchListSlbData(nJ).slb_typ = APSozaiData.slb_typ ''型
                    'APSearchListSlbData(nJ).slb_skin_wei = APSozaiData.slb_skin_wei ''重量（ｽﾗﾌﾞ肌用：黒皮重量）
                    'APSearchListSlbData(nJ).slb_wdth = APSozaiData.slb_wdth ''幅
                    'APSearchListSlbData(nJ).slb_thkns = APSozaiData.slb_thkns ''厚み
                    APSearchListSlbData(nJ).slb_zkai_dte = APSozaiData.slb_zkai_dte ''造塊日（造塊年月日）
                    '**********************************************************'
                    'skjchjdtテーブル
                    'APSozaiData.slb_chno = "12345"        ''チャージNO
                    'APSearchListSlbData(nJ).slb_ccno = APSozaiData.slb_ccno ''CCNO
                    '**********************************************************'
                    
                    '**********************
                End If
                
                '2016/04/20 - TAI - S
                '作業場
                APSearchListSlbData(nJ).slb_works_sky_tok = works_sky_tok
                '2016/04/20 - TAI - E
                
                ReDim Preserve APSearchListSlbData(UBound(APSearchListSlbData) + 1)
                nJ = nJ + 1
'                End If
            Case 1 '修正
            Case 2 '削除
        End Select
        
        For nI = 0 To UBound(APSearchTmpSlbData) - 1
            APSearchListSlbData(nJ) = APSearchTmpSlbData(nI)
            ReDim Preserve APSearchListSlbData(UBound(APSearchListSlbData) + 1)
            nJ = nJ + 1
        Next nI
    
    End If

    MsgWnd.OK_Close
    
    ' 20090116 add by M.Aoyagi    表示速度対応
    MSFlexGrid1.Visible = False
    
    'グリッドへセット
    Call SetMSFlexGrid1
    
    ' 20090116 add by M.Aoyagi    表示速度対応
    MSFlexGrid1.Visible = True
    
    ' 20090115 add by M.Aoyagi    状態キー変更ボタン制御
    If OptInputMode(0).Value Then
        ' 20090115 add by M.Aoyagi    新規時はキー変更モードボタンを無効
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
        imTextColCntNew.Enabled = False
        imTextColCntNew.Text = ""
        Frame_StatusChg.Enabled = False
        lblColCntNew.Enabled = False
    ElseIf OptInputMode(1).Value Then
        ' 20090115 add by M.Aoyagi    修正時のみキー変更モードボタンを有効
        cmdStatChgMode.Enabled = True
'        cmdStatChgFix.Enabled = True
'        imTextColCntNew.Enabled = True
        Frame_StatusChg.Enabled = True
'        lblColCntNew.Enabled = True
    ElseIf OptInputMode(2).Value Then
        ' 20090115 add by M.Aoyagi    削除時はキー変更モードボタンを無効
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
        imTextColCntNew.Enabled = False
        imTextColCntNew.Text = ""
        Frame_StatusChg.Enabled = False
        lblColCntNew.Enabled = False
    End If
    
End Sub

' @(f)
'
' 機能      : スラブ選択ロック／アンロック
'
' 引き数    : ARG1 - True=ロック／False=アンロック フラグ
'
' 返り値    :
'
' 機能説明  : スラブ選択状態の画面ロック／アンロック制御。
'
' 備考      :COLORSYS
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
'        APSlbCont.bProcessing = True 'スラブ選択ロック用処理中フラグ
'        APSlbCont.strSearchInputSlbNumber = imTextSearchSlbNumber.Text '検索スラブＮｏ．
'        If OptSearchMode(0).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 0 '検索オプション指定インデックス番号
'        ElseIf OptSearchMode(1).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 1 '検索オプション指定インデックス番号
'        ElseIf OptSearchMode(2).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 2 '検索オプション指定インデックス番号
'        ElseIf OptSearchMode(3).Value Then
'            APSlbCont.nSearchInputModeSelectedIndex = 3 '検索オプション指定インデックス番号
'        End If
'        'スラブ選択情報保存
'        APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
'        '子スラブはＯＫボタン時に保存
'        'nChildSelectedIndex As Integer '子スラブ指定インデックス番号 0は未指定
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
'        APSlbCont.bProcessing = False 'スラブ選択ロック用処理中フラグ
    End If
    
    Call MSFlexGrid1_Click

    DoEvents

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
    
    Select Case CallNo
    
    Case CALLBACK_USEIMGDATA
        '既に登録データが存在するシナリオ
        If Result = CALLBACK_ncResOK Then          'OK
            'On Error Resume Next
            '登録済みスキャナーイメージ
            
'            Call ImageDataRead
            'イメージファイル読込み
            'Call ImageLoad
            
            
            'On Error GoTo 0
            'Unload Me
        Else
            
        End If
'        cmdSplitChg.Enabled = True
        
    Case CALLBACK_RES_COLORDATA_DBDEL_REQ
        'データ削除の問合せよりOK
        If Result = CALLBACK_ncResOK Then          'OK
            bRet = ColorSlbData_Load(False) '削除処理実行
        Else
            
        End If
        
        cmdOK.Enabled = True 'ボタン有効
        
    Case CALLBACK_RES_COLORDATA_HOSTDEL_REQ
        'データ削除の問合せよりOK（ビジコンへ削除送信シナリオ）
        If Result = CALLBACK_ncResOK Then          'OK
            'ビジコン送信
            
'           '現地にて調整（通信テスト時）
            APResData.slb_fault_u_judg = "0"
            APResData.slb_fault_d_judg = "0"
            
            APResData.slb_no = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
            APResData.host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_dte
            APResData.host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_tme
            APResData.fail_host_wrt_dte = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte
            APResData.fail_host_wrt_tme = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_tme
            
            frmHostSend.SetCallBack Me, CALLBACK_RES_COLORDATA_HOSTDEL_REQ2
            frmHostSend.Show vbModal, Me 'ビジコン送信中は、他の処理を不可とする為、vbModalとする。
        Else
            'キャンセル
            cmdOK.Enabled = True 'ボタン有効
        End If
        
    Case CALLBACK_RES_COLORDATA_HOSTDEL_REQ2
        'ビジコン削除処理よりOK（ビジコンへ削除送信シナリオ）
        If Result = CALLBACK_ncResOK Then          'OK
            bRet = ColorSlbData_Load(False) '削除処理実行
        ElseIf Result = CALLBACK_ncResSKIP Then 'SKIP
            bRet = ColorSlbData_Load(False) '削除処理実行
        Else
            'ビジコンエラー発生
            Call WaitMsgBox(Me, "ビジコン通信エラーが発生した為、ＤＢ削除処理は中断されました。")
        End If
        
        cmdOK.Enabled = True 'ボタン有効
        
    Case CALLBACK_RES_STATECHANGE_COLOR
        '状態キー変更OK
        If Result = CALLBACK_ncResOK Then          'OK
            Call cmdStatChgFixExe
        Else
            cmdStatChgFix.Enabled = True            'ボタン有効
            Call WaitMsgBox(Me, "状態変更処理を中断しました。")
        End If
    
    End Select

End Sub

' @(f)
'
' 機能      : グリッド１初期化
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : グリッド１の初期化を行う。
'
' 備考      :
'
Private Sub InitMSFlexGrid1()

    Dim nJ As Integer
    Dim nRow As Integer
    Dim nCol As Integer

    nMSFlexGrid1_Selected_Row = 0
    lbl_nMSFlexGrid1_Selected_Row.Caption = nMSFlexGrid1_Selected_Row
    
    MSFlexGrid1.SelectionMode = flexSelectionByRow
    MSFlexGrid1.FixedCols = 0
    ' 20090115 modify by M.Aoyagi    画像枚数変更の為加算
'    MSFlexGrid1.Cols = 14 + 1
    
    '2016/04/20 - TAI - S
'    MSFlexGrid1.Cols = 16 + 1
    MSFlexGrid1.Cols = 17 + 1
    '2016/04/20 - TAI - E
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
    MSFlexGrid1.ColWidth(nCol) = 1400
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "鋼種"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1600
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "スラブNo."
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "型"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "向先"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "状態"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "ｶﾗｰ回数"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "ｶﾗｰ実績"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1300                      ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'    MSFlexGrid1.TextMatrix(0, nCol) = "ｶﾗｰｲﾒｰｼﾞ"           ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.TextMatrix(0, nCol) = "ｲﾒｰｼﾞ"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1300                      ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'    MSFlexGrid1.TextMatrix(0, nCol) = "ｶﾗｰPDF"             ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.TextMatrix(0, nCol) = "PDF"
    
    ' 20090115 add by M.Aoyagi    画像枚数追加
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "枚数"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "異常報告"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1300                      ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'    MSFlexGrid1.TextMatrix(0, nCol) = "異常ｲﾒｰｼﾞ"          ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.TextMatrix(0, nCol) = "ｲﾒｰｼﾞ"
    
    nCol = nCol + 1
'    MSFlexGrid1.ColWidth(nCol) = 1300                      ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'    MSFlexGrid1.TextMatrix(0, nCol) = "異常PDF"            ' 20090129 modify by M.Aoyagi    表を詰める為文字を簡略化
    MSFlexGrid1.TextMatrix(0, nCol) = "PDF"
    
    ' 20090115 add by M.Aoyagi    画像枚数追加
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 800
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "枚数"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "処置指示"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1300
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "処置結果"
    
    '2016/04/20 - TAI - S
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "作業場"
    '2016/04/20 - TAI - E

    'タイトル行
    For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000012
        MSFlexGrid1.CellBackColor = &H8000000F
    Next nJ

End Sub

' @(f)
'
' 機能      : グリッド１データセット
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : グリッド１のデータセットを行う。
'
' 備考      :
'
Private Sub SetMSFlexGrid1()
    Dim nJ As Integer
    Dim nCol As Integer
    Dim nRow As Integer
    
    MSFlexGrid1.Rows = 1 + UBound(APSearchListSlbData)
    
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
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_ksh '"鋼種"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignLeftCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_chno & "-" & APSearchListSlbData(nRow - 1).slb_aino '"スラブNo."
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_typ '"型"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_uksk '"向先"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = ConvDpOutStat(conDefine_SYSMODE_COLOR, CInt(APSearchListSlbData(nRow - 1).slb_stat)) '"状態"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_col_cnt '"ｶﾗｰ回数"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.CellPictureAlignment = flexAlignCenterCenter
        
        If APSearchListSlbData(nRow - 1).fail_sys_wrt_dte <> "" Then
            '異常報告が存在する時
            MSFlexGrid1.TextMatrix(nRow, nCol) = "保留"
            Set MSFlexGrid1.CellPicture = PicSigRed.Picture
            
            If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "1" Then
                'ＷＥＢで全完了の場合
                If APSearchListSlbData(nRow - 1).fail_res_host_send <> "2" Then
                    '保留だが、処置を完了し、未送信ではない場合
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte '"ｶﾗｰ実績"（初回登録日）
                    Set MSFlexGrid1.CellPicture = Nothing
                End If
            End If
        ElseIf APSearchListSlbData(nRow - 1).host_send <> "" Then
            'ビジコン通信が送信済みの場合（ＯＫ、ＮＧにかかわらず）
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte '"ｶﾗｰ実績"（初回登録日）
            Set MSFlexGrid1.CellPicture = Nothing
        Else
            If APSearchListSlbData(nRow - 1).sys_wrt_dte <> "" Then
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            Set MSFlexGrid1.CellPicture = Nothing
            Else
                If IsDEBUG("DISP") Then
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).sys_wrt_dte & "?"
                    Set MSFlexGrid1.CellPicture = Nothing
                Else
                    MSFlexGrid1.TextMatrix(nRow, nCol) = ""
                    Set MSFlexGrid1.CellPicture = Nothing
                End If
            End If
        End If
    
        If APSearchListSlbData(nRow - 1).host_send = "0" Then
            'ビジコン通信が異常送信の場合
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "通信ｴﾗｰ"
'            Set MSFlexGrid1.CellPicture = Nothing
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPScanInput, "○", "") '"ｶﾗｰｲﾒｰｼﾞ"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).bAPPdfInput Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "○"
        ElseIf APSearchListSlbData(nRow - 1).sAPPdfInput_ReqDate <> "" Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "△"
        Else
            MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        End If
        'MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPPdfInput, "○", "") '"ｶﾗｰPDF"
    
        ' 20090115 add by M.Aoyagi    画像枚数表示追加の為
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).PhotoImgCnt1
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).fail_host_send <> "" Then
            'ビジコン通信が送信済みの場合（ＯＫ、ＮＧにかかわらず）
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_sys_wrt_dte '"異常報告"（初回登録日）
        Else
            If APSearchListSlbData(nRow - 1).fail_sys_wrt_dte <> "" Then
                If IsDEBUG("DISP") Then
                    '未送信
                    MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_sys_wrt_dte & "?"
                Else
                    MSFlexGrid1.TextMatrix(nRow, nCol) = ""
                End If
            Else
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            End If
        End If
    
        If APSearchListSlbData(nRow - 1).fail_host_send = "0" Then
            'ビジコン通信が異常送信の場合
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "通信ｴﾗｰ"
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPFailScanInput, "○", "") '"異常ｲﾒｰｼﾞ"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).bAPFailPdfInput Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "○"
        ElseIf APSearchListSlbData(nRow - 1).sAPFailPdfInput_ReqDate <> "" Then
            MSFlexGrid1.TextMatrix(nRow, nCol) = "△"
        Else
            MSFlexGrid1.TextMatrix(nRow, nCol) = ""
        End If
        'MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APSearchListSlbData(nRow - 1).bAPFailPdfInput, "○", "") '"異常PDF"
    
        ' 20090115 add by M.Aoyagi    画像枚数表示追加の為
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).PhotoImgCnt2
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_dir_sys_wrt_dte '"処置指示"
    
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "1" Then
            '完了
            MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).fail_res_sys_wrt_dte '"処置結果"
        Else
            If APSearchListSlbData(nRow - 1).fail_res_cmp_flg = "0" Then
                '未完了
                MSFlexGrid1.TextMatrix(nRow, nCol) = "△"
            Else
                '登録無し
                MSFlexGrid1.TextMatrix(nRow, nCol) = ""
            End If
        End If
    
    
        '2016/04/20 - TAI - S
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APSearchListSlbData(nRow - 1).slb_works_sky_tok '"作業場"
        '2016/04/20 - TAI - E

        'ＷＥＢで全完了の場合
        If APSearchListSlbData(nRow - 1).fail_res_host_send = "2" Then
            'ビジコン通信が異常送信の場合
            MSFlexGrid1.TextMatrix(nRow, nCol) = "未送信"
        End If
    
        If APSearchListSlbData(nRow - 1).fail_res_host_send = "0" Then
            'ビジコン通信が異常送信の場合
'            MSFlexGrid1.TextMatrix(nRow, nCol) = "通信ｴﾗｰ"
            MSFlexGrid1.CellForeColor = conDefine_Color_ForColor_HOST_ERROR
        End If
    
    Next nRow

    If MSFlexGrid1.Rows > 1 Then
        MSFlexGrid1.Row = 1
    End If

End Sub

' @(f)
'
' 機能      : 状態決定ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : テーブル、IMGファイル、PDFファイル、SCANファイルの状態キーを変更する。
'
' 備考      :
'
Private Sub cmdStatChgFix_Click()
    Dim bRet      As Boolean
    Dim iRet      As Integer
    Dim strSource As String
    Dim strDestination As String
    Dim sSlbno    As String
    Dim sChno     As String
    Dim sAino     As String
    Dim sStat     As String
    Dim sColCnt   As String
    Dim sStatNew  As String
    Dim sCheckErr As String
    
    Dim fmessage As Object
    Set fmessage = New MessageYN

    ' 状態変更開始ログ書込み
    Call MsgLog(conProcNum_MAIN, "『状態』変更を開始します。")
    cmdStatChgFix.Enabled = False       'ボタン無効

    'スラブ選択チェック ***************************************************************************
    APSlbCont.nListSelectedIndexP1 = nMSFlexGrid1_Selected_Row
    If APSlbCont.nListSelectedIndexP1 = 0 Then
        Call WaitMsgBox(Me, "『状態』変更を行うスラブを選択してください。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    
    '指定カラー回数入力チェック *******************************************************************
    sColCnt = imTextColCntNew.Text
    If Len(sColCnt) <= 0 Then
        Call WaitMsgBox(Me, "変更先『カラー回数』を入力してください。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    sColCnt = Format(CInt(imTextColCntNew.Text), "00")
    
    ' 必要データ設定
    sSlbno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
    sChno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno
    sAino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino
    sStat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
'    sColCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
    If OptStatus(0).Value Then
        sStatNew = "0" '黒皮
    ElseIf OptStatus(1).Value Then
        sStatNew = "1" '1ht後
    ElseIf OptStatus(2).Value Then
        sStatNew = "2" '2ht後
    ElseIf OptStatus(3).Value Then
        sStatNew = "3" '3ht後
    ElseIf OptStatus(4).Value Then
        sStatNew = "4" '4ht後
    ElseIf OptStatus(5).Value Then
        sStatNew = "5" '5ht後
    ElseIf OptStatus(6).Value Then
        sStatNew = "6" '6ht後
    End If
    sCheckErr = ""

    ' 変更先のデータが存在するかチェック **********************************************************
    iRet = DBStatChgCheckCOLOR(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, sStatNew, sStat, sColCnt)
    If iRet = 1 Then
        sCheckErr = "既にデータが存在するので変更できません。"
    ElseIf iRet = 2 Then
        sCheckErr = "指示データが存在するので変更できません。"
    End If
    
    ' IMGフォルダが存在するかチェック *************************************************************
'    PhotoImgCount("SKIN", oDS.Fields("slb_chno").Value, oDS.Fields("slb_aino").Value, oDS.Fields("slb_stat").Value, "00")
    bRet = StatChgFoldCheck("IMG", "COLOR", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にカラーチェック画像ファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にカラーチェック画像ファイルが存在します。"
        End If
    End If
    bRet = StatChgFoldCheck("IMG", "SLBFAIL", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にスラブ異常画像ファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にスラブ異常画像ファイルが存在します。"
        End If
    End If
    
    ' PDFフォルダが存在するかチェック *************************************************************
    bRet = StatChgFoldCheck("PDF", "COLOR", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にカラーチェックPDFファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にカラーチェックPDFファイルが存在します。"
        End If
    End If
    bRet = StatChgFoldCheck("PDF", "SLBFAIL", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にスラブ異常PDFファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にスラブ異常PDFファイルが存在します。"
        End If
    End If
    
    ' SCANフォルダが存在するかチェック ************************************************************
    bRet = StatChgFoldCheck("SCAN", "COLOR", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にカラーチェックスキャンファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にカラーチェックスキャンファイルが存在します。"
        End If
    End If
    bRet = StatChgFoldCheck("SCAN", "SLBFAIL", sChno, sAino, sStatNew, sColCnt)
    If bRet = False Then
        If sCheckErr = "" Then
            sCheckErr = "既にスラブ異常スキャンファイルが存在します。"
        Else
            sCheckErr = sCheckErr & vbCrLf & "既にスラブ異常スキャンファイルが存在します。"
        End If
    End If

    '既存ファイルチェック最終確認 *****************************************************************
    If sCheckErr <> "" Then
        Call WaitMsgBox(Me, sCheckErr)
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If

    '状態変更実行確認 *****************************************************************************
    fmessage.MsgText = "指定データの『状態』キーを変更します。" & vbCrLf & "よろしいですか？"
    fmessage.AutoDelete = False
    fmessage.SetCallBack Me, CALLBACK_RES_STATECHANGE_COLOR, False
    fmessage.Show vbModal, Me '他の処理を不可とする為、vbModalとする。
    Set fmessage = Nothing
End Sub
    
' @(f)
'
' 機能      : 状態決定ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : テーブル、IMGファイル、PDFファイル、SCANファイルの状態キーを変更する。
'
' 備考      :
'
Private Sub cmdStatChgFixExe()
    Dim bRet      As Boolean
    Dim iRet      As Integer
    Dim strSource As String
    Dim strDestination As String
    Dim sSlbno     As String
    Dim sChno      As String
    Dim sAino      As String
    Dim sStat      As String
    Dim sColCnt    As String
    Dim sColCntNew As String
    Dim sStatNew   As String
    
    cmdStatChgFix.Enabled = False       'ボタン無効
    
    Dim MsgWnd As Message
    Set MsgWnd = New Message
    
    ' 状態変更開始ログ書込み
'    Call MsgLog(conProcNum_MAIN, "状態変更を開始します。")
    
    ' 必要データ設定
    sSlbno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no
    sChno = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_chno
    sAino = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_aino
    sStat = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_stat
    sColCnt = APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_col_cnt
    sColCntNew = imTextColCntNew.Text
    sColCntNew = Format(CInt(imTextColCntNew.Text), "00")
    If OptStatus(0).Value Then
        sStatNew = "0" '黒皮
    ElseIf OptStatus(1).Value Then
        sStatNew = "1" '1ht後
    ElseIf OptStatus(2).Value Then
        sStatNew = "2" '2ht後
    ElseIf OptStatus(3).Value Then
        sStatNew = "3" '3ht後
    ElseIf OptStatus(4).Value Then
        sStatNew = "4" '4ht後
    ElseIf OptStatus(5).Value Then
        sStatNew = "5" '5ht後
    ElseIf OptStatus(6).Value Then
        sStatNew = "6" '6ht後
    End If

    ' IMGフォルダ変更 *****************************************************************************
    bRet = StatChgFoldFix("IMG", "COLOR", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "画像ファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    bRet = StatChgFoldFix("IMG", "SLBFAIL", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "画像ファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    
    ' PDFフォルダ変更 *****************************************************************************
    bRet = StatChgFoldFix("PDF", "COLOR", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "PDFファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    bRet = StatChgFoldFix("PDF", "SLBFAIL", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "PDFファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    
    ' SCANフォルダ変更 ****************************************************************************
    bRet = StatChgFoldFix("SCAN", "COLOR", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "スキャンファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If
    bRet = StatChgFoldFix("SCAN", "SLBFAIL", sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "スキャンファイルの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If

    ' テーブル変更 ********************************************************************************
'    bRet = DBStatChgFix(APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).slb_no, sStatNew)
    bRet = DBStatChgFixCOLOR(sSlbno, sChno, sAino, sStat, sStatNew, sColCnt, sColCntNew)
    If bRet = False Then
        Call WaitMsgBox(Me, "データの変更に失敗しました。")
        cmdStatChgFix.Enabled = True            'ボタン有効
        Exit Sub
    End If

    Call WaitMsgBox(Me, "状態変更は正常に終了しました。" & vbCrLf & "PDFの再作成・ビジコン送信は手動で行なって下さい。")
    cmdStatChgFix.Enabled = True        'ボタン有効

    ' 画面更新
    Call cmdSearch_Click

End Sub

' @(f)
'
' 機能      : 状態変更ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 状態変更の為、状態オプションボタンを有効にする。
'
' 備考      :
'
Private Sub cmdStatChgMode_Click()
    Dim i As Integer

    ' 20090115 add by M.Aoyagi    状態を変更可能にする
    Frame_Status.Enabled = True '有効
    For i = 0 To 6
        OptStatus(i).Enabled = True
    Next i

    ' 状態決定ボタン・カラー回数テキストボックスを有効にする
    cmdStatChgFix.Enabled = True
    imTextColCntNew.Enabled = True
    lblColCntNew.Enabled = True

End Sub

Private Sub imTextSearchSlbNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}", True
    End Select
End Sub

' @(f)
'
' 機能      : スラブ番号入力BOXキー押
'
' 引き数    : ARG1 - ASCIIコード
'
' 返り値    :
'
' 機能説明  : スラブ番号入力BOXキー押時の処理を行う。
'
' 備考      :
'
Private Sub imTextSearchSlbNumber_KeyPress(KeyAscii As Integer)
    Dim nI As Integer
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = Asc("-") Then
    ElseIf KeyAscii = Asc("*") Then
        'For nI = 1 To LenB(imTextSearchSlbNumber.Text)
        '    If Mid(imTextSearchSlbNumber.Text, nI, 1) = "*" Then
        '        KeyAscii = 0
        '        Beep
        '    End If
        'Next nI
    Else
        If KeyAscii <> 10 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub



' @(f)
'
' 機能      : グリッド１クリック
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : グリッド１クリック時の処理を行う。
'
' 備考      :
'
Private Sub MSFlexGrid1_Click()
    Dim nJ As Integer
    Dim nNowRow As Integer
    Dim nNowSplitNum As Integer
    Dim nRet As Integer

    ' 20090116 add by M.Aoyagi    表示速度対応
    MSFlexGrid1.Visible = False

    bMouseControl = True

    '現在のRowを一時保存
    nNowRow = MSFlexGrid1.Row

    '以前のセレクト行を未セレクト状態に戻す。
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H80000008
        MSFlexGrid1.CellBackColor = &H80000005
        Next nJ
    Else
        'タイトル行の色を付け直す。
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H80000012
        MSFlexGrid1.CellBackColor = &H8000000F
        Next nJ
    End If

    '現在のセレクト行番号を保存
    nMSFlexGrid1_Selected_Row = nNowRow
    lbl_nMSFlexGrid1_Selected_Row.Caption = nMSFlexGrid1_Selected_Row
    
    '現在の行をセレクト行にする。
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
            MSFlexGrid1.Col = nJ
            If MSFlexGrid1.Enabled Then
                '選択中の色
                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
                    '削除モードの場合
                    If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8080FF
                Else
                    If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8000000D
                End If
                
                '削除モードか？
                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
                    '削除モード
                    cmdDirRes.Enabled = False '禁止！
                Else
                    If APSearchListSlbData(nMSFlexGrid1_Selected_Row - 1).fail_dir_sys_wrt_dte <> "" Then
                        '指示有り
                        cmdDirRes.Enabled = True
                    Else
                        '指示無し
                        cmdDirRes.Enabled = False
                    End If
                End If
            Else
                '選択ロック中の色
                If MSFlexGrid1.CellForeColor <> conDefine_Color_ForColor_HOST_ERROR Then MSFlexGrid1.CellForeColor = &H8000000E
                MSFlexGrid1.CellBackColor = &H808080
            End If
        Next nJ
        If MSFlexGrid1.Enabled Then
            '選択中
        Else
            '選択ロック
        End If
    
        ' 20090130 add by M.Aoyagi    選択行のカラー回数を『状態変更―カラー回数』に設定
        If APSlbCont.nSearchInputModeSelectedIndex = 1 Then
            Me.imTextColCntNew.Text = APSearchListSlbData(nMSFlexGrid1_Selected_Row - 1).slb_col_cnt
        End If
    Else
    End If

    ' 20090116 add by M.Aoyagi    表示速度対応
    MSFlexGrid1.Visible = True

End Sub

' @(f)
'
' 機能      : グリッド１フォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : グリッド１フォーカス取得時の処理を行う。
'
' 備考      :
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
' 機能      : グリッド１セル変更
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : グリッド１セル変更時の処理を行う。
'
' 備考      :
'
Private Sub MSFlexGrid1_SelChange()
    If bMouseControl = False Then
        Call MSFlexGrid1_Click
    End If
    bMouseControl = False
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
    
    Dim nI As Integer
    
    bMouseControl = False
    
'    For nI = 0 To 3
'        lblSearchMAX(nI).Caption = APSysCfgData.nSEARCH_MAX(nI)
'    Next nI
    
    '選択番号表示
    If IsDEBUG("DISP") Then
        lbl_nMSFlexGrid1_Selected_Row.Visible = True
'        lbl_nMSFlexGrid2_Selected_Row.Visible = True
    End If
    
    '2016/04/20 - TAI - S
    '作業場所表示
    If works_sky_tok = WORKS_SKY Then
        lbl_works.Caption = "SKY"               'SKY
        lbl_works.ForeColor = &HFF              '赤
    ElseIf works_sky_tok = WORKS_TOK Then
        lbl_works.Caption = "特鋼"              '特鋼
        lbl_works.ForeColor = &HFF0000          '青
    End If
    '2016/04/20 - TAI - E

'    cmdOK.Enabled = False
    
'    LEAD_SCAN.UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
'    LEAD_SCAN.EnableMethodErrors = False 'False   システムエラーイベントを発生させない
'    LEAD_SCAN.EnableTwainEvent = True
'    LEAD_SCAN.PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
'
'    For nI = 0 To 1
'        LEAD1(nI).UnlockSupport L_SUPPORT_DOCUMENT, L_KEY_DOCUMENT
'        LEAD1(nI).EnableMethodErrors = False 'False   システムエラーイベントを発生させない
'        LEAD1(nI).EnableTwainEvent = True
'        LEAD1(nI).PaintZoomFactor = APSysCfgData.nIMAGE_SIZE(conDefine_SYSMODE_SKIN)
'    Next nI
    
    Call InitMSFlexGrid1

'    If APSlbCont.bProcessing Then 'スラブ選択ロック用処理中フラグ
        imTextSearchSlbNumber.Text = APSlbCont.strSearchInputSlbNumber  '検索スラブＮｏ．
        
        OptInputMode(APSlbCont.nSearchInputModeSelectedIndex).Value = True '入力モード指定インデックス番号
        
        bOptInputModeValue(0) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, True, False)
        bOptInputModeValue(1) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 1, True, False)
        bOptInputModeValue(2) = IIf(APSlbCont.nSearchInputModeSelectedIndex = 2, True, False)
        
        OptStatus(APSlbCont.nSearchInputStatusSelectedIndex).Value = True '状態選択指定インデックス番号
        
        '指示無し
        cmdDirRes.Enabled = False
        
        ' 20090115 add by M.Aoyagi    キー変更モードボタンの初期設定
        cmdStatChgMode.Enabled = False
        cmdStatChgFix.Enabled = False
        imTextColCntNew.Enabled = False
        imTextColCntNew.Text = ""
        Frame_StatusChg.Enabled = False
        lblColCntNew.Enabled = False
        
        'スラブ選択情報
        nMSFlexGrid1_Selected_Row = APSlbCont.nListSelectedIndexP1
        Call SetMSFlexGrid1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        Call MSFlexGrid1_Click
        Call SlbSelLock(True)
        
'    End If

End Sub

' @(f)
'
' 機能      : 入力モードオプションクリック
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 入力モードオプションクリック時の処理を行う。
'
' 備考      :conDefine_ColorActive or conDefine_ColorNotActive
'           :COLORSYS
'
Private Sub OptInputMode_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0 '新規
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
                
                Frame_Status.Enabled = True '有効
                For i = 0 To 6
                    OptStatus(i).Enabled = True
                Next i
            End If
        
            ' 20090115 add by M.Aoyagi    新規時はキー変更モードボタンを無効
            cmdStatChgMode.Enabled = False
            cmdStatChgFix.Enabled = False
            imTextColCntNew.Enabled = False
            imTextColCntNew.Text = ""
            Frame_StatusChg.Enabled = False
            lblColCntNew.Enabled = False
        Case 1 '修正
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
        
                Frame_Status.Enabled = False '無効
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
        Case 2 '削除
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
    
            Frame_Status.Enabled = False '無効
            For i = 0 To 6
                OptStatus(i).Enabled = False
            Next i
    
            ' 20090115 add by M.Aoyagi    削除時はキー変更モードボタンを無効
            cmdStatChgMode.Enabled = False
            cmdStatChgFix.Enabled = False
            imTextColCntNew.Enabled = False
            imTextColCntNew.Text = ""
            Frame_StatusChg.Enabled = False
            lblColCntNew.Enabled = False
    End Select

    If bOptInputModeValue(Index) = False Then
        '変化があった場合
        For i = 0 To 2
            bOptInputModeValue(i) = False
        Next i
        bOptInputModeValue(Index) = True
        
        nMSFlexGrid1_Selected_Row = 0
        APSlbCont.nListSelectedIndexP1 = 0
        
        'スラブ検索リストクリア
        ReDim APSearchListSlbData(0)
        'グリッドへセット
        Call SetMSFlexGrid1
        
    End If

End Sub

' @(f)
'
' 機能      : 入力モードオプションフォーカス取得
'
' 引き数    : ARG1 - インデックス番号
'
' 返り値    :
'
' 機能説明  : 入力モードオプションフォーカス取得時の処理を行う。
'
' 備考      :COLORSYS
'
Private Sub OptInputMode_GotFocus(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0 '新規
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
                
                Frame_Status.Enabled = True '有効
                For i = 0 To 6
                    OptStatus(i).Enabled = True
                Next i
            End If
        Case 1 '修正
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
        
                Frame_Status.Enabled = False '無効
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
        Case 2 '削除
            If OptInputMode(Index).Value Then
                imTextSearchSlbNumber.Enabled = True
                imTextSearchSlbNumber.BackColor = conDefine_ColorActive
            End If
    
                Frame_Status.Enabled = False '無効
                For i = 0 To 6
                    OptStatus(i).Enabled = False
                Next i
    End Select
End Sub

' @(f)
'
' 機能      : 削除問合せ
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 削除問合せ処理。
'
' 備考      :
'
Private Sub ColorDataDel_REQ()

    Dim bHostSendCmp As Boolean

    Dim fmessage As Object
    Set fmessage = New MessageYN

    bHostSendCmp = False '未送信
    If APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).host_wrt_dte <> "" Then
        'カラーチェック検査表で送信済み
        bHostSendCmp = True
    ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_host_wrt_dte <> "" Then
        'スラブ異常報告で送信済み
        bHostSendCmp = True
    ElseIf APSearchListSlbData(APSlbCont.nListSelectedIndexP1 - 1).fail_res_cmp_flg = "1" Then
        '処置結果報告で送信済み
        bHostSendCmp = True
    End If

    If bHostSendCmp Then
        'ビジコンへ削除要求が必要！
        fmessage.MsgText = "選択したデータはビジコン送信済みの為、" & vbCrLf & "ビジコンへ削除を送信後ＤＢを削除します。" & vbCrLf & "よろしいですか？"
    '    fmessage.AutoDelete = True
        fmessage.AutoDelete = False
    '    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SLBFAIL, True
        fmessage.SetCallBack Me, CALLBACK_RES_COLORDATA_HOSTDEL_REQ, False
    Else
        fmessage.MsgText = "選択したデータを削除します。" & vbCrLf & "よろしいですか？"
    '    fmessage.AutoDelete = True
        fmessage.AutoDelete = False
    '    fmessage.SetCallBack Me, CALLBACK_RES_DBSNDDATA_SLBFAIL, True
        fmessage.SetCallBack Me, CALLBACK_RES_COLORDATA_DBDEL_REQ, False
    End If
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

