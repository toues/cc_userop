VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmSrvNextProcess 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "次工程マスタ"
   ClientHeight    =   4185
   ClientLeft      =   720
   ClientTop       =   900
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2340
      TabIndex        =   3
      Top             =   3420
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelNextProcess 
      Caption         =   "削除"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5040
      TabIndex        =   2
      Top             =   2700
      Width           =   1200
   End
   Begin VB.ListBox lstNextProcess 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      ItemData        =   "frmSrvNextProcess.frx":0000
      Left            =   1020
      List            =   "frmSrvNextProcess.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3795
   End
   Begin VB.CommandButton cmdAddNextProcess 
      Caption         =   "追加"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5040
      TabIndex        =   1
      Top             =   180
      Width           =   1200
   End
   Begin imText6Ctl.imText imtxtNextProcCode 
      Height          =   375
      Left            =   1020
      TabIndex        =   5
      Top             =   180
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   661
      Caption         =   "frmSrvNextProcess.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSrvNextProcess.frx":0072
      Key             =   "frmSrvNextProcess.frx":0090
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
      Format          =   "Aa9Ｚ"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   1
      MaxLength       =   32
      LengthAsByte    =   -1
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "入力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmSrvNextProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmSrvNextProcess.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　次工程マスタ追加／削除フォーム
' 　本モジュールは次工程マスタ追加／削除フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

' @(f)
'
' 機能      : 追加ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 追加ボタン処理。
'
' 備考      :
'
Private Sub cmdAddNextProcess_Click()
    Dim nI As Integer
    Dim bSearch As Boolean
    
    If Trim(imtxtNextProcCode.Text) = "" Then Exit Sub
    
    cmdAddNextProcess.Enabled = False
    
    bSearch = False
    
    For nI = 1 To lstNextProcess.ListCount
        If Trim(imtxtNextProcCode.Text) = lstNextProcess.List(nI - 1) Then
            bSearch = True
            Exit For
        End If
    Next nI
    
    If bSearch = False Then
        lstNextProcess.AddItem Trim(imtxtNextProcCode.Text)
    End If
        
    cmdAddNextProcess.Enabled = True

End Sub

' @(f)
'
' 機能      : 削除ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 削除ボタン処理。
'
' 備考      :
'
Private Sub cmdDelNextProcess_Click()

    If lstNextProcess.ListIndex > -1 Then
        lstNextProcess.RemoveItem lstNextProcess.ListIndex
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
' 備考      :
'
Private Sub cmdOK_Click()
    
    Dim nI As Integer

'    ReDim APSysCfgData.Group(0)
'    APSysCfgData.nGroupCount = lstGroup.ListCount
'
'    For nI = 1 To APSysCfgData.nGroupCount
'        APSysCfgData.Group(nI - 1) = lstGroup.List(nI - 1)
'        ReDim Preserve APSysCfgData.Group(UBound(APSysCfgData.Group) + 1)
'    Next nI
'
'    'クローズ時にレジストリへ反映
'    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nGroupCount", APSysCfgData.nGroupCount
'    For nI = 1 To APSysCfgData.nGroupCount
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "Group" & CStr(nI), APSysCfgData.Group(nI - 1)
'    Next nI
'
'
'    ReDim APSysCfgData.StaffNumber(0)
'    ReDim APSysCfgData.StaffName(0)
'    APSysCfgData.nStaffCount = lstStaff.ListCount
'
'    For nI = 1 To APSysCfgData.nStaffCount
'        APSysCfgData.StaffNumber(nI - 1) = Left(lstStaff.List(nI - 1), InStr(lstStaff.List(nI - 1), ":") - 1)
'        APSysCfgData.StaffName(nI - 1) = Mid(lstStaff.List(nI - 1), InStr(lstStaff.List(nI - 1), ":") + 1)
'        ReDim Preserve APSysCfgData.StaffNumber(UBound(APSysCfgData.StaffNumber) + 1)
'        ReDim Preserve APSysCfgData.StaffName(UBound(APSysCfgData.StaffName) + 1)
'    Next nI
'
'    'クローズ時にレジストリへ反映
'    SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nStaffCount", APSysCfgData.nStaffCount
'    For nI = 1 To APSysCfgData.nStaffCount
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "StaffNumber" & CStr(nI), APSysCfgData.StaffNumber(nI - 1)
'        SaveSetting conReg_APPNAME, conReg_APSYSCFG, "StaffName" & CStr(nI), APSysCfgData.StaffName(nI - 1)
'    Next nI
'
'    Unload Me
'
'    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
'    Set cCallBackObject = Nothing
    
    
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''スラブ肌調査入力
            ReDim APNextProcDataSkin(1)
            '****
                '空白はシステムで管理（必ず追加）
                APNextProcDataSkin(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataSkin(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataSkin(UBound(APNextProcDataSkin) + 1)
            Next nI
        
            '次工程マスター保存(SKIN)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountSkin", UBound(APNextProcDataSkin)
            For nI = 1 To UBound(APNextProcDataSkin)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataSkin" & CStr(nI), APNextProcDataSkin(nI - 1).inp_NextProc
            Next nI
        
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            ReDim APNextProcDataColor(1)
            '****
                '空白はシステムで管理（必ず追加）
                APNextProcDataColor(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataColor(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
            Next nI

            '次工程マスター保存(COLOR)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
            For nI = 1 To UBound(APNextProcDataColor)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
            Next nI
        
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            '****
                '空白はシステムで管理（必ず追加）
                APNextProcDataColor(0).inp_NextProc = ""
            '****
            For nI = 1 To lstNextProcess.ListCount
                APNextProcDataColor(nI).inp_NextProc = lstNextProcess.List(nI - 1)
                ReDim Preserve APNextProcDataColor(UBound(APNextProcDataColor) + 1)
            Next nI

            '次工程マスター保存(COLOR)
            SaveSetting conReg_APPNAME, conReg_APSYSCFG, "nNextProcDataCountColor", UBound(APNextProcDataColor)
            For nI = 1 To UBound(APNextProcDataColor)
                SaveSetting conReg_APPNAME, conReg_APSYSCFG, "NextProcDataColor" & CStr(nI), APNextProcDataColor(nI - 1).inp_NextProc
            Next nI

    End Select
    
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
    
    lstNextProcess.Clear
    
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''スラブ肌調査入力
            '空白を除き編集可能
            For nI = 2 To UBound(APNextProcDataSkin)
                lstNextProcess.AddItem APNextProcDataSkin(nI - 1).inp_NextProc
            Next nI
        
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            '空白を除き編集可能
            For nI = 2 To UBound(APNextProcDataColor)
                lstNextProcess.AddItem APNextProcDataColor(nI - 1).inp_NextProc
            Next nI
        
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            '空白を除き編集可能
            For nI = 2 To UBound(APNextProcDataColor)
                lstNextProcess.AddItem APNextProcDataColor(nI - 1).inp_NextProc
            Next nI

    End Select

End Sub

' @(f)
'
' 機能      : 次工程入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 次工程入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtNextProcCode_GotFocus()
    imtxtNextProcCode.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : 次工程入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 次工程入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtNextProcCode_LostFocus()
    imtxtNextProcCode.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : 次工程リストBOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 次工程リストBOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub lstNextProcess_GotFocus()
'    lstNextProcess.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : 次工程リストBOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 次工程リストBOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub lstNextProcess_LostFocus()
'    lstNextProcess.BackColor = conDefine_ColorBKLostFocus
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



