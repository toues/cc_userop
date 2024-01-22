VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmOpRegWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "スタッフ名登録"
   ClientHeight    =   4530
   ClientLeft      =   1635
   ClientTop       =   1245
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7080
   Begin imText6Ctl.imText imtxtStaffName 
      Height          =   375
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   661
      Caption         =   "frmOpRegWnd.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmOpRegWnd.frx":006E
      Key             =   "frmOpRegWnd.frx":008C
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
   Begin VB.CommandButton cmdAddStaff 
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
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
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
      Left            =   3000
      TabIndex        =   4
      Top             =   3780
      Width           =   1200
   End
   Begin VB.ListBox lstStaff 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      ItemData        =   "frmOpRegWnd.frx":00D0
      Left            =   1740
      List            =   "frmOpRegWnd.frx":00D2
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3795
   End
   Begin VB.CommandButton cmdDelStaff 
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
      Left            =   5760
      TabIndex        =   3
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label lblStaffName 
      BackStyle       =   0  '透明
      Caption         =   "スタッフ名"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmOpRegWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmOpRegWnd.Frm                ver 1.00 ( '2008/05 SystEx Ayumi Kikuchi )

' @(s)
' カラーチェック実績ＰＣ　操作員登録表示フォーム
' 　本モジュールはスタッフ名／検査員名／入力者名の登録表示フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

' @(f)
'
' 機能      : 氏名データ追加ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データ追加ボタン処理。
'
' 備考      :
'           :COLORSYS
'
Private Sub cmdAddStaff_Click()
    Dim bRet As Boolean
    Dim nI As Integer
    Dim bSearch As Boolean
    
    If Trim(imtxtStaffName.Text) = "" Then Exit Sub
    
    cmdAddStaff.Enabled = False
    
    bSearch = False
    
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''スラブ肌調査入力
            For nI = 1 To UBound(APStaffData)
                If Trim(imtxtStaffName.Text) = APStaffData(nI - 1).inp_StaffName Then
                    bSearch = True
                    Exit For
                End If
            Next nI
        
            If bSearch = False Then
                bRet = TRTS0060_Write(False, Trim(imtxtStaffName.Text))
                bRet = TRTS0060_Read()
            End If
        
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            For nI = 1 To UBound(APInspData)
                If Trim(imtxtStaffName.Text) = APInspData(nI - 1).inp_InspName Then
                    bSearch = True
                    Exit For
                End If
            Next nI

            If bSearch = False Then
                bRet = TRTS0062_Write(False, Trim(imtxtStaffName.Text))
                bRet = TRTS0062_Read()
            End If
        
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            For nI = 1 To UBound(APInspData)
                If Trim(imtxtStaffName.Text) = APInspData(nI - 1).inp_InspName Then
                    bSearch = True
                    Exit For
                End If
            Next nI
        
            If bSearch = False Then
                bRet = TRTS0062_Write(False, Trim(imtxtStaffName.Text))
                bRet = TRTS0062_Read()
            End If

        Case "frmDirResWnd" ''スラブ異常処置結果報告入力
            For nI = 1 To UBound(APInpData)
                If Trim(imtxtStaffName.Text) = APInpData(nI - 1).inp_InpName Then
                    bSearch = True
                    Exit For
                End If
            Next nI
        
            If bSearch = False Then
                bRet = TRTS0066_Write(False, Trim(imtxtStaffName.Text))
                bRet = TRTS0066_Read()
            End If

    End Select
    
    If bSearch = False Then
        Call InitForm
    End If
    
    cmdAddStaff.Enabled = True

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
' 機能      : 氏名データ削除ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データ削除ボタン処理。
'
' 備考      :
'
Private Sub cmdDelStaff_Click()
    Dim bRet As Boolean
    
    If lstStaff.ListIndex > -1 Then
        '呼出元により、処理分岐
        Select Case cCallBackObject.Name
            
            Case "frmSkinScanWnd" ''スラブ肌調査入力
                bRet = TRTS0060_Write(True, APStaffData(lstStaff.ListIndex).inp_StaffName)
                bRet = TRTS0060_Read()
            
            Case "frmColorScanWnd" ''カラーチェック検査表入力
                bRet = TRTS0062_Write(True, APInspData(lstStaff.ListIndex).inp_InspName)
                bRet = TRTS0062_Read()
            
            Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
                bRet = TRTS0062_Write(True, APInspData(lstStaff.ListIndex).inp_InspName)
                bRet = TRTS0062_Read()
    
            Case "frmDirResWnd" ''スラブ異常処置結果報告入力
                bRet = TRTS0066_Write(True, APInpData(lstStaff.ListIndex).inp_InpName)
                bRet = TRTS0066_Read()
    
        End Select
        
        Call InitForm
'        lstStaff.RemoveItem lstStaff.ListIndex
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
'    Dim nI As Integer
    
'    ReDim APSysCfgData.Group(0)
'    APSysCfgData.nGroupCount = lstGroup.ListCount
    
'    For nI = 1 To APSysCfgData.nGroupCount
'        APSysCfgData.Group(nI - 1) = lstGroup.List(nI - 1)
'        ReDim Preserve APSysCfgData.Group(UBound(APSysCfgData.Group) + 1)
'    Next nI
    
    'ReDim APSysCfgData.Operator(0)
    'APSysCfgData.nOperatorCount = lstOperator.ListCount
    '
    'For nI = 1 To APSysCfgData.nOperatorCount
    '    APSysCfgData.Operator(nI - 1) = lstOperator.List(nI - 1)
    '    ReDim Preserve APSysCfgData.Operator(UBound(APSysCfgData.Operator) + 1)
    'Next nI
    
    '/******************************/
    'ReDim APSysCfgData.StaffNumber(0)
    'ReDim APSysCfgData.StaffName(0)
    'APSysCfgData.nStaffCount = lstStaff.ListCount
    '
    'For nI = 1 To APSysCfgData.nStaffCount
    '    APSysCfgData.StaffNumber(nI - 1) = Left(lstStaff.List(nI - 1), 5)
    '    APSysCfgData.StaffName(nI - 1) = Mid(lstStaff.List(nI - 1), 7)
    '    ReDim Preserve APSysCfgData.StaffNumber(UBound(APSysCfgData.StaffNumber) + 1)
    '    ReDim Preserve APSysCfgData.StaffName(UBound(APSysCfgData.StaffName) + 1)
    'Next nI
    
    
    Unload Me
    
    cCallBackObject.CallBackMessage iCallBackID, CALLBACK_ncResOK
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
'    Dim nI As Integer
'    For nI = 1 To APSysCfgData.nGroupCount
'        If APSysCfgData.Group(nI - 1) <> "" Then
'            lstGroup.AddItem APSysCfgData.Group(nI - 1)
'        End If
'    Next nI
    
'    For nI = 1 To APSysCfgData.nOperatorCount
'        If APSysCfgData.Operator(nI - 1) <> "" Then
'            lstOperator.AddItem APSysCfgData.Operator(nI - 1)
'        End If
'    Next nI

    '/******************************/
    Debug.Print cCallBackObject.Name
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
    
    lstStaff.Clear
    
    '呼出元により、処理分岐
    Select Case cCallBackObject.Name
        
        Case "frmSkinScanWnd" ''スラブ肌調査入力
            Me.Caption = "スタッフ名登録"
            lblStaffName.Caption = "スタッフ名"
            For nI = 1 To UBound(APStaffData)
                lstStaff.AddItem APStaffData(nI - 1).inp_StaffName
            Next nI
        
        Case "frmColorScanWnd" ''カラーチェック検査表入力
            Me.Caption = "検査員名登録"
            lblStaffName.Caption = "検査員名"
            For nI = 1 To UBound(APInspData)
                lstStaff.AddItem APInspData(nI - 1).inp_InspName
            Next nI
        
        Case "frmSlbFailScanWnd" ''スラブ異常報告書入力
            Me.Caption = "検査員名登録"
            lblStaffName.Caption = "検査員名"
            For nI = 1 To UBound(APInspData)
                lstStaff.AddItem APInspData(nI - 1).inp_InspName
            Next nI

        Case "frmDirResWnd" ''スラブ異常処置結果報告入力
            Me.Caption = "入力者名登録"
            lblStaffName.Caption = "入力者名"
            For nI = 1 To UBound(APInpData)
                lstStaff.AddItem APInpData(nI - 1).inp_InpName
            Next nI

    End Select

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

' @(f)
'
' 機能      : 氏名データ入力BOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データ入力BOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub imtxtStaffName_GotFocus()
    imtxtStaffName.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : 氏名データ入力BOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データ入力BOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub imtxtStaffName_LostFocus()
    imtxtStaffName.BackColor = conDefine_ColorBKLostFocus
End Sub

' @(f)
'
' 機能      : 氏名データリストBOXフォーカス取得
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データリストBOXフォーカス取得時の処理を行う。
'
' 備考      :
'
Private Sub lstStaff_GotFocus()
'    lstStaff.BackColor = conDefine_ColorBKGotFocus
End Sub

' @(f)
'
' 機能      : 氏名データリストBOXフォーカス消滅
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 氏名データリストBOXフォーカス消滅時の処理を行う。
'
' 備考      :
'
Private Sub lstStaff_LostFocus()
'    lstStaff.BackColor = conDefine_ColorBKLostFocus
End Sub

