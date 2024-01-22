VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDirResWnd 
   BackColor       =   &H00C0FFFF&
   Caption         =   "処置内容指示確認／結果登録"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   17250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   17250
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdDirPrn 
      Caption         =   "指示印刷"
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
      Left            =   60
      TabIndex        =   51
      Top             =   120
      Width           =   1800
   End
   Begin VB.Frame Frame_Status 
      BackColor       =   &H00C0FFFF&
      Caption         =   "処置結果入力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "適用"
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
         Left            =   6240
         TabIndex        =   45
         Top             =   900
         Width           =   1800
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
         Index           =   3
         ItemData        =   "frmDirResWnd.frx":0000
         Left            =   2160
         List            =   "frmDirResWnd.frx":0002
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   44
         Top             =   1080
         Width           =   3255
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
         Index           =   2
         ItemData        =   "frmDirResWnd.frx":0004
         Left            =   2160
         List            =   "frmDirResWnd.frx":0006
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   41
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         BackStyle       =   0  '透明
         Caption         =   "処置後結果"
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
         Index           =   11
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         BackStyle       =   0  '透明
         Caption         =   "処置状態"
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
         Left            =   1860
         TabIndex        =   39
         Top             =   720
         Width           =   15165
      End
      Begin VB.Label lblDirCmt 
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
         Left            =   1860
         TabIndex        =   38
         Top             =   300
         Width           =   15165
      End
      Begin VB.Label Label6 
         Caption         =   "指示コメント"
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
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1695
      End
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
      ItemData        =   "frmDirResWnd.frx":0008
      Left            =   14580
      List            =   "frmDirResWnd.frx":000A
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   33
      Top             =   780
      Width           =   2595
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "入力者名："
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
         TabIndex        =   31
         Top             =   900
         Width           =   2805
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
         TabIndex        =   30
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
         Index           =   9
         Left            =   8700
         TabIndex        =   29
         Top             =   900
         Width           =   2805
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
         TabIndex        =   28
         Top             =   1440
         Width           =   825
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
         Index           =   10
         Left            =   8700
         TabIndex        =   27
         Top             =   1440
         Width           =   2805
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
         TabIndex        =   26
         Top             =   360
         Width           =   705
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
         Index           =   8
         Left            =   8700
         TabIndex        =   25
         Top             =   360
         Width           =   2805
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
         TabIndex        =   24
         Top             =   360
         Width           =   885
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
         Left            =   4980
         TabIndex        =   23
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "20080129"
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
         Left            =   1200
         TabIndex        =   22
         Top             =   360
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
         TabIndex        =   21
         Top             =   900
         Width           =   885
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
         TabIndex        =   20
         Top             =   1440
         Width           =   2805
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
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblSlb 
         Alignment       =   2  '中央揃え
         BorderStyle     =   1  '実線
         Caption         =   "N304AM"
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
         Index           =   4
         Left            =   1200
         TabIndex        =   18
         Top             =   900
         Width           =   2805
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
         TabIndex        =   17
         Top             =   900
         Width           =   705
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
         Index           =   6
         Left            =   6840
         TabIndex        =   16
         Top             =   900
         Width           =   945
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
         TabIndex        =   15
         Top             =   900
         Width           =   705
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
         Index           =   5
         Left            =   4980
         TabIndex        =   14
         Top             =   900
         Width           =   945
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
         TabIndex        =   13
         Top             =   900
         Width           =   435
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
         TabIndex        =   12
         Top             =   360
         Width           =   1035
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
         Index           =   7
         Left            =   4980
         TabIndex        =   11
         Top             =   1440
         Width           =   2805
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
      Left            =   13020
      TabIndex        =   2
      Top             =   10500
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "送信"
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   5340
      TabIndex        =   50
      Top             =   840
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
      Left            =   4500
      TabIndex        =   49
      Top             =   840
      Width           =   825
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
      Left            =   60
      TabIndex        =   48
      Top             =   10380
      Width           =   1275
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
      Left            =   1380
      TabIndex        =   47
      Top             =   10380
      Width           =   2565
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
      Left            =   4020
      TabIndex        =   46
      Top             =   10380
      Width           =   2565
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
      Index           =   12
      Left            =   1920
      TabIndex        =   35
      Top             =   840
      Width           =   2565
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
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "処置内容指示確認／結果登録"
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
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   17175
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
' カラーチェック検査表入力－スラブ選択表示フォーム
' 　本モジュールはカラーチェック検査表入力－スラブ選択表示フォームで使用する
' 　ためのものである。

Option Explicit

Private cCallBackObject As Object ''コールバックオブジェクト格納
Private iCallBackID As Integer ''コールバックＩＤ格納

Private nMSFlexGrid1_Selected_Row As Integer ''グリッド１選択行番号格納

Private bMouseControl As Boolean ''マウスコントロールフラグ格納

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
    Call SlbSelLock(False)
    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResCANCEL) '2008/09/04 戻り先変更
    Unload Me
End Sub

' @(f)
'
' 機能      : 指示印刷ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 指示印刷ボタン処理。
'
' 備考      :2008/09/04 指示印刷機能
'
Private Sub cmdDirPrn_Click()
    
    Call DirPrnReq

End Sub

' @(f)
'
' 機能      : 適用ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 適用ボタン処理。
'
' 備考      :COLORSYS
'
Private Sub cmdInput_Click()
    
    If nMSFlexGrid1_Selected_Row < 1 Then Exit Sub
    
    ''処置結果入力をリストへ適用
    APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = APDirRes_Stat(cmbRes(2).ListIndex).inp_DirRes_StatCode
    APDirResData(nMSFlexGrid1_Selected_Row - 1).res_aft_stat = APDirRes_Res(cmbRes(3).ListIndex).inp_DirRes_ResCode
    
    '完了の場合
    If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = "1" Then
        If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme <> "" Then
        Else
            '名前が空白の場合
            APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme = cmbRes(0).Text '入力者名リスト
        End If
    
        If APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte <> "" Then
        Else
            '日付が空白の場合
            APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte = Format(Now, "YYYYMMDD")
        End If
    
    Else
        '完了でない場合
        APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_nme = "" '入力者名リスト
        APDirResData(nMSFlexGrid1_Selected_Row - 1).res_wrt_dte = ""
    End If
    
   
    'リスト表示更新
    Call SetMSFlexGrid1

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

    If UBound(APDirResData) < 1 Then Exit Sub

    Call DBSendDataReq_DIRRES

'    Unload Me
'    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) '処理継続 '2008/09/04 戻り先変更

End Sub

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

    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) '2008/09/04 戻り先変更
    Unload Me

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
' 備考      :'2008/09/04 戻り先変更
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
    
    'ビジコン送信有り、ＤＢ登録の登録問い合わせOK
    Case CALLBACK_RES_HOSTSNDDATA_DIRRES
            If Result = CALLBACK_ncResOK Then          'OK

'                ''DB保存準備
'                Call SetAPResData(True)
'
'                'カレント実績入力情報一時保存
'                APResDataBK = APResData

                '無条件で＊：完了を送信
                APResData.slb_fault_u_judg = "9"
                APResData.slb_fault_d_judg = "9"

                'ビジコン送信
                frmHostSend.SetCallBack Me, CALLBACK_HOSTSEND
                frmHostSend.Show vbModal, Me 'ビジコン送信中は、他の処理を不可とする為、vbModalとする。

            Else
                'DB登録キャンセル
            End If
    
    'ビジコン送信無し、ＤＢ登録の登録問い合わせOK
    Case CALLBACK_RES_DBSNDDATA_DIRRES
            If Result = CALLBACK_ncResOK Then          'OK

'                ''DB保存準備
'                Call SetAPResData(True)
'
'                'カレント実績入力情報一時保存
'                APResDataBK = APResData

'                'ビジコン送信
'                frmHostSend.SetCallBack Me, CALLBACK_HOSTSEND
'                frmHostSend.Show vbModal, Me 'ビジコン送信中は、他の処理を不可とする為、vbModalとする。
'                '/* DB登録実行 */
                bRet = DB_SAVE_DIRRES(False)
                
'                bRet = TRTS0022_Write(False)

                If bRet Then
                    '正常終了
                    Unload Me
                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OKで処理終了 '2008/09/04 戻り先変更
                End If

            Else
                'DB登録キャンセル
            End If
    
    '処置結果登録のビジコン通信よりOK
    Case CALLBACK_HOSTSEND
            If Result = CALLBACK_ncResOK Then          'OK
                '正常終了

'                APResData.fail_host_send = "1" '1:正常をセット

'                '/* DB登録実行 */
                bRet = DB_SAVE_DIRRES(False)
'                bRet = TRTS0022_Write(False)

                Call dpDebug

                If bRet Then
                    '正常終了
                    Unload Me
                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OKで処理終了 '2008/09/04 戻り先変更
                End If
            ElseIf Result = CALLBACK_ncResSKIP Then          'SKIP
                '/* DB登録実行 */
                bRet = DB_SAVE_DIRRES(True)  'ビジコンエラー有り
                '処理継続
                'ビジコン通信必須のため、ＤＢ保存は行わない。

'                'ビジコン送信スキップ処理（処理前に戻す。）
'                APResData.host_send = APResDataBK.host_send
'                APResData.host_wrt_dte = APResDataBK.host_wrt_dte
'                APResData.host_wrt_tme = APResDataBK.host_wrt_tme
'

                Call dpDebug

'                '/* DB登録実行 */
'                bRet = DB_SAVE_COLOR()
'
'                If bRet Then
'                    '正常終了
'                    Unload Me
'                    Call fMainWnd.CallBackMessage(iCallBackID, CALLBACK_ncResOK) 'OKで処理終了 '2008/09/04 戻り先変更
'                End If
            
            
            Else
                '/* DB登録実行 */
'                bRet = DB_SAVE_DIRRES(True) '中断
                'キャンセル（エラー発生にて、OKボタンを押した場合、呼出元画面に戻る。）
                'ビジコン通信必須のため、ＤＢ保存は行わない。

'                'ビジコン送信スキップ処理（処理前に戻す。）
'                APResData.host_send = "0" '0:異常をセット
'                APResData.host_wrt_dte = APResDataBK.host_wrt_dte
'                APResData.host_wrt_tme = APResDataBK.host_wrt_tme
                
                Call WaitMsgBox(Me, "送信／ＤＢ保存処理を中断しました。")

                Call dpDebug

            End If

        '指示印刷問合せ 2008/09/04
        Case CALLBACK_RES_DIRPRN_REQ
            If Result = CALLBACK_ncResOK Then          'OK
                frmTRSend.SetCallBack Me, CALLBACK_RES_DIRPRN_SND, "COL02"
                frmTRSend.Show vbModal, Me
            Else
            End If

        '指示印刷要求送信結果 2008/09/04
        Case CALLBACK_RES_DIRPRN_SND
            If Result = CALLBACK_ncResOK Then          'OK
                Call WaitMsgBox(Me, "指示印刷要求は正常終了しました。")
            Else
                Call WaitMsgBox(Me, "指示印刷要求は失敗しました。")
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
    MSFlexGrid1.TextMatrix(0, nCol) = "処置状態"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 10000
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "指示内容"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "処置後結果"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "入力日"
    
    nCol = nCol + 1
    MSFlexGrid1.ColWidth(nCol) = 1500
    'MSFlexGrid1.ColAlignment(nCol) = FlexAlignCenter
    MSFlexGrid1.Row = nRow
    MSFlexGrid1.Col = nCol
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSFlexGrid1.TextMatrix(0, nCol) = "入力者"
    
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
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APDirResData(nRow - 1).res_cmp_flg = "1", "完了", "")  '"処置状態"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignLeftCenter
        '指示内容
        strDir = APDirResData(nRow - 1).dir_nme1 & " " & APDirResData(nRow - 1).dir_val1 & " " & APDirResData(nRow - 1).dir_uni1 & " " & _
        APDirResData(nRow - 1).dir_nme2 & " " & APDirResData(nRow - 1).dir_val2 & " " & APDirResData(nRow - 1).dir_uni2
        MSFlexGrid1.TextMatrix(nRow, nCol) = strDir
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = IIf(APDirResData(nRow - 1).res_aft_stat = "1", "不適合有り", "") '"処置後結果"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APDirResData(nRow - 1).res_wrt_dte '"入力日"
        
        nCol = nCol + 1
        MSFlexGrid1.Row = nRow
        MSFlexGrid1.Col = nCol
        MSFlexGrid1.CellAlignment = flexAlignCenterCenter
        MSFlexGrid1.TextMatrix(nRow, nCol) = APDirResData(nRow - 1).res_wrt_nme '"入力者"
        
    
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
' 機能      : 入力者名登録ボタン
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 入力者名登録ボタン処理。
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

    bMouseControl = True

    '現在のRowを一時保存
    nNowRow = MSFlexGrid1.Row

    '以前のセレクト行を未セレクト状態に戻す。
    If nMSFlexGrid1_Selected_Row <> 0 Then
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000008
        MSFlexGrid1.CellBackColor = &H80000005
        Next nJ
    Else
        'タイトル行の色を付け直す。
        For nJ = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = nMSFlexGrid1_Selected_Row
        MSFlexGrid1.Col = nJ
        MSFlexGrid1.CellForeColor = &H80000012
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
'                If APSlbCont.nSearchInputModeSelectedIndex = 2 Then
'                    '削除モードの場合
'                    MSFlexGrid1.CellForeColor = &H8000000E
'                    MSFlexGrid1.CellBackColor = &H8080FF
'                Else
                    MSFlexGrid1.CellForeColor = &H8000000E
                    MSFlexGrid1.CellBackColor = &H8000000D
'                End If

                ''処置結果入力エリアを更新
                cmbRes(2).ListIndex = IIf(APDirResData(nMSFlexGrid1_Selected_Row - 1).res_cmp_flg = "1", 1, 0)
                cmbRes(3).ListIndex = IIf(APDirResData(nMSFlexGrid1_Selected_Row - 1).res_aft_stat = "1", 1, 0)

            Else
                '選択ロック中の色
                MSFlexGrid1.CellForeColor = &H8000000E
                MSFlexGrid1.CellBackColor = &H808080
            End If
        Next nJ
        If MSFlexGrid1.Enabled Then
            '選択中
        Else
            '選択ロック
        End If
    
    Else
    End If

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
    
    Call GetCurrentAPSlbData
    
    Call InitMSFlexGrid1

'    If APSlbCont.bProcessing Then 'スラブ選択ロック用処理中フラグ
'        imTextSearchSlbNumber.Text = APSlbCont.strSearchInputSlbNumber  '検索スラブＮｏ．
        
'        OptInputMode(APSlbCont.nSearchInputModeSelectedIndex).Value = True '入力モード指定インデックス番号
'        OptStatus(APSlbCont.nSearchInputStatusSelectedIndex).Value = True '状態選択指定インデックス番号
        
        'スラブ選択情報
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

'    lblInputMode.Caption = IIf(APSlbCont.nSearchInputModeSelectedIndex = 0, "【新規】", "【修正】")

    If APResData.fail_host_wrt_dte <> "" Then
        APResData.host_send_flg = "1" '訂正
    Else
        APResData.host_send_flg = "0" '新規
    End If

    lblHostSendFlg.Caption = IIf(APResData.host_send_flg = "0", "0:新規", "1:訂正")

'    '異常報告が存在する場合は、「送信」ボタンを無効にする。
'    If APResData.fail_sys_wrt_dte <> "" Then
'        cmdOK.Enabled = False
'    Else
'        cmdOK.Enabled = True
'    End If

    'カレントスラブ情報ロード

    Call dpDebug

    lblSlb(0).Caption = APResData.slb_chno & "-" & APResData.slb_aino ''スラブNo
    lblSlb(1).Caption = ConvDpOutStat(conDefine_SYSMODE_COLOR, CInt(APResData.slb_stat)) ''状態
    lblSlb(2).Caption = APResData.slb_ccno ''CCNo
    lblSlb(3).Caption = APResData.slb_zkai_dte ''造塊日
    lblSlb(4).Caption = APResData.slb_ksh ''鋼種
    lblSlb(5).Caption = APResData.slb_typ ''型
    lblSlb(6).Caption = APResData.slb_uksk ''向先
    lblSlb(7).Caption = APResData.slb_wei ''重量
    lblSlb(8).Caption = APResData.slb_thkns ''厚み
    lblSlb(9).Caption = APResData.slb_wdth ''幅
    lblSlb(10).Caption = APResData.slb_lngth ''長さ
'    lblSlb(11).Caption = APResData.sys_wrt_dte ''記録日

    lblSlb(12).Caption = Format(CInt(APResData.slb_col_cnt), "00") ''カラー回数

    '入力者名リストBOX設定
    cmbRes(0).Clear
    For nJ = 1 To UBound(APInpData)
        cmbRes(0).AddItem APInpData(nJ - 1).inp_InpName
'        If APDirResData.slb_wrt_nme = APInpData(nJ - 1).inp_InpName Then
'            cmbRes(0).ListIndex = nJ - 1
'        End If
    Next nJ

'    '次工程リストBOX設定
'    cmbRes(1).Clear
'    For nJ = 1 To UBound(APNextProcDataColor)
'        cmbRes(1).AddItem APNextProcDataColor(nJ - 1).inp_NextProc
'        If APResData.slb_nxt_prcs = APNextProcDataColor(nJ - 1).inp_NextProc Then
'            cmbRes(1).ListIndex = nJ - 1
'        End If
'    Next nJ

'    'コメント情報ロード
    lblDirCmt(0).Caption = APDirResData(0).dir_cmt1
    lblDirCmt(1).Caption = APDirResData(0).dir_cmt2

    ''処置状態リストBOX設定
    cmbRes(2).Clear
    For nJ = 1 To UBound(APDirRes_Stat)
        cmbRes(2).AddItem APDirRes_Stat(nJ - 1).inp_DirRes_Stat
    Next nJ
    
    ''処置結果リストBOX設定
    cmbRes(3).Clear
    For nJ = 1 To UBound(APDirRes_Res)
        cmbRes(3).AddItem APDirRes_Res(nJ - 1).inp_DirRes_Res
    Next nJ
    
End Sub

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
Private Sub DBSendDataReq_DIRRES()
    Dim nI As Integer
    Dim bAllCmp As Boolean
    Dim fmessage As Object
    Set fmessage = New MessageYN

    '全て完了かのチェック
    bAllCmp = True '完了
    For nI = 0 To UBound(APDirResData) - 1
        If APDirResData(nI).res_cmp_flg <> "1" Then
            bAllCmp = False '未完了有り
            Exit For
        End If
    Next nI
    
    If bAllCmp Then
        '完了の場合
        fmessage.MsgText = "全て完了となりますので、ビジコンへ完了を送信後、ＤＢへ登録します。" & vbCrLf & "よろしいですか？"
    '    fmessage.AutoDelete = True
        fmessage.AutoDelete = False
        fmessage.SetCallBack Me, CALLBACK_RES_HOSTSNDDATA_DIRRES, False
    Else
        '未完了有り
        fmessage.MsgText = "入力結果をＤＢへ登録します。" & vbCrLf & "よろしいですか？"
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
    fmessage.Show vbModal, Me '他の処理を不可とする為、vbModalとする。
    Set fmessage = Nothing
'    End If

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
Private Function DB_SAVE_DIRRES(ByVal bHostSendError As Boolean) As Boolean
    Dim bNOErrorFlg As Boolean
    Dim bRet As Boolean
    Dim MsgWnd As Message
    Set MsgWnd = New Message

    MsgWnd.MsgText = "データベースサーバーに保存中です。" & vbCrLf & "しばらくお待ちください。"
    MsgWnd.OK.Visible = False
    MsgWnd.Show vbModeless, Me
    MsgWnd.Refresh
    
    bNOErrorFlg = True 'エラー無し

    'ビジコン通信エラー発生時
    If bHostSendError Then
        
'        MsgWnd.OK_Close
'        MsgWnd.MsgText = "ビジコン通信が正常処理出来ない為、" & vbCrLf & "ＤＢ保存を中断しました。"
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
'        bNOErrorFlg = False 'エラー有り
'        DB_SAVE_DIRRES = bNOErrorFlg
'        Exit Function
        
        APResData.fail_res_host_send = "0"       ''/* ビジコン送信結果 */
        APResData.fail_res_host_wrt_dte = APResData.fail_host_wrt_dte    ''/* ビジコン登録日 */
        APResData.fail_res_host_wrt_tme = APResData.fail_host_wrt_tme     ''/* ビジコン登録時刻 */
    
    Else
        APResData.fail_res_host_send = "1"       ''/* ビジコン送信結果 */
        APResData.fail_res_host_wrt_dte = APResData.fail_host_wrt_dte    ''/* ビジコン登録日 */
        APResData.fail_res_host_wrt_tme = APResData.fail_host_wrt_tme     ''/* ビジコン登録時刻 */
    End If

    'TRTS0022 登録
    bRet = TRTS0022_Write(False)
    
    If bRet = False Then
        bNOErrorFlg = False 'エラー有り
        MsgWnd.OK_Close
        MsgWnd.MsgText = "ＤＢ保存エラーが発生しました。" & vbCrLf & "処理を中断しました。"
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
        MsgWnd.MsgText = "ＤＢ保存が正常終了しました。"
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
' 機能      : 指示印刷問い合わせ処理
'
' 引き数    :
'
' 返り値    :
'
' 機能説明  : 指示印刷問い合わせ画面を開く。
'
' 備考      : コールバック有り。2008/09/04
'
Private Sub DirPrnReq()
    Dim fmessage As Object
    Set fmessage = New MessageYN

    fmessage.MsgText = "指示帳票の印刷を行います。" & vbCrLf & "よろしいですか？"
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
    fmessage.Show vbModal, Me '他の処理を不可とする為、vbModalとする。
    Set fmessage = Nothing
'    End If

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

