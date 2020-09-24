VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmWelcome 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   5190
   ClientLeft      =   7455
   ClientTop       =   1785
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   2  'CenterScreen
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmWelcome.frx":0ECA
      PictureMiddle   =   "frmWelcome.frx":1934
      PictureRight    =   "frmWelcome.frx":19D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmWelcome.frx":1A70
      FormBorderLeft  =   "frmWelcome.frx":1AD2
      FormBorderBottom=   "frmWelcome.frx":1B30
      FormBorderRight =   "frmWelcome.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmWelcome.frx":1BF0
      AllowClose      =   0   'False
      CaptionSpacing  =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   8388608
      CaptionScrollToLeft=   0   'False
      CaptionScrollSpeed=   250
      PictureMaximize =   "frmWelcome.frx":2ACA
      PictureMinimize =   "frmWelcome.frx":2E5C
      PictureClose    =   "frmWelcome.frx":31EE
      PictureMinimizeToTray=   "frmWelcome.frx":3580
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmWelcome.frx":3912
      PictureCloseDown=   "frmWelcome.frx":3CA4
      PictureMaximizeDown=   "frmWelcome.frx":4036
      PictureMinimizeDown=   "frmWelcome.frx":43C8
      PictureShrinkDown=   "frmWelcome.frx":475A
      PictureMinimizeToTrayDown=   "frmWelcome.frx":4AEC
      ControlMenu     =   0   'False
      PicturePin      =   "frmWelcome.frx":4E7E
      PicturePinDown  =   "frmWelcome.frx":5210
      PicturePinHover =   "frmWelcome.frx":55A2
      PictureMinimizeToTrayHover=   "frmWelcome.frx":58F4
      PictureShrinkHover=   "frmWelcome.frx":5C46
      PictureMinimizeHover=   "frmWelcome.frx":5F98
      PictureMaximizeHover=   "frmWelcome.frx":62EA
      PictureCloseHover=   "frmWelcome.frx":663C
      TrayTip         =   " docCER> Welcome "
      FormMouseIcon   =   "frmWelcome.frx":698E
      TrayIcon        =   "frmWelcome.frx":71A8
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Welcome"
      TitleDescription=   $"frmWelcome.frx":8082
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   16777215
      BeginProperty TitleDescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleDescriptionForeColor=   0
      Picture         =   "frmWelcome.frx":812F
      BackColor       =   16744576
      PictureBorder   =   1
   End
   Begin prjdocCER.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16761024
      FCOL            =   8388608
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmWelcome.frx":9C81
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.chameleonButton cmdNext 
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Next >>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16761024
      FCOL            =   8388608
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmWelcome.frx":9C9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.chameleonButton cmdBack 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<< &Back"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16761024
      FCOL            =   8388608
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmWelcome.frx":9CB9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.ctlFileBrowser txtProfileFile 
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   635
      Caption         =   "Load profile"
      ButtonCaption   =   "..."
      Filter          =   "docCER profile (*.dp)|*.dp|All files (*.*)|*.*"
      Locked          =   -1  'True
      BackColor       =   16761024
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FileFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjdocCER.chameleonButton cmdClear 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16761024
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmWelcome.frx":9CD5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdClear_Click()
    txtProfileFile.File = ""
    ProfileLoaded = False
End Sub

Private Sub cmdNext_Click()
    If Trim(txtProfileFile.File) <> "" Then SettingLoad txtProfileFile.File
    
    frmDatabase.Show
    Me.Hide
End Sub

