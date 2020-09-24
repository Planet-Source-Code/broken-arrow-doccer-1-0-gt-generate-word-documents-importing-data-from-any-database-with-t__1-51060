VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmDone 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   " Template"
   ClientHeight    =   5190
   ClientLeft      =   5160
   ClientTop       =   2970
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
   Icon            =   "frmDone.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   2  'CenterScreen
   Begin prjdocCER.ctlHyperLink lnkRateMe 
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   4747
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   423
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Rate me..."
      HyperLink       =   "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=51060&lngWId=1"
   End
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmDone.frx":0ECA
      PictureMiddle   =   "frmDone.frx":1934
      PictureRight    =   "frmDone.frx":19D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmDone.frx":1A70
      FormBorderLeft  =   "frmDone.frx":1AD2
      FormBorderBottom=   "frmDone.frx":1B30
      FormBorderRight =   "frmDone.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmDone.frx":1BF0
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
      PictureMaximize =   "frmDone.frx":2ACA
      PictureMinimize =   "frmDone.frx":2E5C
      PictureClose    =   "frmDone.frx":31EE
      PictureMinimizeToTray=   "frmDone.frx":3580
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmDone.frx":3912
      PictureCloseDown=   "frmDone.frx":3CA4
      PictureMaximizeDown=   "frmDone.frx":4036
      PictureMinimizeDown=   "frmDone.frx":43C8
      PictureShrinkDown=   "frmDone.frx":475A
      PictureMinimizeToTrayDown=   "frmDone.frx":4AEC
      ControlMenu     =   0   'False
      PicturePin      =   "frmDone.frx":4E7E
      PicturePinDown  =   "frmDone.frx":5210
      PicturePinHover =   "frmDone.frx":55A2
      PictureMinimizeToTrayHover=   "frmDone.frx":58F4
      PictureShrinkHover=   "frmDone.frx":5C46
      PictureMinimizeHover=   "frmDone.frx":5F98
      PictureMaximizeHover=   "frmDone.frx":62EA
      PictureCloseHover=   "frmDone.frx":663C
      TrayTip         =   " docCER> Template "
      FormMouseIcon   =   "frmDone.frx":698E
      TrayIcon        =   "frmDone.frx":71A8
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Complete!"
      TitleDescription=   $"frmDone.frx":8082
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
      Picture         =   "frmDone.frx":811E
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
      MICON           =   "frmDone.frx":9C70
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
      TX              =   "&Finish"
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
      MICON           =   "frmDone.frx":9C8C
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
      MICON           =   "frmDone.frx":9CA8
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
      Caption         =   "Save to"
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
   Begin prjdocCER.chameleonButton cmdProfileSave 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Save"
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
      MICON           =   "frmDone.frx":9CC4
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
Attribute VB_Name = "frmDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmWordTemplate.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdNext_Click()
    GenerateDOC
    Quit
End Sub

Private Sub cmdProfileSave_Click()
    If Trim(txtProfileFile.File) <> "" Then
        SettingSave txtProfileFile.File
        MsgBox "The profile has been saved.", vbOKOnly + vbInformation, "docCER"
    End If
End Sub

Private Sub Form_Load()
If ProfileLoaded Then txtProfileFile.File = frmWelcome.txtProfileFile.File
End Sub
