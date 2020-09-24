VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmWordTemplate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Template"
   ClientHeight    =   5190
   ClientLeft      =   6390
   ClientTop       =   3585
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
   Icon            =   "frmWordTemplate.frx":0000
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
      TabIndex        =   7
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmWordTemplate.frx":0ECA
      PictureMiddle   =   "frmWordTemplate.frx":1934
      PictureRight    =   "frmWordTemplate.frx":19D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmWordTemplate.frx":1A70
      FormBorderLeft  =   "frmWordTemplate.frx":1AD2
      FormBorderBottom=   "frmWordTemplate.frx":1B30
      FormBorderRight =   "frmWordTemplate.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmWordTemplate.frx":1BF0
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
      PictureMaximize =   "frmWordTemplate.frx":2ACA
      PictureMinimize =   "frmWordTemplate.frx":2E5C
      PictureClose    =   "frmWordTemplate.frx":31EE
      PictureMinimizeToTray=   "frmWordTemplate.frx":3580
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmWordTemplate.frx":3912
      PictureCloseDown=   "frmWordTemplate.frx":3CA4
      PictureMaximizeDown=   "frmWordTemplate.frx":4036
      PictureMinimizeDown=   "frmWordTemplate.frx":43C8
      PictureShrinkDown=   "frmWordTemplate.frx":475A
      PictureMinimizeToTrayDown=   "frmWordTemplate.frx":4AEC
      ControlMenu     =   0   'False
      PicturePin      =   "frmWordTemplate.frx":4E7E
      PicturePinDown  =   "frmWordTemplate.frx":5210
      PicturePinHover =   "frmWordTemplate.frx":55A2
      PictureMinimizeToTrayHover=   "frmWordTemplate.frx":58F4
      PictureShrinkHover=   "frmWordTemplate.frx":5C46
      PictureMinimizeHover=   "frmWordTemplate.frx":5F98
      PictureMaximizeHover=   "frmWordTemplate.frx":62EA
      PictureCloseHover=   "frmWordTemplate.frx":663C
      TrayTip         =   " docCER> Template "
      FormMouseIcon   =   "frmWordTemplate.frx":698E
      TrayIcon        =   "frmWordTemplate.frx":71A8
   End
   Begin VB.TextBox txtFieldEncloser 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Text            =   "%"
      Top             =   1920
      Width           =   735
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Template"
      TitleDescription=   $"frmWordTemplate.frx":8082
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
      Picture         =   "frmWordTemplate.frx":811C
      BackColor       =   16744576
      PictureBorder   =   1
   End
   Begin prjdocCER.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   3
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
      MICON           =   "frmWordTemplate.frx":9C6E
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
      TabIndex        =   2
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
      MICON           =   "frmWordTemplate.frx":9C8A
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
      TabIndex        =   1
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
      MICON           =   "frmWordTemplate.frx":9CA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.ctlFileBrowser txtWordTemplate 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   635
      Caption         =   "Template"
      ButtonCaption   =   "..."
      Filter          =   "Microsoft Word document (*.doc)|*.doc|All files (*.*)|*.*"
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
   Begin VB.Label lblFieldEncloser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field encloser"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1980
      Width           =   1215
   End
End
Attribute VB_Name = "frmWordTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmData.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdNext_Click()
    AppSetting.WordTemplate = txtWordTemplate.File
    AppSetting.FieldEncloser = txtFieldEncloser.Text
    
    Me.Hide
    frmDone.Show
End Sub

Private Sub Form_Load()
If ProfileLoaded Then
    txtWordTemplate.File = AppSetting.WordTemplate
    txtFieldEncloser.Text = AppSetting.FieldEncloser
End If
End Sub
