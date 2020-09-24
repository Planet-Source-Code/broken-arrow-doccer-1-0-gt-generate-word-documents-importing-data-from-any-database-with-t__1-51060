VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmDatabase 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Database"
   ClientHeight    =   5190
   ClientLeft      =   6510
   ClientTop       =   4695
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
   Icon            =   "frmDatabase.frx":0000
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
      TabIndex        =   18
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmDatabase.frx":08CA
      PictureMiddle   =   "frmDatabase.frx":1334
      PictureRight    =   "frmDatabase.frx":13D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmDatabase.frx":1470
      FormBorderLeft  =   "frmDatabase.frx":14D2
      FormBorderBottom=   "frmDatabase.frx":1530
      FormBorderRight =   "frmDatabase.frx":1592
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmDatabase.frx":15F0
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
      PictureMaximize =   "frmDatabase.frx":1ECA
      PictureMinimize =   "frmDatabase.frx":225C
      PictureClose    =   "frmDatabase.frx":25EE
      PictureMinimizeToTray=   "frmDatabase.frx":2980
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmDatabase.frx":2D12
      PictureCloseDown=   "frmDatabase.frx":30A4
      PictureMaximizeDown=   "frmDatabase.frx":3436
      PictureMinimizeDown=   "frmDatabase.frx":37C8
      PictureShrinkDown=   "frmDatabase.frx":3B5A
      PictureMinimizeToTrayDown=   "frmDatabase.frx":3EEC
      ControlMenu     =   0   'False
      PicturePin      =   "frmDatabase.frx":427E
      PicturePinDown  =   "frmDatabase.frx":4610
      PicturePinHover =   "frmDatabase.frx":49A2
      PictureMinimizeToTrayHover=   "frmDatabase.frx":4CF4
      PictureShrinkHover=   "frmDatabase.frx":5046
      PictureMinimizeHover=   "frmDatabase.frx":5398
      PictureMaximizeHover=   "frmDatabase.frx":56EA
      PictureCloseHover=   "frmDatabase.frx":5A3C
      TrayTip         =   " docCER> Database "
      FormMouseIcon   =   "frmDatabase.frx":5D8E
      TrayIcon        =   "frmDatabase.frx":65A8
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   17
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Database"
      TitleDescription=   "Please select the database using the options below to import data from to populate the Microsoft Word document with."
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
      Picture         =   "frmDatabase.frx":6E82
      BackColor       =   16744576
      PictureBorder   =   1
   End
   Begin prjdocCER.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   16
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
      MICON           =   "frmDatabase.frx":89D4
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
      TabIndex        =   15
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
      MICON           =   "frmDatabase.frx":89F0
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
      TabIndex        =   14
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
      MICON           =   "frmDatabase.frx":8A0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.ctlFileBrowser txtAccessDatabase 
      Height          =   360
      Left            =   480
      TabIndex        =   13
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   635
      Caption         =   "Browse"
      ButtonCaption   =   "..."
      Filter          =   "Microsoft Access database (*.mdb)|*.mdb|All files (*.*)|*.*"
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
   Begin VB.OptionButton optDatabaseType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Microsoft Access"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton optDatabaseType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Microsoft SQL Server"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OptionButton optDatabaseType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Other (ODBC)"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtMSSQLServerAddress 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtMSSQLServerDatabase 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtMSSQLServerUserName 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtMSSQLServerPassword 
      Appearance      =   0  'Flat
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtODBCConnectionString 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Connection string"
      Height          =   240
      Left            =   480
      TabIndex        =   12
      Top             =   4260
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Server"
      Height          =   240
      Left            =   480
      TabIndex        =   11
      Top             =   2460
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Username"
      Height          =   240
      Left            =   480
      TabIndex        =   10
      Top             =   3180
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Password"
      Height          =   240
      Left            =   480
      TabIndex        =   9
      Top             =   3540
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Database"
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   2820
      Width           =   825
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmWelcome.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ERROR_HANDLER_cmdNext_Click
    
    AppSetting.Database.MSAccessDatabase = txtAccessDatabase.File
    
    AppSetting.Database.MSSQLServerAddress = txtMSSQLServerAddress.Text
    AppSetting.Database.MSSQLServerUserName = txtMSSQLServerUserName.Text
    AppSetting.Database.MSSQLServerPassword = txtMSSQLServerPassword.Text
    AppSetting.Database.MSSQLServerDatabase = txtMSSQLServerDatabase.Text
    
    AppSetting.Database.ODBCConnectionString = txtODBCConnectionString.Text
    
    SetDatabaseConnection
    
    Me.Hide
    frmQuery.Show

EXIT_cmdNext_Click:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_cmdNext_Click:
    Select Case Err.Number
    Case Else
        If MsgBox("Error in Sub cmdNext_Click() of Form frmDatabase (frmDatabase.frm) of Project prjdocCER (prjdocCER.vbp)" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "Application error!") = vbNo Then Resume EXIT_cmdNext_Click
    End Select
    
    Resume Next
End Sub

Private Sub Form_Load()
    If ProfileLoaded Then
        txtAccessDatabase.File = AppSetting.Database.MSAccessDatabase
        
        txtMSSQLServerAddress.Text = AppSetting.Database.MSSQLServerAddress
        
        txtMSSQLServerUserName.Text = AppSetting.Database.MSSQLServerUserName
        txtMSSQLServerPassword.Text = AppSetting.Database.MSSQLServerPassword
        
        txtMSSQLServerDatabase.Text = AppSetting.Database.MSAccessDatabase
        
        txtODBCConnectionString.Text = AppSetting.Database.ODBCConnectionString
        
        optDatabaseType_Click (AppSetting.DatabaseType)
    Else
        optDatabaseType_Click (0)
    End If
End Sub

Private Sub optDatabaseType_Click(Index As Integer)
    txtAccessDatabase.Enabled = optDatabaseType(0).Value
    
    txtMSSQLServerAddress.Enabled = optDatabaseType(1).Value
    txtMSSQLServerDatabase.Enabled = optDatabaseType(1).Value
    txtMSSQLServerUserName.Enabled = optDatabaseType(1).Value
    txtMSSQLServerPassword.Enabled = optDatabaseType(1).Value
    
    txtODBCConnectionString.Enabled = optDatabaseType(2).Value
    
    AppSetting.DatabaseType = Index
End Sub

