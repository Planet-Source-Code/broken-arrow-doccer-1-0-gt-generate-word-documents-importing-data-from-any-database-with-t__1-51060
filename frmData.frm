VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmData 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Select data"
   ClientHeight    =   5190
   ClientLeft      =   420
   ClientTop       =   450
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
   Icon            =   "frmData.frx":0000
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
      TabIndex        =   9
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmData.frx":030A
      PictureMiddle   =   "frmData.frx":0D74
      PictureRight    =   "frmData.frx":0E12
      PictureRightWidth=   84
      FormBorderTop   =   "frmData.frx":0EB0
      FormBorderLeft  =   "frmData.frx":0F12
      FormBorderBottom=   "frmData.frx":0F70
      FormBorderRight =   "frmData.frx":0FD2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmData.frx":1030
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
      PictureMaximize =   "frmData.frx":134A
      PictureMinimize =   "frmData.frx":16DC
      PictureClose    =   "frmData.frx":1A6E
      PictureMinimizeToTray=   "frmData.frx":1E00
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmData.frx":2192
      PictureCloseDown=   "frmData.frx":2524
      PictureMaximizeDown=   "frmData.frx":28B6
      PictureMinimizeDown=   "frmData.frx":2C48
      PictureShrinkDown=   "frmData.frx":2FDA
      PictureMinimizeToTrayDown=   "frmData.frx":336C
      ControlMenu     =   0   'False
      PicturePin      =   "frmData.frx":36FE
      PicturePinDown  =   "frmData.frx":3A90
      PicturePinHover =   "frmData.frx":3E22
      PictureMinimizeToTrayHover=   "frmData.frx":4174
      PictureShrinkHover=   "frmData.frx":44C6
      PictureMinimizeHover=   "frmData.frx":4818
      PictureMaximizeHover=   "frmData.frx":4B6A
      PictureCloseHover=   "frmData.frx":4EBC
      TrayTip         =   " docCER> Select data "
      FormMouseIcon   =   "frmData.frx":520E
      TrayIcon        =   "frmData.frx":5A28
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   2535
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txtCondition 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Select data"
      TitleDescription=   $"frmData.frx":5D42
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
      Picture         =   "frmData.frx":5DED
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
      MICON           =   "frmData.frx":793F
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
      MICON           =   "frmData.frx":795B
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
      MICON           =   "frmData.frx":7977
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjdocCER.chameleonButton cmdReload 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Reload"
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
      MICON           =   "frmData.frx":7993
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1740
      Width           =   405
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Condition"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   810
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmQuery.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdNext_Click()
    AppSetting.Condition = txtCondition.Text

    Me.Hide
    frmWordTemplate.Show
End Sub

Private Sub cmdReload_Click()
    On Error GoTo ERROR_HANDLER_cmdReload_Click
    
    With Query(GetSQL)
        lvwData.ListItems.Clear
        lvwData.ColumnHeaders.Clear
        
        Dim ColumnCounter As Long, lItem As ListItem
        For ColumnCounter = 0 To .Fields.Count - 1
            lvwData.ColumnHeaders.Add , , .Fields(ColumnCounter).Name
        Next
        
        While Not .EOF
            For ColumnCounter = 0 To .Fields.Count - 1
                If ColumnCounter = 0 Then
                    Set lItem = lvwData.ListItems.Add(, , .Fields(0))
                Else
                    If .Fields(ColumnCounter) <> "" Then lItem.SubItems(ColumnCounter) = .Fields(ColumnCounter)
                End If
            Next
            
            If Not .EOF Then .MoveNext
            
            DoEvents
        Wend
    End With

EXIT_cmdReload_Click:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_cmdReload_Click:
    Select Case Err.Number
    Case Else
        If MsgBox("Error in Sub cmdReload_Click() of Form frmData (frmData.frm) of Project prjdocCER (prjdocCER.vbp)" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "Application error!") = vbNo Then Resume EXIT_cmdReload_Click
    End Select
    
    Resume Next
End Sub

Private Sub Form_Load()
    If ProfileLoaded Then txtCondition.Text = AppSetting.Condition
    
    cmdReload_Click
End Sub

Private Function GetSQL() As String
GetSQL = AppSetting.SQL
If UCase(Left(Trim(AppSetting.SQL), 7)) = "SELECT " Then
    If Trim(txtCondition.Text) <> "" Then GetSQL = AppSetting.SQL & " WHERE " & txtCondition.Text
Else
    If Trim(txtCondition.Text) <> "" Then GetSQL = "SELECT * FROM (" & AppSetting.SQL & ") WHERE " & txtCondition.Text
End If
End Function
