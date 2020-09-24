VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "wowformer.ocx"
Begin VB.Form frmQuery 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Record selection"
   ClientHeight    =   5190
   ClientLeft      =   5550
   ClientTop       =   3960
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
   Icon            =   "frmQuery.frx":0000
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
      TabIndex        =   11
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   503
      PictureLeft     =   "frmQuery.frx":0ECA
      PictureMiddle   =   "frmQuery.frx":1934
      PictureRight    =   "frmQuery.frx":19D2
      PictureRightWidth=   84
      FormBorderTop   =   "frmQuery.frx":1A70
      FormBorderLeft  =   "frmQuery.frx":1AD2
      FormBorderBottom=   "frmQuery.frx":1B30
      FormBorderRight =   "frmQuery.frx":1B92
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      AllowMaximize   =   0   'False
      FormIcon        =   "frmQuery.frx":1BF0
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
      PictureMaximize =   "frmQuery.frx":2ACA
      PictureMinimize =   "frmQuery.frx":2E5C
      PictureClose    =   "frmQuery.frx":31EE
      PictureMinimizeToTray=   "frmQuery.frx":3580
      CaptionPrefix   =   "docCER> "
      PictureShrink   =   "frmQuery.frx":3912
      PictureCloseDown=   "frmQuery.frx":3CA4
      PictureMaximizeDown=   "frmQuery.frx":4036
      PictureMinimizeDown=   "frmQuery.frx":43C8
      PictureShrinkDown=   "frmQuery.frx":475A
      PictureMinimizeToTrayDown=   "frmQuery.frx":4AEC
      ControlMenu     =   0   'False
      PicturePin      =   "frmQuery.frx":4E7E
      PicturePinDown  =   "frmQuery.frx":5210
      PicturePinHover =   "frmQuery.frx":55A2
      PictureMinimizeToTrayHover=   "frmQuery.frx":58F4
      PictureShrinkHover=   "frmQuery.frx":5C46
      PictureMinimizeHover=   "frmQuery.frx":5F98
      PictureMaximizeHover=   "frmQuery.frx":62EA
      PictureCloseHover=   "frmQuery.frx":663C
      TrayTip         =   " docCER> Record selection "
      FormMouseIcon   =   "frmQuery.frx":698E
      TrayIcon        =   "frmQuery.frx":71A8
   End
   Begin prjdocCER.chameleonButton cmdValidateSQL 
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Validate SQL"
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
      MICON           =   "frmQuery.frx":8082
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optRecordSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "SQL"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2460
      Width           =   855
   End
   Begin VB.OptionButton optRecordSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Query"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   855
   End
   Begin VB.OptionButton optRecordSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Table"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox cboQuery 
      Height          =   360
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   6135
   End
   Begin VB.ComboBox cboTable 
      Height          =   360
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   6135
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      Height          =   1560
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2400
      Width           =   6135
   End
   Begin prjdocCER.ctlHeader ctlHeader 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1931
      Title           =   "Record selection"
      TitleDescription=   "Please specify the query/table that contains the data you want to retrieve from the database."
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
      Picture         =   "frmQuery.frx":809E
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
      MICON           =   "frmQuery.frx":9BF0
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
      MICON           =   "frmQuery.frx":9C0C
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
      MICON           =   "frmQuery.frx":9C28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmDatabase.Show
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to exit docCER?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbYes Then Quit
End Sub

Private Sub cmdNext_Click()
    If optRecordSource(0).Value = True Then AppSetting.SQL = cboTable.Text
    If optRecordSource(1).Value = True Then AppSetting.SQL = cboQuery.Text
    If optRecordSource(2).Value = True Then AppSetting.SQL = txtSQL.Text
    
    Me.Hide
    frmData.Show
End Sub

Private Sub cmdValidateSQL_Click()
    On Error GoTo ERROR_HANDLER_cmdValidateSQL_Click

    If Query(txtSQL.Text).RecordCount > -1 Then
        MsgBox "The SQL seems to be OK."
    Else
    
    End If

EXIT_cmdValidateSQL_Click:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_cmdValidateSQL_Click:
    Select Case Err.Number
    Case -2147217865, -2147217900, -2147217908 'Table/query not found, Syntax error, Missing reserved word
        MsgBox "Sorry, the SQL seems to be invalid!"
    Case Else
        If MsgBox("Error in Sub cmdValidateSQL_Click() of Form frmQuery (frmQuery.frm) of Project prjdocCER (prjdocCER.vbp)" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "Application error!") = vbNo Then Resume EXIT_cmdValidateSQL_Click
    End Select
    
    Resume EXIT_cmdValidateSQL_Click
End Sub

Private Sub Form_Load()
    On Error GoTo ERROR_HANDLER_Form_Load
    
    If AppSetting.DatabaseType = 0 Then 'MS Access
        With Query("SELECT Name FROM MSysObjects WHERE Type = 1 AND LEFT(Name, 4) <> 'MSys' ORDER BY Name")
            cboTable.AddItem .Fields!Name
        End With
        
        With Query("SELECT Name FROM MSysObjects WHERE Type = 5 AND LEFT(Name, 4) <> 'MSys' ORDER BY Name")
            cboQuery.AddItem .Fields!Name
        End With
        
        optRecordSource(0).Value = True
        optRecordSource_Click 0
    Else
        optRecordSource(2).Value = True
        optRecordSource_Click 2
        
        optRecordSource(0).Enabled = False
        optRecordSource(1).Enabled = False
    End If
    
Set_Data_Source_Form_Profile:
    If ProfileLoaded Then
        Select Case AppSetting.RecordSelectionType
        Case 0 'Table
            cboTable.Text = AppSetting.SQL
        Case 1 'Query
            cboQuery.Text = AppSetting.SQL
        Case 2 'SQL
            txtSQL.Text = AppSetting.SQL
        End Select
        
        optRecordSource(CInt(AppSetting.RecordSelectionType)).Value = True
        optRecordSource_Click CInt(AppSetting.RecordSelectionType)
    End If
    
EXIT_Form_Load:
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER_Form_Load:
    Select Case Err.Number
    Case 383 ''Text' property is read-only
    Case -2147217911 ''MSysObjects' could not be read
'        MsgBox "The table/query list could not be retrieved from the database, you must provide the SQL to retrieve data.", vbOKOnly + vbInformation, "Access denied!"
        
        optRecordSource(2).Value = True
        optRecordSource_Click 2
        
        optRecordSource(0).Enabled = False
        optRecordSource(1).Enabled = False
        
        Resume Set_Data_Source_Form_Profile
    Case Else
        If MsgBox("Error in Sub Form_Load() of Form frmQuery (frmQuery.frm) of Project prjdocCER (prjdocCER.vbp)" & vbCrLf & vbCrLf & "Error#" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Please check that you didn't specify any wrong value (like alphanumeric input in the numeric field) or missed any required input. If the trouble persists, please press [ALT] + [PRNSCR] on your keboard to take a snapshot of this error message, open PaintBrush from the 'Start menu>Accessories', press '[CTL] + V' to paste the snapshot, save the image and email it to SKJoy2001@Yahoo.Com as an attachment for the resolution." & vbCrLf & vbCrLf & "Do you want to continue the action?", vbCritical + vbYesNo, "Application error!") = vbNo Then Resume ERROR_HANDLER_Form_Load
    End Select
    
    Resume Next
End Sub

Private Sub optRecordSource_Click(Index As Integer)
cboTable.Enabled = False
cboQuery.Enabled = False
txtSQL.Enabled = False

Select Case Index
Case 0 'Table
    cboTable.Enabled = True
Case 1 'Query
    cboQuery.Enabled = True
Case 2 'SQL
    txtSQL.Enabled = True
End Select

AppSetting.RecordSelectionType = Index
End Sub
