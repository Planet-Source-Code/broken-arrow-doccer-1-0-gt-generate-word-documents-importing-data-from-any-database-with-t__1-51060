Attribute VB_Name = "moddocCER"
Option Explicit

Private Type SettingDatabase
    MSAccessDatabase As String
    
    MSSQLServerAddress As String
    MSSQLServerDatabase As String
    MSSQLServerUserName As String
    MSSQLServerPassword As String
    
    ODBCConnectionString As String
End Type

Public Type SettingType
    DatabaseType As Long
    Database As SettingDatabase
    RecordSelectionType As Long
    SQL As String
    Condition As String
    WordTemplate As String
    FieldEncloser As String
End Type

Public AppSetting As SettingType
Public ProfileLoaded As Boolean

Sub Main()
frmWelcome.Show
End Sub

Public Sub SetDatabaseConnection()
    Dim ODBCConnectionString As String

    Select Case AppSetting.DatabaseType
    Case 0 'Microsoft Access
        ODBCConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppSetting.Database.MSAccessDatabase & ";Persist Security Info=False"
    Case 1 'Microsoft SQL Server
        ODBCConnectionString = "Provider=SQLOLEDB.1;Password=" & AppSetting.Database.MSSQLServerPassword & ";Persist Security Info=True;User ID=" & AppSetting.Database.MSSQLServerUserName & ";Initial Catalog=" & AppSetting.Database.MSSQLServerDatabase & ";Data Source=" & AppSetting.Database.MSSQLServerAddress
    Case 2 'Other ODBC
        ODBCConnectionString = AppSetting.Database.ODBCConnectionString
    End Select
    
    SetConnection ODBCConnectionString
End Sub

Public Sub Quit()
Dim FormObject As Form

For Each FormObject In Forms
    Unload FormObject
Next
End Sub

Public Sub GenerateDOC()
Dim FieldCount As Long, RS As Recordset, SQL As String
SQL = AppSetting.SQL
If AppSetting.Condition <> "" Then SQL = "SELECT * FROM (" & AppSetting.SQL & ") WHERE " & AppSetting.Condition
Set RS = Query(SQL)

Dim MSWord As Word.Application, TemplateDOC As Document

Set MSWord = New Word.Application
While Not RS.EOF
    With MSWord
        .Visible = True
        If Not TemplateDOC Is Nothing Then
            .Documents.Open AppSetting.WordTemplate
        Else
            Set TemplateDOC = .Documents.Open(AppSetting.WordTemplate)
        End If
        
        .Selection.WholeStory
        .Selection.Copy
        .Documents.Add DocumentType:=wdNewBlankDocument
        .Selection.PasteAndFormat (wdPasteDefault)
        
        For FieldCount = 0 To RS.Fields.Count - 1
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = AppSetting.FieldEncloser & RS.Fields(FieldCount).Name & AppSetting.FieldEncloser   '"%DistributorName%"
                .Replacement.Text = RS.Fields(FieldCount).Value '"MediaNet Ltd."
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        Next
    End With
    
    RS.MoveNext
    
    DoEvents
Wend

TemplateDOC.Close
End Sub

Public Sub SettingSave(SettingINIFile As String)
INISet "Database", "DatabaseType", AppSetting.DatabaseType, SettingINIFile

INISet "Database", "MSAccessDatabase", AppSetting.Database.MSAccessDatabase, SettingINIFile

INISet "Database", "MSSQLServerAddress", AppSetting.Database.MSSQLServerAddress, SettingINIFile
INISet "Database", "MSSQLServerUserName", AppSetting.Database.MSSQLServerUserName, SettingINIFile
INISet "Database", "MSSQLServerPassword", AppSetting.Database.MSSQLServerPassword, SettingINIFile
INISet "Database", "MSSQLServerDatabase", AppSetting.Database.MSSQLServerDatabase, SettingINIFile

INISet "Database", "ODBCConnectionString", AppSetting.Database.ODBCConnectionString, SettingINIFile

INISet "Database", "SQL", AppSetting.SQL, SettingINIFile
INISet "Database", "Condition", AppSetting.Condition, SettingINIFile

INISet "Setting", "RecordSelectionType", AppSetting.RecordSelectionType, SettingINIFile
INISet "Setting", "FieldEncloser", AppSetting.FieldEncloser, SettingINIFile
INISet "Setting", "WordTemplate", AppSetting.WordTemplate, SettingINIFile
End Sub

Public Sub SettingLoad(SettingINIFile As String)
AppSetting.DatabaseType = INIGetLong("Database", "DatabaseType", 0, SettingINIFile)

AppSetting.Database.MSAccessDatabase = INIGetString("Database", "MSAccessDatabase", "", SettingINIFile)

AppSetting.Database.MSSQLServerAddress = INIGetString("Database", "MSSQLServerAddress", "", SettingINIFile)
AppSetting.Database.MSSQLServerUserName = INIGetString("Database", "MSSQLServerUserName", "", SettingINIFile)
AppSetting.Database.MSSQLServerPassword = INIGetString("Database", "MSSQLServerPassword", "", SettingINIFile)
AppSetting.Database.MSSQLServerDatabase = INIGetString("Database", "MSSQLServerDatabase", "", SettingINIFile)

AppSetting.Database.ODBCConnectionString = INIGetString("Database", "ODBCConnectionString", "", SettingINIFile)

AppSetting.SQL = INIGetString("Database", "SQL", "", SettingINIFile)
AppSetting.Condition = INIGetString("Database", "Condition", "", SettingINIFile)

AppSetting.RecordSelectionType = INIGetLong("Setting", "RecordSelectionType", 0, SettingINIFile)
AppSetting.FieldEncloser = INIGetString("Setting", "FieldEncloser", "%", SettingINIFile)
AppSetting.WordTemplate = INIGetString("Setting", "WordTemplate", "", SettingINIFile)

ProfileLoaded = True
End Sub
