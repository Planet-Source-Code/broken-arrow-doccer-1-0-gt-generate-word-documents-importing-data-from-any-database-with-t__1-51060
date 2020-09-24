Attribute VB_Name = "modDatabase"
Option Explicit

Private DefaultConnectionString As String

Public Function Query(SQL As String, Optional ODBCConnectionString As String = "") As Recordset
    If ODBCConnectionString = "" And DefaultConnectionString = "" Then
        MsgBox "No database connection has been set or provided! Please use SetConnection(ODBCConnectionString As String) to set up a connection or provide the connection string as the optional parameter."
        Exit Function
    End If
    
    If ODBCConnectionString = "" Then ODBCConnectionString = DefaultConnectionString
    
    Dim ADOConn As New ADODB.Connection
    ADOConn.ConnectionString = ODBCConnectionString
    ADOConn.Open
    
    If Left(LTrim(UCase(SQL)), 7) = "INSERT " Then
        Dim PCount As Long
        ADOConn.Execute SQL
        
        Set Query = New Recordset
        Query.Open "SELECT @@IDENTITY AS LATEST_IDENTITY_VALUE", ADOConn, adOpenStatic, adLockOptimistic
    ElseIf Left(LTrim(UCase(SQL)), 7) = "UPDATE " Then 'Or Left(LTrim(UCase(SQL)), 7) = "INSERT " Then
        ADOConn.Execute SQL
    Else
        Set Query = New Recordset
        Query.Open SQL, ADOConn, adOpenStatic, adLockOptimistic
    End If
End Function

Public Sub SetConnection(ODBCConnectionString As String)
    DefaultConnectionString = ODBCConnectionString
End Sub

