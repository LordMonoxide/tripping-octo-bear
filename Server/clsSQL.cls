VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConnection As ADODB.Connection

Private Sub Class_Initialize()
  Set mConnection = New ADODB.Connection
  
  mConnection.ConnectionString = "Provider=MSDASQL;Driver={MySQL ODBC 5.1 Driver};Server=localhost;Database=essence;User=essence;Password=wCK9Leu73nAtQh7P;Option=3;"
  
  Call mConnection.Open
  
  Do While mConnection.State <> ADODB.ObjectStateEnum.adStateOpen
    DoEvents
  Loop
End Sub

Public Function DoSelect(ByRef table As String, Optional ByRef what As String = "*", Optional ByRef where As String) As ADODB.Recordset
  Dim r As ADODB.Recordset
  Set r = New ADODB.Recordset
  
  If LenB(where) <> 0 Then
    Debug.Print "SELECT " & what & " FROM " & table & " WHERE " & where
    Call r.Open("SELECT " & what & " FROM " & table & " WHERE " & where, mConnection)
  Else
    Debug.Print "SELECT " & what & " FROM " & table
    Call r.Open("SELECT " & what & " FROM " & table, mConnection)
  End If
  
  Set DoSelect = r
End Function

Public Function DoInsert(ByRef table As String, ByRef values As String, Optional ByRef columns As String) As Long
  Dim r As ADODB.Recordset
  Set r = New ADODB.Recordset
  
  If LenB(columns) = 0 Then
    Debug.Print "INSERT INTO " & table & " VALUES (" & values & ")"
    Call mConnection.Execute("INSERT INTO " & table & " VALUES (" & values & ")")
  Else
    Debug.Print "INSERT INTO " & table & " (" & columns & ") VALUES (" & values & ")"
    Call mConnection.Execute("INSERT INTO " & table & " (" & columns & ") VALUES (" & values & ")")
  End If
  
  Call r.Open("SELECT LAST_INSERT_ID() AS `id`", mConnection)
  DoInsert = r.fields("id")
End Function

Public Sub DoUpdate(ByRef table As String, ByRef params As String, ByRef where As String)
  Debug.Print "UPDATE " & table & " SET " & params & " WHERE " & where
  Call mConnection.Execute("UPDATE " & table & " SET " & params & " WHERE " & where)
End Sub
