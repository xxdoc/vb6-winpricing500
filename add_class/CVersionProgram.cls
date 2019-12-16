VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVersionProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public VERSION_ID As Long
Public VERSION_NAME As String


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM VERSION_PROGRAM UG "
   
   WhereStr = ""
   If VERSION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (VERSION_ID = " & VERSION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (VERSION_ID = " & VERSION_ID & ") "
      End If
   End If
   
   If Len(VERSION_NAME) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (VERSION_NAME = '" & VERSION_NAME & "') "
      Else
         WhereStr = WhereStr & "AND (VERSION_NAME = '" & VERSION_NAME & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY VERSION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY VERSION_NAME "
   Else
      OrderByStr = "ORDER BY VERSION_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   VERSION_ID = NVLI(Rs("VERSION_ID"), -1)
   VERSION_NAME = NVLS(Rs("VERSION_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT VERSION_PROGRAM " & _
                     "( " & _
                     "   VERSION_ID, " & _
                     "   VERSION_NAME, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(VERSION_NAME)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         VERSION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE VERSION_PROGRAM SET " & _
                     "   VERSION_NAME = '" & ChangeQuote(Trim(VERSION_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE VERSION_ID = " & VERSION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM VERSION_PROGRAM WHERE VERSION_ID = " & VERSION_ID
      m_Conn.Execute (SQL1)
End Function


