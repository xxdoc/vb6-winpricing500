VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public KEY_ID As Long
Public KEY_CODE As String
Public KEY_NAME As String
Public KEY_NAME2 As String
Public MASTER_AREA As Long
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public PARENT_ID As Long
Public PARENT_EX_ID As Long            'LINK GO TO EMPLOYEE
Public PARENT_EX_ID2 As Long           ' LINK GO TO APAR_MAS
Public PARENT_EX_ID3 As Long
Public PARENT_EX_ID4 As Long
Public PARENT_EX_ID5 As Long
Public MASTER_FLAG As String
Public DEALER_ID As Long
Public CASH_DELIVERY_FLAG As String

Public KEY_LOOKUP As String
Public APAR_CODE As String
Public EMP_CODE As String
Public INDEX_LINK As Long
Public SHORT_CODE  As String
Public FROM_LOCATION_ID As String
Public TO_LOCATION_ID As String

Public APAR_NAME As String
Public MasterRefDetails As Collection

Public EMP_NAME As String
Public EMP_LNAME As String
Public DEALER_CODE As String
Public DEALER_NAME  As String
Public DEALER_LNAME As String
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set MasterRefDetails = New Collection
End Sub
Private Sub Class_Terminate()
   Set MasterRefDetails = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM MASTER_REF WHERE (KEY_ID = " & KEY_ID & ")"
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim SubLen As Long
Dim NewStr As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      SelectStr1 = SelectStr1 & ", NM.LONG_NAME EMP_NAME, NM.LAST_NAME EMP_LNAME "
      SelectStr1 = SelectStr1 & ",EM2.EMP_CODE DEALER_CODE,NM2.LONG_NAME DEALER_NAME, NM2.LAST_NAME DEALER_LNAME,MR2.KEY_NAME KEY_NAME2 "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM2 ON (MR.DEALER_ID = EM2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EM2.EMP_ID = EN2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN2.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then     'Patch_2007_03_09_1_jill
      SelectStr1 = "SELECT MR.PARENT_EX_ID2,MR.KEY_ID "
      SelectStr2 = "SELECT COUNT(*) "

      FromStr = "FROM MASTER_REF MR "
      
      TempStr = "WHERE "
      
      WhereStr = TempStr & " (SELECT COUNT(MR1.PARENT_EX_ID2) FROM MASTER_REF MR1 WHERE MR1.MASTER_AREA = 21 AND MR1.PARENT_EX_ID2 = MR.PARENT_EX_ID2 GROUP BY MR1.PARENT_EX_ID2) = 1 "
   ElseIf Ind = 5 Then  'CReportMaster001
      SelectStr1 = "SELECT MR.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DISTINCT MR.PARENT_EX_ID2 , APM.APAR_CODE, NM2.LONG_NAME APAR_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      SelectStr1 = SelectStr1 & ", NM.LONG_NAME EMP_NAME, NM.LAST_NAME EMP_LNAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_REF MR "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   
   End If
   
   
   
   If KEY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.KEY_ID = " & KEY_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(KEY_CODE)) > 0 Then
      If WildCard(KEY_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(KEY_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(KEY_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (MR.KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (MR.KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "')"
         End If
      End If
   End If
   
   
   If Trim(KEY_NAME) <> "" Then
      If WildCard(KEY_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "KEY_NAME", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "KEY_NAME", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (MR.KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (MR.KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "')"
         End If
      End If
   End If
   
   If MASTER_AREA > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.MASTER_AREA = " & MASTER_AREA & ") " & vbCrLf
   End If
   
   If PARENT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_ID = " & PARENT_ID & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID = " & PARENT_EX_ID & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID2 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID2 = " & PARENT_EX_ID2 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID3 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID3 = " & PARENT_EX_ID3 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID4 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID4 = " & PARENT_EX_ID4 & ") " & vbCrLf
   End If
   
   If PARENT_EX_ID5 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.PARENT_EX_ID5 = " & PARENT_EX_ID5 & ") " & vbCrLf
   End If
   
   If INDEX_LINK > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.INDEX_LINK = " & INDEX_LINK & ") " & vbCrLf
   End If
   
   If Len(Trim(MASTER_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR.MASTER_FLAG = '" & ChangeQuote(Trim(MASTER_FLAG)) & "')"
   End If
   
   If Len(FROM_LOCATION_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION_ID)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION_ID)) & "') "
      End If
   End If
   
   If Len(TO_LOCATION_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION_ID)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION_ID)) & "') "
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY KEY_CODE "
   ElseIf Ind = 5 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY KEY_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY KEY_CODE "
   ElseIf Ind = 7 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY KEY_NAME "
      Else
         OrderByStr = "ORDER BY KEY_CODE "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
 '  'debug.print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
      MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      PARENT_EX_ID = NVLI(Rs("PARENT_EX_ID"), -1)
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
      PARENT_EX_ID3 = NVLI(Rs("PARENT_EX_ID3"), -1)
      PARENT_EX_ID4 = NVLI(Rs("PARENT_EX_ID4"), -1)
      PARENT_EX_ID5 = NVLI(Rs("PARENT_EX_ID5"), -1)
      INDEX_LINK = NVLI(Rs("INDEX_LINK"), -1)
      
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "N")
      SHORT_CODE = NVLS(Rs("SHORT_CODE"), "")
      DEALER_ID = NVLI(Rs("DEALER_ID"), -1)
      
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("KEY_CODE"), "")
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
      DEALER_CODE = NVLS(Rs("DEALER_CODE"), "")
      DEALER_NAME = NVLS(Rs("DEALER_NAME"), "")
      DEALER_LNAME = NVLS(Rs("DEALER_LNAME"), "")
      
      CASH_DELIVERY_FLAG = NVLS(Rs("CASH_DELIVERY_FLAG"), "")
      
      KEY_NAME2 = NVLS(Rs("KEY_NAME2"), "")
      
   ElseIf Ind = 2 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
   ElseIf Ind = 5 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
      SHORT_CODE = NVLS(Rs("SHORT_CODE"), "")
   ElseIf Ind = 6 Then
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      KEY_ID = NVLI(Rs("PARENT_EX_ID2"), -1)
      KEY_LOOKUP = NVLS(Rs("APAR_CODE"), "")
   
   ElseIf Ind = 7 Then
      KEY_ID = NVLI(Rs("KEY_ID"), -1)
      KEY_CODE = NVLS(Rs("KEY_CODE"), "")
      KEY_NAME = NVLS(Rs("KEY_NAME"), "")
      MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      PARENT_EX_ID = NVLI(Rs("PARENT_EX_ID"), -1)
      PARENT_EX_ID2 = NVLI(Rs("PARENT_EX_ID2"), -1)
      PARENT_EX_ID3 = NVLI(Rs("PARENT_EX_ID3"), -1)
      PARENT_EX_ID4 = NVLI(Rs("PARENT_EX_ID4"), -1)
      PARENT_EX_ID5 = NVLI(Rs("PARENT_EX_ID5"), -1)
      INDEX_LINK = NVLI(Rs("INDEX_LINK"), -1)
      
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "N")
      SHORT_CODE = NVLS(Rs("SHORT_CODE"), "")
      
      KEY_ID = NVLI(Rs("PARENT_EX_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("EMP_CODE"), "")
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4  As String
Dim TempID5  As String
Dim TempID6  As String
Dim TempID7  As String
Dim TempID8 As String

   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PARENT_EX_ID
   If PARENT_EX_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PARENT_EX_ID2
   If PARENT_EX_ID2 <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PARENT_EX_ID3
   If PARENT_EX_ID3 <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = PARENT_EX_ID4
   If PARENT_EX_ID4 <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = PARENT_EX_ID5
   If PARENT_EX_ID5 <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID7 = DEALER_ID
   If DEALER_ID <= 0 Then
      TempID7 = "NULL"
   End If
   
   TempID8 = CASH_DELIVERY_FLAG
   If Len(CASH_DELIVERY_FLAG) = 0 Then
      TempID8 = "N"
   End If
   
   'เป็น Link ที่ ไม่มี FK ซึ่งตอนนี้ใช้กับประเภทเอกสารคลัง
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MASTER_REF_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MASTER_REF " & _
                     "( " & _
                     "   KEY_ID, " & _
                     "   KEY_CODE, " & _
                     "   KEY_NAME, " & _
                     "   MASTER_AREA, " & _
                     "   PARENT_ID, " & _
                     "   PARENT_EX_ID, " & _
                     "   PARENT_EX_ID2, " & _
                     "   PARENT_EX_ID3, " & _
                     "   PARENT_EX_ID4, " & _
                     "   PARENT_EX_ID5, " & _
                     "   MASTER_FLAG, " & _
                     "   SHORT_CODE, " & _
                     "   INDEX_LINK, " & _
                     "   DEALER_ID, " & _
                     "   CASH_DELIVERY_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(KEY_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(KEY_NAME)) & "'," & _
                        MASTER_AREA & "," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                        TempID5 & "," & _
                        TempID6 & "," & _
                        "'" & ChangeQuote(Trim(MASTER_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(SHORT_CODE)) & "'," & _
                        INDEX_LINK & "," & _
                        TempID7 & "," & _
                        "'" & ChangeQuote(Trim(TempID8)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         KEY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MASTER_REF SET " & _
                     "    KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "', " & _
                     "    KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "', " & _
                     "    MASTER_AREA = " & MASTER_AREA & ", " & _
                     "    PARENT_ID = " & TempID1 & ", " & _
                     "    PARENT_EX_ID = " & TempID2 & ", " & _
                     "    PARENT_EX_ID2 = " & TempID3 & ", " & _
                     "    PARENT_EX_ID3 = " & TempID4 & ", " & _
                     "    PARENT_EX_ID4 = " & TempID5 & ", " & _
                     "    PARENT_EX_ID5 = " & TempID6 & ", " & _
                     "    MASTER_FLAG = '" & ChangeQuote(Trim(MASTER_FLAG)) & "', " & _
                     "    SHORT_CODE = '" & ChangeQuote(Trim(SHORT_CODE)) & "', " & _
                     "    INDEX_LINK = " & INDEX_LINK & ", " & _
                     "    DEALER_ID = " & TempID7 & ", " & _
                     "    CASH_DELIVERY_FLAG= '" & ChangeQuote(Trim(TempID8)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE KEY_ID = " & KEY_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
Public Sub UpdateEmployeeBranch(FromEmp As Long, ToEmp As Long)
Dim SQL As String
   
   SQL = "UPDATE MASTER_REF SET PARENT_EX_ID = " & ToEmp & " WHERE (PARENT_EX_ID = " & FromEmp & ") and (MASTER_AREA =  21)"
   Call m_ADOConn.Execute(SQL)
End Sub
Private Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    
    GenerateSearchLike = WhereStr
End Function
