VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormulaItem"
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

'<--------------------------------------------------------------------------------------------------------
Public FORMULA_ITEM_ID As Long
Public FORMULA_ID As Long
Public BATCH_NO As Long
Public ORDER_NO As Long
Public PART_ITEM_ID As Long
Public FORMULA_ITEM_PERCENT As Double

Public LOCATION_ID  As Long

Public LOST_ID As Long
Public PROBLEM_LIMIT_PERCENT As Double
Public NEXT_FLAG  As String
Public PRODUCTION_TYPE As Long

'-------------------------------------------------------------------------------------------------------->

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public PART_NO As String
Public PART_DESC As String
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public UNIT_ID   As Long

Public PROBLEM_DESC As String
Public SUM_FLAG As String

Public PRODUCTION_TYPE_NAME As String

Public FROM_DATE As Date
Public TO_DATE As Date
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT FI.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.UNIT_CHANGE_ID UNIT_ID "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE LOCATION_NO,MR1.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME PRODUCTION_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM FORMULA_ITEM FI "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = FI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = FI.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = FI.LOST_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (MR3.KEY_ID = FI.PRODUCTION_TYPE) "
   End If
   
   If FORMULA_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FI.FORMULA_ITEM_ID = " & FORMULA_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FI.FORMULA_ITEM_ID = " & FORMULA_ITEM_ID & ") "
      End If
   End If
   
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FI.FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FI.FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY FI.BATCH_NO ASC,FI.ORDER_NO ASC,SC.STOCK_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      FORMULA_ITEM_ID = NVLI(Rs("FORMULA_ITEM_ID"), -1)
      FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
      BATCH_NO = NVLI(Rs("BATCH_NO"), 0)
      ORDER_NO = NVLI(Rs("ORDER_NO"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      FORMULA_ITEM_PERCENT = NVLD(Rs("FORMULA_ITEM_PERCENT"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_LIMIT_PERCENT = NVLD(Rs("PROBLEM_LIMIT_PERCENT"), 0)
      NEXT_FLAG = NVLS(Rs("NEXT_FLAG"), "")
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      PRODUCTION_TYPE_NAME = NVLS(Rs("PRODUCTION_TYPE_NAME"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String

   TempID1 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = LOST_ID
   If LOST_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   TempID5 = PRODUCTION_TYPE
   If PRODUCTION_TYPE <= 0 Then
      TempID5 = "NULL "
   End If
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO FORMULA_ITEM " & _
                     "( " & _
                     "   FORMULA_ITEM_ID, " & _
                     "   FORMULA_ID, " & _
                     "   BATCH_NO, " & _
                     "   ORDER_NO, " & _
                     "   PART_ITEM_ID, " & _
                     "   FORMULA_ITEM_PERCENT, " & _
                     "   LOCATION_ID, " & _
                     "   LOST_ID, " & _
                     "   PROBLEM_LIMIT_PERCENT, " & _
                     "   NEXT_FLAG, " & _
                     "   PRODUCTION_TYPE, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     TempID1 & "," & _
                     BATCH_NO & "," & _
                     ORDER_NO & "," & _
                     TempID2 & "," & _
                     FORMULA_ITEM_PERCENT & "," & _
                     TempID3 & "," & _
                     TempID4 & "," & _
                     PROBLEM_LIMIT_PERCENT & "," & _
                     "'" & ChangeQuote(Trim(NEXT_FLAG)) & "'," & _
                     TempID5 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         FORMULA_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE FORMULA_ITEM SET " & _
                     "  ORDER_NO = " & ORDER_NO & ", " & _
                     "  BATCH_NO = " & BATCH_NO & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & "," & _
                     "  FORMULA_ITEM_PERCENT = " & FORMULA_ITEM_PERCENT & ", " & _
                     "   LOCATION_ID = " & TempID3 & "," & _
                     "   LOST_ID = " & TempID4 & "," & _
                     "  PROBLEM_LIMIT_PERCENT = " & PROBLEM_LIMIT_PERCENT & ", " & _
                     "  NEXT_FLAG = '" & ChangeQuote(Trim(NEXT_FLAG)) & "', " & _
                     "   PRODUCTION_TYPE = " & TempID5 & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_ITEM_ID = " & FORMULA_ITEM_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM FORMULA_ITEM WHERE FORMULA_ITEM_ID = " & FORMULA_ITEM_ID
      m_Conn.Execute (SQL1)
End Function
