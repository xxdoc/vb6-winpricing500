VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCapitalMovement"
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

'<---------------------------------------------------------------------------------------------------------------------
Public CAPITAL_MOVEMENT_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public CAPITAL_AMOUNT As Double
'--------------------------------------------------------------------------------------------------------------------- >

Public CURRENT_AMOUNT As Double                 'ยอดนี้เอามาใช้ออกรายงาน เพื่อดูว่าตอนนี้มีจำนวนคงเหลือเทียบกับต้นทุนเท่าไหร่แต่ว่ายังไม่ได้ เพิ่มเข้า TABLE จริงๆ แต่จริงๆ ก็ควรจะมีนะ

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String

Public FROM_DATE As Date
Public TO_DATE As Date
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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
Dim TempStr As String

   If Ind = 1 Then      'Normal and For LoadCapitalMovement
      SelectStr1 = "SELECT CM.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
   
   End If
   
   If CAPITAL_MOVEMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(Trim(FROM_STOCK_NO)) > 0 Then
      Call AddSearchStockCode(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
   End If
   
   If Len(Trim(TO_STOCK_NO)) > 0 Then
      Call AddSearchStockCode(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
   End If
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.STOCK_NO, PI.STOCK_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.CAPITAL_MOVEMENT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CM.DOCUMENT_DATE "
     Else
         OrderByStr = "ORDER BY CM.CAPITAL_MOVEMENT_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      CAPITAL_MOVEMENT_ID = NVLI(Rs("CAPITAL_MOVEMENT_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
'Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("CAPITAL_MOVEMENT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO CAPITAL_MOVEMENT " & _
                     "( " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   CAPITAL_AMOUNT, " & _
                     "   LOCATION_ID, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                     "" & TempID1 & "," & _
                     "" & CAPITAL_AMOUNT & "," & _
                     "" & TempID2 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'CAPITAL_MOVEMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CAPITAL_MOVEMENT SET " & _
                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   PART_ITEM_ID = " & TempID1 & ", " & _
                     "   CAPITAL_AMOUNT = " & CAPITAL_AMOUNT & ", " & _
                     "   LOCATION_ID = " & TempID2 & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
      m_Conn.Execute (SQL1)
      
End Function
Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
Dim WhereStr2 As String
Dim TempStr  As String
Dim FromStr As String
   
   FromStr = "DELETE FROM CAPITAL_MOVEMENT CM "
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") " & vbCrLf
   End If
      
   If Len(Trim(FROM_STOCK_NO)) > 0 Then
      'Call AddSearchStockCode(FromStr)
      If Len(WhereStr2) > 0 Then
         TempStr = "AND "
      End If
      WhereStr2 = WhereStr2 & TempStr & " (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
   End If
   
   If Len(Trim(TO_STOCK_NO)) > 0 Then
      'Call AddSearchStockCode(FromStr)
      If Len(WhereStr2) > 0 Then
         TempStr = "AND "
      End If
      WhereStr2 = WhereStr2 & TempStr & " (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
   End If
   
   
   If Len(WhereStr2) > 0 Then
      SQL1 = FromStr & WhereStr & " AND CM.PART_ITEM_ID IN (SELECT SC.STOCK_CODE_ID FROM STOCK_CODE SC WHERE " & WhereStr2 & ")"
   Else
      SQL1 = FromStr & WhereStr
   End If
   Call m_Conn.Execute(SQL1)
End Function
Public Sub AddSearchStockCode(FromStr As String)
   If InStr(1, FromStr, "SC") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (CM.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   End If
End Sub

