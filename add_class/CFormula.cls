VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormula"
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

Public FORMULA_ID As Long
Public FORMULA_NO As String
Public FORMULA_DESC As String
Public FORMULA_DATE As Date
Public PART_ITEM_ID As Long
Public PRD_LOCATION_ID As Long
'--------------------------------------------------------------------------------------------------------------------- >

Public PART_NO As String
Public PART_DESC As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date

Public FormulaItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set FormulaItems = New Collection
   
   FROM_DATE = -1
   TO_DATE = -1
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
      SelectStr1 = "SELECT FM.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM FORMULA FM "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = FM.PART_ITEM_ID) "
      
   End If
   
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FM.FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FM.FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
 
    If Trim(FORMULA_NO) <> "" Then
      If WildCard(FORMULA_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FM.FORMULA_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FM.FORMULA_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FM.FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (FM.FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "')"
         End If
      End If
   End If
   
   If PRD_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FM.PRD_LOCATION_ID = " & PRD_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FM.PRD_LOCATION_ID = " & PRD_LOCATION_ID & ") "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (FM.FORMULA_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (FM.FORMULA_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (FM.FORMULA_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (FM.FORMULA_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY FM.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FM.FORMULA_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FM.FORMULA_DATE "
     Else
         OrderByStr = "ORDER BY FM.FORMULA_ID "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
      FORMULA_NO = NVLS(Rs("FORMULA_NO"), "")
      FORMULA_DESC = NVLS(Rs("FORMULA_DESC"), "")
      FORMULA_DATE = InternalDateToDate(NVLS(Rs("FORMULA_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PRD_LOCATION_ID = NVLI(Rs("PRD_LOCATION_ID"), -1)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = PRD_LOCATION_ID
   If PRD_LOCATION_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO FORMULA " & _
                     "( " & _
                     "   FORMULA_ID, " & _
                     "   FORMULA_NO, " & _
                     "   FORMULA_DESC, " & _
                     "   FORMULA_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   PRD_LOCATION_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & ChangeQuote(Trim(FORMULA_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(FORMULA_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(FORMULA_DATE))) & "'," & _
                     TempID2 & "," & _
                     TempID3 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         FORMULA_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE FORMULA SET " & _
                     "   FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "', " & _
                     "  FORMULA_DESC = '" & ChangeQuote(Trim(FORMULA_DESC)) & "', " & _
                     "   FORMULA_DATE= '" & ChangeQuote(Trim(DateToStringInt(FORMULA_DATE))) & "', " & _
                     "   PART_ITEM_ID = " & TempID2 & "," & _
                     "   PRD_LOCATION_ID = " & TempID3 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_ID = " & FORMULA_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM FORMULA_ITEM WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM FORMULA WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
End Function
Private Sub Class_Terminate()
   Set FormulaItems = Nothing
End Sub
