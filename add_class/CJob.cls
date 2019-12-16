VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJob"
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

Public JOB_ID As Long
Public JOB_NO As String
Public JOB_DESC As String
Public JOB_DATE As Date
Public JOB_DATETIME As Date
Public INVENTORY_DOC_ID As Long

Public RESPONSE_BY  As Long
Public COMMIT_FLAG  As String
Public PRD_LOCATION_ID As Long            'สถานที่ผลิต
Public PRODUCT_LOCATION_ID As Long        'สถานที่ที่เบิกวัตถุดิบ

Public PRODUCT_ID As Long           'สินค้าที่ผลิต
Public LOT_ITEM_AMOUNT As Double    ' จำนวนตาม LOT_ITEM
Public PRODUCT_AMOUNT As Double     'เก็บจำนวนตาม ที่เลือกมาผลิต
'--------------------------------------------------------------------------------------------------------------------- >

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date

Public PRODUCT_NO  As String
Public PRODUCT_DESC  As String

Public INVENTORY_REF_NO As String  'เป็นหมายเลขเอกสารนำเข้า/ซื้อเข้าที่ใช้ผลิต
Public INVENTORY_REF_NO_MAIN As String  'เป็นหมายเลขเอกสารนำเข้า/ซื้อเข้าที่ใช้ผลิตขั้นแรก
Public JobInItems As Collection
Public JobOutItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set JobInItems = New Collection
   Set JobOutItems = New Collection
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
      SelectStr1 = "SELECT J.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PRODUCT_NO,SC.STOCK_DESC PRODUCT_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = J.PRODUCT_ID) "
      
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If Trim(JOB_NO) <> "" Then
      If WildCard(JOB_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(J.JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(J.JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         End If
      End If
   End If
   
   If RESPONSE_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.RESPONSE_BY = " & RESPONSE_BY & ") "
      Else
         WhereStr = WhereStr & "AND (J.RESPONSE_BY = " & RESPONSE_BY & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(PRODUCT_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')"
      End If
   End If
   
   If Trim(INVENTORY_REF_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = " WHERE ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID) LEFT OUTER JOIN JOB_ITEM JI ON (JI.EXPORT_LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.JOB_ID = J.JOB_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO)) & "') > 0)"
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID)  LEFT OUTER JOIN JOB_ITEM JI ON (JI.EXPORT_LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.JOB_ID = J.JOB_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO)) & "') > 0)"
      End If
   End If
   
   If Trim(INVENTORY_REF_NO_MAIN) <> "" Then
      If WhereStr = "" Then
         WhereStr = " WHERE ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.MAIN_IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID) LEFT OUTER JOIN JOB_ITEM JI ON (JI.EXPORT_LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.JOB_ID = J.JOB_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO_MAIN)) & "') > 0)"
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.MAIN_IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID)  LEFT OUTER JOIN JOB_ITEM JI ON (JI.EXPORT_LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.JOB_ID = J.JOB_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO_MAIN)) & "') > 0)"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PRODUCT_NO, PI.PRODUCT_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY J.JOB_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY J.JOB_DATE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY J.JOB_DATETIME "
     Else
         OrderByStr = "ORDER BY J.JOB_ID "
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
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DESC = NVLS(Rs("JOB_DESC"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      JOB_DATETIME = InternalDateToDate(NVLS(Rs("JOB_DATETIME"), ""))
      
      RESPONSE_BY = NVLI(Rs("RESPONSE_BY"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PRD_LOCATION_ID = NVLI(Rs("PRD_LOCATION_ID"), -1)
      PRODUCT_LOCATION_ID = NVLI(Rs("PRODUCT_LOCATION_ID"), -1)
      
      PRODUCT_ID = NVLI(Rs("PRODUCT_ID"), -1)
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      PRODUCT_AMOUNT = NVLD(Rs("PRODUCT_AMOUNT"), 0)
      
      PRODUCT_NO = NVLS(Rs("PRODUCT_NO"), "")
      PRODUCT_DESC = NVLS(Rs("PRODUCT_DESC"), "")
      
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

   TempID1 = PRD_LOCATION_ID
   If PRD_LOCATION_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = PRODUCT_ID
   If PRODUCT_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = RESPONSE_BY
   If RESPONSE_BY <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   TempID5 = PRODUCT_LOCATION_ID
   If PRODUCT_LOCATION_ID <= 0 Then
      TempID5 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO JOB " & _
                     "( " & _
                     "   JOB_ID, " & _
                     "   JOB_NO, " & _
                     "   JOB_DESC, " & _
                     "   JOB_DATE, " & _
                     "   JOB_DATETIME, " & _
                     "   LOT_ITEM_AMOUNT, " & _
                     "   RESPONSE_BY, " & _
                     "   COMMIT_FLAG, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PRD_LOCATION_ID, " & _
                     "   PRODUCT_ID, " & _
                     "   PRODUCT_AMOUNT, " & _
                     "   PRODUCT_LOCATION_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & ChangeQuote(Trim(JOB_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(JOB_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(JOB_DATETIME))) & "'," & _
                     LOT_ITEM_AMOUNT & "," & _
                     TempID3 & "," & _
                     "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                     TempID4 & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     PRODUCT_AMOUNT & "," & _
                     TempID5 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         JOB_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE JOB SET " & _
                     "   JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "', " & _
                     "  JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "', " & _
                     "   JOB_DATE= '" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "', " & _
                     "   JOB_DATETIME= '" & ChangeQuote(Trim(DateToStringInt(JOB_DATETIME))) & "', " & _
                     "   RESPONSE_BY = " & TempID3 & "," & _
                     "  LOT_ITEM_AMOUNT = " & LOT_ITEM_AMOUNT & ", " & _
                     "  COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   INVENTORY_DOC_ID = " & TempID4 & "," & _
                     "   PRD_LOCATION_ID = " & TempID1 & "," & _
                     "   PRODUCT_ID = " & TempID2 & "," & _
                     "  PRODUCT_AMOUNT = " & PRODUCT_AMOUNT & ", " & _
                     "   PRODUCT_LOCATION_ID = " & TempID5 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_ID = " & JOB_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_ITEM WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM JOB WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM_LINK  LK WHERE LK.EXPORT_LOT_ITEM_ID IN (SELECT LT.LOT_ITEM_ID FROM LOT_ITEM LT WHERE LT.INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID & ")"
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM INVENTORY_DOC WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
      
End Function
Private Sub Class_Terminate()
   Set JobInItems = Nothing
   Set JobOutItems = Nothing
End Sub
