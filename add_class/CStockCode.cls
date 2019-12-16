VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

'----------------------------------------------------------------------------------------------------------------------------------------------
Public STOCK_CODE_ID As Long
Public STOCK_NO As String
Public STOCK_DESC As String
Public STOCK_TYPE As Long
Public STOCK_TYPE_SUB As Long

Public BILL_CODE As String
Public BILL_DESC As String                      'ชื่อย่อย
Public UNIT_ID As Long                             ' ID หน่วยใหญ่
Public STOCK_AREA As Long
Public UNIT_AMOUNT As Double           'จำนวน หน่วยใหญ่ต่อหน่วยย่อย
Public UNIT_PER_BASKET As Double       'จะนวนต่อตะกร้า
Public REPORT_PRIORITY As Long            'ลำดับรายงาน
Public UNIT_CHANGE_ID As Long                ' ID หน่วยย่อย
Public PART_ITEM_RETURN_ID As Long        'รหัสรับคืน
Public CHK_STD_COST  As String                ' เช็คว่าใช้ Standard Cost หรือเปล่า
Public COST_PER_AMOUNT As Double          ' Standard Cost ต่อหน่วยย่อย
Public LOT_FLAG  As String                                  ' ตัด LOT หรือเปล่า แบบ First In First Out
Public OUTLAY_FLAG  As String                      ' Y เป็นค่าใช้จ่าย ปกติจะไม่ออกในรีพอร์ด N เป็นสินค้า
Public EXCEPTION_FLAG As String                 'ยก เว้น
Public GROUP_COM_ID As Long
Public BARCODE As Long
Public JOINT_CODE As String
Public REBATE_FLAG As String
'----------------------------------------------------------------------------------------------------------------------------------------------

Public GROUP_ID  As Long
Public UNIT_NAME  As String
Public UNIT_CHANGE_NAME  As String
Public STOCK_TYPE_CODE  As String
Public STOCK_TYPE_NAME  As String

Public LOCATION_GROUP_CODE  As String
Public LOCATION_GROUP_NAME  As String

Public STOCK_GROUP_CODE  As String
Public STOCK_GROUP_NAME  As String

Public GROUP_COM_DESC As String
'----------------------------------------------------------------------------------------------------------------------------------------------
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public LOCATION_GROUP_NO As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public KEY_ID  As Long
Public KEY_LOOKUP As String

Public StockCodeChange As Collection
Public StockCodeChangeFt As Collection
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   Set StockCodeChange = New Collection
   Set StockCodeChangeFt = New Collection
End Sub
Private Sub Class_Terminate()
   Set StockCodeChange = Nothing
   Set StockCodeChangeFt = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM STOCK_CODE_CHANGE_FT WHERE  (STOCK_CODE_ID = " & STOCK_CODE_ID & ")"
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM STOCK_CODE_CHANGE WHERE  (STOCK_CODE_ID = " & STOCK_CODE_ID & ")"
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM STOCK_CODE WHERE STOCK_CODE_ID = " & STOCK_CODE_ID
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
Dim NewStr As String
Dim SubLen As Long

   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME,MR4.KEY_NAME GROUP_COM_DESC "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (NM.GROUP_COM_ID = MR4.KEY_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (EN.LOCATION_ID = MR.KEY_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR.PARENT_ID = MR2.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      
      WhereStr = ""
   ElseIf Ind = 3 Then           'CReportInventoryDoc1_1
      SelectStr1 = "SELECT NM.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ", MR4.KEY_CODE STOCK_GROUP_CODE,MR4.KEY_NAME STOCK_GROUP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR1.PARENT_ID = MR4.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 4 Then              'CReportInventoryDoc3_1 ,FOR LOAD TO CReportInventoryDoc3
      SelectStr1 = "SELECT NM.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.UNIT_CHANGE_ID = MR3.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 5 Then              'LoadStockBarcode
      SelectStr1 = "SELECT NM.*, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.UNIT_CHANGE_ID = MR3.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 6 Then           'CReportInventoryDoc9
      SelectStr1 = "SELECT NM.*, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.UNIT_CHANGE_ID = MR3.KEY_ID) "
      
      WhereStr = " WHERE NM.BARCODE > 0 "
   ElseIf Ind = 7 Then              'CReportInventoryDoc3_4
      SelectStr1 = "SELECT NM.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.STOCK_TYPE = MR1.KEY_ID) "
      WhereStr = ""
      
   End If
   
   If STOCK_CODE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.STOCK_CODE_ID = " & STOCK_CODE_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(STOCK_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "')"
      End If
   End If
   
   If BARCODE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BARCODE = " & BARCODE & ") " & vbCrLf
   End If
   
   If Trim(FROM_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (NM.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (NM.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (NM.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (NM.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(STOCK_DESC) <> "" Then
      If WildCard(STOCK_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "STOCK_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "STOCK_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "')"
         End If
      End If
   End If
   
   If Len(LOCATION_GROUP_CODE) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR2.KEY_CODE = '" & LOCATION_GROUP_CODE & "') " & vbCrLf
   End If
   
   If STOCK_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.STOCK_TYPE = " & STOCK_TYPE & ") " & vbCrLf
   End If
   
   If UNIT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.UNIT_ID = " & UNIT_ID & ") " & vbCrLf
   End If
   
   If GROUP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR1.PARENT_ID = " & GROUP_ID & ") " & vbCrLf
   End If
      
   If Len(Trim(EXCEPTION_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "')"
   End If
   
      If Len(Trim(OUTLAY_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.OUTLAY_FLAG = '" & ChangeQuote(Trim(OUTLAY_FLAG)) & "')"
   End If
      
   If Len(Trim(CHK_STD_COST)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHK_STD_COST = '" & ChangeQuote(Trim(CHK_STD_COST)) & "')"
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE,NM.STOCK_NO  "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE,NM.STOCK_NO "
      Else
         OrderByStr = "ORDER BY NM.STOCK_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY NM.STOCK_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE ASC,MR1.KEY_CODE ASC,NM.STOCK_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE,NM.STOCK_NO  "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY NM.BARCODE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY NM.BARCODE "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY NM.STOCK_NO  "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      BILL_CODE = NVLS(Rs("BILL_CODE"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_SUB = NVLI(Rs("STOCK_TYPE_SUB"), -1)
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      UNIT_PER_BASKET = NVLD(Rs("UNIT_PER_BASKET"), 0)
      REPORT_PRIORITY = NVLI(Rs("REPORT_PRIORITY"), 0)
      PART_ITEM_RETURN_ID = NVLI(Rs("PART_ITEM_RETURN_ID"), -1)
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      LOT_FLAG = NVLS(Rs("LOT_FLAG"), "")
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      OUTLAY_FLAG = NVLS(Rs("OUTLAY_FLAG"), "")
      GROUP_COM_ID = NVLI(Rs("GROUP_COM_ID"), -1)
      REBATE_FLAG = NVLS(Rs("REBATE_FLAG"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      GROUP_COM_DESC = NVLS(Rs("GROUP_COM_DESC"), "")
      
      BARCODE = NVLI(Rs("BARCODE"), 0)
      JOINT_CODE = NVLS(Rs("JOINT_CODE"), "")
      
      KEY_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 2 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
   ElseIf Ind = 3 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      BILL_CODE = NVLS(Rs("BILL_CODE"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_SUB = NVLI(Rs("STOCK_TYPE_SUB"), -1)
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      UNIT_PER_BASKET = NVLD(Rs("UNIT_PER_BASKET"), 0)
      REPORT_PRIORITY = NVLI(Rs("REPORT_PRIORITY"), 0)
      PART_ITEM_RETURN_ID = NVLI(Rs("PART_ITEM_RETURN_ID"), -1)
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      LOT_FLAG = NVLS(Rs("LOT_FLAG"), "")
      OUTLAY_FLAG = NVLS(Rs("OUTLAY_FLAG"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 4 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 5 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      UNIT_PER_BASKET = NVLD(Rs("UNIT_PER_BASKET"), 0)
      
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      OUTLAY_FLAG = NVLS(Rs("OUTLAY_FLAG"), "")
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      BARCODE = NVLI(Rs("BARCODE"), 0)
      
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      
      KEY_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("BARCODE"), "")
   ElseIf Ind = 6 Then
      BARCODE = NVLI(Rs("BARCODE"), 0)
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 7 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
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
Dim TempID6 As String
   
   TempID1 = UNIT_ID
   If UNIT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = STOCK_TYPE
   If STOCK_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = UNIT_CHANGE_ID
   If UNIT_CHANGE_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PART_ITEM_RETURN_ID
   If PART_ITEM_RETURN_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = STOCK_TYPE_SUB
   If STOCK_TYPE_SUB <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = GROUP_COM_ID
   If GROUP_COM_ID <= 0 Then
      TempID6 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("STOCK_CODE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STOCK_CODE " & _
                     "( " & _
                     "   STOCK_CODE_ID, " & _
                     "   STOCK_NO, " & _
                     "   STOCK_DESC, " & _
                     "   UNIT_ID, " & _
                     "   STOCK_TYPE, " & _
                     "   BILL_CODE, " & _
                     "   BILL_DESC, " & _
                     "   STOCK_AREA, " & _
                     "   UNIT_AMOUNT, " & _
                     "   UNIT_PER_BASKET, " & _
                     "   REPORT_PRIORITY, " & _
                     "   UNIT_CHANGE_ID, " & _
                     "   PART_ITEM_RETURN_ID, " & _
                     "   CHK_STD_COST, " & _
                     "   COST_PER_AMOUNT, " & _
                     "   LOT_FLAG, " & _
                     "   OUTLAY_FLAG, " & _
                     "   STOCK_TYPE_SUB, " & _
                     "   EXCEPTION_FLAG, "
         SQL1 = SQL1 & _
                     "   GROUP_COM_ID, " & _
                     "   BARCODE, " & _
                     "   JOINT_CODE, " & _
                     "   REBATE_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(STOCK_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(STOCK_DESC)) & "'," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        "'" & ChangeQuote(Trim(BILL_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(BILL_DESC)) & "'," & _
                        STOCK_AREA & "," & _
                        UNIT_AMOUNT & "," & _
                        UNIT_PER_BASKET & "," & _
                        REPORT_PRIORITY & "," & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                        "'" & ChangeQuote(Trim(CHK_STD_COST)) & "'," & _
                        COST_PER_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(LOT_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(OUTLAY_FLAG)) & "'," & _
                        TempID5 & ","
      SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                        TempID6 & "," & _
                        BARCODE & "," & _
                        "'" & ChangeQuote(Trim(JOINT_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(REBATE_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STOCK_CODE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STOCK_CODE SET " & _
                     "    STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "', " & _
                     "    STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "', " & _
                     "    UNIT_ID = " & TempID1 & ", " & _
                     "    STOCK_TYPE = " & TempID2 & ", " & _
                     "    BILL_CODE = '" & ChangeQuote(Trim(BILL_CODE)) & "', " & _
                     "    BILL_DESC = '" & ChangeQuote(Trim(BILL_DESC)) & "', " & _
                     "    STOCK_AREA = " & STOCK_AREA & ", " & _
                     "    UNIT_AMOUNT = " & UNIT_AMOUNT & ", " & _
                     "    UNIT_PER_BASKET = " & UNIT_PER_BASKET & ", " & _
                     "    REPORT_PRIORITY = " & REPORT_PRIORITY & ", " & _
                     "    UNIT_CHANGE_ID = " & TempID3 & ", " & _
                     "    PART_ITEM_RETURN_ID = " & TempID4 & ", " & _
                     "    CHK_STD_COST = '" & ChangeQuote(Trim(CHK_STD_COST)) & "', " & _
                     "    COST_PER_AMOUNT = " & COST_PER_AMOUNT & ", " & _
                     "    LOT_FLAG = '" & ChangeQuote(Trim(LOT_FLAG)) & "', " & _
                     "    OUTLAY_FLAG = '" & ChangeQuote(Trim(OUTLAY_FLAG)) & "', " & _
                     "    STOCK_TYPE_SUB = " & TempID5 & ", " & _
                     "    EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "    GROUP_COM_ID = " & TempID6 & ", " & _
                     "    BARCODE = " & BARCODE & ", " & _
                     "    JOINT_CODE = '" & ChangeQuote(Trim(JOINT_CODE)) & "', "
         SQL1 = SQL1 & _
                     "    REBATE_FLAG = '" & ChangeQuote(Trim(REBATE_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STOCK_CODE_ID = " & STOCK_CODE_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
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
