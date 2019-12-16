VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public ImportExportItems As Collection
Public TransferItems As Collection

' TABLE  INVENTORY_DOC ------------------------------------------------------------------------------------------------------------------------
Public INVENTORY_SUB_TYPE As Long
Public INVENTORY_DOC_ID As Long
' TABLE INVENTORY_DOC ------------------------------------------------------------------------------------------------------------------------
Public INVENTORY_REF_NO_MAIN As String
Public HAVE_DISTINCT_FLAG As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public APAR_MAS_ID As Long
Public DRIVER_ID As Long
Public CAR_LICENSE_ID As Long
Public TRANSPORTOR_ID As Long
Public CUSTOMER_BRANCH As Long

Public APAR_CODE As String
Public APAR_NAME As String
Public DRIVER_NAME  As String
Public CAR_LICENSE_NAME  As String
Public TRANSPORTOR_NAME  As String
Public CUSTOMER_BRANCH_NAME  As String

Public PART_ITEM_ID As Long
Public BILL_DESC As String
Public UNIT_AMOUNT As Double
Public TOTAL_AMOUNT As Double


Public ORDER_BY As Long
Public ORDER_TYPE As Long

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set ImportExportItems = New Collection
   Set TransferItems = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set ImportExportItems = Nothing
   Set TransferItems = Nothing
End Sub

Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FName, FType, FValue, FCat)
   Call m_FieldList.add(F, FName)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function
Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub
Public Sub AddEditData()
Dim SQL As String

Dim Tf As CTableField
   
   Set Tf = m_FieldList(PKField)

   Call PopulateInternalField(ShowMode, Me)
   Call SetFieldValue("YYYYMM", Mid(DateToStringInt(GetFieldValue("DOCUMENT_DATE")), 1, 7))
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   
   SQL = "DELETE FROM LOT_ITEM_LINK  LK WHERE LK.EXPORT_LOT_ITEM_ID IN (SELECT LT.LOT_ITEM_ID FROM LOT_ITEM LT WHERE " & PKField & " = " & Tf.GetValue & ")"
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM LOT_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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

Dim DOCUMENT_NO As String
Dim INVENTORY_DOC_ID As Long
Dim DOCUMENT_TYPE As Long
Dim DEPARTMENT_ID As Long
Dim COMMIT_FLAG As String
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim CANCEL_FLAG As String
Dim INVENTORY_SUB_TYPE As Long
Dim STOCK_CODE_NO  As String

Dim TempDate As String

   INVENTORY_DOC_ID = Val(GetFieldValue("INVENTORY_DOC_ID"))
   DOCUMENT_TYPE = Val(GetFieldValue("DOCUMENT_TYPE"))
   DEPARTMENT_ID = Val(GetFieldValue("DEPARTMENT_ID"))
   DOCUMENT_NO = GetFieldValue("DOCUMENT_NO")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   COMMIT_FLAG = GetFieldValue("COMMIT_FLAG")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   CANCEL_FLAG = GetFieldValue("CANCEL_FLAG")
   INVENTORY_SUB_TYPE = Val(GetFieldValue("INVENTORY_SUB_TYPE"))
   STOCK_CODE_NO = GetFieldValue("STOCK_CODE_NO")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.KEY_NAME DEPARTMENT_NAME "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE, NM2.SHORT_NAME APAR_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.DEPARTMENT_ID = MR1.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      WhereStr = ""
   
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
      End If
   End If
         
   If INVENTORY_SUB_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If DEPARTMENT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DEPARTMENT_ID = " & DEPARTMENT_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(COMMIT_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(COMMIT_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.COMMIT_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "')"
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      If Len(Trim(STOCK_CODE_NO)) > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         If WildCard(STOCK_CODE_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (LI.INVENTORY_DOC_ID = NM.INVENTORY_DOC_ID) AND (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
         Else
            WhereStr = WhereStr & TempStr & " (((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (LI.INVENTORY_DOC_ID = NM.INVENTORY_DOC_ID) AND SC.STOCK_NO = '" & ChangeQuote(Trim(STOCK_CODE_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If Trim(INVENTORY_REF_NO_MAIN) <> "" Then
      If WhereStr = "" Then
         WhereStr = " WHERE ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.MAIN_IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID) LEFT OUTER JOIN LOT_ITEM JI ON (JI.LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.INVENTORY_DOC_ID = NM.INVENTORY_DOC_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO_MAIN)) & "') > 0)"
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM_LINK LKK LEFT OUTER JOIN LOT_ITEM LIT ON (LKK.MAIN_IMPORT_LOT_ITEM_ID = LIT.LOT_ITEM_ID) LEFT OUTER JOIN INVENTORY_DOC INDC ON (INDC.INVENTORY_DOC_ID = LIT.INVENTORY_DOC_ID)  LEFT OUTER JOIN LOT_ITEM JI ON (JI.LOT_ITEM_ID = LKK.EXPORT_LOT_ITEM_ID)  WHERE (JI.INVENTORY_DOC_ID = NM.INVENTORY_DOC_ID) AND INDC.DOCUMENT_NO = '" & ChangeQuote(Trim(INVENTORY_REF_NO_MAIN)) & "') > 0)"
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_NO "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub QueryData2(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
Dim TempDate As String

   If Ind = 2 Then 'LoadDistinctTripFormExportType
      SelectStr1 = "SELECT  DISTINCT IND.APAR_MAS_ID "
      SelectStr1 = SelectStr1 & ",IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID,IND.CUSTOMER_BRANCH "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE, NM2.SHORT_NAME APAR_NAME "
      
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME DRIVER_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CAR_LICENSE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME TRANSPORTOR_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME CUSTOMER_BRANCH_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (IND.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (IND.DRIVER_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (IND.CAR_LICENSE_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (IND.TRANSPORTOR_ID = MR3.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (IND.CUSTOMER_BRANCH = MR4.KEY_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "
   ElseIf Ind = 3 Then  'GetDistinctForTranSport
      SelectStr1 = "SELECT  DISTINCT LI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (IND.INVENTORY_DOC_ID = LI.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "
      '-------------------------------------------------------------------------------------------------------------------------------WHERE เฉพาะ ของ ind = 3
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If DRIVER_ID > 0 Then
         WhereStr = WhereStr & TempStr & " (IND.DRIVER_ID = " & DRIVER_ID & ") " & vbCrLf
      Else
         WhereStr = WhereStr & TempStr & " (IND.DRIVER_ID IS NULL) " & vbCrLf
      End If
   
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If CAR_LICENSE_ID > 0 Then
         WhereStr = WhereStr & TempStr & " (IND.CAR_LICENSE_ID = " & CAR_LICENSE_ID & ") " & vbCrLf
      Else
         WhereStr = WhereStr & TempStr & " (IND.CAR_LICENSE_ID IS NULL) " & vbCrLf
      End If
      
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If TRANSPORTOR_ID > 0 Then
         WhereStr = WhereStr & TempStr & " (IND.TRANSPORTOR_ID = " & TRANSPORTOR_ID & ") " & vbCrLf
      Else
         WhereStr = WhereStr & TempStr & " (IND.TRANSPORTOR_ID IS NULL) " & vbCrLf
      End If
      '------------------------------------------------------------------------------------------------------------------------------WHERE เฉพาะ ของ ind = 3
   ElseIf Ind = 4 Then  'GetAmountForTranSport
      SelectStr1 = "SELECT SUM(LI.TX_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",IND.APAR_MAS_ID,IND.CUSTOMER_BRANCH,LI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (IND.INVENTORY_DOC_ID = LI.INVENTORY_DOC_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "
   ElseIf Ind = 5 Then 'LoadDistinctTripFormExportTypeGroup
      SelectStr1 = "SELECT  DISTINCT IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID "
      
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME DRIVER_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CAR_LICENSE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME TRANSPORTOR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (IND.DRIVER_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (IND.CAR_LICENSE_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (IND.TRANSPORTOR_ID = MR3.KEY_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "
   ElseIf Ind = 6 Then  'GetAmountForTranSport2
      SelectStr1 = "SELECT SUM(LI.TX_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",LI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (IND.INVENTORY_DOC_ID = LI.INVENTORY_DOC_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "
   ElseIf Ind = 7 Then  'GetDistinctForTranSport2
      SelectStr1 = "SELECT  DISTINCT LI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IND "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (IND.INVENTORY_DOC_ID = LI.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = "WHERE IND.DOCUMENT_TYPE = " & EXPORT_DOCTYPE & " AND " & " IND.TRIP_FLAG = 'Y' "

   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 4 Then
      GroupByStr = " GROUP BY IND.APAR_MAS_ID,IND.CUSTOMER_BRANCH,LI.PART_ITEM_ID,IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID  "
   ElseIf Ind = 6 Then
      GroupByStr = " GROUP BY LI.PART_ITEM_ID,IND.DRIVER_ID,IND.CAR_LICENSE_ID,IND.TRANSPORTOR_ID  "
   End If
   
   If Ind = 2 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE ,MR2.KEY_CODE,MR3.KEY_CODE,APM.APAR_CODE,MR4.KEY_CODE "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE ,MR2.KEY_CODE,MR3.KEY_CODE "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   End If
      
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
'   Debug.Print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("INVENTORY_DOC_ID", NVLI(Rs("INVENTORY_DOC_ID"), -1))
      Call SetFieldValue("DOCUMENT_NO", NVLS(Rs("DOCUMENT_NO"), ""))
      Call SetFieldValue("DOCUMENT_DATE", InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), "")))
      Call SetFieldValue("DOCUMENT_DESC", NVLS(Rs("DOCUMENT_DESC"), ""))
      Call SetFieldValue("BILL_NO", NVLS(Rs("BILL_NO"), ""))
      Call SetFieldValue("DO_NO", NVLS(Rs("DO_NO"), ""))
      Call SetFieldValue("TRUCK_NO", NVLS(Rs("TRUCK_NO"), ""))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("DELIVERY_ID", NVLI(Rs("DELIVERY_ID"), -1))
      Call SetFieldValue("DELIVERY_FEE", NVLD(Rs("DELIVERY_FEE"), 0))
      Call SetFieldValue("SENDER_NAME", NVLS(Rs("SENDER_NAME"), ""))
      Call SetFieldValue("RECEIVE_NAME", NVLS(Rs("RECEIVE_NAME"), ""))
      Call SetFieldValue("DOCUMENT_TYPE", NVLI(Rs("DOCUMENT_TYPE"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("COMMIT_FLAG", NVLS(Rs("COMMIT_FLAG"), ""))
      Call SetFieldValue("SALE_FLAG", NVLS(Rs("SALE_FLAG"), ""))
      Call SetFieldValue("REASON_ID", NVLI(Rs("REASON_ID"), -1))
      Call SetFieldValue("QUE_NO", NVLS(Rs("QUE_NO"), ""))
      Call SetFieldValue("ENTRY_DATE", InternalDateToDate(NVLS(Rs("ENTRY_DATE"), "")))
      Call SetFieldValue("EXIT_DATE", InternalDateToDate(NVLS(Rs("EXIT_DATE"), "")))
      Call SetFieldValue("EXCEPTION_FLAG", NVLS(Rs("EXCEPTION_FLAG"), ""))
      Call SetFieldValue("DEPARTMENT_ID", NVLI(Rs("DEPARTMENT_ID"), -1))
      Call SetFieldValue("ADJUST_FLAG", NVLS(Rs("ADJUST_FLAG"), ""))
      Call SetFieldValue("ENP_ID", NVLI(Rs("ENP_ID"), -1))
      Call SetFieldValue("CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))
      Call SetFieldValue("INVENTORY_SUB_TYPE", NVLI(Rs("INVENTORY_SUB_TYPE"), -1))
      Call SetFieldValue("COUNT_AMOUNT", NVLS(Rs("COUNT_AMOUNT"), "Y"))
      
      Call SetFieldValue("DEPARTMENT_NAME", NVLS(Rs("DEPARTMENT_NAME"), ""))
      
      Call SetFieldValue("OLD_CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))
      
      Call SetFieldValue("BARCODE_JOB_FLAG", NVLS(Rs("BARCODE_JOB_FLAG"), ""))
      Call SetFieldValue("TEMP_SALE_FLAG", NVLS(Rs("TEMP_SALE_FLAG"), ""))
      
      Call SetFieldValue("CUSTOMER_BRANCH", NVLI(Rs("CUSTOMER_BRANCH"), -1))
      Call SetFieldValue("DRIVER_ID", NVLI(Rs("DRIVER_ID"), -1))
      Call SetFieldValue("CAR_LICENSE_ID", NVLI(Rs("CAR_LICENSE_ID"), -1))
      Call SetFieldValue("TRANSPORTOR_ID", NVLI(Rs("TRANSPORTOR_ID"), -1))
      Call SetFieldValue("TRIP_FLAG", NVLS(Rs("TRIP_FLAG"), ""))
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
   End If
End Sub
Public Sub PopulateFromRS2(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 2 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DRIVER_ID = NVLI(Rs("DRIVER_ID"), -1)
      CAR_LICENSE_ID = NVLI(Rs("CAR_LICENSE_ID"), -1)
      TRANSPORTOR_ID = NVLI(Rs("TRANSPORTOR_ID"), -1)
      CUSTOMER_BRANCH = NVLI(Rs("CUSTOMER_BRANCH"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      DRIVER_NAME = NVLS(Rs("DRIVER_NAME"), "")
      CAR_LICENSE_NAME = NVLS(Rs("CAR_LICENSE_NAME"), "")
      TRANSPORTOR_NAME = NVLS(Rs("TRANSPORTOR_NAME"), "")
      CUSTOMER_BRANCH_NAME = NVLS(Rs("CUSTOMER_BRANCH_NAME"), "")
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DRIVER_ID = NVLI(Rs("DRIVER_ID"), -1)
      CAR_LICENSE_ID = NVLI(Rs("CAR_LICENSE_ID"), -1)
      TRANSPORTOR_ID = NVLI(Rs("TRANSPORTOR_ID"), -1)
      CUSTOMER_BRANCH = NVLI(Rs("CUSTOMER_BRANCH"), -1)
   ElseIf Ind = 5 Then
      DRIVER_ID = NVLI(Rs("DRIVER_ID"), -1)
      CAR_LICENSE_ID = NVLI(Rs("CAR_LICENSE_ID"), -1)
      TRANSPORTOR_ID = NVLI(Rs("TRANSPORTOR_ID"), -1)
      
      DRIVER_NAME = NVLS(Rs("DRIVER_NAME"), "")
      CAR_LICENSE_NAME = NVLS(Rs("CAR_LICENSE_NAME"), "")
      TRANSPORTOR_NAME = NVLS(Rs("TRANSPORTOR_NAME"), "")
   ElseIf Ind = 6 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DRIVER_ID = NVLI(Rs("DRIVER_ID"), -1)
      CAR_LICENSE_ID = NVLI(Rs("CAR_LICENSE_ID"), -1)
      TRANSPORTOR_ID = NVLI(Rs("TRANSPORTOR_ID"), -1)
   ElseIf Ind = 7 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   End If
End Sub
Public Sub RegisterFields()
   TableName = "INVENTORY_DOC"
   SequenceName = "INVENTORY_DOC_SEQ"
   PKField = "INVENTORY_DOC_ID"
      
   If MasterInd = "1" Then
      Call AddField("INVENTORY_DOC_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("DOCUMENT_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("DOCUMENT_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("DOCUMENT_DESC", STRING_TYPE, "", DATA_CAT)
      Call AddField("BILL_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("DO_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("TRUCK_NO", STRING_TYPE, -1, DATA_CAT)
      Call AddField("APAR_MAS_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("DELIVERY_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("DELIVERY_FEE", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("SENDER_NAME", STRING_TYPE, "", DATA_CAT)
      Call AddField("RECEIVE_NAME", STRING_TYPE, "", DATA_CAT)
      Call AddField("DOCUMENT_TYPE", INT_TYPE, -1, DATA_CAT)
      Call AddField("EMP_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("COMMIT_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("SALE_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("REASON_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("QUE_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("ENTRY_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("EXIT_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("EXCEPTION_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("DEPARTMENT_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("ADJUST_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("ENP_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("CANCEL_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("YYYYMM", STRING_TYPE, "", DATA_CAT)
      Call AddField("INVENTORY_SUB_TYPE", INT_TYPE, -1, DATA_CAT)
      Call AddField("COUNT_AMOUNT", STRING_TYPE, "Y", DATA_CAT)         'เป็น Flag ที่ใช้ นับ แสดงยอด LOT ที่ใช้คิด LOT คงเหลือ
      Call AddField("BARCODE_JOB_FLAG", STRING_TYPE, "N", DATA_CAT)  'Flag สำหรับบอกว่ามาจากระบบการยิง Barcode
      Call AddField("TEMP_SALE_FLAG", STRING_TYPE, "N", DATA_CAT)    ' Flag สำหรับบอกว่าการโอนนี้เป็นการฝากขาย
      
      Call AddField("TRIP_FLAG", STRING_TYPE, "N", DATA_CAT)    ' Flag สำหรับบอกว่าเป็นการนำไป Query ลง รายงาน Trip ด้วย
      Call AddField("CUSTOMER_BRANCH", INT_TYPE, -1, DATA_CAT)
      Call AddField("DRIVER_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("CAR_LICENSE_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("TRANSPORTOR_ID", INT_TYPE, -1, DATA_CAT)
      
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
      Call AddField("STOCK_CODE_NO", STRING_TYPE, "", TEMP_CAT)
      Call AddField("DEPARTMENT_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
      
      Call AddField("CONFIG_DOC_TYPE", INT_TYPE, "", TEMP_CAT)
      Call AddField("RUNNING_NO", INT_TYPE, "", TEMP_CAT)
      
      Call AddField("OLD_CANCEL_FLAG", STRING_TYPE, "", TEMP_CAT)
   End If
End Sub
Public Function UndoCommit() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   COMMIT_FLAG = '" & ChangeQuote(Trim(GetFieldValue("COMMIT_FLAG"))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & Val(GetFieldValue("INVENTORY_DOC_ID"))

   Call m_ADOConn.Execute(SQL1)
End Function
Public Function UpdateInventorySubType() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & Val(INVENTORY_DOC_ID)

   Call m_ADOConn.Execute(SQL1)
End Function
Public Function UpdateCountAmount(ToDate As Date) As Boolean
Dim ServerDateTime As String
Dim SQL As String
Dim TempDate  As String
Dim WhereStr As String
Dim TempStr As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ToDate > 0 Then
      TempDate = DateToStringIntHi(ToDate)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   SQL = "UPDATE INVENTORY_DOC SET COUNT_AMOUNT = 'N' " & WhereStr
   m_ADOConn.Execute (SQL)
      
   'SQL = "DELETE FROM LOT_ITEM_LINK LK WHERE LK.EXPORT_LOT_ITEM_ID IN (SELECT LI.LOT_ITEM_ID FROM INVENTORY_DOC IVD LEFT OUTER JOIN LOT_ITEM LI ON (IVD.INVENTORY_DOC_ID = LI.INVENTORY_DOC_ID) " & WhereStr & ")"
   
   SQL = "DELETE FROM DOC_ITEM_LINK "
   m_ADOConn.Execute (SQL)
   
   SQL = "DELETE FROM LOT_ITEM_LINK "
   m_ADOConn.Execute (SQL)
   
End Function
