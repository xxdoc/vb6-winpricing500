VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocItem"
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
Public PrintLabels As Collection
Public LotItemLinkCollection As Collection


'Public GENERATE_LOT_FLAG As Boolean 'สำหรับตรวจสอบว่า เมื่อไม่มี LotitemLinkCollection แล้ว ให้ทำการสร้างการตัด Lot อย่างอัตโนมัติ

'DOC_ITEM
Public DOC_ITEM_ID As Long
Public CAPITAL_AMOUNT As Double                       'ต้นทุนขาย @
Public TOTAL_INCLUDE_PRICE As Double             'ต้นทุนขาย
'DOC_ITEM

Public EXTRA_DISCOUNT_AMOUNT_DIF As Double         'สำหรับใช้ตรวจสอบว่า ยอด ส่วนลดของหน้าบิลกับที่กระจายให้ item ลูกรวมกันแล้วต่างกันเท่าไหร่

Public LOCATION_NO As String

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set PrintLabels = New Collection
   Set LotItemLinkCollection = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set PrintLabels = Nothing
   Set LotItemLinkCollection = Nothing
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

   Call PopulateInternalField(ShowMode, Me)
   'Call SetFieldValue("TRANSACTION_SEQ", 999)
   
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
   
   SQL = "DELETE FROM DOC_ITEM_LINK WHERE DOC_ITEM_ID = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM PRINT_LABEL WHERE DOC_ITEM_ID = " & Tf.GetValue
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

Dim DOC_ITEM_ID As Long
Dim BILLING_DOC_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim PO_ID As Long
Dim DOC_ITEM_TYPE As Long

   DOC_ITEM_ID = Val(GetFieldValue("DOC_ITEM_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   PO_ID = Val(GetFieldValue("PO_ID"))
   DOC_ITEM_TYPE = Val(GetFieldValue("DOC_ITEM_TYPE"))
   
   If DOC_ITEM_TYPE <= 0 Then
      DOC_ITEM_TYPE = 0
   End If
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EN.*, SC.STOCK_NO,SC.STOCK_DESC,BD.DISCOUNT_VALUE, MR1.KEY_CODE LOCATION_NO,MR1.KEY_NAME LOCATION_NAME, MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_TRAN_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (EN.UNIT_TRAN_ID = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(EN.ITEM_AMOUNT) ITEM_AMOUNT,EN.PART_ITEM_ID  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT EN.*,BD.EXT_DISCOUNT_PERCENT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = "WHERE BD.EXT_DISCOUNT_PERCENT > 0 "
   End If
   
   If DOC_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.DOC_ITEM_ID = " & DOC_ITEM_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If PO_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.PO_ID = " & PO_ID & ") " & vbCrLf
   End If
   
   If DOC_ITEM_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.DOC_ITEM_TYPE = " & DOC_ITEM_TYPE & ") " & vbCrLf
   End If
   
   If Ind = 2 Then
      GroupByStr = "GROUP BY EN.PART_ITEM_ID "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY DOC_ITEM_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EN.PART_ITEM_ID "
   Else
      OrderByStr = "ORDER BY DOC_ITEM_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("AVG_PRICE", NVLD(Rs("AVG_PRICE"), 0))
      Call SetFieldValue("AVG_PRICE_EX", NVLD(Rs("AVG_PRICE_EX"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("CRDR_AMOUNT", NVLD(Rs("CRDR_AMOUNT"), 0))
      Call SetFieldValue("RECEIPT_AMOUNT", NVLD(Rs("RECEIPT_AMOUNT"), 0))
      Call SetFieldValue("GUI_ID", NVLI(Rs("GUI_ID"), -1))
      Call SetFieldValue("LINK_ID", NVLI(Rs("LINK_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_ITEM_RETURN_ID", NVLI(Rs("PART_ITEM_RETURN_ID"), -1))
      Call SetFieldValue("LOCATION_ID", NVLI(Rs("LOCATION_ID"), -1))
      Call SetFieldValue("ITEM_DESC", NVLS(Rs("ITEM_DESC"), ""))
      Call SetFieldValue("DISCOUNT_PERCENT", NVLD(Rs("DISCOUNT_PERCENT"), 0))
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
     ' Call SetFieldValue("DISCOUNT_VALUE", NVLD(Rs("DISCOUNT_VALUE"), 0))
      Call SetFieldValue("SELL_TYPE", NVLI(Rs("SELL_TYPE"), -1))
      Call SetFieldValue("SELL_TYPE_OLD", NVLD(Rs("SELL_TYPE"), 0))
      Call SetFieldValue("STOCK_TYPE", NVLI(Rs("STOCK_TYPE"), -1))
      Call SetFieldValue("PO_NO", NVLS(Rs("PO_NO"), ""))
      Call SetFieldValue("PO_ID", NVLI(Rs("PO_ID"), -1))
      Call SetFieldValue("DOC_ITEM_TYPE", NVLD(Rs("DOC_ITEM_TYPE"), 0))
      Call SetFieldValue("DOC_ITEM_TYPE_OLD", NVLD(Rs("DOC_ITEM_TYPE"), 0))
      Call SetFieldValue("PRODUCT_DETAIL", NVLS(Rs("PRODUCT_DETAIL"), ""))
      Call SetFieldValue("FREE_FLAG", NVLS(Rs("FREE_FLAG"), ""))
      Call SetFieldValue("EXTRA_DISCOUNT_AMOUNT", NVLD(Rs("EXTRA_DISCOUNT_AMOUNT"), 0))
      
      Call SetFieldValue("UNIT_TRAN_ID", NVLI(Rs("UNIT_TRAN_ID"), -1))
      Call SetFieldValue("UNIT_MULTIPLE", NVLD(Rs("UNIT_MULTIPLE"), 0))
      
      Call SetFieldValue("WEIGHT_AMOUNT", NVLD(Rs("WEIGHT_AMOUNT"), 0))
      
      Call SetFieldValue("CAPITAL_AMOUNT", NVLD(Rs("CAPITAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_INCLUDE_PRICE", NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0))
      Call SetFieldValue("LOT_ITEM_FLAG", NVLS(Rs("LOT_ITEM_FLAG"), "Y"))
      
      Call SetFieldValue("STOCK_DESC", NVLS(Rs("STOCK_DESC"), ""))
      Call SetFieldValue("STOCK_NO", NVLS(Rs("STOCK_NO"), ""))
      
      Call SetFieldValue("UNIT_NAME", NVLS(Rs("UNIT_NAME"), ""))
      Call SetFieldValue("OLD_DISCOUNT_PERCENT", NVLD(Rs("DISCOUNT_PERCENT"), 0))
      Call SetFieldValue("UNIT_TRAN_NAME", NVLS(Rs("UNIT_TRAN_NAME"), ""))
      Call SetFieldValue("LOCATION_NAME", NVLS(Rs("LOCATION_NAME"), ""))
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
   ElseIf Ind = 2 Then
   
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
   
   ElseIf Ind = 3 Then
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("AVG_PRICE", NVLD(Rs("AVG_PRICE"), 0))
      Call SetFieldValue("AVG_PRICE_EX", NVLD(Rs("AVG_PRICE_EX"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("CRDR_AMOUNT", NVLD(Rs("CRDR_AMOUNT"), 0))
      Call SetFieldValue("RECEIPT_AMOUNT", NVLD(Rs("RECEIPT_AMOUNT"), 0))
      Call SetFieldValue("GUI_ID", NVLI(Rs("GUI_ID"), -1))
      Call SetFieldValue("LINK_ID", NVLI(Rs("LINK_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_ITEM_RETURN_ID", NVLI(Rs("PART_ITEM_RETURN_ID"), -1))
      Call SetFieldValue("LOCATION_ID", NVLI(Rs("LOCATION_ID"), -1))
      Call SetFieldValue("ITEM_DESC", NVLS(Rs("ITEM_DESC"), ""))
      Call SetFieldValue("DISCOUNT_PERCENT", NVLD(Rs("DISCOUNT_PERCENT"), 0))
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("SELL_TYPE", NVLI(Rs("SELL_TYPE"), -1))
      Call SetFieldValue("SELL_TYPE_OLD", NVLD(Rs("SELL_TYPE"), 0))
      Call SetFieldValue("STOCK_TYPE", NVLI(Rs("STOCK_TYPE"), -1))
      Call SetFieldValue("PO_NO", NVLS(Rs("PO_NO"), ""))
      Call SetFieldValue("PO_ID", NVLI(Rs("PO_ID"), -1))
      Call SetFieldValue("DOC_ITEM_TYPE", NVLD(Rs("DOC_ITEM_TYPE"), 0))
      Call SetFieldValue("DOC_ITEM_TYPE_OLD", NVLD(Rs("DOC_ITEM_TYPE"), 0))
      Call SetFieldValue("PRODUCT_DETAIL", NVLS(Rs("PRODUCT_DETAIL"), ""))
      Call SetFieldValue("FREE_FLAG", NVLS(Rs("FREE_FLAG"), ""))
      
      Call SetFieldValue("UNIT_TRAN_ID", NVLI(Rs("UNIT_TRAN_ID"), -1))
      Call SetFieldValue("UNIT_MULTIPLE", NVLD(Rs("UNIT_MULTIPLE"), 0))
   
      Call SetFieldValue("WEIGHT_AMOUNT", NVLD(Rs("WEIGHT_AMOUNT"), 0))
      
      Call SetFieldValue("EXT_DISCOUNT_PERCENT", NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "DOC_ITEM"
   SequenceName = "DOC_ITEM_SEQ"
   PKField = "DOC_ITEM_ID"
   
   Call AddField("DOC_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BILLING_DOC_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_RETURN_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("LOCATION_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("AVG_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("AVG_PRICE_EX", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("CRDR_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("RECEIPT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("GUI_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("LINK_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("DISCOUNT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
'   Call AddField("DISCOUNT_VALUE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("DISCOUNT_PERCENT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("ITEM_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("STOCK_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("SELL_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("PO_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("PO_ID", INT_TYPE, "-1", DATA_CAT)
   Call AddField("DOC_ITEM_TYPE", INT_TYPE, 1, DATA_CAT)
   Call AddField("PRODUCT_DETAIL", STRING_TYPE, "", DATA_CAT)
   Call AddField("EXTRA_DISCOUNT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("FREE_FLAG", STRING_TYPE, "N", DATA_CAT)
   Call AddField("UNIT_TRAN_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("UNIT_MULTIPLE", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("WEIGHT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("CAPITAL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)     'ต้นทุนขาย@
   Call AddField("TOTAL_INCLUDE_PRICE", MONEY_TYPE, 0, DATA_CAT)     'ต้นทุนขาย
   
   Call AddField("LOT_ITEM_FLAG", STRING_TYPE, "Y", DATA_CAT)     'เป็น Flag ที่เช็คว่ามีผลต่อคลังหรือไม่ เมื่อขายหรือรับคืน
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("STOCK_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("STOCK_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("STOCK_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("LOCATION_NAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("EXT_DISCOUNT_PERCENT", MONEY_TYPE, 0, TEMP_CAT)
   Call AddField("DOC_ITEM_TYPE_OLD", INT_TYPE, 1, TEMP_CAT)
   Call AddField("SELL_TYPE_OLD", INT_TYPE, 1, TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   Call AddField("UNIT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("OLD_DISCOUNT_PERCENT", MONEY_TYPE, 0, TEMP_CAT)
   Call AddField("UNIT_TRAN_NAME", STRING_TYPE, "", TEMP_CAT)
End Sub

Public Function ShowDescText() As String
   If GetFieldValue("SELL_TYPE") = 1 Then
      ShowDescText = GetFieldValue("STOCK_DESC") & " " & GetFieldValue("PRODUCT_DETAIL") & " " & AddDoubleFrontEnd(GetFieldValue("WEIGHT_AMOUNT"), "(นน. ", ")")
   ElseIf GetFieldValue("SELL_TYPE") = 2 Then
      ShowDescText = GetFieldValue("STOCK_DESC") & " " & GetFieldValue("PRODUCT_DETAIL") & " " & AddDoubleFrontEnd(GetFieldValue("WEIGHT_AMOUNT"), "(นน. ", ")")
   ElseIf GetFieldValue("SELL_TYPE") = 3 Then
      ShowDescText = GetFieldValue("ITEM_DESC") & " " & AddDoubleFrontEnd(GetFieldValue("WEIGHT_AMOUNT"), "(นน. ", ")")
   End If
   
End Function

Public Function ShowDescCode() As String
   If GetFieldValue("SELL_TYPE") = 1 Then
      ShowDescCode = GetFieldValue("STOCK_NO")
   ElseIf GetFieldValue("SELL_TYPE") = 2 Then
      ShowDescCode = GetFieldValue("STOCK_NO")
   ElseIf GetFieldValue("SELL_TYPE") = 3 Then
      ShowDescCode = ""
   End If
End Function
Public Function CopyObject(Ind As Long, TempObj As CDocItem)
      'Call SetFieldValue("DOC_ITEM_ID", TempObj.GetFieldValue("DOC_ITEM_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
      'Call SetFieldValue("BILLING_DOC_ID", TempObj.GetFieldValue("BILLING_DOC_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
      Call SetFieldValue("ITEM_AMOUNT", TempObj.GetFieldValue("ITEM_AMOUNT"))
      Call SetFieldValue("AVG_PRICE", TempObj.GetFieldValue("AVG_PRICE"))
      Call SetFieldValue("TOTAL_PRICE", TempObj.GetFieldValue("TOTAL_PRICE"))
      Call SetFieldValue("CRDR_AMOUNT", TempObj.GetFieldValue("CRDR_AMOUNT"))
      Call SetFieldValue("RECEIPT_AMOUNT", TempObj.GetFieldValue("RECEIPT_AMOUNT"))
      Call SetFieldValue("GUI_ID", TempObj.GetFieldValue("GUI_ID"))
      'Call SetFieldValue("LINK_ID", TempObj.GetFieldValue("LINK_ID"))                       'ไม่ต้อง COPY มา เนื่องจากเดี่ยวมันจะรันเพิ่ม
      Call SetFieldValue("PART_ITEM_ID", TempObj.GetFieldValue("PART_ITEM_ID"))
      Call SetFieldValue("PART_ITEM_RETURN_ID", TempObj.GetFieldValue("PART_ITEM_RETURN_ID"))
      Call SetFieldValue("LOCATION_ID", TempObj.GetFieldValue("LOCATION_ID"))
      Call SetFieldValue("ITEM_DESC", TempObj.GetFieldValue("ITEM_DESC"))
      Call SetFieldValue("DISCOUNT_PERCENT", TempObj.GetFieldValue("DISCOUNT_PERCENT"))
      Call SetFieldValue("DISCOUNT_AMOUNT", TempObj.GetFieldValue("DISCOUNT_AMOUNT"))
      Call SetFieldValue("SELL_TYPE", TempObj.GetFieldValue("SELL_TYPE"))
      Call SetFieldValue("DOC_ITEM_TYPE", TempObj.GetFieldValue("DOC_ITEM_TYPE"))
      Call SetFieldValue("STOCK_TYPE", TempObj.GetFieldValue("STOCK_TYPE"))
      Call SetFieldValue("PO_NO", TempObj.GetFieldValue("PO_NO"))
      Call SetFieldValue("PRODUCT_DETAIL", TempObj.GetFieldValue("PRODUCT_DETAIL"))
      Call SetFieldValue("FREE_FLAG", TempObj.GetFieldValue("FREE_FLAG"))
      
      Call SetFieldValue("UNIT_TRAN_ID", TempObj.GetFieldValue("UNIT_TRAN_ID"))
      Call SetFieldValue("UNIT_MULTIPLE", TempObj.GetFieldValue("UNIT_MULTIPLE"))
      
      Call SetFieldValue("STOCK_DESC", TempObj.GetFieldValue("STOCK_DESC"))
      Call SetFieldValue("STOCK_NO", TempObj.GetFieldValue("STOCK_NO"))
      
      Call SetFieldValue("PO_ID", TempObj.GetFieldValue("PO_ID"))
      Call SetFieldValue("UNIT_TRAN_NAME", TempObj.GetFieldValue("UNIT_TRAN_NAME"))
      Call SetFieldValue("CAPITAL_AMOUNT", TempObj.GetFieldValue("CAPITAL_AMOUNT"))
      Call SetFieldValue("TOTAL_INCLUDE_PRICE", TempObj.GetFieldValue("TOTAL_INCLUDE_PRICE"))
      
      Dim Pl As CPrintLabel
      Dim PLA As CPrintLabel
      For Each Pl In TempObj.PrintLabels
         Set PLA = New CPrintLabel
         PLA.Flag = "A"
         'Call PLA.SetFieldValue("PRINT_LABEL_ID", PL.GetFieldValue("PRINT_LABEL_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
         'Call PLA.SetFieldValue("DOC_ITEM_ID", PL.GetFieldValue("DOC_ITEM_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
         Call PLA.SetFieldValue("BLOCK_ID", Pl.GetFieldValue("BLOCK_ID"))
         Call PLA.SetFieldValue("SALE_ID", Pl.GetFieldValue("SALE_ID"))
         Call PLA.SetFieldValue("SALE_LONG_NAME", Pl.GetFieldValue("SALE_LONG_NAME"))
         Call PLA.SetFieldValue("SALE_LAST_NAME", Pl.GetFieldValue("SALE_LAST_NAME"))
         Call PLA.SetFieldValue("BRANCH_ID", Pl.GetFieldValue("BRANCH_ID"))
         Call PLA.SetFieldValue("ITEM_AMOUNT", Pl.GetFieldValue("ITEM_AMOUNT"))
         Call PLA.SetFieldValue("PACK_AMOUNT", Pl.GetFieldValue("PACK_AMOUNT"))
         Call PLA.SetFieldValue("TOTAL_PRICE", Pl.GetFieldValue("TOTAL_PRICE"))
         Call PLA.SetFieldValue("TOTAL_AMOUNT", Pl.GetFieldValue("TOTAL_AMOUNT"))
         
         Call PLA.SetFieldValue("BLOCK_NAME", Pl.GetFieldValue("BLOCK_NAME"))
         Call PLA.SetFieldValue("BRANCH_NAME", Pl.GetFieldValue("BRANCH_NAME"))
         Call PLA.SetFieldValue("BILLING_DOC_ID", Pl.GetFieldValue("BILLING_DOC_ID"))
         Call PLA.SetFieldValue("LABEL_DESC", Pl.GetFieldValue("LABEL_DESC"))
             
         Call PLA.SetFieldValue("PART_ITEM_ID", Pl.GetFieldValue("PART_ITEM_ID"))
         Call PLA.SetFieldValue("PART_NO", Pl.GetFieldValue("PART_NO"))
         Call PLA.SetFieldValue("PART_DESC", Pl.GetFieldValue("PART_DESC"))
         
         Call PLA.SetFieldValue("UNIT_TRAN_ID", Pl.GetFieldValue("UNIT_TRAN_ID"))
         Call PLA.SetFieldValue("UNIT_MULTIPLE", Pl.GetFieldValue("UNIT_MULTIPLE"))
      
         Call PrintLabels.add(PLA)
      Next Pl
      Set Pl = Nothing
      Set PLA = Nothing
End Function
Public Function UpdateCapitalSell()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE DOC_ITEM SET " & _
                     "   CAPITAL_AMOUNT = " & CAPITAL_AMOUNT & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DOC_ITEM_ID = " & DOC_ITEM_ID
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
