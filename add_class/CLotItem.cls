VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLotItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String

'<--------------------------------------------------------------------------------------------------------------------------------------------
'Table LOT_ITEM
Public LOT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public FROM_LOCATION As String
Public TO_LOCATION As String
Public TX_AMOUNT  As Double
Public AVG_PRICE As Double
Public TOTAL_INCLUDE_PRICE  As Double
Public TX_TYPE As String
Public MULTIPLIER As Long

Public LINK_ID As Long
Public CANCEL_FLAG As String
Public FREE_FLAG As String

Public UNIT_TRAN_ID  As Long
Public UNIT_MULTIPLE  As Double

Public WEIGHT_AMOUNT  As Double
Public SUM_AMOUNT As Double      'เป็นจำนวนที่ใช้ในการรวมยอดเพื่อออกยอดยกมา

Public CREATE_DATE As Date
Public MODIFY_DATE As Date

Public LOT_ITEM_AMOUNT As Double
Public NEW_AVG_PRICE   As Double
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table INVENTORY_DOC
Public DOCUMENT_TYPE As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public INVENTORY_SUB_TYPE As Long
Public INVENTORY_SUB_TYPE_CODE As String
Public INVENTORY_SUB_TYPE_NAME As String

Public COUNT_AMOUNT    As String
Public TEMP_SALE_FLAG As String
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table STOCK_CODE
Public PART_NO As String
Public PART_DESC As String
Public PART_CODE As String
Public PART_TYPE As Long
Public UNIT_ID As Long
Public UNIT_CHANGE_NAME As String
Public BILL_DESC  As String
Public CHK_STD_COST As String
Public COST_PER_AMOUNT  As Double
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table PART_TYPE
Public PART_GROUP As Long
Public PART_TYPE_CODE As String
Public PART_TYPE_NAME As String
Public PART_GROUP_CODE As String
Public PART_GROUP_NAME As String
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table LOT_ITEM_LINK
Public IMPORT_LOT_ITEM_ID   As Long
Public MAIN_IMPORT_LOT_ITEM_ID As Long
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table APAR_MAS
Public APAR_NAME As String
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table MASTER_REF
Public LOCATION_NO  As String
Public LOCATION_NAME As String
Public LOCATION_CODE As String
Public UNIT_TRAN_NAME  As String
'-------------------------------------------------------------------------------------------------------------------------------------------->
'Table DOC_ITEM
Public DOC_ITEM_ID As Long
'-------------------------------------------------------------------------------------------------------------------------------------------->
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public OUTLAY_FLAG As String
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public FROM_LOCATION_ID As Long
Public TO_LOCATION_ID As Long
Public FROM_LOCATION_CODE As String
Public TO_LOCATION_CODE  As String
Public FROM_LOCATION_CODE2  As String
Public TO_LOCATION_CODE2  As String
Public CONSIGNMENT As Long
Public YYYYMM As String

Public LOCATION_GROUP_NO As String

Public LOCATION_GROUP_CODE  As String
Public LOCATION_GROUP_NAME  As String

Public STOCK_GROUP_CODE  As String
Public STOCK_GROUP_NAME  As String

Public STOCK_CODE_ID As Long
Public STOCK_NO As String
Public STOCK_DESC As String
Public STOCK_TYPE As Long
Public STOCK_TYPE_SUB As Long
Public STOCK_TYPE_CODE As String
Public STOCK_TYPE_NAME As String

Public EXPORT_NO As String
Public EXPORT_DATE As Date
Public LotItemLink As Collection
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   Set LotItemLink = New Collection
   
   CANCEL_FLAG = "N"
   FREE_FLAG = "N"
End Sub
Private Sub Class_Terminate()
   Set LotItemLink = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
      
   SQL = "DELETE FROM LOT_ITEM_LINK WHERE EXPORT_LOT_ITEM_ID  = " & LOT_ITEM_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM LOT_ITEM WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long, MustCount As Boolean)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
   
   If Ind = 1 Then  ' NORMAL USE
      SelectStr1 = "SELECT EN.*, SC.*, MR1.KEY_NAME LOCATION_NAME, MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_TRAN_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (EN.UNIT_TRAN_ID = MR3.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then   ' FOR LOAD TO FRMPROCESSCOMMIT
      SelectStr1 = "SELECT EN.*,IND.DOCUMENT_DATE,IND.DOCUMENT_NO,IND.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DOP ON (IND.DOCUMENT_TYPE = DOP.DOCUMENT_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE  (SC.CHK_STD_COST = 'N') "
      
   ElseIf Ind = 5 Then ' FOR LOAD TO CREPORTINVENTORYDOC2
      SelectStr1 = "SELECT EN.*,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.CHK_STD_COST "
      SelectStr1 = SelectStr1 & ",PT.KEY_CODE PART_TYPE_CODE,PT.KEY_NAME PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",EN.CREATE_DATE,EN.MODIFY_DATE "
            
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DOP ON (IND.DOCUMENT_TYPE = DOP.DOCUMENT_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (SC.STOCK_TYPE = PT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.UNIT_CHANGE_ID = MR4.KEY_ID) "
   ElseIf Ind = 6 Then           'FrmAddLotItem
      SelectStr1 = "SELECT IND.DOCUMENT_NO,IND.DOCUMENT_DATE,EN.LOT_ITEM_ID,EN.LOCATION_ID,SC.UNIT_CHANGE_ID UNIT_ID "
      SelectStr1 = SelectStr1 & ",EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT LOT_ITEM_AMOUNT,SUM(LK.IMPORT_AMOUNT) TX_AMOUNT,EN.AVG_PRICE "
      SelectStr1 = SelectStr1 & ",IND.INVENTORY_DOC_ID,IND.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 "
      WhereStr = WhereStr & " AND ((EN.TX_AMOUNT > (SELECT SUM(LK1.IMPORT_AMOUNT)  FROM LOT_ITEM_LINK LK1 "
      WhereStr = WhereStr & " WHERE EN.LOT_ITEM_ID =  LK1.IMPORT_LOT_ITEM_ID)) "
      WhereStr = WhereStr & " OR (SELECT COUNT(*) FROM LOT_ITEM_LINK LK2  "
      WhereStr = WhereStr & " WHERE  EN.LOT_ITEM_ID =  LK2.IMPORT_LOT_ITEM_ID) = 0 ) "
   ElseIf Ind = 7 Then ' FOR LOAD TO CREPORTINVENTORYDOC2_1
      SelectStr1 = "SELECT EN.*,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",PT.KEY_CODE PART_TYPE_CODE,PT.KEY_NAME PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",NM2.LONG_NAME APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DOP ON (IND.DOCUMENT_TYPE = DOP.DOCUMENT_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (SC.STOCK_TYPE = PT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.UNIT_CHANGE_ID = MR4.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (IND.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN1 ON (APM.APAR_MAS_ID = EN1.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN1.NAME_ID = NM2.NAME_ID) "
      
   ElseIf Ind = 8 Then ' GetNextLotItemID
      SelectStr1 = "SELECT EN.LOT_ITEM_ID,LK.IMPORT_LOT_ITEM_ID,LK.MAIN_IMPORT_LOT_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = EN.LOT_ITEM_ID) "
      
      WhereStr = " WHERE LK.IMPORT_LOT_ITEM_ID IS NOT NULL AND EN.TX_TYPE = 'E' "
   ElseIf Ind = 9 Then '     CReportInventory006_2
      SelectStr1 = "SELECT  DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE PART_TYPE_CODE,MR2.KEY_NAME PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR5.KEY_CODE PART_GROUP_CODE,MR5.KEY_NAME PART_GROUP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.UNIT_CHANGE_ID = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR2.PARENT_ID = MR5.KEY_ID) "
   ElseIf Ind = 10 Then 'GetSumMovementPartItemType จิวแก้เพิ่มให้แบ่งตามฝากขายด้วย 27-08-2555 09:51:00
      SelectStr1 = "SELECT EN.PART_ITEM_ID, EN.TX_TYPE, IND.DOCUMENT_TYPE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",IND.TEMP_SALE_FLAG "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 12 Then                'GetSumTransferPartItemDocDate
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE (EN.LINK_ID = LT1.LINK_ID) AND EN.TX_TYPE = 'I' AND LT1.TX_TYPE = 'E' "
      WhereStr = WhereStr & " AND EN.LOCATION_ID = " & TO_LOCATION_ID
      
      If FROM_LOCATION_ID > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         WhereStr = WhereStr & TempStr & " (LT1.LOCATION_ID = " & FROM_LOCATION_ID & ") " & vbCrLf
      End If
   ElseIf Ind = 13 Then          'GetDistinctTransferPartItem
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE (EN.LINK_ID = LT1.LINK_ID) AND EN.TX_TYPE = 'I' AND LT1.TX_TYPE = 'E' "
      WhereStr = WhereStr & " AND EN.LOCATION_ID = " & TO_LOCATION_ID
      
      If FROM_LOCATION_ID > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         WhereStr = WhereStr & TempStr & " (LT1.LOCATION_ID = " & FROM_LOCATION_ID & ") " & vbCrLf
      End If
   ElseIf Ind = 14 Then       'GetDistinctPartItemByDocTypeDocSubType
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 15 Then    'GetSumAmountByPartItemDate
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 16 Then             'CReportInventoryDoc5,CReportInventoryDoc7
      SelectStr1 = "SELECT DISTINCT IND.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
   ElseIf Ind = 17 Then             'CReportInventoryDoc6
      SelectStr1 = "SELECT DISTINCT IND.DOCUMENT_DATE,IND.INVENTORY_SUB_TYPE,MR1.KEY_CODE INVENTORY_SUB_TYPE_CODE,MR1.KEY_NAME INVENTORY_SUB_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (IND.INVENTORY_SUB_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 18 Then    'GetSumAmountByPartItemDateIndSub
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE,IND.INVENTORY_SUB_TYPE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 19 Then     'Patch_2007_09_22_1_jill
      SelectStr1 = "SELECT IND.DOCUMENT_DATE,IND.DOCUMENT_NO,EN.PART_ITEM_ID,EN.LOT_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
   ElseIf Ind = 20 Then    'สร้าง BalanceAccum โดยใช้ LotItem แทน CReportInventoryDoc1_1 ,CReportInventoryDoc1_2,LoadLeftAmountLotItem
      SelectStr1 = "SELECT SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr1 = SelectStr1 & ",EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.CHK_STD_COST ,SC.COST_PER_AMOUNT "
      SelectStr1 = SelectStr1 & ",LC.KEY_ID LOCATION_ID,LC.KEY_CODE LOCATION_NO,LC.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",PT.KEY_CODE PART_TYPE_CODE,PT.KEY_NAME PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",UN.KEY_NAME UNIT_CHANGE_NAME "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (LC.KEY_ID = EN.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = EN.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (PT.KEY_ID = SC.STOCK_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF UN ON (SC.UNIT_CHANGE_ID = UN.KEY_ID) "
      
   ElseIf Ind = 22 Then           'AdjustLotItemLink
      SelectStr1 = "SELECT  EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,EN.LOCATION_ID "
      SelectStr1 = SelectStr1 & ",IND.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT,EN.LOT_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'E' AND LK.IMPORT_LOT_ITEM_ID IS NULL "
      
   ElseIf Ind = 23 Then       'GetDocItemIDLinkLotItemID
      SelectStr1 = "SELECT  EN.LOT_ITEM_ID,DI.DOC_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = " WHERE  (DI.LINK_ID = EN.LINK_ID) AND (IND.DOCUMENT_TYPE IN ('10','21','30','31')) "
      
   ElseIf Ind = 24 Then           'GetBalanceByLotItemLinkDate
      SelectStr1 = "SELECT  IND.DOCUMENT_DATE,EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT LOT_ITEM_AMOUNT,SUM(LK.IMPORT_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",EN.AVG_PRICE,LC.KEY_CODE LOCATION_NO "
      SelectStr1 = SelectStr1 & ",PT.KEY_CODE PART_TYPE_CODE,PT.KEY_NAME PART_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (LI.LOT_ITEM_ID = LK.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (LC.KEY_ID = EN.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (PT.KEY_ID = SC.STOCK_TYPE) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 "
      WhereStr = WhereStr & " AND ((EN.TX_AMOUNT > (SELECT SUM(LK1.IMPORT_AMOUNT)  FROM LOT_ITEM_LINK LK1 "
      WhereStr = WhereStr & " WHERE EN.LOT_ITEM_ID =  LK1.IMPORT_LOT_ITEM_ID)) "
      WhereStr = WhereStr & " OR (SELECT COUNT(*) FROM LOT_ITEM_LINK LK2  "
      WhereStr = WhereStr & " WHERE  EN.LOT_ITEM_ID =  LK2.IMPORT_LOT_ITEM_ID) = 0 ) "
   ElseIf Ind = 25 Then           'frmAdjustLotItemLink
      SelectStr1 = "SELECT EN.LOT_ITEM_ID "
      SelectStr1 = SelectStr1 & ",EN.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT LOT_ITEM_AMOUNT,SUM(LK.IMPORT_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",IND.INVENTORY_DOC_ID,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 "
      WhereStr = WhereStr & " AND ((EN.TX_AMOUNT > (SELECT SUM(LK1.IMPORT_AMOUNT)  FROM LOT_ITEM_LINK LK1 "
      WhereStr = WhereStr & " WHERE EN.LOT_ITEM_ID =  LK1.IMPORT_LOT_ITEM_ID)) "
      WhereStr = WhereStr & " OR (SELECT COUNT(*) FROM LOT_ITEM_LINK LK2  "
      WhereStr = WhereStr & " WHERE  EN.LOT_ITEM_ID =  LK2.IMPORT_LOT_ITEM_ID) = 0 ) "
   ElseIf Ind = 26 Then           'CReportInventoryDoc1_3
      SelectStr1 = "SELECT IND.DOCUMENT_NO,IND.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",LC.KEY_CODE LOCATION_NO,LC.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT LOT_ITEM_AMOUNT,SUM(LK.IMPORT_AMOUNT) TX_AMOUNT,EN.AVG_PRICE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (LC.KEY_ID = EN.LOCATION_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 "
      WhereStr = WhereStr & " AND ((EN.TX_AMOUNT > (SELECT SUM(LK1.IMPORT_AMOUNT)  FROM LOT_ITEM_LINK LK1 "
      WhereStr = WhereStr & " WHERE EN.LOT_ITEM_ID =  LK1.IMPORT_LOT_ITEM_ID)) "
      WhereStr = WhereStr & " OR (SELECT COUNT(*) FROM LOT_ITEM_LINK LK2  "
      WhereStr = WhereStr & " WHERE  EN.LOT_ITEM_ID =  LK2.IMPORT_LOT_ITEM_ID) = 0 ) "
   ElseIf Ind = 27 Then           'CReportInventoryDoc1_4
      SelectStr1 = "SELECT IND.DOCUMENT_NO,IND.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",IND1.DOCUMENT_NO EXPORT_NO,IND1.DOCUMENT_DATE EXPORT_DATE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",LC.KEY_CODE LOCATION_NO,LC.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",EN.TX_AMOUNT SUM_AMOUNT "
      SelectStr1 = SelectStr1 & ",LI.TX_AMOUNT LOT_ITEM_AMOUNT,LK.IMPORT_AMOUNT TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (EN.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
            
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (LI.LOT_ITEM_ID = LK.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND1 ON (LI.INVENTORY_DOC_ID = IND1.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (LC.KEY_ID = EN.LOCATION_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 "
   ElseIf Ind = 28 Then
      SelectStr1 = "SELECT EN.*,IND.DOCUMENT_DATE,IND.DOCUMENT_NO,MR1.KEY_NAME INVENTORY_SUB_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (IND.INVENTORY_SUB_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.UNIT_CHANGE_ID = MR4.KEY_ID) "
   ElseIf Ind = 29 Then    'เช็คว่ามียอดเพียงพอหรือไม่
      SelectStr1 = "SELECT SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      
      FromStr = "FROM LOT_ITEM EN "
      
      If LOT_ITEM_ID > 0 Then
         WhereStr = " WHERE (EN.LOT_ITEM_ID <> " & LOT_ITEM_ID & ") " & vbCrLf
         LOT_ITEM_ID = -1
      End If
   ElseIf Ind = 30 Then    'LoadLeftAmount         จะไม่เพิ่มการกรุ๊ปเพื่อให้เร็วในการ Query
      SelectStr1 = "SELECT EN.PART_ITEM_ID, SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
   
   ElseIf Ind = 31 Then 'GetLotItemPartTxType
      SelectStr1 = "SELECT EN.PART_ITEM_ID, EN.TX_TYPE,  SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 32 Then      'LoadLeftAmountLocation         จะไม่เพิ่มการกรุ๊ปเพื่อให้เร็วในการ Query
       SelectStr1 = "SELECT EN.PART_ITEM_ID,EN.LOCATION_ID, SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
   ElseIf Ind = 33 Then    'สำหรับคิดต้นทุน ณ วันที่นั้นๆ ขอ วัตถุดิบ ที่คิดต้นทุนมาตรฐาน
      SelectStr1 = "SELECT EN.PART_ITEM_ID,EN.LOCATION_ID,EN.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'E' "
   ElseIf Ind = 34 Then       'GetSumAmountByPartItemIndSub
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.INVENTORY_SUB_TYPE, SUM(EN.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 36 Then       'GetDistinctPartItemByProduction
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 AND IND.BARCODE_JOB_FLAG = 'Y' "     'เอาเฉพาะยอดที่รับจากผลิต
   ElseIf Ind = 37 Then    'GetSumAmountByPartItemDateProduction
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 AND IND.BARCODE_JOB_FLAG = 'Y' "     'เอาเฉพาะยอดที่รับจากผลิต
   ElseIf Ind = 38 Then             'CReportInventoryDoc11
      SelectStr1 = "SELECT DISTINCT IND.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' AND EN.TX_AMOUNT > 0 AND IND.BARCODE_JOB_FLAG = 'Y' "     'เอาเฉพาะยอดที่รับจากผลิต
   ElseIf Ind = 39 Or Ind = 40 Then     'CReportInventoryDoc3_2, CReportInventoryDoc3_3
      SelectStr1 = "SELECT DISTINCT EN.LOCATION_ID LOCATION_NO, MR1.KEY_NAME LOCATION_NAME,MR1.KEY_CODE LOCATION_CODE, SC.STOCK_CODE_ID STOCK_CODE_ID, SC.STOCK_NO STOCK_NO, SC.STOCK_DESC STOCK_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
           
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
   ElseIf Ind = 41 Then 'GetLotItemPartTxType2 for CReportInventoryDoc3_2, CReportInventoryDoc3_3
      SelectStr1 = "SELECT EN.PART_ITEM_ID,EN.LOCATION_ID LOCATION_NO , EN.TX_TYPE,  SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_ID = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR.PARENT_ID = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 42 Then    'CReportInventoryDoc2_2
      SelectStr1 = "SELECT IND.DOCUMENT_DATE , EN.PART_ITEM_ID, EN.TX_TYPE, IND.DOCUMENT_TYPE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",IND.TEMP_SALE_FLAG "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 43 Then    'CReportInventoryDoc2_2
      SelectStr1 = "SELECT DISTINCT IND.DOCUMENT_DATE , EN.PART_ITEM_ID , SC.CHK_STD_COST , SC.STOCK_NO , SC.STOCK_DESC , SC.COST_PER_AMOUNT  "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.UNIT_CHANGE_ID = MR4.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 44 Then
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & " INNER JOIN "
      FromStr = FromStr & " ( "
      FromStr = FromStr & " SELECT DISTINCT IND_T.INVENTORY_DOC_ID "
      FromStr = FromStr & " FROM LOT_ITEM EN_T "
      FromStr = FromStr & " LEFT OUTER JOIN INVENTORY_DOC IND_T ON (IND_T.INVENTORY_DOC_ID = EN_T.INVENTORY_DOC_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN STOCK_CODE   SC_T ON (SC_T.STOCK_CODE_ID = EN_T.PART_ITEM_ID) "
      FromStr = FromStr & " WHERE EN_T.TX_TYPE = 'E' "
      
      If Len(Trim(FROM_STOCK_NO)) > 0 Then
         FromStr = FromStr & " AND " & " (SC_T.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
         FROM_STOCK_NO = ""
      End If
      If Len(Trim(TO_STOCK_NO)) > 0 Then
         FromStr = FromStr & " AND " & " (SC_T.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
         TO_STOCK_NO = ""
      End If
      
      FromStr = FromStr & " ) "
      FromStr = FromStr & " AS P ON (IND.INVENTORY_DOC_ID = P.INVENTORY_DOC_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'I' "
   ElseIf Ind = 45 Then
      SelectStr1 = "SELECT DISTINCT IND.INVENTORY_DOC_ID,IND.DOCUMENT_NO,IND.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'E' "
   ElseIf Ind = 46 Then       'GetDistinctPartInputByInput
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE EN.TX_TYPE = 'E'  "
   ElseIf Ind = 47 Then    'GetSumAmountByInputOutput
      SelectStr1 = "SELECT EN.PART_ITEM_ID,EN.TX_TYPE,IND.INVENTORY_DOC_ID, SUM(EN.TX_AMOUNT) TX_AMOUNT  "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   
      WhereStr = ""
   ElseIf Ind = 48 Then          'GetDistinctTransferPartItemConsignment
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,IND.DOCUMENT_DATE,IND.DOCUMENT_NO,SC.STOCK_DESC PART_DESC,SC.STOCK_NO PART_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (LT1.LOCATION_ID = MR2.KEY_ID) "
      
      WhereStr = " WHERE (EN.LINK_ID = LT1.LINK_ID) AND EN.TX_TYPE = 'I' AND LT1.TX_TYPE = 'E' "
      If CONSIGNMENT > 0 Then
         WhereStr = WhereStr & " AND IND.TEMP_SALE_FLAG = 'Y' "            'ฝากขาย
      End If
      If Len(Trim(TO_LOCATION_CODE)) > 0 Or Len(Trim(TO_LOCATION_CODE2)) > 0 Or Len(Trim(FROM_LOCATION_CODE)) > 0 Or Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
         WhereStr = WhereStr & " AND (  "
          If Len(Trim(TO_LOCATION_CODE)) > 0 Then
            WhereStr = WhereStr & " MR1.KEY_CODE >= " & TO_LOCATION_CODE
            If Len(Trim(TO_LOCATION_CODE2)) > 0 Or Len(Trim(FROM_LOCATION_CODE)) > 0 Or Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(TO_LOCATION_CODE2)) > 0 Then
            WhereStr = WhereStr & " MR1.KEY_CODE <= " & TO_LOCATION_CODE2
            If Len(Trim(FROM_LOCATION_CODE)) > 0 Or Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(FROM_LOCATION_CODE)) > 0 Then
            WhereStr = WhereStr & " MR2.KEY_CODE >= " & FROM_LOCATION_CODE
            If Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
            WhereStr = WhereStr & " MR2.KEY_CODE <= " & FROM_LOCATION_CODE2
         End If
         WhereStr = WhereStr & " OR  "
         If Len(Trim(FROM_LOCATION_CODE)) > 0 Then
            WhereStr = WhereStr & " MR1.KEY_CODE >= " & FROM_LOCATION_CODE
            If Len(Trim(FROM_LOCATION_CODE2)) > 0 Or Len(Trim(TO_LOCATION_CODE)) > 0 Or Len(Trim(TO_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
            WhereStr = WhereStr & " MR1.KEY_CODE <= " & FROM_LOCATION_CODE2
            If Len(Trim(TO_LOCATION_CODE)) > 0 Or Len(Trim(TO_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(TO_LOCATION_CODE)) > 0 Then
            WhereStr = WhereStr & " MR2.KEY_CODE >= " & TO_LOCATION_CODE
            If Len(Trim(TO_LOCATION_CODE2)) > 0 Then
               WhereStr = WhereStr & " AND  "
            End If
         End If
         If Len(Trim(TO_LOCATION_CODE2)) > 0 Then
            WhereStr = WhereStr & " MR2.KEY_CODE <= " & TO_LOCATION_CODE2
         End If
         WhereStr = WhereStr & " )"
      End If
   ElseIf Ind = 49 Then                'GetTransferPartItemDocDateConsignment
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE,IND.DOCUMENT_NO, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (LT1.LOCATION_ID = MR2.KEY_ID) "
      
      WhereStr = " WHERE (EN.LINK_ID = LT1.LINK_ID) AND EN.TX_TYPE = 'I' AND LT1.TX_TYPE = 'E' "
      If CONSIGNMENT > 0 Then
         WhereStr = WhereStr & " AND IND.TEMP_SALE_FLAG = 'Y' "            'ฝากขาย
      End If
      If Len(Trim(TO_LOCATION_CODE)) > 0 Then
         WhereStr = WhereStr & " AND MR1.KEY_CODE >= " & TO_LOCATION_CODE
      End If
      If Len(Trim(TO_LOCATION_CODE2)) > 0 Then
         WhereStr = WhereStr & " AND MR1.KEY_CODE <= " & TO_LOCATION_CODE2
      End If
      If Len(Trim(FROM_LOCATION_CODE)) > 0 Then
         WhereStr = WhereStr & " AND MR2.KEY_CODE >= " & FROM_LOCATION_CODE
      End If
      If Len(Trim(FROM_LOCATION_CODE2)) > 0 Then
         WhereStr = WhereStr & " AND MR2.KEY_CODE <= " & FROM_LOCATION_CODE2
      End If
   ElseIf Ind = 50 Then      'frmAdjustStockAmount ,frmAddEditBalanceVerify,frmProcessEndYear
      SelectStr1 = "SELECT EN.PART_ITEM_ID,EN.LOCATION_ID, SUM(EN.SUM_AMOUNT) SUM_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
''   ElseIf Ind = 51 Then 'LoadDistinctTripFromExportType
'      ElseIf Ind = 51 Then    'GetSumAmountByPartItemMonthly
'      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.DOCUMENT_DATE, SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE,IND.YYYYMM  "
'      SelectStr2 = "SELECT COUNT(*)  "
'
'      FromStr = "FROM LOT_ITEM EN "
'      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
'
'    ElseIf Ind = 52 Then              'CReportInventoryDoc5,CReportInventoryDoc7
'      SelectStr1 = "SELECT DISTINCT IND.DOCUMENT_DATE,IND.YYYYMM,SC.BILL_DESC ,SC.STOCK_NO PART_NO, EN.PART_ITEM_ID  "
'      SelectStr2 = "SELECT COUNT(*)  "
'
'      FromStr = "FROM LOT_ITEM EN "
'      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
'
      
   ElseIf Ind = 51 Then    'GetSumAmountByPartItemDate : CReportInventoryDoc6_3
      SelectStr1 = "SELECT EN.PART_ITEM_ID,IND.YYYYMM,SUM(EN.TX_AMOUNT) TX_AMOUNT, SUM(EN.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 52 Then             'CReportInventoryDoc6_3
      SelectStr1 = "SELECT DISTINCT IND.YYYYMM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (EN.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 53 Then              'CReportInventoryDoc3_5_1
      SelectStr1 = "SELECT DISTINCT EN.PART_ITEM_ID,MR2.KEY_CODE LOCATION_GROUP_CODE,MR2.KEY_NAME LOCATION_GROUP_NAME,MR4.KEY_CODE LOCATION_CODE,MR4.KEY_NAME LOCATION_NAME,SC.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr2 = "SELECT COUNT(*) "

      FromStr = "FROM LOT_ITEM EN "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (SC.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (EN.LOCATION_ID = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR4.PARENT_ID = MR2.KEY_ID) "
      WhereStr = ""
   End If
   
   ' EN - LOT_ITEM
   ' MR1 - LOCATION
   'PT - PART_TYPE
   'SC STOCK_TYPE
   'IND - INVENTORY_DOC
   'DOP - DOCUMENT_PRIORITY
   'PG -PART_GROUP
   
   If LOT_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.LOT_ITEM_ID = " & LOT_ITEM_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(OUTLAY_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.OUTLAY_FLAG = '" & ChangeQuote(Trim(OUTLAY_FLAG)) & "')"
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.PART_ITEM_ID = " & PART_ITEM_ID & ") " & vbCrLf
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If INVENTORY_SUB_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
  '  WhereStr = WhereStr & TempStr & " (IND.INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ") " & vbCrLf
      If INVENTORY_SUB_TYPE = 999999999 Then
         WhereStr = WhereStr & TempStr & " (IND.INVENTORY_SUB_TYPE   is  null  ) " & vbCrLf       '------------------------แก้เอกสารการเบิก ST007
      Else
        WhereStr = WhereStr & TempStr & " (IND.INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ") " & vbCrLf
      End If
      
   End If
   
    If Len(LOCATION_CODE) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR4.KEY_CODE = '" & LOCATION_CODE & "') " & vbCrLf
   End If
   
   If Len(LOCATION_GROUP_CODE) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR2.KEY_CODE = '" & LOCATION_GROUP_CODE & "') " & vbCrLf
   End If
   
   If FROM_DOC_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If Len(Trim(PART_NO)) > 0 Then
      Call AddSearchStockCode(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PART_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (SC.STOCK_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
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

   If Len(Trim(PART_DESC)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PART_DESC, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(SC.STOCK_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (SC.STOCK_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
      End If
   End If
   
   If PART_TYPE > 0 Then
      Call AddSearchPartType(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PT.KEY_ID = " & PART_TYPE & ") " & vbCrLf
   End If
   
   If PART_GROUP > 0 Then
      Call AddSearchPartGroup(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PG.KEY_ID = " & PART_GROUP & ") " & vbCrLf
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.LOCATION_ID = " & LOCATION_ID & ") " & vbCrLf
   End If
   
   If Len(LOCATION_GROUP_NO) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR2.KEY_CODE = " & LOCATION_GROUP_NO & ") " & vbCrLf
   End If
   
   
   If Len(FROM_LOCATION) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR1.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR1.KEY_CODE >= '" & Trim(ChangeQuote(FROM_LOCATION)) & "') "
      End If
   End If
   
   If Len(TO_LOCATION) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR1.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION)) & "') "
      Else
         WhereStr = WhereStr & "AND (MR1.KEY_CODE <= '" & Trim(ChangeQuote(TO_LOCATION)) & "') "
      End If
   End If

   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EN.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (EN.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(COUNT_AMOUNT)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(COUNT_AMOUNT, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(IND.COUNT_AMOUNT,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (IND.COUNT_AMOUNT = '" & ChangeQuote(Trim(COUNT_AMOUNT)) & "')"
      End If
   End If
   
   If Len(Trim(CHK_STD_COST)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.CHK_STD_COST = '" & ChangeQuote(Trim(CHK_STD_COST)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 6 Then
      GroupByStr = " GROUP BY IND.DOCUMENT_NO,IND.DOCUMENT_DATE,EN.LOT_ITEM_ID,EN.PART_ITEM_ID,SC.STOCK_NO,SC.STOCK_DESC "
      GroupByStr = GroupByStr & ",EN.TX_AMOUNT,EN.LOCATION_ID,SC.UNIT_CHANGE_ID,IND.INVENTORY_DOC_ID,IND.DOCUMENT_TYPE,EN.AVG_PRICE  "
   ElseIf Ind = 10 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID, EN.TX_TYPE, IND.DOCUMENT_TYPE,IND.TEMP_SALE_FLAG "
   ElseIf Ind = 12 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.DOCUMENT_DATE "
   ElseIf Ind = 15 Or Ind = 37 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.DOCUMENT_DATE "
   ElseIf Ind = 18 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.DOCUMENT_DATE,IND.INVENTORY_SUB_TYPE "
   ElseIf Ind = 20 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,SC.STOCK_NO,SC.STOCK_DESC,LC.KEY_ID,LC.KEY_CODE,LC.KEY_NAME "
      GroupByStr = GroupByStr & ",PT.KEY_CODE,PT.KEY_NAME,UN.KEY_NAME,SC.CHK_STD_COST ,SC.COST_PER_AMOUNT "
   ElseIf Ind = 24 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,SC.STOCK_NO,SC.STOCK_DESC,PT.KEY_CODE,PT.KEY_NAME "
      GroupByStr = GroupByStr & ",EN.TX_AMOUNT,EN.AVG_PRICE,LC.KEY_CODE,IND.DOCUMENT_DATE "
   ElseIf Ind = 25 Then
      GroupByStr = " GROUP BY EN.LOT_ITEM_ID,EN.PART_ITEM_ID "
      GroupByStr = GroupByStr & ",EN.TX_AMOUNT,IND.INVENTORY_DOC_ID,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 26 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,LC.KEY_CODE,LC.KEY_NAME,EN.AVG_PRICE "
      GroupByStr = GroupByStr & ",EN.TX_AMOUNT,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 30 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID "
   ElseIf Ind = 31 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,EN.TX_TYPE "
   ElseIf Ind = 32 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,EN.LOCATION_ID "
   ElseIf Ind = 34 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.INVENTORY_SUB_TYPE "
   ElseIf Ind = 41 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,EN.LOCATION_ID,EN.TX_TYPE "
   ElseIf Ind = 42 Then
      GroupByStr = " GROUP BY IND.DOCUMENT_DATE,EN.PART_ITEM_ID, EN.TX_TYPE, IND.DOCUMENT_TYPE,IND.TEMP_SALE_FLAG "
   ElseIf Ind = 47 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID , EN.TX_TYPE, IND.INVENTORY_DOC_ID "
   ElseIf Ind = 49 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 50 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,SC.STOCK_NO,EN.LOCATION_ID "
      
    ElseIf Ind = 51 Then
      GroupByStr = " GROUP BY EN.PART_ITEM_ID,IND.YYYYMM  "
   
   
   
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EN.LOT_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EN.LOT_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
         OrderByStr = "ORDER BY IND.DOCUMENT_DATE ASC,DOP.PRIORITY1 ASC,IND.DOCUMENT_NO ASC,EN.TX_TYPE "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PT.KEY_CODE ASC,SC.STOCK_NO ASC,IND.DOCUMENT_DATE ASC,DOP.PRIORITY1 ASC,IND.DOCUMENT_NO ASC,EN.TX_TYPE "
      If ORDER_BY = 1 Then
         OrderByStr = OrderByStr & ",SC.STOCK_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = OrderByStr & ",SC.STOCK_DESC "
      End If
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE ASC,IND.DOCUMENT_NO ASC,SC.STOCK_NO "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PT.KEY_CODE ASC,SC.STOCK_NO ASC,IND.DOCUMENT_DATE ASC,DOP.PRIORITY1 ASC,IND.DOCUMENT_NO ASC,EN.TX_TYPE "
      If ORDER_BY = 1 Then
         OrderByStr = OrderByStr & ",SC.STOCK_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = OrderByStr & ",SC.STOCK_DESC "
      End If
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY EN.LOT_ITEM_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY MR5.KEY_CODE,MR2.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 13 Or Ind = 44 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 14 Or Ind = 46 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 16 Or Ind = 38 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE ASC,MR1.KEY_CODE "
   ElseIf Ind = 19 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE ASC,IND.DOCUMENT_NO ASC,EN.PART_ITEM_ID "
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY PT.KEY_CODE,SC.STOCK_NO,LC.KEY_CODE "
   ElseIf Ind = 22 Then
      OrderByStr = "ORDER BY EN.LOCATION_ID,SC.STOCK_NO,IND.DOCUMENT_DATE "
   ElseIf Ind = 24 Then
      OrderByStr = "ORDER BY PT.KEY_CODE ASC,SC.STOCK_NO ASC,IND.DOCUMENT_DATE ASC,LC.KEY_CODE"
   ElseIf Ind = 25 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE ASC,IND.DOCUMENT_NO "
   ElseIf Ind = 26 Then
      OrderByStr = "ORDER BY LC.KEY_CODE,SC.STOCK_NO,IND.DOCUMENT_DATE ASC,IND.DOCUMENT_NO "
   ElseIf Ind = 27 Then
      OrderByStr = "ORDER BY LC.KEY_CODE,SC.STOCK_NO,IND.DOCUMENT_DATE ASC,IND.DOCUMENT_NO,IND1.DOCUMENT_DATE,IND1.DOCUMENT_NO "
   ElseIf Ind = 28 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 33 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE "
      ORDER_TYPE = 2
   ElseIf Ind = 36 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 39 Then
      OrderByStr = "ORDER BY EN.LOCATION_ID,SC.STOCK_NO,EN.PART_ITEM_ID "
   ElseIf Ind = 40 Then
      OrderByStr = "ORDER BY EN.PART_ITEM_ID,EN.LOCATION_ID "
   ElseIf Ind = 43 Then
      OrderByStr = "ORDER BY SC.STOCK_NO,IND.DOCUMENT_DATE "
   ElseIf Ind = 45 Then
      OrderByStr = "ORDER BY IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 48 Then
      OrderByStr = "ORDER BY SC.STOCK_NO,IND.DOCUMENT_DATE,IND.DOCUMENT_NO "
   ElseIf Ind = 50 Then
      OrderByStr = "ORDER BY SC.STOCK_NO ASC,EN.LOCATION_ID "
   ElseIf Ind = 52 Then
      OrderByStr = "ORDER BY  IND.YYYYMM "
   ElseIf Ind = 53 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE ASC,MR4.KEY_CODE ASC,SC.STOCK_NO  "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   If Ind = 10 Or Ind = 12 Or Ind = 15 Or Ind = 18 Or Ind = 23 Or Ind = 29 Or Ind = 30 Or Ind = 31 Or Ind = 32 Or Ind = 34 Or Ind = 37 Or Ind = 41 Or Ind = 42 Or Ind = 47 Or Ind = 49 Or Ind = 51 Then
      OrderByStr = ""
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)

End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      ' In table -------------------------------------------------------------
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      MULTIPLIER = NVLI(Rs("MULTIPLIER"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      FREE_FLAG = NVLS(Rs("FREE_FLAG"), "")
      
      UNIT_TRAN_ID = NVLI(Rs("UNIT_TRAN_ID"), -1)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
      
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      NEW_AVG_PRICE = NVLD(Rs("NEW_AVG_PRICE"), 0)
      ' In table -------------------------------------------------------------
      
      PART_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      PART_NO = NVLS(Rs("STOCK_NO"), "")
      PART_DESC = NVLS(Rs("STOCK_DESC"), "")
      PART_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
   ElseIf Ind = 2 Then
      ' In table -------------------------------------------------------------
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      MULTIPLIER = NVLI(Rs("MULTIPLIER"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
      NEW_AVG_PRICE = NVLD(Rs("NEW_AVG_PRICE"), 0)
      ' In table -------------------------------------------------------------
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 3 Then
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   ElseIf Ind = 5 Then
      ' In table -------------------------------------------------------------
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      MULTIPLIER = NVLI(Rs("MULTIPLIER"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      CREATE_DATE = InternalDateToDate(NVLS(Rs("CREATE_DATE"), ""))
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
      ' In table -------------------------------------------------------------
      
      NEW_AVG_PRICE = NVLD(Rs("NEW_AVG_PRICE"), 0)
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      
   ElseIf Ind = 6 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 7 Then
      ' In table -------------------------------------------------------------
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      MULTIPLIER = NVLI(Rs("MULTIPLIER"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      ' In table -------------------------------------------------------------
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
   ElseIf Ind = 8 Then
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      IMPORT_LOT_ITEM_ID = NVLI(Rs("IMPORT_LOT_ITEM_ID"), -1)
      MAIN_IMPORT_LOT_ITEM_ID = NVLI(Rs("MAIN_IMPORT_LOT_ITEM_ID"), -1)
   ElseIf Ind = 9 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_GROUP_CODE = NVLS(Rs("PART_GROUP_CODE"), "")
      PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TEMP_SALE_FLAG = NVLS(Rs("TEMP_SALE_FLAG"), "")
      
   ElseIf Ind = 16 Or Ind = 38 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 12 Or Ind = 15 Or Ind = 37 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 13 Or Ind = 14 Or Ind = 36 Or Ind = 44 Or Ind = 46 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
   ElseIf Ind = 17 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      INVENTORY_SUB_TYPE = NVLI(Rs("INVENTORY_SUB_TYPE"), -1)
      INVENTORY_SUB_TYPE_CODE = NVLS(Rs("INVENTORY_SUB_TYPE_CODE"), "")
      INVENTORY_SUB_TYPE_NAME = NVLS(Rs("INVENTORY_SUB_TYPE_NAME"), "")
   ElseIf Ind = 18 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      INVENTORY_SUB_TYPE = NVLI(Rs("INVENTORY_SUB_TYPE"), -1)
   ElseIf Ind = 19 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
   ElseIf Ind = 20 Then
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 22 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 23 Then
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      DOC_ITEM_ID = NVLI(Rs("DOC_ITEM_ID"), -1)
   ElseIf Ind = 24 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      PART_TYPE_CODE = NVLS(Rs("PART_TYPE_CODE"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 25 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 26 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 27 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      
      EXPORT_NO = NVLS(Rs("EXPORT_NO"), "")
      EXPORT_DATE = InternalDateToDate(NVLS(Rs("EXPORT_DATE"), ""))
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)       ' ยอดจากจำนวนใบรับเข้า
      
      LOT_ITEM_AMOUNT = NVLD(Rs("LOT_ITEM_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 28 Then
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      INVENTORY_SUB_TYPE_NAME = NVLS(Rs("INVENTORY_SUB_TYPE_NAME"), "")
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 29 Then
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
   ElseIf Ind = 30 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
   ElseIf Ind = 31 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
   ElseIf Ind = 32 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
   ElseIf Ind = 33 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 34 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      INVENTORY_SUB_TYPE = NVLI(Rs("INVENTORY_SUB_TYPE"), -1)
   ElseIf Ind = 39 Or Ind = 40 Then
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), -1)
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_CODE = NVLS(Rs("LOCATION_CODE"), "")
      PART_CODE = NVLS(Rs("STOCK_CODE_ID"), "")
      PART_NO = NVLS(Rs("STOCK_NO"), "")
      PART_DESC = NVLS(Rs("STOCK_DESC"), "")
   ElseIf Ind = 41 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
   ElseIf Ind = 42 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TEMP_SALE_FLAG = NVLS(Rs("TEMP_SALE_FLAG"), "")
   ElseIf Ind = 43 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      PART_NO = NVLS(Rs("STOCK_NO"), "")
      PART_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 45 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   ElseIf Ind = 47 Then
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
   ElseIf Ind = 48 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
  ElseIf Ind = 49 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 50 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      SUM_AMOUNT = NVLD(Rs("SUM_AMOUNT"), 0)
      
    ElseIf Ind = 51 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
'      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      
    ElseIf Ind = 52 Then
'      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      YYYYMM = NVLS(Rs("YYYYMM"), "")
'     BILL_DESC = NVLS(Rs("BILL_DESC"), "")
'      PART_NO = NVLS(Rs("PART_NO"), "")
'      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   ElseIf Ind = 53 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
'
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      
      CHK_STD_COST = NVLS(Rs("CHK_STD_COST"), "")
      COST_PER_AMOUNT = NVLD(Rs("COST_PER_AMOUNT"), 0)
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      LOCATION_GROUP_CODE = NVLS(Rs("LOCATION_GROUP_CODE"), "")
      LOCATION_GROUP_NAME = NVLS(Rs("LOCATION_GROUP_NAME"), "")
      
      LOCATION_CODE = NVLS(Rs("LOCATION_CODE"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   End If
End Sub

Public Sub AddSearchStockCode(FromStr As String)
   If InStr(1, FromStr, "SC") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (EN.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   End If
End Sub
Public Sub AddSearchPartType(FromStr As String)
   If InStr(1, FromStr, "PT") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (SC.STOCK_TYPE = PT.KEY_ID) "
   End If
End Sub
Public Sub AddSearchPartGroup(FromStr As String)
   If InStr(1, FromStr, "PT") <= 0 Then
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PT ON (SC.STOCK_TYPE = PT.KEY_ID) "
   End If
   If InStr(1, FromStr, "PG") <= 0 Then
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PG ON (PT.PARENT_ID = PG.KEY_ID) "
   End If
End Sub
Public Sub UpDateStdCost(DocumentTypeSet As String, Optional PlusCost As Double = 0, Optional UpDateStockCode As Boolean = False, Optional PlusFlag As Boolean = False)
Dim ServerDateTime As String
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
Dim WhereStr2 As String
Dim TempStr As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If FROM_DOC_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(DocumentTypeSet) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IND.DOCUMENT_TYPE IN " & DocumentTypeSet & ") " & vbCrLf
   End If
   
   If Len(Trim(FROM_STOCK_NO)) > 0 Then
      If Len(WhereStr2) > 0 Then
         TempStr = "AND "
      Else
         TempStr = " "
      End If
      WhereStr2 = WhereStr2 & TempStr & " (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
   End If
   
   If Len(Trim(TO_STOCK_NO)) > 0 Then
      If Len(WhereStr2) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "  "
      End If
      WhereStr2 = WhereStr2 & TempStr & " (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
   End If
   
   If Len(WhereStr2) > 0 Then
      TempStr = "AND "
   Else
      TempStr = " "
   End If
   WhereStr2 = WhereStr2 & TempStr & " (SC.CHK_STD_COST = 'Y')"
   
   If PlusFlag Then
      SQL1 = "UPDATE LOT_ITEM LI SET " & _
                     "   LI.AVG_PRICE = (SELECT SC.COST_PER_AMOUNT FROM STOCK_CODE SC WHERE SC.STOCK_CODE_ID = LI.PART_ITEM_ID) + " & PlusCost & ", " & _
                     "    LI.MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    LI.MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LI.INVENTORY_DOC_ID IN (SELECT IND.INVENTORY_DOC_ID FROM INVENTORY_DOC IND " & WhereStr & ") AND LI.PART_ITEM_ID IN (SELECT SC.STOCK_CODE_ID FROM STOCK_CODE SC WHERE " & WhereStr2 & ")"
                     Call m_ADOConn.Execute(SQL1)
   Else
      SQL1 = "UPDATE LOT_ITEM LI SET " & _
                     "   LI.AVG_PRICE = " & Val(AVG_PRICE) & ", " & _
                    "    LI.MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    LI.MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LI.INVENTORY_DOC_ID IN (SELECT IND.INVENTORY_DOC_ID FROM INVENTORY_DOC IND " & WhereStr & ") AND LI.PART_ITEM_ID IN (SELECT SC.STOCK_CODE_ID FROM STOCK_CODE SC WHERE " & WhereStr2 & ")"
                     Call m_ADOConn.Execute(SQL1)
   End If
   
   '---------------------------------------------------------------------------------------------------------------
   Dim Li As CLotItem
   Dim m_Rs As ADODB.Recordset
   Dim itemcount As Long
   Dim SQL2  As String
   Dim SelectStr1  As String
   Dim FromStr As String
   
   Set Li = New CLotItem
   Set m_Rs = New ADODB.Recordset
   
   SelectStr1 = "SELECT LI.TX_AMOUNT,LI.AVG_PRICE,LI.LOT_ITEM_ID "
   
   FromStr = "FROM LOT_ITEM LI "
   FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IND ON (LI.INVENTORY_DOC_ID = IND.INVENTORY_DOC_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & " AND " & WhereStr2
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, False, itemcount, m_Rs, glbErrorLog)
   
   While Not m_Rs.EOF
      Li.LOT_ITEM_ID = NVLI(m_Rs("LOT_ITEM_ID"), -1)
      Li.TX_AMOUNT = NVLD(m_Rs("TX_AMOUNT"), 0)
      Li.AVG_PRICE = NVLD(m_Rs("AVG_PRICE"), 0)
      
      Li.TOTAL_INCLUDE_PRICE = Li.TX_AMOUNT * Li.AVG_PRICE
      Li.NEW_AVG_PRICE = Li.AVG_PRICE
      
      Li.ShowMode = SHOW_EDIT
      Call Li.UpdateAvg
      
      m_Rs.MoveNext
   Wend
   
   '----------------------------------------------------------------------------------------------------------------
   
   If UpDateStockCode Then
      If PlusFlag Then
         SQL1 = "UPDATE STOCK_CODE SC SET " & _
                     "   SC.COST_PER_AMOUNT = SC.COST_PER_AMOUNT + " & PlusCost & ", " & _
                     "    SC.MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    SC.MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SC.STOCK_CODE_ID IN (SELECT SC.STOCK_CODE_ID FROM STOCK_CODE SC WHERE " & WhereStr2 & ")"
      Else
         SQL1 = "UPDATE STOCK_CODE SC SET " & _
                     "   SC.COST_PER_AMOUNT = " & Val(AVG_PRICE) & ", " & _
                     "    SC.MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    SC.MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SC.STOCK_CODE_ID IN (SELECT SC.STOCK_CODE_ID FROM STOCK_CODE SC WHERE " & WhereStr2 & ")"
      End If
      Call m_ADOConn.Execute(SQL1)
   End If
End Sub
Public Function UpdateAvg()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE LOT_ITEM SET " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "   NEW_AVG_PRICE = " & NEW_AVG_PRICE & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
   
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = UNIT_TRAN_ID
   If UNIT_TRAN_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SUM_AMOUNT = TX_AMOUNT * MULTIPLIER
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LOT_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO LOT_ITEM " & _
                     "( " & _
                     "   LOT_ITEM_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   TX_AMOUNT, " & _
                     "   AVG_PRICE, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   TX_TYPE, " & _
                     "   MULTIPLIER, " & _
                     "   LINK_ID, " & _
                     "   CANCEL_FLAG, " & _
                     "   FREE_FLAG, " & _
                     "   UNIT_TRAN_ID, " & _
                     "   UNIT_MULTIPLE, " & _
                     "   WEIGHT_AMOUNT, " & _
                     "   SUM_AMOUNT, " & _
                     "   NEW_AVG_PRICE, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     "" & ID & "," & _
                     "" & INVENTORY_DOC_ID & "," & _
                     "" & TempID1 & "," & _
                     "" & TempID2 & "," & _
                     "" & TX_AMOUNT & "," & _
                     "" & AVG_PRICE & "," & _
                     "" & TOTAL_INCLUDE_PRICE & "," & _
                     "'" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                     "" & MULTIPLIER & "," & _
                     "" & LINK_ID & "," & _
                     "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & _
                     "'" & ChangeQuote(Trim(FREE_FLAG)) & "'," & _
                     "" & TempID3 & "," & _
                     "" & UNIT_MULTIPLE & "," & _
                     "" & WEIGHT_AMOUNT & "," & _
                     "" & SUM_AMOUNT & "," & _
                     "" & NEW_AVG_PRICE & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         LOT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LOT_ITEM SET " & _
                     "   PART_ITEM_ID = " & TempID1 & ", " & _
                     "   LOCATION_ID = " & TempID2 & ", " & _
                     "   TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "   TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "', " & _
                     "   MULTIPLIER = " & MULTIPLIER & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "', " & _
                     "   FREE_FLAG = '" & ChangeQuote(Trim(FREE_FLAG)) & "', " & _
                     "   UNIT_TRAN_ID = " & TempID3 & ", " & _
                     "   UNIT_MULTIPLE = " & UNIT_MULTIPLE & ", " & _
                     "   WEIGHT_AMOUNT = " & WEIGHT_AMOUNT & ", " & _
                     "   SUM_AMOUNT = " & SUM_AMOUNT & ", " & _
                     "   NEW_AVG_PRICE = " & NEW_AVG_PRICE & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
            
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
   
End Function
