VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobItem"
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

'<------------------------------------------------------------------------------------------------------------------------------------

Public JOB_ITEM_ID As Long
Public JOB_ID As Long
Public BATCH_NO As Long
Public ORDER_NO As Long
Public PART_ITEM_ID As Long
Public TX_TYPE As String
Public TX_AMOUNT As Double
Public TX_AMOUNT_STD As Double
Public TX_WEIGHT As Double
Public TX_WEIGHT_STD As Double
Public LINK_ID As Long

Public LOCATION_ID  As Long         ' คลัง

Public LOST_ID As Long                       ' ID ข้อผลิดพลาด
Public PROBLEM_LIMIT_PERCENT As Double             ' เปอร์เซ็นต์ข้อผิดพลาดมากสุด
Public NEXT_FLAG  As String                  'เป็น Flag ที่บอกว่าจะไปยังการผลิตต่อไปครับ
Public PRODUCTION_TYPE As Long      'ประเภทผลิต

Public EXPORT_LOT_ITEM_ID As Long 'สร้างเพื่อ LINK ไปยัง LOT_ITEM เพื่อความเร็วในการ QUERY ของ frmAddLotItem  (LINK ไปยัง EXPORT_LOT_ITEM_ID)
'------------------------------------------------------------------------------------------------------------------------------------>

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public JOB_NO As String

Public UNIT_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public LOCATION_NO As String
Public LOCATION_NAME As String

Public WEIGHT_AMOUNT  As Double
Public PROBLEM_DESC As String
Public SUM_FLAG As String
Public UNIT_AMOUNT As Double
Public INPUT_ID As Long
Public OUTPUT_ID As Long
Public PRD_LOCATION_ID As Long
Public BILL_DESC  As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public JOB_DATE  As Date
Public BATCH_NO_SET  As String

Public FROM_STOCK_NO As String     'จากวัตถุดิบ
Public TO_STOCK_NO As String           'ถึงวัตถุดิบ
Public FROM_STOCK_NO1 As String     'จากผลิตภัณฑ์
Public TO_STOCK_NO1 As String           'ถึงผลิตภัณฑ์

Public DOCUMENT_DATE As Date           'วันที่ของวัตถุดิบที่นำเข้ามาผลิต
Public DOCUMENT_NO As String              'หมายเลขเอกสารของวัตถุดิบที่นำเข้ามาผลิต
Public FROM_IVD_DATE As Date
Public TO_IVD_DATE As Date

Public PROBLEM_SHORT_CODE As String   ' จาก MasterRef
Public PRD_TYPE_SHORT_CODE As String   ' จาก MasterRef
Public PRD_TYPE_DETAIL_FLAG   As String

Public IMPORT_AMOUNT   As Double ' จาก table LotItemLink
Public MainLocationPrdFlag  As Boolean
Public NoMainLocationPrdFlag  As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
   
   'ถ้าเป็น EXPORT จะใช้ SC
   'ถ้าเป็น IMPORT จะใช้ SC1
   
   If Ind = 1 Then
      SelectStr1 = "SELECT JO.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE LOCATION_NO,MR1.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
      SelectStr1 = SelectStr1 & ",LK.IMPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = JO.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JO.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LT.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID) "
      
   ElseIf Ind = 2 Then                    'CReportProduct001
      SelectStr1 = "SELECT DISTINCT JO.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",JO.BATCH_NO,JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      
   ElseIf Ind = 3 Then     'GetDistinctJobInItem
      SelectStr1 = "SELECT DISTINCT JI.PART_ITEM_ID,SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      WhereStr = " WHERE JI.TX_TYPE = 'I' "
      Call GenerateBatchString(WhereStr)
   ElseIf Ind = 4 Then     'GetSumJobInOutLostItem
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,JI.PART_ITEM_ID INPUT_ID,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID,JO.BATCH_NO "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (JI.JOB_ID = J.JOB_ID) "
      
      WhereStr = " WHERE JO.TX_TYPE =  'E' AND JI.TX_TYPE =  'I' "
   ElseIf Ind = 5 Then     'GetSumJobInItem
      SelectStr1 = "SELECT SUM(JI.TX_AMOUNT) TX_AMOUNT,JI.PART_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      Call GenerateBatchString(WhereStr)
   ElseIf Ind = 6 Then                    'CReportProduct002 ,'CReportProduct003
      SelectStr1 = "SELECT J.JOB_DATE,SUM(JI.TX_AMOUNT) TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      Call GenerateBatchString(WhereStr)
   ElseIf Ind = 7 Then     'GetSumJobOutLostItemByDate
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,J.JOB_DATE,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID,JO.BATCH_NO "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      
      WhereStr = " WHERE JO.TX_TYPE =  'E'  "
      
   ElseIf Ind = 8 Then                'GetDistinctJobOutItem
      SelectStr1 = "SELECT DISTINCT JO.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",JO.BATCH_NO,JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG,MR2.SHORT_CODE PROBLEM_SHORT_CODE "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      
      WhereStr = " WHERE (JO.TX_AMOUNT <> 0 AND JO.TX_TYPE = 'E')  "
   ElseIf Ind = 9 Then                'GetDistinctJobOutItemByPrdType
      SelectStr1 = "SELECT DISTINCT JO.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",JO.BATCH_NO,JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG,MR2.SHORT_CODE PROBLEM_SHORT_CODE "
      SelectStr1 = SelectStr1 & ",JO.PRODUCTION_TYPE,MR3.SHORT_CODE PRD_TYPE_SHORT_CODE "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (MR3.KEY_ID = JO.PRODUCTION_TYPE) "
      
      WhereStr = " WHERE (MR2.MASTER_FLAG <>  'Y' OR MR2.MASTER_FLAG IS NULL) AND (JO.TX_TYPE = 'E' ) AND (JO.TX_AMOUNT <> 0) "
      
   ElseIf Ind = 10 Then                'GetDistinctJobOutItemByPrdTypeEx
      SelectStr1 = "SELECT DISTINCT JO.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG,MR2.SHORT_CODE PROBLEM_SHORT_CODE "
      SelectStr1 = SelectStr1 & ",JO.PRODUCTION_TYPE,MR3.SHORT_CODE PRD_TYPE_SHORT_CODE,MR3.MASTER_FLAG PRD_TYPE_DETAIL_FLAG "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (MR3.KEY_ID = JO.PRODUCTION_TYPE) "
      
      WhereStr = " WHERE (MR2.MASTER_FLAG <>  'Y' OR MR2.MASTER_FLAG IS NULL) AND JO.TX_TYPE = 'E' "
      WhereStr = WhereStr & " AND JO.TX_AMOUNT <> 0 "
      'MR2.MASTER_FLAG =  'Y' เป็น Flag ที่ บอกว่าเป็นบรรทัดรวม
   ElseIf Ind = 11 Then    'GetSumJobInOutLostItemByInPutDocDate
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,JI.PART_ITEM_ID INPUT_ID,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_DATE,LT.WEIGHT_AMOUNT "
      SelectStr1 = SelectStr1 & ",JO.PRODUCTION_TYPE "              'เพิ่มเข้ามา เนื่องจากยอดในตัว Premium กับตัว ByProduct มันรวมกัน
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (J.JOB_ID = JI.JOB_ID) "
           
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LT.LOT_ITEM_ID = LK.MAIN_IMPORT_LOT_ITEM_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID)"
      
      If MainLocationPrdFlag Or NoMainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JO.TX_TYPE =  'E'  AND JI.TX_TYPE =  'I' "
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      ElseIf NoMainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'N' "
      End If
   ElseIf Ind = 12 Then                    'CReportProduct004  'CReportProduct006'CReportProduct007
      SelectStr1 = "SELECT JI.PART_ITEM_ID,IVD.DOCUMENT_DATE,SUM(DISTINCT LT.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC,SC1.BILL_DESC,LT.WEIGHT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LK.MAIN_IMPORT_LOT_ITEM_ID = LT.LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      If MainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      
      Call GenerateBatchString(WhereStr)
      
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      End If
      
   ElseIf Ind = 13 Then                    'GetJobInputByPartItemInPutDate
      SelectStr1 = "SELECT JI.PART_ITEM_ID,SC1.BILL_DESC,IVD.DOCUMENT_DATE,SUM(JI.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",LT.WEIGHT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LT.LOT_ITEM_ID = LK.MAIN_IMPORT_LOT_ITEM_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      If MainLocationPrdFlag Or NoMainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' AND JI.PART_ITEM_ID IS NOT NULL "
      
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      ElseIf NoMainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'N' "
      End If
      
   ElseIf Ind = 14 Then                    'CReportProduct003_1
      SelectStr1 = "SELECT J.JOB_DATE,SUM(JI.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",JI.PART_ITEM_ID,SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC,SC1.BILL_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      Call GenerateBatchString(WhereStr)
   ElseIf Ind = 15 Then    'GetSumJobInOutLostItemByInPutJobDate
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,JI.PART_ITEM_ID INPUT_ID,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",J.JOB_DATE "
      SelectStr1 = SelectStr1 & ",JO.PRODUCTION_TYPE "              'เพิ่มเข้ามา เนื่องจากยอดในตัว Premium กับตัว ByProduct มันรวมกัน
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (J.JOB_ID = JI.JOB_ID) "
      If MainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JO.TX_TYPE =  'E'  AND JI.TX_TYPE =  'I'  "
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      End If
   
   ElseIf Ind = 17 Then                    'CReportProduct005
      SelectStr1 = "SELECT J.JOB_NO,JI.PART_ITEM_ID,J.JOB_DATE "
      SelectStr1 = SelectStr1 & ",SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC,SC1.BILL_DESC "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LK.MAIN_IMPORT_LOT_ITEM_ID = LT.LOT_ITEM_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      If MainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      
      Call GenerateBatchString(WhereStr)
      
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      End If
   
   ElseIf Ind = 19 Then    'GetSumJobInOutLostItemByJobNo
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,J.JOB_NO,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID,JO.PRODUCTION_TYPE "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      
      If MainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JO.TX_TYPE =  'E'   "
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      End If
   ElseIf Ind = 20 Then                    'GetJobInputByJobNo
      SelectStr1 = "SELECT J.JOB_NO,JI.PART_ITEM_ID,SUM(JI.TX_AMOUNT) TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      If MainLocationPrdFlag Then
         FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      End If
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' AND JI.PART_ITEM_ID IS NOT NULL "
      
      Call GenerateBatchString(WhereStr)
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      End If
   ElseIf Ind = 21 Then                'GetDistinctJobOutItemByPrdTypeIvdDate
      SelectStr1 = "SELECT DISTINCT JO.PART_ITEM_ID,SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC,SC.BILL_DESC "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME PROBLEM_DESC,MR2.MASTER_FLAG SUM_FLAG,MR2.SHORT_CODE PROBLEM_SHORT_CODE "
      SelectStr1 = SelectStr1 & ",JO.PRODUCTION_TYPE,MR3.SHORT_CODE PRD_TYPE_SHORT_CODE,MR3.MASTER_FLAG PRD_TYPE_DETAIL_FLAG "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = JO.LOST_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (MR3.KEY_ID = JO.PRODUCTION_TYPE) "
      
      FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LK.MAIN_IMPORT_LOT_ITEM_ID = LT.LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID) "
      
      WhereStr = " WHERE (MR2.MASTER_FLAG <>  'Y' OR MR2.MASTER_FLAG IS NULL) AND JO.TX_TYPE = 'E'  AND JI.TX_TYPE = 'I' "
      WhereStr = WhereStr & " AND JO.TX_AMOUNT > 0 "
      'MR2.MASTER_FLAG =  'Y' เป็น Flag ที่ บอกว่าเป็นบรรทัดรวม
   ElseIf Ind = 22 Then    'GetSumJobInOutLostItemByInPutDocDateMainFlag
      SelectStr1 = "SELECT SUM(JO.TX_AMOUNT) TX_AMOUNT,JI.PART_ITEM_ID INPUT_ID,JO.PART_ITEM_ID OUTPUT_ID "
      SelectStr1 = SelectStr1 & ",JO.LOST_ID "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_DATE,LT.WEIGHT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JO "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JO.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (J.JOB_ID = JI.JOB_ID) "
           
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_LINK LK ON (LK.EXPORT_LOT_ITEM_ID = JI.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT ON (LT.LOT_ITEM_ID = LK.MAIN_IMPORT_LOT_ITEM_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = LT.INVENTORY_DOC_ID)"
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = J.PRD_LOCATION_ID) "
      
      WhereStr = " WHERE JO.TX_TYPE =  'E'  AND JI.TX_TYPE =  'I'  "
      If MainLocationPrdFlag Then
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'Y' "
      Else
         WhereStr = WhereStr & " AND MR1.MASTER_FLAG =  'N' "
      End If
   ElseIf Ind = 23 Then     'CReportProduct008
      SelectStr1 = "SELECT SUM(JI.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",JI.PART_ITEM_ID,SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC,SC1.BILL_DESC "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (J.JOB_ID = JI.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      
      WhereStr = " WHERE JI.TX_TYPE =  'I' "
      Call GenerateBatchString(WhereStr)
   End If
      
   If JOB_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JO.JOB_ITEM_ID = " & JOB_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JO.JOB_ITEM_ID = " & JOB_ITEM_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JO.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JO.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (J.JOB_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (J.JOB_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If FROM_IVD_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_IVD_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If TO_IVD_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_IVD_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Trim(TX_TYPE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (JO.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
      Else
         WhereStr = WhereStr & "AND (JO.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
      End If
   End If
   
   If PRD_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PRD_LOCATION_ID = " & PRD_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.PRD_LOCATION_ID = " & PRD_LOCATION_ID & ") "
      End If
   End If
   
   If Trim(FROM_STOCK_NO) <> "" Then
      Call GenerateSearchStockNo(FromStr, "I", WhereStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC1.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC1.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO) <> "" Then
      Call GenerateSearchStockNo(FromStr, "I", WhereStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC1.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC1.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(FROM_STOCK_NO1) <> "" Then
      Call GenerateSearchStockNo(FromStr, "E", WhereStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO1)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO1)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO1) <> "" Then
      Call GenerateSearchStockNo(FromStr, "E", WhereStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO1)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO1)) & "')"
      End If
   End If
   
   If Trim(BATCH_NO_SET) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (JO.BATCH_NO IN " & ChangeQuote(Trim(BATCH_NO_SET)) & ")"
      Else
         WhereStr = WhereStr & "AND (JO.BATCH_NO IN " & ChangeQuote(Trim(BATCH_NO_SET)) & ")"
      End If
   End If
   
   If PRODUCTION_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JO.PRODUCTION_TYPE = " & PRODUCTION_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (JO.PRODUCTION_TYPE = " & PRODUCTION_TYPE & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 4 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,JO.PART_ITEM_ID,JO.LOST_ID,JO.BATCH_NO "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY J.JOB_DATE "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY J.JOB_DATE,JO.PART_ITEM_ID,JO.LOST_ID,JO.BATCH_NO "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,IVD.DOCUMENT_DATE,JO.PART_ITEM_ID,JO.LOST_ID,LT.WEIGHT_AMOUNT,JO.PRODUCTION_TYPE "
   ElseIf Ind = 22 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,IVD.DOCUMENT_DATE,JO.PART_ITEM_ID,JO.LOST_ID,LT.WEIGHT_AMOUNT "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,SC1.STOCK_NO ,SC1.STOCK_DESC ,SC1.BILL_DESC,IVD.DOCUMENT_DATE,LT.WEIGHT_AMOUNT "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,SC1.BILL_DESC,IVD.DOCUMENT_DATE,LT.WEIGHT_AMOUNT "
   ElseIf Ind = 14 Then
      GroupByStr = "GROUP BY J.JOB_DATE,JI.PART_ITEM_ID,SC1.STOCK_NO,SC1.STOCK_DESC,SC1.BILL_DESC "
   ElseIf Ind = 15 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,J.JOB_DATE,JO.PART_ITEM_ID,JO.LOST_ID,JO.PRODUCTION_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY J.JOB_NO,JO.PART_ITEM_ID,JO.LOST_ID,JO.PRODUCTION_TYPE "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY J.JOB_NO,JI.PART_ITEM_ID "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY JI.PART_ITEM_ID,SC1.STOCK_NO,SC1.STOCK_DESC,SC1.BILL_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      OrderByStr = "ORDER BY JO.JOB_ITEM_ID "
      OrderType = 1
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY JO.BATCH_NO ASC,JO.ORDER_NO "
      OrderType = 1
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY SC1.STOCK_NO "
      OrderType = 1
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY J.JOB_DATE "
      OrderType = 1
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY JO.BATCH_NO ASC,MR2.KEY_CODE ASC,SC.STOCK_NO "
      OrderType = 1
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY JO.BATCH_NO ASC,JO.PRODUCTION_TYPE ASC,SC.STOCK_NO ASC,MR2.KEY_CODE "
      OrderType = 1
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY JO.PRODUCTION_TYPE ASC,SC.STOCK_NO ASC,MR2.KEY_CODE "
      OrderType = 1
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY SC1.STOCK_NO ASC,IVD.DOCUMENT_DATE "
      OrderType = 1
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY SC1.STOCK_NO ASC,J.JOB_DATE "
      OrderType = 1
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY J.JOB_NO ASC,J.JOB_DATE ASC,SC1.STOCK_NO  "
      OrderType = 1
   ElseIf Ind = 21 Then
      OrderByStr = "ORDER BY JO.PRODUCTION_TYPE ASC,SC.STOCK_NO ASC,MR2.KEY_CODE "
      OrderType = 1
   ElseIf Ind = 23 Then
      OrderByStr = "ORDER BY SC1.STOCK_NO "
      OrderType = 1
   End If
   
   If OrderType = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf OrderType = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   If Ind = 4 Or Ind = 5 Or Ind = 7 Or Ind = 11 Or Ind = 13 Or Ind = 15 Or Ind = 19 Or Ind = 20 Or Ind = 22 Then
      OrderByStr = ""
   End If
      
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      '<------------------------------------------------------------------------------------------------------------------------------------
      JOB_ITEM_ID = NVLI(Rs("JOB_ITEM_ID"), -1)
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
      ORDER_NO = NVLI(Rs("ORDER_NO"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_AMOUNT_STD = NVLD(Rs("TX_AMOUNT_STD"), 0)
      TX_WEIGHT = NVLD(Rs("TX_WEIGHT"), 0)
      TX_WEIGHT_STD = NVLD(Rs("TX_WEIGHT_STD"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_LIMIT_PERCENT = NVLD(Rs("PROBLEM_LIMIT_PERCENT"), 0)
      NEXT_FLAG = NVLS(Rs("NEXT_FLAG"), "")
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
      EXPORT_LOT_ITEM_ID = NVLI(Rs("EXPORT_LOT_ITEM_ID"), -1)
      '------------------------------------------------------------------------------------------------------------------------------------>
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 4 Then
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 6 Then
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 7 Then
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
   ElseIf Ind = 8 Then
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      PROBLEM_SHORT_CODE = NVLS(Rs("PROBLEM_SHORT_CODE"), "")
   ElseIf Ind = 9 Then
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      PROBLEM_SHORT_CODE = NVLS(Rs("PROBLEM_SHORT_CODE"), "")
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
      PRD_TYPE_SHORT_CODE = NVLS(Rs("PRD_TYPE_SHORT_CODE"), "")
   ElseIf Ind = 10 Or Ind = 21 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      PROBLEM_DESC = NVLS(Rs("PROBLEM_DESC"), "")
      SUM_FLAG = NVLS(Rs("SUM_FLAG"), "")
      PROBLEM_SHORT_CODE = NVLS(Rs("PROBLEM_SHORT_CODE"), "")
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
      PRD_TYPE_SHORT_CODE = NVLS(Rs("PRD_TYPE_SHORT_CODE"), "")
      PRD_TYPE_DETAIL_FLAG = NVLS(Rs("PRD_TYPE_DETAIL_FLAG"), "")
   ElseIf Ind = 11 Then
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
   ElseIf Ind = 22 Then
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
   ElseIf Ind = 14 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 15 Then
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
   ElseIf Ind = 17 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      JOB_NO = NVLS(Rs("JOB_NO"), "")
   ElseIf Ind = 19 Then
      OUTPUT_ID = NVLI(Rs("OUTPUT_ID"), -1)
      LOST_ID = NVLI(Rs("LOST_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PRODUCTION_TYPE = NVLI(Rs("PRODUCTION_TYPE"), -1)
   ElseIf Ind = 20 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
   ElseIf Ind = 23 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String

   TempID1 = JOB_ID
   If JOB_ID <= 0 Then
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
   
   TempID5 = LOST_ID
   If LOST_ID <= 0 Then
      TempID5 = "NULL "
   End If
   
   TempID6 = PRODUCTION_TYPE
   If PRODUCTION_TYPE <= 0 Then
      TempID6 = "NULL "
   End If
   
   TempID7 = EXPORT_LOT_ITEM_ID
   If EXPORT_LOT_ITEM_ID <= 0 Then
      TempID7 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO JOB_ITEM " & _
                     "( " & _
                     "   JOB_ITEM_ID, " & _
                     "   JOB_ID, " & _
                     "   LINK_ID, " & _
                     "   BATCH_NO, " & _
                     "   ORDER_NO, " & _
                     "   PART_ITEM_ID, " & _
                     "   TX_TYPE, " & _
                     "   TX_AMOUNT, " & _
                     "   TX_AMOUNT_STD, " & _
                     "   TX_WEIGHT, " & _
                     "   TX_WEIGHT_STD, " & _
                     "   LOCATION_ID, " & _
                     "   LOST_ID, " & _
                     "   PROBLEM_LIMIT_PERCENT, " & _
                     "   NEXT_FLAG, " & _
                     "   PRODUCTION_TYPE, " & _
                     "   EXPORT_LOT_ITEM_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     TempID1 & "," & _
                     LINK_ID & "," & _
                     BATCH_NO & "," & _
                     ORDER_NO & "," & _
                     TempID2 & "," & _
                     "'" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                     TX_AMOUNT & "," & _
                     TX_AMOUNT_STD & "," & _
                     TX_WEIGHT & "," & _
                     TX_WEIGHT_STD & "," & _
                     TempID3 & "," & _
                     TempID5 & "," & _
                     PROBLEM_LIMIT_PERCENT & "," & _
                     "'" & ChangeQuote(Trim(NEXT_FLAG)) & "'," & _
                     TempID6 & "," & _
                     TempID7 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         JOB_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE JOB_ITEM SET " & _
                     "  LINK_ID = " & LINK_ID & ", " & _
                     "  BATCH_NO = " & BATCH_NO & ", " & _
                     "  ORDER_NO = " & ORDER_NO & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & "," & _
                     "  TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "', " & _
                     "  TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "  TX_AMOUNT_STD = " & TX_AMOUNT_STD & ", " & _
                     "  TX_WEIGHT = " & TX_WEIGHT & ", " & _
                     "  TX_WEIGHT_STD = " & TX_WEIGHT_STD & ", " & _
                     "   LOCATION_ID = " & TempID3 & "," & _
                     "   LOST_ID = " & TempID5 & "," & _
                     "  PROBLEM_LIMIT_PERCENT = " & PROBLEM_LIMIT_PERCENT & ", " & _
                     "  NEXT_FLAG = '" & ChangeQuote(Trim(NEXT_FLAG)) & "', " & _
                     "   PRODUCTION_TYPE = " & TempID6 & "," & _
                     "   EXPORT_LOT_ITEM_ID = " & TempID7 & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_ITEM_ID = " & JOB_ITEM_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_ITEM WHERE JOB_ITEM_ID = " & JOB_ITEM_ID
      m_Conn.Execute (SQL1)
End Function
Public Sub GenerateBatchString(WhereStr As String)
Dim Str As String
Dim Ins As Byte
Dim ManyBatch As Boolean
Dim iCount As Byte
   ManyBatch = False
   iCount = 0
   While (BATCH_NO_SET <> "")
      If Left(BATCH_NO_SET, 1) = "(" Then
         BATCH_NO_SET = Mid(BATCH_NO_SET, 2)
      ElseIf Left(BATCH_NO_SET, 1) = ")" Then
         BATCH_NO_SET = ""
         If ManyBatch Then
            WhereStr = WhereStr & " ) "
         End If
         Exit Sub
      Else
         Ins = InStr(1, BATCH_NO_SET, ",")
         iCount = iCount + 1
         If Ins = 0 Then
            Str = Left(BATCH_NO_SET, Len(BATCH_NO_SET) - 1)
            BATCH_NO_SET = ")"
         Else
            ManyBatch = True
            Str = Left(BATCH_NO_SET, Ins - 1)
            BATCH_NO_SET = Mid(BATCH_NO_SET, Ins + 1)
         End If
         
         If Not (ManyBatch) Then
            If Len(WhereStr) > 0 Then
               WhereStr = WhereStr & " AND  ( " & " ( " & Str & " )  IN (SELECT DISTINCT JOO.BATCH_NO FROM JOB_ITEM JOO WHERE JOO.JOB_ID = J.JOB_ID AND JOO.TX_TYPE = 'E' )  ) "
            Else
               WhereStr = WhereStr & " WHERE  ( " & " ( " & Str & " )  IN (SELECT DISTINCT JOO.BATCH_NO FROM JOB_ITEM JOO WHERE JOO.JOB_ID = J.JOB_ID AND JOO.TX_TYPE = 'E' )  ) "
            End If
         Else
            If iCount = 1 Then
               If Len(WhereStr) > 0 Then
                  WhereStr = WhereStr & " AND  ( ( " & " ( " & Str & " )  IN (SELECT DISTINCT JOO.BATCH_NO FROM JOB_ITEM JOO WHERE JOO.JOB_ID = J.JOB_ID AND JOO.TX_TYPE = 'E' )  ) "
               Else
                  WhereStr = WhereStr & " WHERE  ( ( " & " ( " & Str & " )  IN (SELECT DISTINCT JOO.BATCH_NO FROM JOB_ITEM JOO WHERE JOO.JOB_ID = J.JOB_ID AND JOO.TX_TYPE = 'E' )  ) "
               End If
            Else
               WhereStr = WhereStr & " OR   ( " & " ( " & Str & " )  IN (SELECT DISTINCT JOO.BATCH_NO FROM JOB_ITEM JOO WHERE JOO.JOB_ID = J.JOB_ID AND JOO.TX_TYPE = 'E' )  ) "
            End If
         End If
      End If
   Wend
End Sub
Public Sub GenerateSearchStockNo(FromStr As String, TxType As String, WhereStr As String)
   If TxType = "I" Then
      If InStr(1, FromStr, "JI.JOB_ID") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JI ON (JI.JOB_ID = J.JOB_ID) "
        If WhereStr = "" Then
         WhereStr = "WHERE (JI.TX_TYPE =  'I')"
         Else
         WhereStr = WhereStr & "AND (JI.TX_TYPE =  'I')"
         End If
      End If
      If InStr(1, FromStr, "SC1.STOCK_CODE_ID") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (SC1.STOCK_CODE_ID = JI.PART_ITEM_ID) "
      End If
   ElseIf TxType = "E" Then
      If InStr(1, FromStr, "JO.JOB_ID") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN JOB_ITEM JO ON (JO.JOB_ID = J.JOB_ID) "
        If WhereStr = "" Then
         WhereStr = "WHERE (JO.TX_TYPE =  'E')"
         Else
         WhereStr = WhereStr & "AND (JO.TX_TYPE =  'E')"
         End If
      End If
      If InStr(1, FromStr, "SC.STOCK_CODE_ID") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = JO.PART_ITEM_ID) "
      End If
   End If
End Sub

