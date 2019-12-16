VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuiConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GuiConfigs As Collection                                                                    '1

Public DOAddMenuItems As Collection                                                          '4
Public POAddMenuItems As Collection                                                          '5

Public ReceiptAddMenuItems As Collection                                                    '6
Public ReportMenuItems As Collection                                                            '7
Public QuoatationPrintMenuItems As Collection                                          '8
Public POPrintMenuItems As Collection                                                        '9
Public DOPrintMenuItems As Collection                                                        '10

Public IVPrintMenuItems As Collection                                                        '11
Public RCPrintMenuItems As Collection                                                        '12
Public RCPrintMenuItems3 As Collection
Public RCPrintMenuItems4 As Collection

Public RCPrintMenuItems2 As Collection                                                        '12
Public RCPackPrintMenuItems As Collection                                                        '12
Public RTPrintMenuItems As Collection                                           '13
Public RTPrintMenuItems2 As Collection
Public CdnDbnPrintMenuItems As Collection                                              '14
Public BSPrintMenuItems As Collection                                                        '15
Public RoPrintMenuItems As Collection                                                        '15

Public PkglistPrintMenuItems As Collection                                                '16
Public MasterCommandMenuItems As Collection                                      '17
Public MainCommandMenuItems As Collection                                        '18
Public StockCommandMenuItems As Collection                                       '19
Public ProdCommandMenuItems As Collection                                          '20

Public LedgerCommandMenuItems As Collection                                       '21
Public GLCommandMenuItems As Collection                                             '22
Public CMCommandMenuItems As Collection                                             '23
Public PackageCommandMenuItems As Collection                                     '24
Public CostCommandMenuItems As Collection                                     '24

Public AdminCommandMenuItems As Collection                                       '26
Public TagetCommandMenuItems As Collection                                     '27
Public SellSubMenuItems As Collection                                                                    '28
Public PoSubMenuItems As Collection
Public SellReturnMenuItems As Collection                                                                    '29
Public CashMenuItems As Collection                                                                                 '30

Private Sub Class_Initialize()
   Set m_GuiConfigs = New Collection
   Set ReportMenuItems = New Collection
   Set DOAddMenuItems = New Collection
   
   Set POAddMenuItems = New Collection
   Set ReceiptAddMenuItems = New Collection
   Set QuoatationPrintMenuItems = New Collection
   Set POPrintMenuItems = New Collection
   Set DOPrintMenuItems = New Collection
   
   Set IVPrintMenuItems = New Collection
   Set RCPrintMenuItems = New Collection
   Set RCPrintMenuItems2 = New Collection
   Set RCPrintMenuItems3 = New Collection
   Set RCPrintMenuItems4 = New Collection
   Set RCPackPrintMenuItems = New Collection
   Set RTPrintMenuItems = New Collection
   Set RTPrintMenuItems2 = New Collection
   Set CdnDbnPrintMenuItems = New Collection
   Set BSPrintMenuItems = New Collection
   Set RoPrintMenuItems = New Collection
   
   Set PkglistPrintMenuItems = New Collection
   Set CMCommandMenuItems = New Collection
   Set MasterCommandMenuItems = New Collection
   Set MainCommandMenuItems = New Collection
   Set StockCommandMenuItems = New Collection
   
   Set ProdCommandMenuItems = New Collection
   Set LedgerCommandMenuItems = New Collection
   Set PackageCommandMenuItems = New Collection
   Set AdminCommandMenuItems = New Collection
   Set CostCommandMenuItems = New Collection
   
   Set GLCommandMenuItems = New Collection
   Set TagetCommandMenuItems = New Collection
   Set SellSubMenuItems = New Collection
   Set PoSubMenuItems = New Collection
   Set SellReturnMenuItems = New Collection
   Set CashMenuItems = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_GuiConfigs = Nothing
   
   Set ReportMenuItems = Nothing
   Set DOAddMenuItems = Nothing
   Set POAddMenuItems = Nothing
   Set ReceiptAddMenuItems = Nothing
   Set QuoatationPrintMenuItems = Nothing
   Set POPrintMenuItems = Nothing
   Set DOPrintMenuItems = Nothing
   Set IVPrintMenuItems = Nothing
   Set RCPrintMenuItems = Nothing
   Set RCPrintMenuItems2 = Nothing
   Set RCPrintMenuItems3 = Nothing
   Set RCPrintMenuItems4 = Nothing
   Set RCPackPrintMenuItems = Nothing
   Set CdnDbnPrintMenuItems = Nothing
   Set BSPrintMenuItems = Nothing
   Set RoPrintMenuItems = Nothing
   
   Set CMCommandMenuItems = Nothing
   Set MasterCommandMenuItems = Nothing
   Set MainCommandMenuItems = Nothing
   Set StockCommandMenuItems = Nothing
   Set ProdCommandMenuItems = Nothing
   Set LedgerCommandMenuItems = Nothing
   Set PackageCommandMenuItems = Nothing
   Set AdminCommandMenuItems = Nothing
   Set CostCommandMenuItems = Nothing
   
   Set RTPrintMenuItems = Nothing
   Set RTPrintMenuItems2 = Nothing
   Set PkglistPrintMenuItems = Nothing
   Set GLCommandMenuItems = Nothing
   Set TagetCommandMenuItems = Nothing
   Set SellSubMenuItems = Nothing
   Set PoSubMenuItems = Nothing
   Set SellReturnMenuItems = Nothing
   Set CashMenuItems = Nothing
End Sub

Private Sub CreateGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = New CGuiConfig
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
   
   Call m_GuiConfigs.add(TempC, Key)
   Set TempC = Nothing
End Sub

Private Sub SetGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(Key)
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
End Sub

Private Sub GenerateDOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("เลือกจากใบ PO")
   TempID(4) = 4
   TempKey(5) = MapText("เลือกจากใบเสนอราคา")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'lMenuChosen = oMenu.Popup("พิมพ์บนกระดาษเปล่า", "ปรับค่าหน้ากระดาษ", "-", "พิมพ์บนฟอร์มเปล่า (เหมือนจริง)", "พิมพ์บนฟอร์มเปล่า (ข้อมูล)", "ปรับค่าหน้ากระดาษ")
Private Sub GeneratePrintMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(50) As String
Dim TempID(50) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("พิมพ์บนกระดาษเปล่า")
  TempID(1) = 1
   TempKey(2) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("พิมพ์ PREVIEW (F10) (LETTER)")
   TempID(4) = 4
   TempKey(5) = MapText("พิมพ์ ทันที (F11) (LETTER)")
   TempID(5) = 5
   TempKey(6) = MapText("ปรับค่าหน้ากระดาษ (LETTER)")
   TempID(6) = 6
   TempKey(7) = MapText("-")
   TempID(7) = 7
   TempKey(8) = MapText("พิมพ์  ลาเบล ส่งของ")
   TempID(8) = 8
   TempKey(9) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(9) = 9
   TempKey(10) = MapText("-")
   TempID(10) = 10
   TempKey(11) = MapText("รายงานใบแยกสถานที่จัดส่ง")
   TempID(11) = 11
   TempKey(12) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(12) = 12
   TempKey(13) = MapText("รายงานใบรับของ")
   TempID(13) = 13
   TempKey(14) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(14) = 14
   TempKey(15) = MapText("-")
   TempID(15) = 15
   TempKey(16) = MapText("ใบสำคัญรับ (PREVIEW)")
   TempID(16) = 16
   TempKey(17) = MapText("ใบสำคัญรับ")
   TempID(17) = 17
   TempKey(18) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(18) = 18
   TempKey(19) = MapText("-")
   TempID(19) = 19
   TempKey(20) = MapText("ใบเสร็จทั้งหมดจากเอกสารเป็นชุด (PREVIEW)")
   TempID(20) = 20
   TempKey(21) = MapText("ใบเสร็จทั้งหมดจากเอกสารเป็นชุด")
   TempID(21) = 21
   TempKey(22) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(22) = 22
   TempKey(23) = MapText("-")
   TempID(23) = 23
   TempKey(24) = MapText("ใบวางบิลสำหรับ MK")
   TempID(24) = 24
   TempKey(25) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(25) = 25
   TempKey(26) = MapText("-")
   TempID(26) = 26
   TempKey(27) = MapText("พิมพ์  ลาเบล TOP ส่งของ(สติ๊กเกอร์)")
   TempID(27) = 27
   TempKey(28) = MapText("พิมพ์  ลาเบล Big C ส่งของ(สติ๊กเกอร์)")
   TempID(28) = 28
   TempKey(29) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(29) = 29
   TempKey(30) = MapText("-")
   TempID(30) = 30
   TempKey(31) = MapText("พิมพ์  ลาเบล อิออน ส่งของ(สติ๊กเกอร์)")
   TempID(31) = 31
   TempKey(32) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(32) = 32
   TempKey(33) = MapText("พิมพ์บนกระดาษฉบับย่อ")
   TempID(33) = 33
   TempKey(34) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(34) = 34
   TempKey(35) = MapText("-")
   TempID(35) = 35
   TempKey(36) = MapText("พิมพ์ PREVIEW (F10) (1/2 LETTER)")
   TempID(36) = 36
   TempKey(37) = MapText("พิมพ์ ทันที (F11) (1/2 LETTER)")
   TempID(37) = 37
   TempKey(38) = MapText("ปรับค่าหน้ากระดาษ (1/2 LETTER)")
   TempID(38) = 38
   TempKey(39) = MapText("-")
   TempID(39) = 39

   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("เพิ่มรายการใหม่ (ทีละรายการ)", "เพิ่มรายการใหม่ (เป็นชุด)")
Private Sub GeneratePOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("สร้างรายการใหม่", "-", "สร้างจากใบแจ้งหนี้", "-", "สร้างจากใบกำกับภาษี")
Private Sub GenerateAddReceiptMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("สร้างรายการใหม่ (ขายสด)")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("สร้างจากใบส่งของ (รับชำระหนี้)")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("สร้างจากใบกำกับภาษี")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateReportMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long

   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = ROOT_TREE & " A-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/ลูกค้า (SL001)")
   D.KEYWORD = ROOT_TREE & " 5-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '3
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวันตามลูกค้า (SL002)")
   D.KEYWORD = ROOT_TREE & " 5-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '4
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวัน (SL003)")
   D.KEYWORD = ROOT_TREE & " 5-3"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '5
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายรายวัน (SL004)")
   D.KEYWORD = ROOT_TREE & " 5-4"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '6
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงาน (SL005)")
   D.KEYWORD = ROOT_TREE & " 5-5"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '7
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ/ลูกค้า (SL006)")
'   D.KEYWORD = ROOT_TREE & " 5-6"
'   D.PARENT_KEY = ROOT_TREE & " A-1"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ (SL006)")
   D.KEYWORD = ROOT_TREE & " A-1-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบลูกหนี้")
   D.KEYWORD = ROOT_TREE & " A-2"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานลูกหนี้ค้างชำระรายตัว (AR001)")
   D.KEYWORD = ROOT_TREE & " A-2-1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปลูกหนี้ค้างชำระรายตัว (AR002)")
   D.KEYWORD = ROOT_TREE & " A-2-2"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปลูกหนี้ค้างชำระรายตัว เรียงตามพนักงานขาย (AR002.2)")
   D.KEYWORD = ROOT_TREE & " A-2-2-2"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการ์ดลูกหนี้ (AR003)")
   D.KEYWORD = ROOT_TREE & " A-2-3"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุลูกหนี้ (AR004)")
   D.KEYWORD = ROOT_TREE & " A-2-4"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการนำฝากธนาคาร (AR005)")
   D.KEYWORD = ROOT_TREE & " A-2-5"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเงินสดคงเหลือ (AR006)")
   D.KEYWORD = ROOT_TREE & " A-2-6"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateMasterCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักส่วนกลาง")
   D.KEYWORD = "MASTER_MAIN"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบบัญชีแยกประเภท")
   D.KEYWORD = "MASTER_GL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบบริหารบัญชี")
   D.KEYWORD = "MASTER_LEDGER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบคลัง")
   D.KEYWORD = "MASTER_INVENTORY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
            
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบการผลิต")
   D.KEYWORD = "MASTER_PRODUCTION"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลหลัก")
   D.KEYWORD = "MASTER_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateMainCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
         
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลบริษัท")
   D.KEYWORD = "MAIN_ENTERPRISE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลลูกค้า")
   D.KEYWORD = "MAIN_CUSTOMER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลซัพพลายเออร์")
   D.KEYWORD = "MAIN_SUPPLIER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลพนักงาน")
   D.KEYWORD = "MAIN_EMPLOYEE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลกลาง")
   D.KEYWORD = "MAIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateProdCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("สูตรการผลิต")
   D.KEYWORD = "PRODUCT_FORMULA"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการผลิต")
   D.KEYWORD = "PRODUCT_JOB"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เป้าการผลิต")
   D.KEYWORD = "PRODUCT_TAGET"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบการผลิต")
   D.KEYWORD = "PRODUCT_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
         
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateStockCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลรหัสคลัง")
   D.KEYWORD = "INVENTORY_PART"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลเอกสารคลัง")
   D.KEYWORD = "INVENTORY_DOC"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("บาร์โค๊ด")
   D.KEYWORD = "INVENTORY_BARCODE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบคลัง")
   D.KEYWORD = "INVENTORY_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateAdminCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลกลุ่มผู้ใช้งาน")
   D.KEYWORD = "ADMIN_GROUP"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_USER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateLedgerCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = "LEDGER_SELL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานซื้อ")
   D.KEYWORD = "LEDGER_BUY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบข้อมูลการเงิน")
   D.KEYWORD = "LEDGER_CASH"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบข้อมูลสินทรัพย์")
   D.KEYWORD = "LEDGER_CASH"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบบัญชี")
   D.KEYWORD = "LEDGER_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("โปรแกรมบัญชีการเงินอื่นๆ")
   D.KEYWORD = "LEDGER_PROGRAM"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateGLCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลผังบัญชี")
   D.KEYWORD = "GL_CHART"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลสมุดรายวัน")
   D.KEYWORD = "GL_JOURNAL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบบัญชีแยกประเภท")
   D.KEYWORD = "GL_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateCMCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ตารางคอมมิตชั่น")
   D.KEYWORD = "COMMISSION_TABLE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ตารางหักรับคืนสินค้า")
   D.KEYWORD = "RETURN_TABLE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
            
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ตารางเพิ่มค่าคอม")
   D.KEYWORD = "COMMISSION_TABLE_EX"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("แผนภูมิการคิดคอมมิตชั่น")
   D.KEYWORD = "COMMISSION_CHART"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("แผนภูมิพนักงานขาย")
   D.KEYWORD = "SALE_ORGANIZE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ปรับประเภทตัวแทน")
   D.KEYWORD = "ADJUST_DEALER_TYPE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานคอมมิตชั่น")
   D.KEYWORD = "COMMISSION_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GeneratePackageCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการตั้งราคาสินค้า")
   D.KEYWORD = "PACKAGE_DATA"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateTagetCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลเป้าการขายลูกค้า")
   D.KEYWORD = "TAGET_CUSTOMER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเป้าการขาย")
   D.KEYWORD = "TAGET_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateCashMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem
   
   TempKey(1) = CashDocTypeToText(CHEQUE_REV)
   TempID(1) = CHEQUE_REV
   TempKey(2) = MapText("-")
   TempID(2) = -1
   
   TempKey(3) = CashDocTypeToText(CHEQUE_PAY)
   TempID(3) = CHEQUE_PAY
   TempKey(4) = MapText("-")
   TempID(4) = -1
   
   TempKey(5) = CashDocTypeToText(CASH_DEPOSIT)
   TempID(5) = CASH_DEPOSIT
   TempKey(6) = MapText("-")
   TempID(6) = -1
   
   TempKey(7) = CashDocTypeToText(POST_CHEQUE)
   TempID(7) = POST_CHEQUE
   TempKey(8) = MapText("-")
   TempID(8) = -1
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateSellSubExMenuItems(MenuMask As String, Col As Collection, Optional AddPoItem As Boolean = False)
Dim Mask As String
Dim I As Long
Dim TempKey(30) As String
Dim TempID(30) As Long
Dim D As CMenuItem
Dim Mr As CMasterRef
Dim TempColl As Collection

   Set TempColl = New Collection
   
   Call LoadMaster(Nothing, TempColl, , , MASTER_INVOICE_SUB)
   
   I = 0
   
   For Each Mr In TempColl
      I = I + 1
      TempKey((2 * I) - 1) = Mr.KEY_NAME
      TempID((2 * I) - 1) = Mr.KEY_ID
      
      If I <> TempColl.Count Then
         TempKey(2 * I) = MapText("-")
         TempID(2 * I) = 0
      End If
   Next Mr
   
   Set Mr = Nothing
   
   If AddPoItem Then
      Set D = New CMenuItem
      D.KEYWORD = "ขายสด"
      D.KEY_ID = 1
      Call Col.add(D)
      Set D = Nothing
      
      Set D = New CMenuItem
      D.KEYWORD = "-"
      D.KEY_ID = 0
      Call Col.add(D)
      Set D = Nothing
   End If
   
   For I = 1 To ((2 * TempColl.Count) - 1)
      Set D = New CMenuItem
      D.KEYWORD = TempKey(I)
      D.KEY_ID = TempID(I)
      Call Col.add(D)
      Set D = Nothing
   Next I
   
   Set TempColl = Nothing
End Sub
Private Sub GenerateSellReturnExMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(30) As String
Dim TempID(30) As Long
Dim D As CMenuItem
Dim TempColl As Collection
Dim Mr As CMasterRef
   
   Set TempColl = New Collection
   
   Call LoadMaster(Nothing, TempColl, , , MASTER_INVOICE_RETURN)
   
   I = 0
   
   For Each Mr In TempColl
      I = I + 1
      TempKey((2 * I) - 1) = Mr.KEY_NAME
      TempID((2 * I) - 1) = Mr.KEY_ID
      
      If I <> TempColl.Count Then
         TempKey(2 * I) = MapText("-")
         TempID(2 * I) = 0
      End If
   Next Mr
   
   Set Mr = Nothing
   
   For I = 1 To ((2 * TempColl.Count) - 1)
      Set D = New CMenuItem
      D.KEYWORD = TempKey(I)
      D.KEY_ID = TempID(I)
      Call Col.add(D)
      Set D = Nothing
   Next I
   
   Set TempColl = Nothing
End Sub
Private Sub GenerateCostCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ปรับต้นทุนมาตรฐาน")
   D.KEYWORD = "COST_STD"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("คำนวณต้นทุนคงเหลือและต้นทุนขาย")
   D.KEYWORD = "COST_CAPITAL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ปรับยอด STOCK แบบเป็นชุด")
   D.KEYWORD = "COST_STOCK-AMOUNT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Public Sub CreateGuiConfig(Programowner As String)
    
   '====
    Call CreateGuiConfigObj("LOGO_VIEW", False)
    Call CreateGuiConfigObj("AGENT_SERVER_CONNECT", False)
   
    Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)
   
   '====
   Call CreateGuiConfigObj("GL_VIEW", True)
   Call CreateGuiConfigObj("HR_VIEW", True)
   Call CreateGuiConfigObj("PRODUCTION_VIEW", True)
   Call CreateGuiConfigObj("PACKAGE_VIEW", True)
   Call CreateGuiConfigObj("LEDGER_VIEW", True)
   Call CreateGuiConfigObj("INVENTORY_VIEW", True)
   Call CreateGuiConfigObj("COMMISSION_VIEW", True)
   Call CreateGuiConfigObj("ENTERPRISE_INFO_EDIT", True)
   Call CreateGuiConfigObj("TAGET_VIEW", True)
   Call CreateGuiConfigObj("COST_VIEW", True)
   
   '====
   Call SetGuiConfigObj("HR_VIEW", False)
   Call SetGuiConfigObj("LEDGER_VIEW", True)
   Call SetGuiConfigObj("COMMISSION_VIEW", True)
   Call SetGuiConfigObj("PACKAGE_VIEW", True)
   Call SetGuiConfigObj("TAGET_VIEW", True)
   
    Call SetGuiConfigObj("GL_VIEW", False)
   
   Call GenerateCashMenuItems("YYYYYYY", CashMenuItems, Programowner)
   
   Call GenerateSellSubExMenuItems("Y", SellSubMenuItems)
   Call GenerateSellSubExMenuItems("Y", PoSubMenuItems, True)
   Call GenerateSellReturnExMenuItems("Y", SellReturnMenuItems, Programowner)
   
   Call GenerateReportMenuItems("YYYYYYYYYYYYYYYYYY", ReportMenuItems, Programowner)
   Call GenerateDOAddMenuItems("YYNNN", DOAddMenuItems, Programowner)
   Call GeneratePOAddMenuItems("YY", POAddMenuItems, Programowner)
   Call GenerateAddReceiptMenuItems("YYYNN", ReceiptAddMenuItems, Programowner)
   
   Call GeneratePrintMenuItems("YYNNNN", QuoatationPrintMenuItems, Programowner)
   'Call GeneratePrintMenuItems("YYYNNNNYYYYYNNNNNNNNNNNNNYYYY", POPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YYYNNNNYYYYYNNNNNNNNNNNNNYYYYYYY", POPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YYYYYYYYYNNNNNNNNNNNNNNNNNNNNNNNNNYYYYY", DOPrintMenuItems, Programowner)   'NNNYYYYYY
   Call GeneratePrintMenuItems("YYNNNN", IVPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YYYYYYNNNNNNNNNNNN", RCPrintMenuItems2, Programowner)
'  Call GeneratePrintMenuItems("YYYYYYNNNNNNNNYYYYYNNNNNNNNNNNNNYY", RCPrintMenuItems, Programowner)
  Call GeneratePrintMenuItems("YYYYYYNNNNNNNNYYYYYNNNNNNNNNNNNNYYYYYYY", RCPrintMenuItems, Programowner)
  Call GeneratePrintMenuItems("YYYNNNNNNNNNNNNNNNNNNNNNNNNNNNNN", RCPrintMenuItems3, Programowner) ' ระบบบัญชีและการเงิน--->-ขาย---->ใบลดหนี้     CN_DOCTYPE =7  สำหรับ popup การกดปุ่ม print
   Call GeneratePrintMenuItems("YYYNNNNNNNNNNNNNNNNNNNNNNNNNNNNN", RCPrintMenuItems4, Programowner)  ' ระบบบัญชีและการเงิน--->ขาย---->ใบเพิ่มหนี้     DN_DOCTYPE =8  สำหรับ popup การกดปุ่ม print
 'Call GeneratePrintMenuItems("YYYYYYNNNNNNNNYYYY", RCPrintMenuItems, Programowner)
  
   Call GeneratePrintMenuItems("NNNNNNNNNNNNNNNYYYYYYY", RCPackPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YYNNNNNNNNNN", RTPrintMenuItems, Programowner)
  Call GeneratePrintMenuItems("YYNNNNNNNNNN", RTPrintMenuItems2, Programowner)
   Call GeneratePrintMenuItems("YYNNNN", CdnDbnPrintMenuItems, Programowner)
  Call GeneratePrintMenuItems("YYYNNNNNNNNNNNNNNNNNNNNYY", BSPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("NNNNNNNNNNNNYY", RoPrintMenuItems, Programowner)  'พิมพ์ใบรับสินค้า
   
   
   Call GenerateMainCommandMenuItems("YYYYY", MainCommandMenuItems, Programowner)
   
   Call GenerateMasterCommandMenuItems("YNYYY", MasterCommandMenuItems, Programowner)
   
   Call GenerateAdminCommandMenuItems("YYN", AdminCommandMenuItems, Programowner)
   Call GenerateProdCommandMenuItems("YYY", ProdCommandMenuItems, Programowner)
   
   Call GenerateLedgerCommandMenuItems("YYYNY", LedgerCommandMenuItems, Programowner)
   
   Call GenerateStockCommandMenuItems("YYYYYY", StockCommandMenuItems, Programowner)
   Call GenerateGLCommandMenuItems("YYY", GLCommandMenuItems, Programowner)
   Call GenerateCMCommandMenuItems("YYYYYY", CMCommandMenuItems, Programowner)
   Call GeneratePackageCommandMenuItems("Y", PackageCommandMenuItems, Programowner)
   Call GenerateCostCommandMenuItems("YYY", CostCommandMenuItems, Programowner)
   Call GenerateTagetCommandMenuItems("Y", TagetCommandMenuItems, Programowner)
End Sub
Public Function VerifyGuiConfig(KeyCode As String, Optional PrintError As Boolean = False) As Boolean
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(KeyCode)
   VerifyGuiConfig = TempC.ENABLE_FLAG
   
   If Not TempC.ENABLE_FLAG Then
      If PrintError Then
         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
         glbErrorLog.ShowUserError
      End If
   End If
End Function
Public Function ShowWindowCaption(Programowner As String) As String
   ShowWindowCaption = MapText("ระบบบริหารงาน ExClusive System Software (Xivess)")
End Function
