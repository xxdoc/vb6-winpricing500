Attribute VB_Name = "modLoadData"
Option Explicit

Public Sub LoadMasterKeyID(C As ComboBox, Optional Cl As Collection = Nothing, Optional KEY_ID As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = KEY_ID
   Call TempData.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(7, Rs)
   
   
         If Not (C Is Nothing) Then
            C.AddItem (TempData.EMP_NAME & " " & TempData.EMP_LNAME)
            C.ItemData(I) = TempData.PARENT_EX_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.PARENT_EX_ID)))
         End If
'      If Not (C Is Nothing) Then
'         If ShowType = 2 Then
'            C.AddItem (TempData.KEY_NAME & "  (" & TempData.KEY_CODE & ")")
'         Else
'            C.AddItem (TempData.KEY_NAME)
'         End If
'         C.ItemData(I) = TempData.KEY_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.KEY_ID)))
'      End If
      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = -1, Optional ShowType As Long = -1, Optional MasterArea As MASTER_TYPE, Optional ParentExID2 As Long = -1, Optional ParentID As Long = -1, Optional ParentExID As Long = -1, Optional IndexLink As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   TempData.PARENT_EX_ID2 = ParentExID2
   TempData.PARENT_EX_ID = ParentExID
   TempData.PARENT_ID = ParentID
   TempData.INDEX_LINK = IndexLink
   Call TempData.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
       Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         If ShowType = 2 Then
            C.AddItem (TempData.KEY_NAME & "  (" & TempData.KEY_CODE & ")")
         Else
            C.AddItem (TempData.KEY_NAME)
         
         End If
         C.ItemData(I) = TempData.KEY_ID
      End If
         
      If Not (Cl Is Nothing) Then
        If ShowType = 2 Then
         Call Cl.add(TempData, Trim(TempData.KEY_CODE))
       Else
         Call Cl.add(TempData, Trim(Str(TempData.KEY_ID)))
       End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   If MasterArea = 34 Then
'       I = I + 1
'     Set TempData = New CMasterRef
'     Call TempData.PopulateFromRS(1, Rs)
''   C.AddItem ("ใบเอกสารเบิกอื่นๆ")
''   TempData.KEY_ID = -1
''   C.ItemData(I) = TempData.KEY_ID
'   C.AddItem ("ใบเอกสารเบิกอื่นๆ")
'   TempData.KEY_ID = 999999999
'   C.ItemData(I) = TempData.KEY_ID
'   ' TempData.KEY_ID = 999999999
'    'C.ItemData(I) = TempData.KEY_ID
'   End If

'   If MasterArea = 34 Then ' เพิ่มในเอกสารการเบิก ST007  เพื่อให้ค่า InventoryDoc_Sub_Type =null (ค่าที่ออกมา เป็น -1) คือ การเบิกที่ไม่ได้เบิกจากDatabase ในที่นี้จะเบิกจาก barcode
'       I = I + 1
'     Set TempData = New CMasterRef
'     Call TempData.PopulateFromRS(1, Rs)
'       If Not (C Is Nothing) Then
'           TempData.KEY_ID = 999999999
'          C.AddItem ("ใบเอกสารเบิกอื่นๆ")
'         C.ItemData(I) = TempData.KEY_ID
'      End If
'End If
     
    
    If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterID(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterArea As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   Call TempData.QueryData(5, Rs, itemcount, True)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.KEY_CODE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterFromToCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = -1, Optional ShowType As Long = -1, Optional MasterArea As MASTER_TYPE, Optional FromLocationID As String, Optional ToLocationID As String, Optional ParentExID2 As Long = -1, Optional ParentID As Long = -1, Optional ParentExID As Long = -1, Optional IndexLink As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   TempData.PARENT_EX_ID2 = ParentExID2
   TempData.PARENT_EX_ID = ParentExID
   TempData.PARENT_ID = ParentID
   TempData.INDEX_LINK = IndexLink
   TempData.FROM_LOCATION_ID = FromLocationID
   TempData.TO_LOCATION_ID = ToLocationID
   Call TempData.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         If ShowType = 2 Then
            C.AddItem (TempData.KEY_NAME & "  (" & TempData.KEY_CODE & ")")
         Else
            C.AddItem (TempData.KEY_NAME)
         End If
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterFromToType As MASTER_COMMISSION_AREA, Optional KeyType As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterFromTo
Dim I As Long
Dim D As CMasterFromTo

   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
      
   Set D = New CMasterFromTo
   Call D.SetFieldValue("MASTER_FROMTO_TYPE", MasterFromToType)
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterFromTo
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("MASTER_FROMTO_DESC"))
         C.ItemData(I) = TempData.GetFieldValue("MASTER_FROMTO_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("MASTER_FROMTO_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCustomerFromLocationSale(C As ComboBox, Optional Cl As Collection = Nothing, Optional ShowType As Long = -1, Optional MasterArea As MASTER_TYPE, Optional EmpId As Long = -1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   TempData.PARENT_EX_ID = EmpId
   Call TempData.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
         If ShowType = 2 Then
            C.AddItem (TempData.APAR_NAME & "  (" & TempData.APAR_CODE & ")")
         Else
            C.AddItem (TempData.APAR_NAME)
         End If
         C.ItemData(I) = TempData.PARENT_EX_ID2       ' รหัสซัพพลายเออร์
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PARENT_EX_ID2)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
      
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadApArMasConsignment(Cl As Collection)
On Error GoTo ErrorHandler
Dim APM As CAPARMas
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPARMas
Dim I As Long

   Set Rs = New ADODB.Recordset
 Set APM = New CAPARMas
   
   APM.CONSIGNMENT_FLAG = "Y"
   Call APM.QueryData(5, Rs, itemcount, True)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAPARMas
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE))
         Debug.Print TempData.APAR_CODE
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
         Debug.Print TempData.APAR_CODE
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEnterprise(Ep As CEnterprise, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CEnterprise
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEnterprise
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ep.SetFieldValue("ENTERPRISE_ID", -1)
   Call Ep.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEnterprise
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("ENTERPRISE_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("ENTERPRISE_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("ENTERPRISE_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadUserGroup(Ug As CUserGroup, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ug.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("GROUP_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("GROUP_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("GROUP_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadEmployee(Optional Emp As CEmployee = Nothing, Optional C As ComboBox = Nothing, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Rs = New ADODB.Recordset
      
      Emp.EMP_ID = -1
      Emp.ORDER_TYPE = 1
      
      Call Emp.QueryData(1, Rs, itemcount, False)
      
      While Not Rs.EOF
         I = I + 1
         Set TempData = New CEmployee
         Call TempData.PopulateFromRS(1, Rs)
      
         If Not (C Is Nothing) Then
            C.AddItem (TempData.EMP_NAME & " " & TempData.EMP_LNAME)
            C.ItemData(I) = TempData.EMP_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.EMP_ID)))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
      
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
   ElseIf Not (C Is Nothing) Then
      For Each TempData In m_EmployeeColl
         I = I + 1
         C.AddItem (TempData.EMP_NAME & " " & TempData.EMP_LNAME)
         C.ItemData(I) = TempData.EMP_ID
      Next TempData
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadApArMas(Optional ApAr As CAPARMas = Nothing, Optional C As ComboBox = Nothing, Optional Cl As Collection = Nothing, Optional ApArInd As Long = 1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPARMas
Dim I As Long
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Rs = New ADODB.Recordset
      
      ApAr.APAR_MAS_ID = -1
      ApAr.ORDER_TYPE = 1
      
      Call ApAr.QueryData(1, Rs, itemcount, False)
      
      While Not Rs.EOF
         I = I + 1
         Set TempData = New CAPARMas
         Call TempData.PopulateFromRS(1, Rs)
      
         If Not (C Is Nothing) Then
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.APAR_MAS_ID)))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
      
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
   ElseIf Not (C Is Nothing) Then
      If ApArInd = 1 Then
         For Each TempData In m_CustomerColl
            I = I + 1
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         Next TempData
      ElseIf ApArInd = 2 Then
         For Each TempData In m_SupplierColl
            I = I + 1
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         Next TempData
      End If
      
   End If
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadApArAddress(Cl As Collection, Optional ApArInd As Long = 1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim ApAr As CAPARMas
Dim TempData As CAPARMas
Dim I As Long
   
   Set Rs = New ADODB.Recordset
   
   Set ApAr = New CAPARMas
   
   ApAr.APAR_MAS_ID = -1
   ApAr.ORDER_TYPE = 1
   ApAr.APAR_IND = ApArInd
   Call ApAr.QueryData(6, Rs, itemcount, False)
   
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAPARMas
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (Cl Is Nothing) Then
         Set ApAr = GetObject("CAPARMas", Cl, Trim(Str(TempData.APAR_MAS_ID)), False)
         If ApAr Is Nothing Then
            Call Cl.add(TempData, Trim(Str(TempData.APAR_MAS_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub InitEnterpriseOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสบริษัท"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อบริษัท"))
   C.ItemData(2) = 2
End Sub

Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผู้ค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ค้า"))
   C.ItemData(2) = 2
End Sub
Public Sub InitUnitOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("บาท"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เปอร์เซ็นต์"))
   C.ItemData(2) = 2
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub
Public Sub InitMasterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัส"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("รายละเอียด"))
   C.ItemData(2) = 2
End Sub

Public Sub InitJournalOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(1) = 1
End Sub

Public Sub InitEmployeeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อพนักงาน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitInventoryDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitBillingDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub
Public Sub InitBillingDocConsignmentFlag(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1
   
   C.AddItem (MapText("แสดงเฉพาะฝากขาย"))
   C.ItemData(1) = 1

   C.AddItem (MapText("แสดงเฉพาะไม่ฝากขาย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitCommissionOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เริ่มใช้"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่สิ้นสุด"))
   C.ItemData(3) = 3

End Sub

Public Sub InitPartItemOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสต็อค"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อรายการสต็อค"))
   C.ItemData(2) = 2
End Sub

Public Sub InitEmptyCombo(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("น้อยไปมาก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มากไปน้อย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReportUnitType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หน่วยหลัก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("หน่วยย่อย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitAdjustType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปรับลด"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ปรับเพิ่ม"))
   C.ItemData(2) = 2
End Sub
Public Sub InitTransportDetailOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("คนขับ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ทะเบียน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ขนส่ง"))
   C.ItemData(3) = 3
End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub

Public Sub InitAPAR(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ลูกค้า")
   C.ItemData(1) = 1

   C.AddItem ("ผู้ค้า")
   C.ItemData(2) = 2
End Sub
Public Sub InitShortCode(C As ComboBox)
   C.Clear
   
   C.AddItem ("เฉพาะห้างสรรพสินค้า")
   C.ItemData(0) = 1
   
   C.AddItem ("ไม่รวมห้างสรรพสินค้า")
   C.ItemData(1) = 2

   C.AddItem ("แสดงทั้งหมด")
   C.ItemData(2) = 0
End Sub

Public Sub InitDrCr(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("เดบิต")
   C.ItemData(1) = 1

   C.AddItem ("เครดิต")
   C.ItemData(2) = 2
End Sub

Public Sub InitSellType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("สินค้า")
   C.ItemData(1) = 1

   C.AddItem ("บริการ")
   C.ItemData(2) = 2

   C.AddItem ("กำหนดเอง")
   C.ItemData(3) = 3
End Sub
Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("GROUP_RIGHT_ID", -1)
   Call D.SetFieldValue("GROUP_ID", GroupID)
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("RIGHT_ITEM_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("GROUP_RIGHT_ID")
      End If
      
      'Debug.Print TempData.GetFieldValue("RIGHT_ID") & "-" & TempData.GetFieldValue("RIGHT_ITEM_NAME")
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadGLAccount(Mr As CGLAccount, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CGLAccount
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGLAccount
Dim I As Long

   Set Rs = New ADODB.Recordset

   Call Mr.SetFieldValue("GL_ACCOUNT_ID", -1)
   Call Mr.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGLAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("ACC_CODE") & "-" & TempData.GetFieldValue("ACC_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("GL_ACCOUNT_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("GL_ACCOUNT_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional StockType As Long = -1, Optional ChkStd As String = "")
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCode
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCode
   
   TempData.STOCK_CODE_ID = -1
   TempData.STOCK_TYPE = StockType
   TempData.EXCEPTION_FLAG = "N"
   TempData.CHK_STD_COST = ChkStd
   Call TempData.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCode
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STOCK_DESC)
         C.ItemData(I) = TempData.STOCK_CODE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.STOCK_CODE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockCodeFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromStock As String, Optional ToStock As String, Optional TempType As Long = 0, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCode
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCode
   
   TempData.FROM_STOCK_NO = FromStock
   TempData.TO_STOCK_NO = ToStock
   If OUTLAY_FLAG = 0 Then
      TempData.OUTLAY_FLAG = "N"
   End If
   Call TempData.QueryData(7, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCode
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STOCK_DESC)
         C.ItemData(I) = TempData.STOCK_CODE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If TempType = 1 Then
            Call Cl.add(TempData, Trim(TempData.STOCK_NO))
         Else
            Call Cl.add(TempData)
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockBarcode(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCode
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCode
   
   TempData.STOCK_CODE_ID = -1
   TempData.EXCEPTION_FLAG = "N"
   Call TempData.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
     
      Set TempData = New CStockCode
      Call TempData.PopulateFromRS(5, Rs)
   
      If TempData.BARCODE > 0 Then
          I = I + 1
         If Not (C Is Nothing) Then
            C.AddItem (TempData.STOCK_DESC)
            C.ItemData(I) = TempData.STOCK_CODE_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.STOCK_CODE_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockCodeChange(Cl As Collection)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCodeChange
Dim I As Long
   
   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCodeChange
   
   TempData.STOCK_CODE_CHANGE_ID = -1
   Call TempData.QueryData(1, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCodeChange
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_ID & "-" & TempData.STOCK_CODE_ID & "-" & TempData.UNIT_CHANGE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadStockCodeChange2(Cl As Collection)
On Error Resume Next
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCodeChange
Dim I As Long
   
   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCodeChange
   
   TempData.STOCK_CODE_CHANGE_ID = -1
   Call TempData.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCodeChange
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_CODE_ID & "-" & TempData.UNIT_CHANGE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = err.Description
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEnterpriseAddress(Ad As CAddress, C As ComboBox, Optional Cl As Collection = Nothing, Optional ShowFirst As Boolean = True)
Dim D As CAddress
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long

   TempIndex = 0
   Set Rs = New ADODB.Recordset
   
   Call Ad.SetFieldValue("ADDRESS_ID", -1)
   Call Ad.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.GetFieldValue("ADDRESS_ID")
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadApArMasAddress(Ad As CAddress, C As ComboBox, Optional Cl As Collection = Nothing, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long
Dim Mask As Long
   TempIndex = 0
   Mask = 0
   Set Rs = New ADODB.Recordset
   
   Call Ad.SetFieldValue("ADDRESS_ID", -1)
   Call Ad.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.GetFieldValue("ADDRESS_ID")
         If TempData.GetFieldValue("MAIN_FLAG") = "Y" Then
            Mask = I
            C.ListIndex = I
         End If
      End If
      If (I > 0) And ShowFirst And Mask = 0 And Not (C Is Nothing) Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("ADDRESS_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportChequeBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เช็ค"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่เช็ค"))
   C.ItemData(2) = 2
End Sub

Public Sub LoadUpdateRcpCnDn(Rcp As CRcpCnDn_Item, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CRcpCnDn_Item
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRcpCnDn_Item
Dim I As Long

   Set Rs = New ADODB.Recordset

   
   Call Rcp.QueryData(2, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRcpCnDn_Item
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PAID_AMOUNT"))
         C.ItemData(I) = TempData.GetFieldValue("DOC_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("DOC_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportS_1_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("คนขับ และ ขนส่ง"))
   C.ItemData(1) = 2
End Sub
Public Sub InitReportS_1_1_6Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1
   C.AddItem (MapText("วันที่"))
   C.ItemData(2) = 2
   C.AddItem (MapText("สาขา"))
   C.ItemData(3) = 3
End Sub
Public Sub InitReportS_1_3Orderby(C As ComboBox)   'สำหรับการพิมพ์เอกสารเป็นชุด เช่น ใช้ใน class CReportNormalRcp001_3 ใบรับคืนสินค้าเป็นชุด
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 4
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 3
End Sub

Public Sub InitReportS_2_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("สาขา"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("ยอดขาย"))
   C.ItemData(2) = 2

End Sub

Public Sub InitReportS_2_21Orderby(C As ComboBox) ' สำหรับ  ROOT_TREE & " S-2-21" / InitReportS_2_15   / CReportBilling021
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
'   C.AddItem (MapText("ตามวันที่"))
'   C.ItemData(1) = 1
   
  
      
   C.AddItem (MapText("ตามวันที่ "))
   C.ItemData(1) = 1
    C.AddItem (MapText("ตามเลขที่"))
   C.ItemData(2) = 2
  C.AddItem (MapText("ตามเอกสารย่อย/เลขที่/เลขที่ "))
  C.ItemData(3) = 5

End Sub
Public Sub InitReport6_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("รหัสสินค้า/วัตถุดิบ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ชื่อสินค้า/วัตถุดิบ"))
   C.ItemData(2) = 2
   
End Sub
Public Sub LoadDistinctBlockBranchAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long

   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BILLING_DOC_ID", ID)
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("BLOCK_ID"))
         C.ItemData(I) = TempData.GetFieldValue("BLOCK_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("BLOCK_ID") & "-" & TempData.GetFieldValue("PART_ITEM_ID") & "-" & TempData.GetFieldValue("BRANCH_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctBlockBranch(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long

   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BILLING_DOC_ID", ID)
   Call D.QueryData(3, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("BLOCK_ID"))
         C.ItemData(I) = TempData.GetFieldValue("BLOCK_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("BLOCK_ID") & "-" & TempData.GetFieldValue("BRANCH_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctBlock(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long

   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BILLING_DOC_ID", ID)
   Call D.QueryData(6, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("BLOCK_ID"))
         C.ItemData(I) = TempData.GetFieldValue("BLOCK_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("BLOCK_ID") & "-" & TempData.GetFieldValue("PART_ITEM_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadDistinctLabelPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long

   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("BILLING_DOC_ID", ID)
   Call D.QueryData(4, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PART_ITEM_ID"))
         C.ItemData(I) = TempData.GetFieldValue("PART_ITEM_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("PART_ITEM_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctLabelPartItemEx(Cl As Collection, Optional FromDate As Date = 1, Optional ToDate As Date = -1, Optional InCludeFree As Long = -1, Optional EmpId As Long = -1, Optional FromAparCode As String = "", Optional ToAparCode As String = "")
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long
   
   MasterInd = "10"
   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.ORDER_TYPE = 2
   D.FREE_FLAG = StringToFreeFlag(InCludeFree)
   D.EMP_ID = EmpId
   D.FROM_APAR_CODE = FromAparCode
   D.TO_APAR_CODE = ToAparCode
   Call D.QueryDataReport(10, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumBranchPartItem(Cl As Collection, Optional FromDate As Date = 1, Optional ToDate As Date = -1, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long

   MasterInd = "11"
   
   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call D.QueryDataReport(11, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(11, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMP_ID & "-" & TempData.BRANCH_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitSaleType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("พนักงานขายคิดจากราคาขายเป็นหลัก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("พนักงานขายคิดจากจำนวนขายเป็นหลัก"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("หัวหน้าพนักงานขาย"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("พนักงานขายคิดจากเฉพาะจำนวนขาย"))
   C.ItemData(4) = 4
   
End Sub
Public Function IdToSaleType(ID As Long) As String
   If ID = 1 Then
      IdToSaleType = "พนักงานขายคิดจากราคาขายเป็นหลัก"
   ElseIf ID = 2 Then
      IdToSaleType = "พนักงานขายคิดจากจำนวนขายเป็นหลัก"
   ElseIf ID = 3 Then
      IdToSaleType = "หัวหน้าพนักงานขาย"
   ElseIf ID = 4 Then
      IdToSaleType = "พนักงานขายคิดจากเฉพาะจำนวนขาย"
   End If
End Function
Public Sub LoadCommissionChart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FK_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionChart
Dim I As Long
   If FK_ID <= 0 Then
      Exit Sub
   End If
   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("COMMISSION_CHART_ID", -1)
   Call D.SetFieldValue("MASTER_FROMTO_ID", FK_ID)
   Call D.SetFieldValue("ORDER_TYPE", 1)
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("EMP_NAME") & " " & TempData.GetFieldValue("EMP_LNAME"))
         C.ItemData(I) = TempData.GetFieldValue("COMMISSION_CHART_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("COMMISSION_CHART_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCommissionChartEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FK_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionChart
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionChart
Dim I As Long
   If FK_ID <= 0 Then
      Exit Sub
   End If
   Set D = New CCommissionChart
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("COMMISSION_CHART_ID", -1)
   Call D.SetFieldValue("MASTER_FROMTO_ID", FK_ID)
   Call D.SetFieldValue("ORDER_TYPE", 1)
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionChart
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("EMP_NAME") & " " & TempData.GetFieldValue("EMP_LNAME"))
         C.ItemData(I) = TempData.GetFieldValue("COMMISSION_CHART_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("OLD_PK"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitPackageOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("รหัสการตั้งราคา"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("รายละเอียด"))
   C.ItemData(2) = 2

End Sub
Public Sub LoadPackage(Package As CPackage, C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterFlag As String = "N")
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackage
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Package.SetFieldValue("PACKAGE_ID", -1)
   Call Package.SetFieldValue("PACKAGE_MASTER_FLAG", MasterFlag)
   Call Package.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackage
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PACKAGE_NO") & " " & TempData.GetFieldValue("PACKAGE_DESC"))
         C.ItemData(I) = TempData.GetFieldValue("PACKAGE_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("PACKAGE_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPackageDetail(PackageDetail As CPackageDetail, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackageDetail
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call PackageDetail.SetFieldValue("PACKAGE_DETAIL_ID", -1)
   Call PackageDetail.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackageDetail
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PACKAGE_DETAIL_ID"))
         C.ItemData(I) = TempData.GetFieldValue("PACKAGE_DETAIL_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportS_2_16Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("ตามปริมาณขาย"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("ตามยอดขาย"))
   C.ItemData(2) = 2

End Sub
Public Sub InitReportNullOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
End Sub

Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5

C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7

C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8

C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9

C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10

C.AddItem (MapText("พฤศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub

Public Sub LoadCommissionTotalSum(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DocumentTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   Call D.QueryData(5, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_BY & "-" & TempData.GROUP_COM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCommissionTabelPrice(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommissionType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMasterFromToDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterFromToDetail
Dim I As Long

   Set D = New CMasterFromToDetail
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("VALID_FROM", FromDate)
   Call D.SetFieldValue("VALID_TO", ToDate)
   Call D.SetFieldValue("MASTER_FROMTO_TYPE", CommissionType)
   
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterFromToDetail
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(I)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCommissionTabelEx(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CommissionType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMasterFromToEx
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterFromToEx
Dim I As Long
   
   Set D = New CMasterFromToEx
   Set Rs = New ADODB.Recordset
   
   D.VALID_FROM = FromDate
   D.VALID_TO = ToDate
   D.MASTER_FROMTO_TYPE = CommissionType
   
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterFromToEx
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDocItem(Mr As CDocItem, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDocItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDocItem
Dim I As Long

   Set Rs = New ADODB.Recordset

   Call Mr.SetFieldValue("DOC_ITEM_ID", -1)
   Call Mr.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDocItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("DOC_ITEM_ID"))
         C.ItemData(I) = TempData.GetFieldValue("DOC_ITEM_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("DOC_ITEM_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumTop(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = 1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long
   
   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(12, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("EMP_ID") & "-" & TempData.GetFieldValue("GROUP_COM_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumPrintLabelReturn(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = 1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CPrintLabel
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrintLabel
Dim I As Long
   
   Set D = New CPrintLabel
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(100, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrintLabel
      Call TempData.PopulateFromRS(100, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("EMP_ID") & "-" & TempData.GetFieldValue("GROUP_COM_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitTagetOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เดือนปี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ชื่อพนักงาน"))
   C.ItemData(3) = 3
End Sub
Public Sub InitTagetJobOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เดือนปี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รายละเอียดวัตถุดิบ"))
   C.ItemData(3) = 3
End Sub

Public Sub LoadTaget(Cl As Collection, YYYYMM As String, Optional TopFlag As String = "")
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Set TempData = New CTagetDetail
   
   Call TempData.SetFieldValue("TAGET_ID", -1)
   Call TempData.SetFieldValue("YYYYMM", YYYYMM)
   TempData.TOP_FLAG = TopFlag
   Call TempData.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("EMP_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTagetByEmpGroupcom(Cl As Collection, YYYYMM As String, Optional TopFlag As String = "")
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Set TempData = New CTagetDetail
   
   Call TempData.SetFieldValue("TAGET_ID", -1)
   Call TempData.SetFieldValue("YYYYMM", YYYYMM)
   TempData.TOP_FLAG = TopFlag
   Call TempData.QueryData(13, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("EMP_ID") & "-" & TempData.GetFieldValue("GROUP_COM_ID")))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitImportType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("นำเข้าลูกหนี้ยกมา")
   C.ItemData(1) = 1
   
'   C.AddItem ("นำเข้าสต็อคยกมา")
'   C.ItemData(2) = 2
'
'   C.AddItem ("นำเข้าข้อมูลลูกหนี้")
'   C.ItemData(3) = 3
'
'   C.AddItem ("นำเข้าข้อมูล BALANCE ACCUM")
'   C.ItemData(4) = 4
End Sub
Public Sub LoadConfigDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CConfigDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CConfigDoc
Dim I As Long

   Set D = New CConfigDoc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CConfigDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      TempData.Flag = "I"
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.GetFieldValue("CONFIG_DOC_CODE"))
         C.ItemData(I) = TempData.GetFieldValue("CONFIG_DOC_TYPE")
      End If

      If Not (Cl Is Nothing) Then
         ''debug.print (Trim(Str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
      End If
            
      Set TempData = Nothing
      
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitDocItemType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปกติ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ไม่แสดง"))
   C.ItemData(2) = 2
   
'   C.AddItem (MapText("ไม่แสดง"))
'   C.ItemData(3) = 3
End Sub
Public Sub LoadSaleAmount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "12"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   Call BD.QueryData(12, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_MAS_ID & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctBranchEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(40, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(41, Rs1, itemcount)

   MasterInd = "41"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 40, Rs1, 41, 4)
      
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctBranchEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(43, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
      
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(44, Rs1, itemcount)

   MasterInd = "44"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 43, Rs1, 44, 5)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long


   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(47, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "47"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(47, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
      
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub GetSaleAmountBranchEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(37, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "37"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(37, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountBranchEmpAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(42, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "42"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(42, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(45, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "45"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(45, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(39, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "39"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(39, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "39"
End Sub
Public Sub GetSaleAmountBranchEmpStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(46, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "46"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(46, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTagetDetailBranchEmpAparStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(3, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "3"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BRANCH_CODE & "-" & TempData.EMPLOYEE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailBranchEmpStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(6, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "6"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BRANCH_CODE & "-" & TempData.EMPLOYEE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String = "", Optional ToSaleCode As String = "", Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(7, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = 7
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailStockType(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String = "", Optional ToSaleCode As String = "", Optional FromStockNo As String, Optional ToStockNo As String, Optional userAccess As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.SALE_CODE_ACCESS = userAccess
   
   Call BD.QueryData(15, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = 15
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(15, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_TYPE_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctBranchAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(48, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(49, Rs1, itemcount)
   
   MasterInd = "49"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 48, Rs1, 49, 6)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailBranchAparStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(8, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "8"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(8, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BRANCH_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(50, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "50"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(50, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(51, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "51"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(51, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountBranchAparStockCodeDocTypeFreeyyyymm(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(110, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
      
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(110, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchAparStockCodeDocTypeFreeExyyyymm(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(111, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(111, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctBranchStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(52, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(53, Rs1, itemcount)
   
   MasterInd = "53"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 52, Rs1, 53, 7)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailBranchStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(9, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "9"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BRANCH_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(54, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "54"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(54, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountBranchStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(55, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "55"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(55, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(56, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(57, Rs1, itemcount)

   MasterInd = "57"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 56, Rs1, 57, 3)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(147, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate

   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(148, Rs1, itemcount)

   MasterInd = "148"
   
   Set BD = Nothing
   
   Call GetDataToRsPopulate2(Cl, Rs, 147, Rs1, 148, 16)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpDocTypeFreeForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim TempData As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(159, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = Nothing
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(159, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_JOINT_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDebtorAging(Cl As Collection, Cl2 As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional ShortCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset

'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.BILLING_DOC_ID = -1
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   If ShortCode = 0 Then
      ShortCode = ""
   End If
   BD.SHORT_CODE = ShortCode
   Call BD.QueryData(145, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.BILLING_DOC_ID = -1
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.SHORT_CODE = ShortCode
   Call BD.QueryData(146, Rs1, itemcount)

   MasterInd = "146"
   
   Set BD = Nothing

   If Rs.RecordCount = 0 Then
      While Not Rs1.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(146, Rs1)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData)
         End If
         
         Set TempData = Nothing
         Rs1.MoveNext
      Wend
   ElseIf Rs1.RecordCount = 0 Then
      While Not Rs.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(145, Rs)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData)
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   Else
      Call GetDataToRs(Cl, Rs, 145, Rs1, 146, 15)
   End If
   'สำหรับตรวจสอบ
   Set Rs = Nothing
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.BILLING_DOC_ID = -1
'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.SHORT_CODE = ShortCode
    Call BD.QueryData(13, Rs, itemcount)
   
   If Not (Cl2 Is Nothing) Then
      Set Cl2 = Nothing
      Set Cl2 = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl2 Is Nothing) Then
         Call Cl2.add(TempData, Trim(TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(135, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(136, Rs1, itemcount)

   MasterInd = "136"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 135, Rs1, 136, 11)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpGroupStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(151, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(152, Rs1, itemcount)

   MasterInd = "136"
   
   Set BD = Nothing
   
   Call GetDataToRsPopulate2(Cl, Rs, 151, Rs1, 152, 17)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadNote(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String = "", Optional ToAparCode As String = "", Optional FromSaleCode As String = "", Optional ToSaleCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim TempData As CBillingDoc
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(130, Rs, itemcount)
   
   Set BD = Nothing
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(144, Rs1, itemcount)
   
   Set BD = Nothing
   Call GetDataToRs(Cl, Rs, 130, Rs1, 144, 14)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpAparDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1, Optional FromStockNo As String, Optional ToStockNo As String, Optional userAccess As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.SALE_CODE_ACCESS = userAccess
   Call BD.QueryData(124, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(128, Rs1, itemcount)

   MasterInd = "128"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 124, Rs1, 128, 10)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctStockGroupDocTypeFree(Cl As Collection, Cl2 As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1, Optional FromStockNo As String, Optional ToStockNo As String, Optional userAccess As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc
Dim TempStockType As CTagetDetail

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   BD.SALE_CODE_ACCESS = userAccess
   Call BD.QueryData(125, Rs, itemcount)
   
   Set BD = Nothing

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(139, Rs1, itemcount)
   
   Call GetDataToRs(Cl, Rs, 125, Rs1, 139, 12)
   
   For Each TempStockType In Cl2
      Set TempData = GetObject("CBillingDoc", Cl, TempStockType.STOCK_TYPE_CODE, False)
      If TempData Is Nothing Then
         Set TempData = New CBillingDoc
         TempData.STOCK_TYPE_CODE = TempStockType.STOCK_TYPE_CODE
         TempData.STOCK_GROUP_NAME = TempStockType.STOCK_GROUP_NAME
         Call Cl.add(TempData, Trim(TempData.STOCK_TYPE_CODE))
         Set TempData = Nothing
      End If
   Next TempStockType

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctStockCodeDocTypeFree2(Cl As Collection, Cl2 As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc
Dim TempStock As CTagetDetail

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(140, Rs, itemcount)
   
   Set BD = Nothing

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(141, Rs1, itemcount)
   
   Call GetDataToRs(Cl, Rs, 140, Rs1, 141, 13)

   For Each TempStock In Cl2
      Set TempData = GetObject("CBillingDoc", Cl, TempStock.STOCK_NO, False)
      If TempData Is Nothing Then
         Set TempData = New CBillingDoc
         TempData.STOCK_DESC = TempStock.STOCK_DESC
         TempData.BILL_DESC = TempStock.BILL_DESC
         TempData.STOCK_NO = TempStock.STOCK_NO
         Call Cl.add(TempData, Trim(TempData.STOCK_NO))
         Set TempData = Nothing
      End If
   Next TempStock

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeGroupName(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(118, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(119, Rs1, itemcount)

   MasterInd = "119"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 118, Rs1, 119, 8)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeGroupName2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(120, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(121, Rs1, itemcount)

   MasterInd = "121"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 120, Rs1, 121, 9)
   
   MasterInd = "1"

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailEmpAparStock(Cl As Collection, Optional YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   Call BD.QueryData(10, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "10"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMPLOYEE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailEmp(Cl As Collection, Optional YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   Call BD.QueryData(16, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "16"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(16, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMPLOYEE_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(58, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "58"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(58, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(59, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "59"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(59, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "86"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(86, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(86, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional userAccess As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "126"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.SALE_CODE_ACCESS = userAccess
   Call BD.QueryData(126, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(126, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_TYPE_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "142"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(142, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(142, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "87"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(87, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(87, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional userAccess As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "127"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.SALE_CODE_ACCESS = userAccess
   Call BD.QueryData(127, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(127, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_TYPE_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "143"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(143, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(143, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
'   Dim TempSum As Double
'   Dim TempSum2 As Double
'   Dim TempSum3 As Double
   
   MasterInd = "116"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(116, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(116, Rs)
      
'      TempSum = TempSum + TempData.TOTAL_PRICE
'      TempSum2 = TempSum2 + TempData.DISCOUNT_AMOUNT
'      TempSum3 = TempSum3 + TempData.EXT_DISCOUNT_AMOUNT
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   'debug.print (TempSum & "------" & TempSum2 & "----------" & TempSum3)
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmp(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "149"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(149, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(149, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "160"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(160, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(160, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_JOINT_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeDateFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "137"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(137, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(137, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional CheckQMC_To_FAndB As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "153"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If ConsignmentFlag = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.DOCUMENT_TYPE = PO_DOCTYPE
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
      BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   End If
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(153, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(153, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeDateForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "153"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If ConsignmentFlag = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.DOCUMENT_TYPE = PO_DOCTYPE
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
      BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   End If
   BD.ORDER_TYPE = 9999
'   BD.FROM_APAR_CODE = FromAparCode
'   BD.TO_APAR_CODE = ToAparCode
'   BD.FROM_STOCK_NO = FromStockNo
'   BD.TO_STOCK_NO = ToStockNo
'   BD.FROM_SALE_CODE = FromSaleCode
'   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(153, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(153, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
'   Dim TempSum As Double
'   Dim TempSum2 As Double
'   Dim TempSum3 As Double
   
   MasterInd = "122"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(122, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(122, Rs)
      
'      TempSum = TempSum + TempData.TOTAL_PRICE
'      TempSum2 = TempSum2 + TempData.DISCOUNT_AMOUNT
'      TempSum3 = TempSum3 + TempData.EXT_DISCOUNT_AMOUNT
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   ''debug.print (TempSum & "------" & TempSum2 & "----------" & TempSum3)
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "117"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(117, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(117, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "150"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(150, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(150, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpDateFreeForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "161"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(161, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(161, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_JOINT_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "138"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(138, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(138, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeExDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional CheckQMC_To_FAndB As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "154"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   If ConsignmentFlag = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.DOCUMENT_TYPE = PO_DOCTYPE
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
      BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   End If
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(154, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(154, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeExDateForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "154"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   If ConsignmentFlag = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.DOCUMENT_TYPE = PO_DOCTYPE
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
      BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   End If
   BD.ORDER_TYPE = 9999
'   BD.FROM_APAR_CODE = FromAparCode
'   BD.TO_APAR_CODE = ToAparCode
'   BD.FROM_STOCK_NO = FromStockNo
'   BD.TO_STOCK_NO = ToStockNo
'   BD.FROM_SALE_CODE = FromSaleCode
'   BD.TO_SALE_CODE = ToSaleCode
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(154, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(154, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "123"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(123, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(123, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(60, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(61, Rs1, itemcount)
   
   MasterInd = "61"
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 60, Rs1, 61, 1)
   
   MasterInd = "1"
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctEmpStockCodeDocTypeFreeForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(158, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(158, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDataToRs(Cl As Collection, Rs As ADODB.Recordset, Ind1 As Long, Rs1 As ADODB.Recordset, Ind2 As Long, KeyType As Long)
Dim Af As Boolean
Dim Bf As Boolean
Dim TempData As CBillingDoc
Dim TempData1 As CBillingDoc
Dim TempDataTest As CBillingDoc

Dim I As Long
Dim j As Long
   Af = True
   Bf = True
    I = 1
    j = 1
    
   While Not (Rs.EOF And Rs1.EOF)
      If Af Then
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(Ind1, Rs)
      End If
      If Bf Then
         Set TempData1 = New CBillingDoc
         Call TempData1.PopulateFromRS(Ind2, Rs1)
      End If

      Af = False
      Bf = False
      
'      If I = 23 And j = 47 Then
'         Debug.Print
'      End If
      
      If Rs.RecordCount >= I And Rs1.RecordCount < j Then
        Set TempDataTest = GetObject("CBillingDoc", Cl, GetDataKeyRs(TempData, KeyType), False)
        If TempDataTest Is Nothing Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
         End If
         Set TempData = Nothing
         If Not (Rs.EOF) Then
            Rs.MoveNext
            Af = True
            I = I + 1
         End If
      ElseIf Rs.RecordCount < I And Rs1.RecordCount >= j Then
         Set TempDataTest = GetObject("CBillingDoc", Cl, GetDataKeyRs(TempData1, KeyType), False)
        If TempDataTest Is Nothing Then
            Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
         End If
         Set TempData1 = Nothing
         If Not (Rs1.EOF) Then
            Rs1.MoveNext
            Bf = True
            j = j + 1
         End If
      Else
         If GetDataKeyRs(TempData, KeyType) < GetDataKeyRs(TempData1, KeyType) Then
            Set TempDataTest = GetObject("CBillingDoc", Cl, GetDataKeyRs(TempData, KeyType), False)
            If TempDataTest Is Nothing Then
               Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            End If
            Set TempData = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) > GetDataKeyRs(TempData1, KeyType) Then
            Set TempDataTest = GetObject("CBillingDoc", Cl, GetDataKeyRs(TempData1, KeyType), False)
            If TempDataTest Is Nothing Then
               Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
            End If
            Set TempData1 = Nothing
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) = GetDataKeyRs(TempData1, KeyType) Then
            Set TempDataTest = GetObject("CBillingDoc", Cl, GetDataKeyRs(TempData, KeyType), False)
            If TempDataTest Is Nothing Then
               Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            End If
            Set TempData = Nothing
            Set TempData1 = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         End If
      End If
      'Debug.Print I & "-" & j
   Wend

End Sub
Public Sub GetDataToRsPopulate2(Cl As Collection, Rs As ADODB.Recordset, Ind1 As Long, Rs1 As ADODB.Recordset, Ind2 As Long, KeyType As Long)
Dim Af As Boolean
Dim Bf As Boolean
Dim TempData As CBillingDoc
Dim TempData1 As CBillingDoc
Dim I As Long
Dim j As Long
   Af = True
   Bf = True
    I = 1
    j = 1
    
   While Not (Rs.EOF And Rs1.EOF)
      If Af Then
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS2(Ind1, Rs)
      End If
      If Bf Then
         Set TempData1 = New CBillingDoc
         Call TempData1.PopulateFromRS2(Ind2, Rs1)
      End If
      Af = False
      Bf = False
      If Rs.RecordCount >= I And Rs1.RecordCount < j Then
         Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
         Set TempData = Nothing
         If Not (Rs.EOF) Then
            Rs.MoveNext
            Af = True
            I = I + 1
         End If
      ElseIf Rs.RecordCount < I And Rs1.RecordCount >= j Then
         Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
         Set TempData1 = Nothing
         If Not (Rs1.EOF) Then
            Rs1.MoveNext
            Bf = True
            j = j + 1
         End If
      Else
         If GetDataKeyRs(TempData, KeyType) < GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) > GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
            Set TempData1 = Nothing
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) = GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            Set TempData1 = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         End If
      End If
   Wend

End Sub
Public Function GetDataKeyRs(DataGet As Object, KeyType As Long) As String
   If KeyType = 1 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 2 Then
      GetDataKeyRs = Trim(DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 3 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.STOCK_NO)
'      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 4 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 5 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.SALE_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 6 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 7 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 8 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 9 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.SALE_CODE & "-" & DataGet.STOCK_GROUP_NAME & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 10 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.APAR_CODE)
   ElseIf KeyType = 11 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 12 Then
      GetDataKeyRs = Trim(DataGet.STOCK_TYPE_CODE)
   ElseIf KeyType = 13 Then
      GetDataKeyRs = Trim(DataGet.STOCK_NO)
   ElseIf KeyType = 14 Then
      GetDataKeyRs = Trim(DataGet.DOCUMENT_NO & "-" & DataGet.DOCUMENT_DATE)
   ElseIf KeyType = 15 Then
'      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.DOC_ID_BILLS_NO & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.DOCUMENT_NO & "-" & DataGet.DUE_DATE)
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.DOCUMENT_NO & "-" & DataGet.Due_Date)
   ElseIf KeyType = 16 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE)
   ElseIf KeyType = 17 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.STOCK_TYPE_CODE & "-" & DataGet.STOCK_NO)
   Else
   End If
End Function
Public Sub LoadTagetDetailEmpStock(Cl As Collection, Optional YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   Call BD.QueryData(11, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "11"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(11, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMPLOYEE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailSaleCode(Cl As Collection, Optional YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   Call BD.QueryData(17, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "17"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(17, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMPLOYEE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailSaleCode2(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(18, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "18"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(18, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMPLOYEE_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailEmpStockGroupName(Cl As Collection, Optional YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   Call BD.QueryData(14, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "14"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(14, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.EMPLOYEE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode

   Call BD.QueryData(62, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "62"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(62, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode

   Call BD.QueryData(63, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "59"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(63, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim TempData As CBillingDoc
Dim Rs As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(64, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(64, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   Set BD = Nothing
   
   MasterInd = "1"
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadTagetDetailAparStock(Cl As Collection, Optional YYYYMM As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CTagetDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetDetail
Dim I As Long
   
   Set BD = New CTagetDetail
   Set Rs = New ADODB.Recordset
   
   Call BD.SetFieldValue("YYYYMM", YYYYMM)
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(12, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "12"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetDetail
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(66, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "66"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(66, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctBillingAddition(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingAddition
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingAddition
Dim I As Long
   
   MasterInd = 2
   Set BD = New CBillingAddition
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingAddition
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumBillingAdditionID(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingAddition
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingAddition
Dim I As Long
   
   MasterInd = 3
   Set BD = New CBillingAddition
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(3, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingAddition
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.ADDITION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctBillingSubTract(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingSubTract
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingSubTract
Dim I As Long
   
   MasterInd = 2
   Set BD = New CBillingSubTract
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingSubTract
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumMovementPartItemType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "10"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.LOCATION_GROUP_NO = LocationGroupNo
   Call Lt.QueryData(10, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_TYPE & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.TEMP_SALE_FLAG))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumMovementPartItemTypeDocDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "10"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(42, Rs, itemcount, True)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(42, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_TYPE & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.TEMP_SALE_FLAG & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctTransferPartItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationID As Long = -1, Optional ToLocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "13"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.FROM_LOCATION_ID = FromLocationID
   Lt.TO_LOCATION_ID = ToLocationID
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If OUTLAY_FLAG = 0 Then
      Lt.OUTLAY_FLAG = "N"
   End If
   Call Lt.QueryData(13, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctTransferPartItemConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationCode As String = "", Optional FromLocationCode2 As String = "", Optional ToLocationCode As String = "", Optional ToLocationCode2 As String = "", Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional Consign As Long = -1, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "48"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.FROM_LOCATION_CODE = FromLocationCode
   Lt.TO_LOCATION_CODE = ToLocationCode
   Lt.FROM_LOCATION_CODE2 = FromLocationCode2
   Lt.TO_LOCATION_CODE2 = ToLocationCode2
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.CONSIGNMENT = Consign
   If OUTLAY_FLAG = 0 Then
      Lt.OUTLAY_FLAG = "N"
   End If
   Call Lt.QueryData(48, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(48, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumTransferPartItemDocDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationID As Long = -1, Optional ToLocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "12"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.FROM_LOCATION_ID = FromLocationID
   Lt.TO_LOCATION_ID = ToLocationID
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If OUTLAY_FLAG = 0 Then
      Lt.OUTLAY_FLAG = "N"
   End If
   Call Lt.QueryData(12, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetTransferPartItemDocDateConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationCode As String = "", Optional FromLocationCode2 As String = "", Optional ToLocationCode As String = "", Optional ToLocationCode2 As String = "", Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional Consign As Long = -1, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "49"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.FROM_LOCATION_CODE = FromLocationCode
   Lt.TO_LOCATION_CODE = ToLocationCode
   Lt.FROM_LOCATION_CODE2 = FromLocationCode2
   Lt.TO_LOCATION_CODE2 = ToLocationCode2
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.CONSIGNMENT = Consign
   If OUTLAY_FLAG = 0 Then
      Lt.OUTLAY_FLAG = "N"
   End If
   Call Lt.QueryData(49, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(49, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub GetSumBillingSubTractID(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim Bs As CBillingSubTract
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingSubTract
Dim I As Long
   
   MasterInd = "3"
   Set Bs = New CBillingSubTract
   Set Rs = New ADODB.Recordset
   
   Bs.FROM_DATE = FromDate
   Bs.TO_DATE = ToDate
   Call Bs.QueryDataReport(3, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingSubTract
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.SUBTRACT_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Function CashDocTypeToText(Ct As CASH_DOC_TYPE) As String
   If Ct = CHEQUE_REV Then
      CashDocTypeToText = "เช็ครับ"
   ElseIf Ct = CHEQUE_PAY Then
      CashDocTypeToText = "เช็คจ่าย"
   ElseIf Ct = CASH_DEPOSIT Then
      CashDocTypeToText = "ใบนำฝาก"
   ElseIf Ct = POST_CHEQUE Then
      CashDocTypeToText = "ใบ POST เช็ค"
   End If
End Function
Public Sub InitJobOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ JOB"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่ JOB"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันเวลา JOB"))
   C.ItemData(3) = 3
   
End Sub
Public Sub InitBalanceVerifyOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลข"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่"))
   C.ItemData(2) = 2
      
End Sub

Public Sub InitFormulaOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ สูตร"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่ สูตร"))
   C.ItemData(2) = 2
   
End Sub
Public Sub GetDistinctJobInItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional BatchNoSet As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.BATCH_NO_SET = BatchNoSet
   Call Ji.QueryData(3, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctJobOutItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(8, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(8, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctJobOutItemByPrdType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(9, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctJobOutItemByPrdTypeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(10, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetSumJobInOutLostItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Call Ji.QueryData(4, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BATCH_NO & "-" & TempData.INPUT_ID & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumJobOutLostItemByDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(7, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_DATE & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID & "-" & TempData.BATCH_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetSumJobInItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional BatchNoSet As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.BATCH_NO_SET = BatchNoSet
   Call Ji.QueryData(5, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumJobInOutLostItemByInPutDocDate(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromIvdDate As Date = -1, Optional ToIvdDate As Date = -1, Optional NoMainLocationPrdFlag As Boolean = False, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.FROM_IVD_DATE = FromIvdDate
   Ji.TO_IVD_DATE = ToIvdDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Ji.NoMainLocationPrdFlag = NoMainLocationPrdFlag
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(11, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(11, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INPUT_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID & "-" & TempData.WEIGHT_AMOUNT & "-" & TempData.PRODUCTION_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumJobInOutLostItemByInPutJobDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(15, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(15, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INPUT_ID & "-" & TempData.JOB_DATE & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID & "-" & TempData.PRODUCTION_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetJobInputByPartItemInPutDate(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromIvdDate As Date = -1, Optional ToIvdDate As Date = -1, Optional NoMainLocationPrdFlag As Boolean = False, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.FROM_IVD_DATE = FromIvdDate
   Ji.TO_IVD_DATE = ToIvdDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Ji.NoMainLocationPrdFlag = NoMainLocationPrdFlag
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Call Ji.QueryData(13, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.WEIGHT_AMOUNT))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetBalanceVerifyByDateLocationPartItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional StockNo As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationId As Long = -1)
On Error GoTo ErrorHandler
Dim Ji As CBalanceVerifyDeTail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceVerifyDeTail
Dim I As Long
   
   Set Ji = New CBalanceVerifyDeTail
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PART_NO = StockNo
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.LOCATION_ID = LocationId
   Call Ji.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceVerifyDeTail
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BALANCE_VERIFY_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctPartItemByDocTypeDocSubType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "14"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(14, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(14, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDateExportPartItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional ShowOutlay As Long)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   

   MasterInd = "52"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If ShowOutlay = 0 Then
       Lt.OUTLAY_FLAG = "N"
    End If
    

   
   
   Call Lt.QueryData(52, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(52, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData)
'      End If
'
    If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.YYYYMM))
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
  '       Debug.Print (Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.YYYYMM))
      End If

      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctPartItemByProduction(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional TxType = "")
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "14"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If TxType = "I" Then
      Call Lt.QueryData(36, Rs, itemcount, False)
   ElseIf TxType = "E" Then
      'ยังไม่มีใช้
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      If TxType = "I" Then
         Call TempData.PopulateFromRS(36, Rs)
      Else
         Call TempData.PopulateFromRS(36, Rs)
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumAmountByPartItemDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "15"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(15, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(15, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumAmountByPartItemMonthly(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "51"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(51, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(51, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.YYYYMM & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub GetSumAmountByPartItemDateProduction(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional TxType = "")
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "15"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If TxType = "I" Then
      Call Lt.QueryData(37, Rs, itemcount, False)
   ElseIf TxType = "E" Then
      'ยังไม่มีใช้
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      If TxType = "I" Then
         Call TempData.PopulateFromRS(37, Rs)
      Else
      
      End If
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumAmountByPartItemDateIndSub(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "18"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(18, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(18, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.INVENTORY_SUB_TYPE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub InitInventoryDocType(C As ComboBox)
Dim I As Long

   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   For I = IMPORT_DOCTYPE To ADJUST_DOCTYPE
      C.AddItem (Doctype2Text(I))
      C.ItemData(I) = I
   Next I
   
End Sub
Public Sub GetJobInputByJobNo(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Call Ji.QueryData(20, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(20, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_NO & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumJobInOutLostItemByJobNo(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(19, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(19, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_NO & "-" & TempData.PRODUCTION_TYPE & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctJobOutItemByPrdTypeIvdDate(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional FromIvdDate As Date = -1, Optional ToIvdDate As Date = -1, Optional FromStockNo1 As String, Optional ToStockNo1 As String)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.FROM_IVD_DATE = FromIvdDate
   Ji.TO_IVD_DATE = ToIvdDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.FROM_STOCK_NO1 = FromStockNo1
   Ji.TO_STOCK_NO1 = ToStockNo1
   Call Ji.QueryData(21, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(21, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumJobInOutLostItemByInPutDocDateMainFlag(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional PrdLocation As Long = -1, Optional BatchNoSet As String, Optional MainLocationPrdFlag As Boolean = False, Optional FromIvdDate As Date = -1, Optional ToIvdDate As Date = -1)
On Error GoTo ErrorHandler
Dim Ji As CJobItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobItem
Dim I As Long
   
   Set Ji = New CJobItem
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.FROM_IVD_DATE = FromIvdDate
   Ji.TO_IVD_DATE = ToIvdDate
   Ji.PRD_LOCATION_ID = PrdLocation
   Ji.BATCH_NO_SET = BatchNoSet
   Ji.MainLocationPrdFlag = MainLocationPrdFlag
   Call Ji.QueryData(22, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobItem
      Call TempData.PopulateFromRS(22, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INPUT_ID & "-" & TempData.DOCUMENT_DATE & "-" & TempData.OUTPUT_ID & "-" & TempData.LOST_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCodeDocType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockCode As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "75"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockCode
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(77, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(77, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountStockCodeDocTypeForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockCode As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "155"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
'   BD.FROM_STOCK_NO = FromStockCode
'   BD.TO_STOCK_NO = ToStockNo
'   BD.FROM_APAR_CODE = FromAparCode
'   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(155, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(155, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub GetDistinctStockCode_Y(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "80"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   BD.APM_CONSIGNMENT_FLAG = "Y"
   Call BD.QueryData(165, Rs, itemcount) '  '80
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(165, Rs)
      'Call TempData.PopulateFromRS(80, Rs)
            
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub GetDistinctStockCodeForFAndB_Y(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "156"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.APM_CONSIGNMENT_FLAG = "Y"
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   Call BD.QueryData(166, Rs, itemcount, , 2) ' 156 '
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(166, Rs)  ' 156 '166
      
      If Not (Cl Is Nothing) Then
'         If TempData.JOINT_CODE <> "" Then
'            Debug.Print TempData.JOINT_CODE
'         End If
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub


Public Sub GetDistinctStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "80"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(80, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(80, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctStockCodeForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "156"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If CONSIGNMENT = 1 Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
      BD.CONSIGNMENT_FLAG = "Y"
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(156, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(156, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountStockCodeDocTypeDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "75"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(85, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(85, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountAparTypeStockCodeDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   MasterInd = "84"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(84, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(84, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_TYPE_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountAparGroupStockCodeDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   MasterInd = "91"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(91, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(91, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctBillBankAccount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
      
   MasterInd = "3"
   
   Set BD = New CCashTran
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(3, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetTransferAmountBillBankAccount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CCashTran
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
      
   MasterInd = "4"
   
   Set BD = New CCashTran
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(4, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.BANK_ACCOUNT))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetBalanceByLotItemLinkDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "24"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.COUNT_AMOUNT = "Y"
   Call Lt.QueryData(24, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(24, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDocItemIDLinkLotItemID(Cl As Collection, Optional LocationId As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional PartItem As Long, Optional ChkStd As String, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "23"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.LOCATION_ID = LocationId
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.CHK_STD_COST = ChkStd
   Lt.PART_ITEM_ID = PartItem
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Call Lt.QueryData(23, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(23, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.LOT_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountMonthByStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "24"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   
   Call BD.QueryData(24, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(24, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
        
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountMonth(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "21"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT2_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   Call BD.QueryData(21, Rs, itemcount)
      
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(21, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_MAS_ID & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountMonthByCustomerStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1, Optional ApArInd As Byte = 1)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   MasterInd = "27"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.APAR_IND = ApArInd
   If ApArInd = 1 Then
      BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   ElseIf ApArInd = 2 Then
      BD.DOCUMENT_TYPE_SET = "(" & S_INVOICE_DOCTYPE & "," & S_RECEIPT1_DOCTYPE & "," & S_RETURN_DOCTYPE & ")"
   End If
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(27, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(27, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   MasterInd = "1"
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional UnitType As Long = -1, Optional TotalSalePrice As Double, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim CreditStockCode As CCreditStockCode
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(25, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(25, Rs)
      
      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT
            
            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT
            
            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE
            
            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT
   
         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT
            
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE
            
            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE
            
            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT
            
         End If
      End If
      
      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCode2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional UnitType As Long = -1, Optional TotalSalePrice As Double, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1, Optional SaleCode As String = "", Optional FromAparCode As String = "", Optional ToAparCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc
Dim CreditStockCode As CCreditStockCode
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(131, Rs, itemcount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = Nothing

   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(132, Rs1, itemcount)

   MasterInd = "132"
   
   Set BD = Nothing

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(131, Rs)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT

            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs.MoveNext
   Wend

   While Not Rs1.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(132, Rs1)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT

            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs1.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountAparCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional UnitType As Long = -1, Optional TotalSalePrice As Double, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1, Optional SaleCode As String = "", Optional FromAparCode As String = "", Optional ToAparCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc
Dim CreditStockCode As CCreditStockCode
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(133, Rs, itemcount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = Nothing

   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(134, Rs1, itemcount)

   MasterInd = "134"
   
   Set BD = Nothing

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(133, Rs)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.APAR_NAME = TempData.APAR_NAME
            CreditStockCode.CUSTOMER_BRANCH_NAME = TempData.CUSTOMER_BRANCH_NAME
            CreditStockCode.CUSTOMER_BRANCH_CODE = TempData.CUSTOMER_BRANCH_CODE

            Call Cl.add(CreditStockCode, Trim(TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH))
         End If
         CreditStockCode.APAR_CODE = TempData.APAR_CODE
         CreditStockCode.CUSTOMER_BRANCH = TempData.CUSTOMER_BRANCH
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs.MoveNext
   Wend

   While Not Rs1.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(134, Rs1)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.APAR_NAME = TempData.APAR_NAME
            CreditStockCode.CUSTOMER_BRANCH_NAME = TempData.CUSTOMER_BRANCH_NAME
            CreditStockCode.CUSTOMER_BRANCH_CODE = TempData.CUSTOMER_BRANCH_CODE

            Call Cl.add(CreditStockCode, Trim(TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH))
         End If
         CreditStockCode.APAR_CODE = TempData.APAR_CODE
         CreditStockCode.CUSTOMER_BRANCH = TempData.CUSTOMER_BRANCH
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs1.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   MasterInd = "1"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(29, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(29, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.FREE_FLAG))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(36, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   MasterInd = "36"
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(36, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.FREE_FLAG))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountAparStockCodeDocTypeFreeDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(169, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   MasterInd = "169"
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(169, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.FREE_FLAG & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetSaleAmountAparGroupDocTypeStockCodeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   MasterInd = "75"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(75, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(75, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountAparTypeStockCodeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   MasterInd = "78"
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(78, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(78, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_TYPE_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountDocumentNo(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromLocationNo As String, Optional ToLocationNo As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1) ', Optional orderBy As Long = 1
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_LOCATION_NO = FromLocationNo
   BD.TO_LOCATION_NO = ToLocationNo
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
'  BD.ORDER_BY = orderBy   ' เพิ่มตอนแก้ไขCReportBillingDoc021 /ROOT_TREE&"S-2-21"/InitReportS_2_21
   Call BD.QueryData(95, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(95, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description & " " & TempData.DOCUMENT_NO & " ซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำ"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetAmountForTranSport(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentTypeSet As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim Ind As CInventoryDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim TempData2 As CInventoryDoc
Dim I As Long
      
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If DocumentTypeSet = "(" & PO_DOCTYPE & ")" Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   Call BD.QueryData(97, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(97, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID & "-" & TempData.APAR_MAS_ID & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Ind = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   
   Ind.FROM_DATE = FromDate
   Ind.TO_DATE = ToDate
   Call Ind.QueryData2(4, Rs, itemcount)
   
   While Not Rs.EOF
      Set TempData2 = New CInventoryDoc
      Call TempData2.PopulateFromRS2(4, Rs)
      
      Set TempData = GetObject("CBillingDoc", Cl, Trim(TempData2.DRIVER_ID & "-" & TempData2.CAR_LICENSE_ID & "-" & TempData2.TRANSPORTOR_ID & "-" & TempData2.APAR_MAS_ID & "-" & TempData2.CUSTOMER_BRANCH & "-" & TempData2.PART_ITEM_ID), False)
      If TempData Is Nothing Then
         Set TempData = New CBillingDoc
         TempData.TOTAL_AMOUNT = TempData2.TOTAL_AMOUNT
         TempData.PART_ITEM_ID = TempData2.PART_ITEM_ID
         TempData.APAR_MAS_ID = TempData2.APAR_MAS_ID
         TempData.DRIVER_ID = TempData2.DRIVER_ID
         TempData.CAR_LICENSE_ID = TempData2.CAR_LICENSE_ID
         TempData.TRANSPORTOR_ID = TempData2.TRANSPORTOR_ID
         TempData.CUSTOMER_BRANCH = TempData2.CUSTOMER_BRANCH
      
         Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID & "-" & TempData.APAR_MAS_ID & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.PART_ITEM_ID))
      Else
         TempData.TOTAL_AMOUNT = TempData.TOTAL_AMOUNT + TempData2.TOTAL_AMOUNT
      End If
      
      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctForTranSport(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentTypeSet As String, Optional DriverID As Long = -1, Optional CarLicenseID As Long = -1, Optional TranSportorID As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim Ind As CInventoryDoc
Dim TempData2 As CInventoryDoc
Dim I As Long
      
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If DocumentTypeSet = "(" & PO_DOCTYPE & ")" Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.DRIVER_ID = DriverID
   BD.CAR_LICENSE_ID = CarLicenseID
   BD.TRANSPORTOR_ID = TranSportorID
   Call BD.QueryData(99, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(99, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = New ADODB.Recordset
   Set Ind = New CInventoryDoc
   
   Ind.FROM_DATE = FromDate
   Ind.TO_DATE = ToDate
   Ind.DRIVER_ID = DriverID
   Ind.CAR_LICENSE_ID = CarLicenseID
   Ind.TRANSPORTOR_ID = TranSportorID
   Call Ind.QueryData2(3, Rs, itemcount)
   
   While Not Rs.EOF
      Set TempData2 = New CInventoryDoc
      Call TempData2.PopulateFromRS2(3, Rs)
      
      Set TempData = GetObject("CBillingDoc", Cl, Trim(Str(TempData2.PART_ITEM_ID)), False)
      If TempData Is Nothing Then
         Set TempData = New CBillingDoc
         TempData.PART_ITEM_ID = TempData2.PART_ITEM_ID
         TempData.BILL_DESC = TempData2.BILL_DESC
         TempData.UNIT_AMOUNT = TempData2.UNIT_AMOUNT
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctForTranSport2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentTypeSet As String, Optional DriverID As Long = -1, Optional CarLicenseID As Long = -1, Optional TranSportorID As Long = -1, Optional ReportType As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim Ind As CInventoryDoc
Dim TempData2 As CInventoryDoc
Dim I As Long
      
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If DocumentTypeSet = "(" & PO_DOCTYPE & ")" Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   If ReportType = 1 Then  'ขายจริง
      BD.CONSIGNMENT_FLAG = "N"
   ElseIf ReportType = 2 Then 'ฝากขาย
      BD.CONSIGNMENT_FLAG = "Y"
   End If
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.DRIVER_ID = DriverID
   BD.CAR_LICENSE_ID = CarLicenseID
   BD.TRANSPORTOR_ID = TranSportorID
   If Not (ReportType = 3) Then 'ส่งเสริมจัดทริป
      Call BD.QueryData(114, Rs, itemcount)
   
      While Not Rs.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS2(114, Rs)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   End If
   
   Set Rs = New ADODB.Recordset
   Set Ind = New CInventoryDoc
   
   Ind.FROM_DATE = FromDate
   Ind.TO_DATE = ToDate
   Ind.DRIVER_ID = DriverID
   Ind.CAR_LICENSE_ID = CarLicenseID
   Ind.TRANSPORTOR_ID = TranSportorID
   If Not (ReportType = 1 Or ReportType = 2) Then  'เป็นทั้งหมด หรือเป็นส่งเสริมจัดทริปถึงสั่ง Query
      Call Ind.QueryData2(7, Rs, itemcount)
      
      While Not Rs.EOF
         Set TempData2 = New CInventoryDoc
         Call TempData2.PopulateFromRS2(7, Rs)
         
         Set TempData = GetObject("CBillingDoc", Cl, Trim(Str(TempData2.PART_ITEM_ID)), False)
         If TempData Is Nothing Then
            Set TempData = New CBillingDoc
            TempData.PART_ITEM_ID = TempData2.PART_ITEM_ID
            TempData.BILL_DESC = TempData2.BILL_DESC
            TempData.UNIT_AMOUNT = TempData2.UNIT_AMOUNT
            Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
         End If
         
         Set TempData = Nothing
         Set TempData2 = Nothing
         Rs.MoveNext
      Wend
   End If
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctForTranSport3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentTypeSet As String, Optional DriverID As Long = -1, Optional CarLicenseID As Long = -1, Optional TranSportorID As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If DocumentTypeSet = "(" & PO_DOCTYPE & ")" Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   BD.DRIVER_ID = DriverID
   BD.CAR_LICENSE_ID = CarLicenseID
   BD.TRANSPORTOR_ID = TranSportorID
   Call BD.QueryData(115, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(115, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub LoadMasterTypeName(C As ComboBox)
Dim I As MASTER_TYPE
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   For I = MASTER_COUNTRY To MASTER_CAR_LICENSE
      C.AddItem (MasterType2String(I))
      C.ItemData(I) = I
   Next I
End Sub
Public Function MasterType2String(M As MASTER_TYPE) As String
   MasterType2String = ""
   If M = MASTER_COUNTRY Then
      MasterType2String = "ประเภท"
   ElseIf M = MASTER_SEX Then
      MasterType2String = "เพศ"
   ElseIf M = MASTER_CUSTYPE Then
      MasterType2String = "ประเภทลูกค้า"
   ElseIf M = MASTER_CUSGRADE Then
      MasterType2String = "ระดับลูกค้า"
   ElseIf M = MASTER_SUPTYPE Then
      MasterType2String = "ประเภทซัพพลายเออร์"
   ElseIf M = MASTER_SUPGRADE Then
      MasterType2String = "ระดับซัพพลายเออร์"
   ElseIf M = MASTER_POSITION Then
      MasterType2String = "ตำแหน่ง"
   ElseIf M = MASTER_PREFIX Then
      MasterType2String = "คำนำหน้าชื่อ"
   ElseIf M = MASTER_JOURNAL Then
      MasterType2String = "สมุดรายวัน"
   ElseIf M = MASTER_DEPARTMENT Then
      MasterType2String = "แผนก"
   ElseIf M = MASTER_UNIT Then
      MasterType2String = "หน่วย"
   ElseIf M = MASTER_STOCKTYPE Then
      MasterType2String = "ประเภทสต็อค"
   ElseIf M = MASTER_STOCKGROUP Then
      MasterType2String = "กลุ่มสต็อค"
   ElseIf M = MASTER_LOCATION Then
      MasterType2String = "คลัง"
   ElseIf M = MASTER_DOCTYPE Then
      MasterType2String = "ประเภทเอกสาร"
   ElseIf M = MASTER_BANK Then
      MasterType2String = "ธนาคาร"
   ElseIf M = MASTER_BBRANCH Then
      MasterType2String = "สาขาธนาคาร"
   ElseIf M = MASTER_CHEQUE_TYPE Then
      MasterType2String = "ประเภทเช็ค"
   ElseIf M = MASTER_CNDN_REASON Then
      MasterType2String = "เหตุผลรับคืน"
   ElseIf M = MASTER_LOCATION_SALE Then
      MasterType2String = "เขตการขาย/สาขา"
   ElseIf M = MASTER_APARMAS_BRANCH Then
      MasterType2String = "สาขาลูกค้า"
   ElseIf M = MASTER_CUSTOMER_BLOCK Then
      MasterType2String = "บล็อคลูกค้า"
   ElseIf M = MASTER_INVOICE_SUB Then
      MasterType2String = "ใบส่งสินค้าย่อย"
   ElseIf M = MASTER_INVOICE_RETURN Then
      MasterType2String = "ใบส่งสินค้าคืน"
   ElseIf M = MASTER_SUBTRACT Then
      MasterType2String = "ส่วนหัก"
   ElseIf M = MASTER_BANK_ACCOUNT Then
      MasterType2String = "เลขที่บัญชี"
   ElseIf M = MASTER_BACCOUNT_TYPE Then
      MasterType2String = "ประเภทบัญชี"
   ElseIf M = MASTER_ADDITION Then
      MasterType2String = "ส่วนเพิ่ม"
   ElseIf M = MASTER_PRODUCTION_LOST Then
      MasterType2String = "ประเภทสูญเสีย"
   ElseIf M = MASTER_PRODUCTION_LOCATION Then
      MasterType2String = "สถานที่ผลิต"
   ElseIf M = MASTER_PRODUCTION_TYPE Then
      MasterType2String = "ประเภทผลิต"
   ElseIf M = MASTER_CUSGROUP Then
      MasterType2String = "กลุ่มลูกค้า"
   ElseIf M = MASTER_STOCKTYPE_SUB Then
      MasterType2String = "ประเภทสต็อคย่อย"
   ElseIf M = MASTER_INVENTORY_SUB_TYPE Then
      MasterType2String = "ประเภทเอกสารย่อย"
   ElseIf M = MASTER_INVENTORY_SALE_GROUP Then
      MasterType2String = "กลุ่มสถานที่จัดเก็บ"
   ElseIf M = MASTER_DRIVER Then
      MasterType2String = "คนขับรถ"
   ElseIf M = MASTER_TRANSPORTOR Then
      MasterType2String = "สำนักงานขนส่ง"
   ElseIf M = MASTER_CAR_LICENSE Then
      MasterType2String = "ทะเบียนรถ"
   End If
   
End Function
Public Sub LoadTranSportByDriverCarTranSportTor(Cl As Collection)
On Error GoTo ErrorHandler
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTranSportDetail
Dim I As Long
   
   Set Rs = New ADODB.Recordset
   Set TempData = New CTranSportDetail
   
   TempData.TRANSPORT_DETAIL_ID = -1
   Call TempData.QueryData(1, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTranSportDetail
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description & " " & Trim(TempData.DRIVER_NAME & "-" & TempData.CAR_LICENSE_NAME & "-" & TempData.TRANSPORTOR_NAME)
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDisTinctBillingDocID(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DocumentNo As String, Optional DocumentType As Long, Optional FromDueDate As Date = -1, Optional ToDueDate As Date = -1, Optional OrderBy As Long = -1, Optional DocumentTypeSet As String, Optional DriverID As Long, Optional TranSportorID As Long, Optional BillingDocPack As Long, Optional ConsignFlag As String = "", Optional FromDocumentNo As String, Optional ToDocumentNo As String)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_DUE_DATE = FromDueDate
   D.TO_DUE_DATE = ToDueDate
   D.DOCUMENT_NO = DocumentNo
   D.FROM_DOCUMENT_NO = FromDocumentNo
   D.TO_DOCUMENT_NO = ToDocumentNo
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   D.ORDER_BY = OrderBy
   D.DRIVER_ID = DriverID
   D.TRANSPORTOR_ID = TranSportorID
   D.BILLING_DOC_PACK = BillingDocPack
   D.CONSIGNMENT_FLAG = ConsignFlag
   Call D.QueryData(100, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(100, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDisTinctBillingDocID2(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional FromDocumentNo As String, Optional ToDocumentNo As String, Optional DocumentType As Long, Optional FromDueDate As Date = -1, Optional ToDueDate As Date = -1, Optional OrderBy As Long = -1, Optional OrderType As Long = -1, Optional DocumentTypeSet As String, Optional DriverID As Long, Optional TranSportorID As Long, Optional BillingDocPack As Long, Optional ConsignFlag As String = "")
' LoadDisTinctBillingDocID2 สำหรับการพิมพ์เอกสารเป็นชุด เช่น ใช้ใน class CReportNormalRcp001_3 ใบรับคืนสินค้าเป็นชุด
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset

   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_DUE_DATE = FromDueDate
   D.TO_DUE_DATE = ToDueDate
   D.FROM_DOCUMENT_NO = FromDocumentNo
   D.TO_DOCUMENT_NO = ToDocumentNo
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   D.ORDER_BY = OrderBy
  D.ORDER_TYPE = OrderType
   D.DRIVER_ID = DriverID
   D.TRANSPORTOR_ID = TranSportorID
   D.BILLING_DOC_PACK = BillingDocPack
   D.CONSIGNMENT_FLAG = ConsignFlag
   Call D.QueryData(100, Rs, itemcount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(100, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDisTinctPOID(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DocumentTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   Call D.QueryData(101, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(101, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpDocType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(102, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(102, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpDocTypeBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(103, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(103, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparBranchCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(104, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(104, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpAparBranchCodeBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(105, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(105, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.CUSTOMER_BRANCH_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTagetJobByType(Cl As Collection, Optional MonthID As Long = -1, Optional YearNo As String)
On Error GoTo ErrorHandler
Dim D As CTagetJobDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetJobDetail
Dim I As Long
   
   Set D = New CTagetJobDetail
   Set Rs = New ADODB.Recordset
   
   D.MONTH_ID = MonthID
   D.YEAR_NO = YearNo
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetJobDetail
      Call TempData.PopulateFromRS(2, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BATCH_NO & "-" & TempData.INPUT_ID & "-" & TempData.OUTPUT_TYPE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTagetJobInputByType(Cl As Collection, Optional MonthID As Long = -1, Optional YearNo As String)
On Error GoTo ErrorHandler
Dim D As CTagetJob
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTagetJob
Dim I As Long
   
   Set D = New CTagetJob
   Set Rs = New ADODB.Recordset
   
   D.MONTH_ID = MonthID
   D.YEAR_NO = YearNo
   Call D.QueryData(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTagetJob
      Call TempData.PopulateFromRS(2, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.INPUT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetLotItemPartTxType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "31"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.LOCATION_GROUP_NO = LocationGroupNo
   Call Lt.QueryData(31, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(31, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetLotItemPartTxType2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim m_Data As CMasterRef
Dim I As Long

   Set m_Data = New CMasterRef
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
         MasterInd = "41"
         Set Lt = New CLotItem
         Set Rs = New ADODB.Recordset
         
         Lt.FROM_DOC_DATE = FromDate
         Lt.TO_DOC_DATE = ToDate
         Lt.FROM_STOCK_NO = FromStockNo
         Lt.TO_STOCK_NO = ToStockNo
         Lt.LOCATION_GROUP_NO = LocationGroupNo
         Call Lt.QueryData(41, Rs, itemcount, False)
         I = 0
         While Not Rs.EOF
            I = I + 1
            Set TempData = New CLotItem
            Call TempData.PopulateFromRS(41, Rs)
            
            If Not (Cl Is Nothing) Then
               Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.LOCATION_NO))
            End If
            
            Set TempData = Nothing
            Rs.MoveNext
         Wend
      MasterInd = "1"
      
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetCountAllCycleDriver(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
Dim TempBd As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = DocTypeSet
   Call BD.QueryData(106, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(106, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempBd = GetObject("CBillingDoc", Cl, Trim(Str(TempData.DRIVER_ID)), False)
         If TempBd Is Nothing Then
            TempData.TRANSPORT_CYCLE = 1
            Call Cl.add(TempData, Trim(Str(TempData.DRIVER_ID)))
         Else
            TempBd.TRANSPORT_CYCLE = TempBd.TRANSPORT_CYCLE + 1
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetCountAllCycleTranSportor(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
Dim TempBd As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = DocTypeSet
   Call BD.QueryData(107, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(107, Rs)
      
      If Not (Cl Is Nothing) Then
         Set TempBd = GetObject("CBillingDoc", Cl, Trim(Str(TempData.TRANSPORTOR_ID)), False)
         If TempBd Is Nothing Then
            TempData.TRANSPORT_CYCLE = 1
            Call Cl.add(TempData, Trim(Str(TempData.TRANSPORTOR_ID)))
         Else
            TempBd.TRANSPORT_CYCLE = TempBd.TRANSPORT_CYCLE + 1
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumAmountByDriver(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
Dim TempBd As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = DocTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(108, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(108, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.DRIVER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumAmountByTranSportor(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
Dim TempBd As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = DocTypeSet
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(109, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(109, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.TRANSPORTOR_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSumAmountByPartItemIndSub(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(34, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(34, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INVENTORY_SUB_TYPE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetAmountForTranSport2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentTypeSet As String, Optional ReportType As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim Ind As CInventoryDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim TempData2 As CInventoryDoc

Dim I As Long
      
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   If DocumentTypeSet = "(" & PO_DOCTYPE & ")" Then
      BD.FROM_DUE_DATE = FromDate
      BD.TO_DUE_DATE = ToDate
   Else
      BD.FROM_DATE = FromDate
      BD.TO_DATE = ToDate
   End If
   If ReportType = 1 Then  'ขายจริง
      BD.CONSIGNMENT_FLAG = "N"
   ElseIf ReportType = 2 Then 'ฝากขาย
      BD.CONSIGNMENT_FLAG = "Y"
   End If
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   If Not (ReportType = 3) Then 'ส่งเสริมจัดทริป
      Call BD.QueryData(113, Rs, itemcount)
   
      While Not Rs.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS2(113, Rs)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID & "-" & TempData.PART_ITEM_ID))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   End If
      
      
   Set Ind = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   
   Ind.FROM_DATE = FromDate
   Ind.TO_DATE = ToDate
   If Not (ReportType = 1 Or ReportType = 2) Then  'เป็นทั้งหมด หรือเป็นส่งเสริมจัดทริปถึงสั่ง Query
      Call Ind.QueryData2(6, Rs, itemcount)
   
      While Not Rs.EOF
         Set TempData2 = New CInventoryDoc
         Call TempData2.PopulateFromRS2(6, Rs)
         
         Set TempData = GetObject("CBillingDoc", Cl, Trim(TempData2.DRIVER_ID & "-" & TempData2.CAR_LICENSE_ID & "-" & TempData2.TRANSPORTOR_ID & "-" & TempData2.PART_ITEM_ID), False)
         If TempData Is Nothing Then
            Set TempData = New CBillingDoc
            TempData.TOTAL_AMOUNT = TempData2.TOTAL_AMOUNT
            TempData.PART_ITEM_ID = TempData2.PART_ITEM_ID
            TempData.DRIVER_ID = TempData2.DRIVER_ID
            TempData.CAR_LICENSE_ID = TempData2.CAR_LICENSE_ID
            TempData.TRANSPORTOR_ID = TempData2.TRANSPORTOR_ID
            
            Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID & "-" & TempData.PART_ITEM_ID))
         Else
            TempData.TOTAL_AMOUNT = TempData.TOTAL_AMOUNT + TempData2.TOTAL_AMOUNT
         End If
         
         Set TempData = Nothing
         Set TempData2 = Nothing
         Rs.MoveNext
      Wend
   End If
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctPartOutputByInput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "44"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
      
   Call Lt.QueryData(44, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      
      Call TempData.PopulateFromRS(44, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetDistinctPartInputByInput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "44"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   
      
   Call Lt.QueryData(46, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      
      Call TempData.PopulateFromRS(46, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumAmountByInputOutput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   MasterInd = "47"
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationId
   Lt.DOCUMENT_TYPE = DocumentType
'   Lt.FROM_STOCK_NO = FromStockNo
'   Lt.TO_STOCK_NO = ToStockNo
   
   Call Lt.QueryData(47, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      
      Call TempData.PopulateFromRS(47, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INVENTORY_DOC_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSumBillByCustomerBranch(Cl As Collection, Optional SumBillID As Long)
On Error GoTo ErrorHandler
Dim SBD As CBillDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillDetail
Dim I As Long
   
   Set SBD = New CBillDetail
   Set Rs = New ADODB.Recordset
   
   SBD.SUM_BILL_ID = SumBillID
   Call SBD.QueryData(3, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillDetail
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.CUSTOMER_BRANCH)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetCountSumBillByCustomerBranch(Cl As Collection, Optional SumBillID As Long)
On Error GoTo ErrorHandler
Dim SBD As CBillDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillDetail
Dim I As Long
   
   Set SBD = New CBillDetail
   Set Rs = New ADODB.Recordset
   
   SBD.SUM_BILL_ID = SumBillID
   Call SBD.QueryData(4, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillDetail
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_BRANCH & "-" & TempData.DOC_ID_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetBillDetailReceipted(Cl As Collection, Optional FromSummaryDocDate As Date = -1, Optional ToSummaryDocDate As Date = -1)
On Error GoTo ErrorHandler
Dim SBD As CBillDetail
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillDetail
Dim I As Long
   
   Set SBD = New CBillDetail
   Set Rs = New ADODB.Recordset
   
   SBD.FROM_SUMMARY_DOC_DATE = FromSummaryDocDate
   SBD.TO_SUMMARY_DOC_DATE = ToSummaryDocDate
   Call SBD.QueryData(5, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillDetail
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.BILLING_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctEmpStockCodeDocTypeFreeConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional Database As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE = PO_DOCTYPE
   BD.CONSIGNMENT_FLAG = "Y"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(158, Rs, itemcount, , Database)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(158, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "149"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE = PO_DOCTYPE
   BD.CONSIGNMENT_FLAG = "Y"
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(149, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(149, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpDateFreeConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional CheckQMC_To_FAndB As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "150"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE = PO_DOCTYPE
   BD.CONSIGNMENT_FLAG = "Y"
   If CheckQMC_To_FAndB = 1 Then
      BD.CheckQMC_To_FAndB = True
   End If
   Call BD.QueryData(150, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(150, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub GetSaleAmountEmpConsignmentForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "160"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE = PO_DOCTYPE
   BD.CONSIGNMENT_FLAG = "Y"
   Call BD.QueryData(160, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(160, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_JOINT_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpDateFreeConsignmentForFAndB(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   MasterInd = "161"
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE = PO_DOCTYPE
   BD.CONSIGNMENT_FLAG = "Y"
   Call BD.QueryData(161, Rs, itemcount, , 2)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(161, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_JOINT_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub LoadSaleChart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FK_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CSaleChart
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSaleChart
Dim I As Long
   If FK_ID <= 0 Then
      Exit Sub
   End If
   Set D = New CSaleChart
   Set Rs = New ADODB.Recordset
   
   D.SALE_CHART_ID = -1
   D.MASTER_FROMTO_ID = FK_ID
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSaleChart
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SALE_NAME & "(" & TempData.SALE_CODE & ")")
         C.ItemData(I) = TempData.SALE_CHART_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.SALE_CHART_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctTripFormExportType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CInventoryDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim I As Long
      
   Set BD = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryData2(2, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS2(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.APAR_MAS_ID & "-" & TempData.CUSTOMER_BRANCH & "-" & TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID)
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadDistinctTripFormExportTypeGroup(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CInventoryDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim I As Long
      
   Set BD = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryData2(5, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS2(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DRIVER_ID & "-" & TempData.CAR_LICENSE_ID & "-" & TempData.TRANSPORTOR_ID))
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub

Public Sub LoadDealerType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (DealerTypeToString(SILVER))
   C.ItemData(1) = SILVER
   
   C.AddItem (DealerTypeToString(SILVER_PLUS))
   C.ItemData(2) = SILVER_PLUS
   
   C.AddItem (DealerTypeToString(SILVER_PLUS_PLUS))
   C.ItemData(3) = SILVER_PLUS_PLUS
   
   C.AddItem (DealerTypeToString(GOLD_MUNUS))
   C.ItemData(4) = GOLD_MUNUS
   
   C.AddItem (DealerTypeToString(GOLD))
   C.ItemData(5) = GOLD
   
   C.AddItem (DealerTypeToString(GOLD_PLUS))
   C.ItemData(6) = GOLD_PLUS
   
   C.AddItem (DealerTypeToString(GOLD_PLUS_PLUS))
   C.ItemData(7) = GOLD_PLUS_PLUS
   
   C.AddItem (DealerTypeToString(PLATINUM_MUNUS))
   C.ItemData(8) = PLATINUM_MUNUS
   
   C.AddItem (DealerTypeToString(PLATINUM))
   C.ItemData(9) = PLATINUM
   
   C.AddItem (DealerTypeToString(HEADER_GROUP))
   C.ItemData(10) = HEADER_GROUP
End Sub
Public Sub GetSaleAmountDealerDocType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FREE_FLAG = StringToFreeFlag(InCludeFree)
   Call BD.QueryData(167, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(167, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetEmpDealerTypeYYYYMM(Cl As Collection, YYYYMM As String)
On Error GoTo ErrorHandler
Dim BD As CEmployeeDealer
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmployeeDealer
Dim I As Long
   
   Set BD = New CEmployeeDealer
   Set Rs = New ADODB.Recordset
   
   BD.YYYYMM = YYYYMM
   Call BD.QueryData(1, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmployeeDealer
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.EMP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetDistinctCheckBillingMovement(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long


   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   Call BD.QueryData(171, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   MasterInd = "171"
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(171, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   MasterInd = "1"
      
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadInvoidByPo(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim TempData2 As CBillingDoc
Dim Cl2 As Collection
Dim I As Long
Dim m_Billing As CBillingDoc


   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE = INVOICE_DOCTYPE
   BD.CANCEL_FLAG = "N"
'   BD.BILLING_DOC_ID = 473407
   Call BD.QueryData(175, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set Cl2 = New Collection
   Dim PrevKey1 As String
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(175, Rs)
'      If TempData.PO_NO = "PO6210/0011" Then
'        Debug.Print
'      End If
        If Not (Cl Is Nothing) Then
            
         Set m_Billing = GetObject("CBillingDoc", Cl, Trim(TempData.PO_NO), False)
         If m_Billing Is Nothing Then
            Set m_Billing = New CBillingDoc
'            m_Billing.DOCUMENT_NO = TempData.DOCUMENT_NO
'            mRcpSub.SALE_CODE = TempData.SALE_CODE
'            mRcpSub.SALE_NAME = TempData.SALE_NAME
            Call Cl.add(m_Billing, Trim(TempData.PO_NO))
         End If
         
         Call m_Billing.collBillSub.add(TempData)
      End If
      
      
'       If Not (Cl2 Is Nothing) Then
'            Call Cl2.add(TempData)
'      End If
'
'      If PrevKey1 <> Trim(TempData.PO_NO) Or I = 1 Then
'          If Not (Cl Is Nothing) Then
'            Call Cl.add(Cl2, Trim(TempData.PO_NO))
'
'            Set Cl2 = Nothing
'            Set Cl2 = New Collection
'          End If
'      End If
'      PrevKey1 = Trim(TempData.PO_NO)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
