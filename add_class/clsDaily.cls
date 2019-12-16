VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing
   
   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub
'==
Public Function QueryEnterprise(Ua As CEnterprise, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryEnterprise"
   QueryEnterprise = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cEnpname As CEnterpriseName
      Set cEnpname = New CEnterpriseName
      Call cEnpname.SetFieldValue("ENTERPRISE_ID", Ua.GetFieldValue("ENTERPRISE_ID"))
      Call cEnpname.SetFieldValue("ORDER_BY", 1)
      Call cEnpname.QueryData(1, m_Rs1, iCount)
      Set cEnpname = Nothing
      
      Set Ua.EnpNames = Nothing
      Set Ua.EnpNames = New Collection
      While Not m_Rs1.EOF
         Set cEnpname = New CEnterpriseName
         Call cEnpname.PopulateFromRS(1, m_Rs1)

         Dim CName As CName
         Dim iCount2 As Long
         Set CName = New CName
         Call CName.SetFieldValue("NAME_ID", cEnpname.GetFieldValue("NAME_ID"))
         Call CName.QueryData(1, m_Rs2, iCount2)
         Set CName = Nothing
         While Not m_Rs2.EOF
            Set CName = New CName
            Call CName.PopulateFromRS(1, m_Rs2)
            CName.Flag = "I"
            Call cEnpname.Names.add(CName)
            Set CName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpname.Flag = "I"
         Call Ua.EnpNames.add(cEnpname)
         Set cEnpname = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
      
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      Call cEnpAddr.SetFieldValue("ENTERPRISE_ID", Ua.GetFieldValue("ENTERPRISE_ID"))
      Call cEnpAddr.SetFieldValue("ORDER_BY", 1)
      Call cEnpAddr.QueryData(1, m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.EnpAddresses = Nothing
      Set Ua.EnpAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CEnterpriseAddress
         Call cEnpAddr.PopulateFromRS(1, m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         Call cAddr.SetFieldValue("ADDRESS_ID", cEnpAddr.GetFieldValue("ADDRESS_ID"))
         Call cAddr.QueryData(1, m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(1, m_Rs2)
            cAddr.Flag = "I"
            Call cEnpAddr.Addresses.add(cAddr)
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.EnpAddresses.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If

   QueryEnterprise = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryEnterprise = False
End Function

Public Function AddEditEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Name As CName
Dim Addr1 As CAddress

   RName = "AddEditEnterprise"
   AddEditEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
         
   Dim EnpName As CEnterpriseName
   Dim EmpAddr As CEnterpriseAddress
   Dim Addr As CAddress
   
   For Each EmpAddr In Ug.EnpAddresses
      If EmpAddr.Flag = "A" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.ShowMode = SHOW_ADD
         Call Addr.AddEditData
                  
         Call EmpAddr.SetFieldValue("ENTERPRISE_ID", Ug.GetFieldValue("ENTERPRISE_ID"))
         Call EmpAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
         EmpAddr.ShowMode = SHOW_ADD
         EmpAddr.AddEditData
      ElseIf EmpAddr.Flag = "E" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.ShowMode = SHOW_EDIT
         Call Addr.AddEditData
         
      ElseIf EmpAddr.Flag = "D" Then
         Call EmpAddr.DeleteData
         
         Set Addr = EmpAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next EmpAddr
   
   For Each EnpName In Ug.EnpNames
      If EnpName.Flag = "A" Then
         Set Name = EnpName.Names(1)
         Name.ShowMode = SHOW_ADD
         Call Name.AddEditData
                  
         Call EnpName.SetFieldValue("ENTERPRISE_ID", Ug.GetFieldValue("ENTERPRISE_ID"))
         Call EnpName.SetFieldValue("NAME_ID", Name.GetFieldValue("NAME_ID"))
         EnpName.ShowMode = SHOW_ADD
         EnpName.AddEditData
      ElseIf EnpName.Flag = "E" Then
         Set Name = EnpName.Names(1)
         Name.ShowMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next EnpName
      
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEnterprise = False
End Function

Public Function DeleteEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteEnterprise"
   DeleteEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_MSGBOX)
   
   IsOK = False

   DeleteEnterprise = True
End Function

Public Function QueryCustomer(Ua As CAPARMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomer"
   QueryCustomer = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CApArName
      Set cCstName = New CApArName
      Call cCstName.SetFieldValue("APAR_NAME_ID", -1)
      Call cCstName.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
      Call cCstName.SetFieldValue("ORDER_BY", 1)
      Call cCstName.QueryData(1, m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CApArName
         Call cCstName.PopulateFromRS(1, m_Rs1)

         Dim CName As CName
         Dim iCount2 As Long
         Set CName = New CName
         Call CName.SetFieldValue("NAME_ID", cCstName.GetFieldValue("NAME_ID"))
         Call CName.QueryData(1, m_Rs2, iCount2)
         Set CName = Nothing
         While Not m_Rs2.EOF
            Set CName = New CName
            Call CName.PopulateFromRS(1, m_Rs2)
            CName.Flag = "I"
            Set cCstName.Name = CName
            Set CName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CApArAddress
      Set cEnpAddr = New CApArAddress
      Call cEnpAddr.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
      Call cEnpAddr.SetFieldValue("ORDER_BY", 1)
      Call cEnpAddr.QueryData(1, m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddresses = Nothing
      Set Ua.CstAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CApArAddress
         Call cEnpAddr.PopulateFromRS(1, m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         Call cAddr.SetFieldValue("ADDRESS_ID", cEnpAddr.GetFieldValue("ADDRESS_ID"))
         Call cAddr.QueryData(1, m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(1, m_Rs2)
'            Call cAddr.SetFieldValue("ADDRESS_TYPE", cEnpAddr.GetFieldValue("ADDRESS_TYPE"))
            cAddr.Flag = "I"
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddresses.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
      
      Dim Br As CMasterRef
      Set Br = New CMasterRef
      Br.PARENT_EX_ID2 = Ua.APAR_MAS_ID
      Call Br.QueryData(1, m_Rs1, iCount)
      Set Br = Nothing
      
      Set Ua.CstBranchs = Nothing
      Set Ua.CstBranchs = New Collection

      While Not m_Rs1.EOF
         Set Br = New CMasterRef
         Call Br.PopulateFromRS(1, m_Rs1)
         Br.Flag = "I"
         Call Ua.CstBranchs.add(Br)
         Set Br = Nothing

         m_Rs1.MoveNext
      Wend
      Set Br = Nothing
      
   End If
      
   QueryCustomer = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomer = False
End Function

Public Function AddEditCustomer(Ug As CAPARMas, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim CstName As CApArName
   Dim Name As CName
   Dim CstAddr As CApArAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.CstAddresses
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.ShowMode = SHOW_ADD
         Call Addr.AddEditData

         Call CstAddr.SetFieldValue("APAR_MAS_ID", Ug.APAR_MAS_ID)
         Call CstAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
'         Call CstAddr.SetFieldValue("ADDRESS_TYPE", -1)
         CstAddr.ShowMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.ShowMode = SHOW_EDIT
         Call Addr.AddEditData

         Call CstAddr.SetFieldValue("APAR_MAS_ID", Ug.APAR_MAS_ID)
         Call CstAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
'         Call CstAddr.SetFieldValue("ADDRESS_TYPE", Addr.GetFieldValue("ADDRESS_TYPE"))
         CstAddr.ShowMode = SHOW_EDIT
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set Name = CstName.Name
         Name.ShowMode = SHOW_ADD
         Call Name.AddEditData

         Call CstName.SetFieldValue("APAR_MAS_ID", Ug.APAR_MAS_ID)
         Call CstName.SetFieldValue("NAME_ID", Name.GetFieldValue("NAME_ID"))
         CstName.ShowMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set Name = CstName.Name
         Name.ShowMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next CstName
            
   Dim Br As CMasterRef
   For Each Br In Ug.CstBranchs
      Br.PARENT_EX_ID2 = Ug.APAR_MAS_ID
      If Br.Flag = "A" Then
         Br.ShowMode = SHOW_ADD
         Call Br.AddEditData
      ElseIf Br.Flag = "E" Then
         Br.ShowMode = SHOW_EDIT
         Call Br.AddEditData
      ElseIf Br.Flag = "D" Then
         Call Br.DeleteData
      End If
   Next Br
   
   Dim CUS As CAPARMas
   Set CUS = New CAPARMas
   CUS.APAR_IND = 1
   Call LoadApArMas(CUS, Nothing, m_CustomerColl)
   Set CUS = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomer = False
End Function

Public Function DeleteCustomer(Ug As CAPARMas, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteCustomer = False
End Function

Public Function DeleteMasterRef(Ug As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteMasterRef = True
End Function

Public Function QueryEmployee(Ua As CEmployee, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

Dim Qp As CEmployee

   RName = "QueryEmployee"
   QueryEmployee = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      Set Ua.EmpName = Nothing
      Set Ua.EmpName = New CEmployeeName
      Call Ua.EmpName.SetFieldValue("EMP_NAME_ID", -1)
      Call Ua.EmpName.SetFieldValue("EMP_ID", Ua.EMP_ID)
      Call Ua.EmpName.QueryData(1, m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EmpName.PopulateFromRS(1, m_Rs1)
      End If
      
      Set Ua.EName = Nothing
      Set Ua.EName = New CName
      Call Ua.EName.SetFieldValue("NAME_ID", Ua.EmpName.GetFieldValue("NAME_ID"))
      Call Ua.EName.QueryData(1, m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EName.PopulateFromRS(1, m_Rs1)
      End If
      
      
      Dim EmpDl As CEmployeeDealer
      Set EmpDl = New CEmployeeDealer
      EmpDl.EMP_ID = Ua.EMP_ID
      Call EmpDl.QueryData(1, m_Rs1, iCount)
      Set EmpDl = Nothing

      Set Ua.CollEmpDealer = Nothing
      Set Ua.CollEmpDealer = New Collection
      
      While Not m_Rs1.EOF
         Set EmpDl = New CEmployeeDealer
         Call EmpDl.PopulateFromRS(1, m_Rs1)
         
         EmpDl.Flag = "I"
         Call Ua.CollEmpDealer.add(EmpDl)
         Set EmpDl = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryEmployee = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployee = False
End Function

Public Function DeleteEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteEmployee"
   DeleteEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteEmployee = True
End Function

Public Function AddEditEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmployee"
   AddEditEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Call Ug.EName.AddEditData
   
   Call Ug.EmpName.SetFieldValue("EMP_ID", Ug.EMP_ID)
   Call Ug.EmpName.SetFieldValue("NAME_ID", Ug.EName.GetFieldValue("NAME_ID"))
   Call Ug.EmpName.AddEditData
            
   Dim EmpDl As CEmployeeDealer
   For Each EmpDl In Ug.CollEmpDealer
      EmpDl.EMP_ID = Ug.EMP_ID
      If EmpDl.Flag = "A" Then
         EmpDl.AddEditMode = SHOW_ADD
         Call EmpDl.AddEditData
      ElseIf EmpDl.Flag = "E" Then
         EmpDl.AddEditMode = SHOW_EDIT
         Call EmpDl.AddEditData
      ElseIf EmpDl.Flag = "D" Then
         Call EmpDl.DeleteData
      End If
   Next EmpDl
   
    Dim TempEmp As CEmployee
    Set TempEmp = New CEmployee
    Call LoadEmployee(TempEmp, Nothing, m_EmployeeColl)
    Set TempEmp = Nothing
   
   
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEmployee = False
End Function

Public Function QueryUserAccount(Ua As CUserAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryUserAccount"
   QueryUserAccount = False

   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryUserAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserAccount = False
End Function

Public Function QueryGLAccount(Ua As CGLAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryGLAccount"
   QueryGLAccount = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryGLAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryGLAccount = False
End Function

Public Function QueryInventoryDoc(Ua As CInventoryDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryInventoryDoc"
   QueryInventoryDoc = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CLotItem
      Dim Lk As CLotItemLink
      Set Gr = New CLotItem
      Gr.CANCEL_FLAG = ""
      Gr.INVENTORY_DOC_ID = Ua.GetFieldValue("INVENTORY_DOC_ID")
      Call Gr.QueryData(1, m_Rs1, iCount, False)
      Set Gr = Nothing
      
      Set Ua.ImportExportItems = Nothing
      Set Ua.ImportExportItems = New Collection
      
      While Not m_Rs1.EOF
         Set Gr = New CLotItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         If Gr.TX_TYPE = "E" Then
            Set Lk = New CLotItemLink
            Lk.EXPORT_LOT_ITEM_ID = Gr.LOT_ITEM_ID
            Call Lk.QueryData(1, m_Rs2, ItemCount)
            Set Lk = Nothing
            
            Set Gr.LotItemLink = Nothing
            Set Gr.LotItemLink = New Collection
            While Not m_Rs2.EOF
               Set Lk = New CLotItemLink
               Call Lk.PopulateFromRS(1, m_Rs2)
               Call Gr.LotItemLink.add(Lk, Trim(Str(Lk.IMPORT_LOT_ITEM_ID)))
               
               m_Rs2.MoveNext
            Wend
            Set Lk = Nothing
         End If
         
         Call Ua.ImportExportItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If

   QueryInventoryDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryDoc = False
End Function

Public Function QueryStockCode(Ua As CStockCode, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
   RName = "QueryStockCode"
   QueryStockCode = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      Dim Ji As CStockCodeChange
      Set Ji = New CStockCodeChange
      Ji.STOCK_CODE_ID = Ua.STOCK_CODE_ID
      Call Ji.QueryData(1, m_Rs1, iCount)
      Set Ji = Nothing
      
      Set Ua.StockCodeChange = Nothing
      Set Ua.StockCodeChange = New Collection
      
      While Not m_Rs1.EOF
         Set Ji = New CStockCodeChange
         Call Ji.PopulateFromRS(1, m_Rs1)
         
         Call Ua.StockCodeChange.add(Ji)
         
         Set Ji = Nothing
         
         m_Rs1.MoveNext
      Wend
      
      Dim JiFt As CStockCodeChangeFt
      Set JiFt = New CStockCodeChangeFt
      JiFt.STOCK_CODE_ID = Ua.STOCK_CODE_ID
      Call JiFt.QueryData(1, m_Rs1, iCount)
      Set JiFt = Nothing
      
      Set Ua.StockCodeChangeFt = Nothing
      Set Ua.StockCodeChangeFt = New Collection
      
      While Not m_Rs1.EOF
         Set JiFt = New CStockCodeChangeFt
         Call JiFt.PopulateFromRS(1, m_Rs1)
         
         Call Ua.StockCodeChangeFt.add(JiFt)
         
         Set JiFt = Nothing
         
         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryStockCode = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryStockCode = False
End Function

Public Function AddEditUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserAccount"
   AddEditUserAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserAccount = False
End Function

Public Function AddEditGLAccount(Ua As CGLAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditGLAccount"
   AddEditGLAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditGLAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditGLAccount = False
End Function

Public Function AddEditInventoryDoc(Ua As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditInventoryDoc"
   AddEditInventoryDoc = False
   HasBegin = False
   
'   If Ua.GetFieldValue("DOCUMENT_DATE") < DateSerial(2008, 4, 1) Then
'      glbErrorLog.LocalErrorMsg = "ไม่อนุญาติให้แก้ไขเอกสารคลัง ก่อนวันที่ 01/04/2551"
'      glbErrorLog.ShowUserError
'      IsOK = False
'      AddEditInventoryDoc = False
'      Exit Function
'   End If
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cd As CConfigDoc
   If Ua.ShowMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.GetFieldValue("RUNNING_NO"))
      Call Cd.SetFieldValue("LAST_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.GetFieldValue("CONFIG_DOC_TYPE"))
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Gr As CLotItem
   Dim Lk As CLotItemLink
   For Each Gr In Ua.ImportExportItems
      Gr.INVENTORY_DOC_ID = Ua.GetFieldValue("INVENTORY_DOC_ID")
      Gr.CANCEL_FLAG = Ua.GetFieldValue("CANCEL_FLAG")
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
         
         For Each Lk In Gr.LotItemLink
            If Lk.Flag = "A" Then
               Lk.AddEditMode = SHOW_ADD
               Lk.EXPORT_LOT_ITEM_ID = Gr.LOT_ITEM_ID
               Call Lk.AddEditData
            ElseIf Lk.Flag = "E" Then
               Lk.AddEditMode = SHOW_EDIT
               Call Lk.AddEditData
            ElseIf Lk.Flag = "D" Then
               Call Lk.DeleteData
            End If
         Next Lk
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
         
         For Each Lk In Gr.LotItemLink
            If Lk.Flag = "A" Then
               Lk.AddEditMode = SHOW_ADD
               Lk.EXPORT_LOT_ITEM_ID = Gr.LOT_ITEM_ID
               Call Lk.AddEditData
            ElseIf Lk.Flag = "E" Then
               Lk.AddEditMode = SHOW_EDIT
               Call Lk.AddEditData
            ElseIf Lk.Flag = "D" Then
               Call Lk.DeleteData
            End If
         Next Lk
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If (Ua.GetFieldValue("OLD_CANCEL_FLAG") <> Ua.GetFieldValue("CANCEL_FLAG")) And Len(Ua.GetFieldValue("OLD_CANCEL_FLAG")) > 0 Then
      For Each Gr In Ua.ImportExportItems
         Gr.ShowMode = SHOW_EDIT
         Gr.CANCEL_FLAG = Ua.GetFieldValue("CANCEL_FLAG")
         Call Gr.AddEditData
      Next Gr
   End If
   Set Gr = Nothing
   
'   If Ua.GetFieldValue("DOCUMENT_DATE") = DateSerial(Year(Now), Month(Now), Day(Now)) Then
'      Call UpdateCostAndRemainAmount
'      Call VerifyStockBalance(Ua.ImportExportItems, glbErrorLog, Ua.GetFieldValue("DOCUMENT_DATE"))
'   End If
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditInventoryDoc = False
End Function

Public Function AddEditStockCode(Ua As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStockCode"
   AddEditStockCode = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CStockCodeChange
   For Each Gr In Ua.StockCodeChange
      Gr.STOCK_CODE_ID = Ua.STOCK_CODE_ID
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   Dim GrFt As CStockCodeChangeFt
   For Each GrFt In Ua.StockCodeChangeFt
      GrFt.STOCK_CODE_ID = Ua.STOCK_CODE_ID
      If GrFt.Flag = "A" Then
         GrFt.ShowMode = SHOW_ADD
         Call GrFt.AddEditData
      ElseIf GrFt.Flag = "E" Then
         GrFt.ShowMode = SHOW_EDIT
         Call GrFt.AddEditData
      ElseIf GrFt.Flag = "D" Then
         Call GrFt.DeleteData
      End If
   Next GrFt
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStockCode = False
End Function

Public Function AddEditMasterRef(Ua As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterRef = False
End Function

Public Function DeleteUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserAccount"
   DeleteUserAccount = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteUserAccount = True
End Function

Public Function DeleteGLAccount(Ua As CGLAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteGLAccount"
   DeleteGLAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteGLAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteGLAccount = True
End Function

Public Function DeleteInventoryDoc(Ua As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteInventoryDoc"
   DeleteInventoryDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteInventoryDoc = True
End Function

Public Function DeleteStockCode(Ua As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteStockCode"
   DeleteStockCode = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteStockCode = True
End Function

Public Function DBLogin(UserName As String, Password As String, CanLogin As Boolean, User As CUser, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim ItemCount As Long
Dim LoginID As Long

   RName = "Login"
   DBLogin = False

   SQL1 = "SELECT COUNT(*) FROM USER_ACCOUNT "
   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If NVLI(m_Rs1(0), 0) <= 0 Then 'ถ้าไม่มีผู้ใช้อยุในดาตาเบสให้ล็อคอินอะไรก็ได้
      If UserName <> "" Then
         ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

         DBLogin = True
         CanLogin = False
         Exit Function
      End If
      
      User.USER_ID = 0
      User.GROUP_ID = 0
      User.GROUP_NAME = ""
      User.USER_NAME = UserName

      DBLogin = True
      CanLogin = True
      Exit Function
   End If

   SQL1 = "SELECT UD.*, GD.* FROM USER_ACCOUNT UD, USER_GROUP GD "
   SQL1 = SQL1 & "WHERE (GD.GROUP_ID = UD.GROUP_ID) "
   SQL1 = SQL1 & "AND (UD.USER_NAME = '" & ChangeQuote(UserName) & "' ) "
   SQL1 = SQL1 & "AND (UD.USER_PASSWORD = '" & ChangeQuote(EncryptText(Password)) & "' ) "

   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   If NVLS(m_Rs1("USER_STATUS"), "Y") <> "Y" Then
      ErrorObj.LocalErrorMsg = "รายชื่อของท่านขณะนี้ไม่สามารถเข้าใช้งานระบบได้ โปรดติดต่อผู้บริหารระบบ "

      DBLogin = True
      CanLogin = False
      Exit Function
   End If
   
   User.GROUP_ID = NVLI(m_Rs1("GROUP_ID"), -1)
   User.GROUP_NAME = NVLS(m_Rs1("GROUP_NAME"), "")
   User.USER_NAME = UserName
   User.REAL_USER_ID = NVLI(m_Rs1("USER_ID"), -1)
   User.USER_ID = NVLI(m_Rs1("USER_ID"), -1)
   User.USER_ACCESS = NVLS(m_Rs1("USER_ACCESS"), "")
   
'   m_ADOConn.BeginTrans
'
'   glbLoginTracking.LOGIN_DATE = Now
'   glbLoginTracking.LOGOUT_DATE = -1
'   glbLoginTracking.IP_ADDRESS = glbDatabaseMngr.m_Winsock.LocalIP
'   glbLoginTracking.USER_NAME = User.USER_NAME
'   glbLoginTracking.USER_ID = User.USER_ID
'   glbLoginTracking.AddEditMode = SHOW_ADD
'   Call glbLoginTracking.AddEditData
'
'   m_ADOConn.CommitTrans
   
   CanLogin = True
   DBLogin = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   DBLogin = False
End Function
Public Function QueryUserGroup(Ua As CUserGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryUserGroup"
   QueryUserGroup = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CGroupRight
      Set Gr = New CGroupRight
      Call Gr.SetFieldValue("GROUP_ID", Ua.GetFieldValue("GROUP_ID"))
      Call Gr.QueryData(2, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.RightItems = Nothing
      Set Ua.RightItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CGroupRight
         Call Gr.PopulateFromRS(2, m_Rs1)
         Call Ua.RightItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryUserGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserGroup = False
End Function

Public Function AddEditUserGroup(Ua As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserGroup"
   AddEditUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CGroupRight
   For Each Gr In Ua.RightItems
      Call Gr.SetFieldValue("GROUP_ID", Ua.GetFieldValue("GROUP_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserGroup = False
End Function

Public Function DeleteUserGroup(Ug As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserGroup"
   DeleteUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   
   IsOK = False

   DeleteUserGroup = True
End Function

Public Function ChangePassword(UserID As Long, OldPassword As String, NewPassword As String, IsCanChange As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim UserName As String
Dim ItemCount As Long

   RName = "ChangePassword"
   ChangePassword = False
   HasBegin = False

   SQL = "SELECT * FROM USER_ACCOUNT WHERE (USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs1, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่มีผู้ใช้นี้อยู่ในฐานข้อมูล "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If
   UserName = NVLS(m_Rs1("USER_NAME"), "")

   SQL = "SELECT * FROM USER_ACCOUNT "
   SQL = SQL & "WHERE (USER_PASSWORD = '" & ChangeQuote(EncryptText(OldPassword)) & "') AND "
   SQL = SQL & "(USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs1, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ท่านใส่รหัสผ่านเก่าไม่ถูกต้อง "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If

   m_ADOConn.BeginTrans
   HasBegin = True

   SQL = "UPDATE USER_ACCOUNT SET "
   SQL = SQL & "USER_PASSWORD = '" & ChangeQuote(EncryptText(NewPassword)) & "' "
   SQL = SQL & "WHERE (USER_ID = " & UserID & ") "
   Call m_ADOConn.Execute(SQL)

   Call m_ADOConn.CommitTrans
   HasBegin = False

   IsCanChange = True
   ChangePassword = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description & SQL
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = ""
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   ChangePassword = False
End Function

Public Function AddEditJournal(Ua As CJournal, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditJournal"
   AddEditJournal = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CJournalItem
   For Each Gr In Ua.JournalItems
      Call Gr.SetFieldValue("JOURNAL_ID", Ua.GetFieldValue("JOURNAL_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditJournal = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditJournal = False
End Function

Public Function DeleteJournal(Ug As CJournal, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteJournal"
   DeleteJournal = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteJournal = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteJournal = True
End Function

Public Function QueryJournal(Ua As CJournal, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryJournal"
   QueryJournal = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CJournalItem
      Set Gr = New CJournalItem
      Call Gr.SetFieldValue("JOURNAL_ID", Ua.GetFieldValue("JOURNAL_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.JournalItems = Nothing
      Set Ua.JournalItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CJournalItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         Call Ua.JournalItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryJournal = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryJournal = False
End Function

Public Sub CreateTransferItems(Ivd As CInventoryDoc)
Dim Ti As CTransferItem
Dim O As Object
Dim Ei As CLotItem
Dim II As CLotItem
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.ImportExportItems.Count \ 2
   Count2 = Ivd.ImportExportItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.ImportExportItems(j)
         If (O.TX_TYPE = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.TX_TYPE = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function QueryBillingDoc(Ua As CBillingDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBillingDoc"
   QueryBillingDoc = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CDocItem
      Set Gr = New CDocItem
      
      Dim Plb As CPrintLabel
      Dim iCount2 As Long
      Set Plb = Nothing
      
      Dim Dik As CDocItemLink
      
      Call Gr.SetFieldValue("DOC_ITEM_TYPE", Ua.DOC_ITEM_TYPE)
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DocItems = Nothing
      Set Ua.DocItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CDocItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Set Plb = Nothing
         Set Plb = New CPrintLabel
         
         Call Plb.SetFieldValue("DOC_ITEM_ID", Gr.GetFieldValue("DOC_ITEM_ID"))
         Call Plb.QueryData(1, m_Rs2, iCount2)
         Set Plb = Nothing
         While Not m_Rs2.EOF
            Set Plb = New CPrintLabel
            Call Plb.PopulateFromRS(1, m_Rs2)
            Plb.Flag = "I"
            Call Gr.PrintLabels.add(Plb)
            Set Plb = Nothing
            
            m_Rs2.MoveNext
         Wend
         
         Set Dik = Nothing
         Set Dik = New CDocItemLink
         Dik.DOC_ITEM_ID = Gr.GetFieldValue("DOC_ITEM_ID")
         Call Dik.QueryData(1, m_Rs3, iCount2, True)
         Set Dik = Nothing
         While Not m_Rs3.EOF
            Set Dik = Nothing
            Set Dik = New CDocItemLink
            Call Dik.PopulateFromRS(1, m_Rs3)
            Dik.Flag = "I"
            Call Gr.LotItemLinkCollection.add(Dik)
            Set Dik = Nothing
            
            m_Rs3.MoveNext
         Wend
         
         Gr.Flag = "I"
         Call Ua.DocItems.add(Gr)
         Set Gr = Nothing
         
         m_Rs1.MoveNext
      Wend
      Set Plb = Nothing
      Set Gr = Nothing
      
      Dim Rcp As CRcpCnDn_Item
      Set Rcp = New CRcpCnDn_Item
      Rcp.CheckBillsFlag = Ua.CHECK_BILLS_FLAG
      Call Rcp.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Rcp.QueryData(1, m_Rs1, iCount)
      Set Rcp = Nothing
      
      Set Ua.RcpCnDnItems = Nothing
      Set Ua.RcpCnDnItems = New Collection

      While Not m_Rs1.EOF
         Set Rcp = New CRcpCnDn_Item
         Call Rcp.PopulateFromRS(1, m_Rs1)
         Rcp.Flag = "I"
         'Debug.Print (Trim(Str(Rcp.GetFieldValue("DOC_ID"))))
         Call Ua.RcpCnDnItems.add(Rcp, Trim(Str(Rcp.GetFieldValue("DOC_ID"))))
         
         Set Rcp = Nothing

         m_Rs1.MoveNext
      Wend
      Set Rcp = Nothing
   
      Dim BSub As CBillingSubTract
      Set BSub = New CBillingSubTract
      Call BSub.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BSub.QueryData(1, m_Rs1, iCount)
      Set BSub = Nothing
      
      Set Ua.BillingSubTracts = Nothing
      Set Ua.BillingSubTracts = New Collection

      While Not m_Rs1.EOF
         Set BSub = New CBillingSubTract
         Call BSub.PopulateFromRS(1, m_Rs1)
         BSub.Flag = "I"
         Call Ua.BillingSubTracts.add(BSub)
         Set BSub = Nothing

         m_Rs1.MoveNext
      Wend
      Set BSub = Nothing
      
      Dim BAdd As CBillingAddition
      Set BAdd = New CBillingAddition
      Call BAdd.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BAdd.QueryData(1, m_Rs1, iCount)
      Set BAdd = Nothing
      
      Set Ua.BillingAdditions = Nothing
      Set Ua.BillingAdditions = New Collection
      
      While Not m_Rs1.EOF
         Set BAdd = New CBillingAddition
         Call BAdd.PopulateFromRS(1, m_Rs1)
         BAdd.Flag = "I"
         Call Ua.BillingAdditions.add(BAdd)
         Set BAdd = Nothing

         m_Rs1.MoveNext
      Wend
      Set BAdd = Nothing
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.Payments = Nothing
      Set Ua.Payments = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.Payments.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
       Dim Ced As CExtraDiscount
      Set Ced = New CExtraDiscount
      Call Ced.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ced.SetFieldValue("ORDER_BY", 1)
      Call Ced.SetFieldValue("ORDER_TYPE", 1)
      Call Ced.QueryData(1, m_Rs1, iCount)
      Set Ced = Nothing
      
      Set Ua.ExtraDiscounts = Nothing
      Set Ua.ExtraDiscounts = New Collection

      While Not m_Rs1.EOF
         Set Ced = New CExtraDiscount
         Call Ced.PopulateFromRS(1, m_Rs1)
         Ced.Flag = "I"
         Call Ua.ExtraDiscounts.add(Ced)
         Set Ced = Nothing

         m_Rs1.MoveNext
      Wend
      Set Ced = Nothing
      
   End If
   
   QueryBillingDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingDoc = False
End Function
Public Function QueryBillingDocOnlyDo(Ua As CBillingDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim iCount As Long

   RName = "QueryBillingDocOnlyDo"
   QueryBillingDocOnlyDo = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CDocItem
      Set Gr = New CDocItem
      
      Dim Plb As CPrintLabel
      Dim iCount2 As Long
      Set Plb = Nothing
      
      Call Gr.SetFieldValue("DOC_ITEM_TYPE", Ua.DOC_ITEM_TYPE)
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DocItems = Nothing
      Set Ua.DocItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CDocItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Set Plb = Nothing
         Set Plb = New CPrintLabel
         
         Call Plb.SetFieldValue("DOC_ITEM_ID", Gr.GetFieldValue("DOC_ITEM_ID"))
         Call Plb.QueryData(1, m_Rs2, iCount2)
         Set Plb = Nothing
         While Not m_Rs2.EOF
            Set Plb = New CPrintLabel
            Call Plb.PopulateFromRS(1, m_Rs2)
            Plb.Flag = "I"
            Call Gr.PrintLabels.add(Plb)
            Set Plb = Nothing
            
            m_Rs2.MoveNext
         Wend
                  
         Gr.Flag = "I"
         Call Ua.DocItems.add(Gr)
         Set Gr = Nothing
         
         m_Rs1.MoveNext
      Wend
      Set Plb = Nothing
      Set Gr = Nothing
      
   End If
   
   
   
   QueryBillingDocOnlyDo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryBillingDocOnlyDo = False
End Function

Public Function QueryCheque(Ua As CCheque, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCheque"
   QueryCheque = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryCheque = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCheque = False
End Function

Public Function DeleteCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCheque"
   DeleteCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
'   ErrorObj.SystemErrorMsg = Err.Description
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   
   IsOK = False

   DeleteCheque = True
End Function

Public Function DeleteBillingDoc(Ua As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim BD As CBillingDoc
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "DeleteBillingDoc"
   DeleteBillingDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Set BD = New CBillingDoc
   Set m_Rs = New ADODB.Recordset
   
   BD.BILLING_DOC_ID = Ua.BILLING_DOC_ID
   BD.QueryFlag = 1
   Call QueryBillingDoc(BD, m_Rs, ItemCount, IsOK, glbErrorLog)
   
   Call BD.PopulateFromRS(1, m_Rs)
   
   Ua.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
   Call Ua.DeleteData
   
   'Call UpDateBDRcpCnDnItem(BD) 'ปิดไว้เนื่องจากว่ามนไม่จำเป็นที่จะต้องดูที่หน้านี้ดูจากรายงานน่าจะดีกว่า
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteBillingDoc = True
End Function

Public Function AddEditBillingDoc(Ua As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PaymentType As Long
Dim OldPaymentType As Long
Dim I As Long
   RName = "AddEditBillingDoc"
   AddEditBillingDoc = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cd As CConfigDoc
    If Ua.ShowMode = SHOW_ADD And Ua.RUNNING_NO > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ua.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Gr As CDocItem
   Dim Plb As CPrintLabel
   Dim Dik As CDocItemLink
   For Each Gr In Ua.DocItems
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
      For Each Plb In Gr.PrintLabels
         Call Plb.SetFieldValue("DOC_ITEM_ID", Gr.GetFieldValue("DOC_ITEM_ID"))
         If Ua.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Call Plb.SetFieldValue("TOTAL_AMOUNT_SUM", -1 * Plb.GetFieldValue("TOTAL_AMOUNT"))
            Call Plb.SetFieldValue("TOTAL_PRICE_SUM", -1 * Plb.GetFieldValue("TOTAL_PRICE"))
            Call Plb.SetFieldValue("DISCOUNT_AMOUNT_SUM", -1 * Plb.GetFieldValue("DISCOUNT_AMOUNT"))
            Call Plb.SetFieldValue("EXTRA_DISCOUNT_AMOUNT_SUM", -1 * Plb.GetFieldValue("EXTRA_DISCOUNT_AMOUNT"))
         Else
            Call Plb.SetFieldValue("TOTAL_AMOUNT_SUM", Plb.GetFieldValue("TOTAL_AMOUNT"))
            Call Plb.SetFieldValue("TOTAL_PRICE_SUM", Plb.GetFieldValue("TOTAL_PRICE"))
            Call Plb.SetFieldValue("DISCOUNT_AMOUNT_SUM", Plb.GetFieldValue("DISCOUNT_AMOUNT"))
            Call Plb.SetFieldValue("EXTRA_DISCOUNT_AMOUNT_SUM", Plb.GetFieldValue("EXTRA_DISCOUNT_AMOUNT"))
         End If
         If Plb.Flag = "A" Then
            Plb.ShowMode = SHOW_ADD
            Call Plb.AddEditData
         ElseIf Plb.Flag = "E" Then
            Plb.ShowMode = SHOW_EDIT
            Call Plb.AddEditData
         ElseIf Plb.Flag = "D" Then
            Call Plb.DeleteData
         End If
      Next Plb
      
      For Each Dik In Gr.LotItemLinkCollection
         Dik.DOC_ITEM_ID = Gr.GetFieldValue("DOC_ITEM_ID")
         If Dik.Flag = "A" Then
            Dik.AddEditMode = SHOW_ADD
            Call Dik.AddEditData
         ElseIf Dik.Flag = "E" Then
            Dik.AddEditMode = SHOW_EDIT
            Call Dik.AddEditData
         ElseIf Dik.Flag = "D" Then
            Call Dik.DeleteData
         End If
      Next Dik
   Next Gr
   
   Set Gr = Nothing
   Set Plb = Nothing
   
   Dim Rc As CRcpCnDn_Item
   For Each Rc In Ua.RcpCnDnItems
      Call Rc.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Rc.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If Rc.Flag = "A" Then
         Rc.ShowMode = SHOW_ADD
         Call Rc.AddEditData
      ElseIf Rc.Flag = "E" Then
         Rc.ShowMode = SHOW_EDIT
         Call Rc.AddEditData
      ElseIf Rc.Flag = "D" Then
         Call Rc.DeleteData
      End If
   Next Rc
   Set Rc = Nothing
   
   Dim BSub As CBillingSubTract
   For Each BSub In Ua.BillingSubTracts
      Call BSub.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BSub.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If BSub.Flag = "A" Then
         BSub.ShowMode = SHOW_ADD
         Call BSub.AddEditData
      ElseIf BSub.Flag = "E" Then
         BSub.ShowMode = SHOW_EDIT
         Call BSub.AddEditData
      ElseIf BSub.Flag = "D" Then
         Call BSub.DeleteData
      End If
   Next BSub
   Set BSub = Nothing
   
   Dim BAdd As CBillingAddition
   For Each BAdd In Ua.BillingAdditions
      Call BAdd.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BAdd.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If BAdd.Flag = "A" Then
         BAdd.ShowMode = SHOW_ADD
         Call BAdd.AddEditData
      ElseIf BAdd.Flag = "E" Then
         BAdd.ShowMode = SHOW_EDIT
         Call BAdd.AddEditData
      ElseIf BAdd.Flag = "D" Then
         Call BAdd.DeleteData
      End If
   Next BAdd
   Set BAdd = Nothing
   
   Dim Edc As CExtraDiscount
  For Each Edc In Ua.ExtraDiscounts
   Call Edc.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
   If Edc.Flag = "A" Then
      Edc.ShowMode = SHOW_ADD
      Call Edc.AddEditData
   ElseIf Edc.Flag = "E" Then
      Edc.ShowMode = SHOW_EDIT
      Call Edc.AddEditData
   ElseIf Edc.Flag = "D" Then
      Call Edc.DeleteData
   End If
Next Edc
Set Edc = Nothing
   
   
   Dim Ct As CCashTran
   Dim TempChequeID As Long
   For Each Ct In Ua.Payments
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
      Call Ct.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
      Call Ct.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
      Call Ct.SetFieldValue("EMP_ID", Ua.SALE_BY)
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Or _
         Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "I" Then 'เช็ค
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_TYPE", Ua.DOCUMENT_TYPE)
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
         End If
         If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "E" Then 'เช็คออกจากมือ
            Call Ct.SetFieldValue("CHEQUE_ID", TempChequeID)
         Else
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            TempChequeID = Ct.GetFieldValue("CHEQUE_ID")
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And Not (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT) _
         And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) And Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "E" Then 'เช็คออกจากมือ
               Call Ct.SetFieldValue("CHEQUE_ID", TempChequeID)
            Else
               Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
               TempChequeID = Ct.GetFieldValue("CHEQUE_ID")
            End If
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT) _
         And Not (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHEQUE_ID", -1)
            Call Ct.AddEditData
            
            Call Ct.Cheque.DeleteData
         ElseIf ((PaymentType = OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT)) _
         Or ((PaymentType <> OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) _
         And (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT)) Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If Ct.GetFieldValue("OLD_CHEQUE_ID") > 0 Then 'เช็ค
            Call Ct.Cheque.DeleteData
         End If
      End If
   Next Ct

   'ทำเพื่อเซต APAR_ID ให้ตรงกันกับที่ BILLING_DOC
   For Each Ct In Ua.Payments
      If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
         Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
         Call Ct.Cheque.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
         Call Ct.Cheque.UpdateApArID
      End If
   Next Ct
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBillingDoc = False
End Function

Public Function AddEditCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCheque"
   AddEditCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCheque = False
End Function

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData(Ug.ShowMode)

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function
Public Function DO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
Dim Ei As CLotItem
Dim PartItemID As Long
Dim DocItemLink As CDocItemLink
Dim LotItemLink As CLotItemLink
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   DO2InventoryDoc = True
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Call Ivd.SetFieldValue("INVENTORY_DOC_ID", BD.INVENTORY_DOC_ID)
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
            
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Call Ivd.SetFieldValue("DOCUMENT_DATE", BD.DOCUMENT_DATE)
   Call Ivd.SetFieldValue("DOCUMENT_NO", BD.DOCUMENT_NO)
   Call Ivd.SetFieldValue("COMMIT_FLAG", BD.COMMIT_FLAG)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call Ivd.SetFieldValue("APAR_MAS_ID", BD.APAR_MAS_ID)
   
   Call Ivd.SetFieldValue("DELIVERY_FEE", 0)
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("DEPARTMENT_ID", BD.DEPARTMENT_ID)
   Call Ivd.SetFieldValue("CANCEL_FLAG", BD.CANCEL_FLAG)

   For Each Di In BD.DocItems
      If Di.GetFieldValue("SELL_TYPE") = 1 Then
         PartItemID = Di.GetFieldValue("PART_ITEM_ID")
         If Di.Flag = "A" Or Di.GetFieldValue("SELL_TYPE_OLD") <> 1 Then
            Set Ei = New CLotItem
            
            If Area = 1 Then
               Ei.TX_TYPE = "E"
               Ei.MULTIPLIER = -1
               If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                  Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
               End If
            ElseIf Area = 2 Then
               Ei.TX_TYPE = "I"
               Ei.MULTIPLIER = 1
               Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
               Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
               Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
            End If
            
            Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
            Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                        
            Ei.Flag = "A"
            
            Ei.PART_ITEM_ID = PartItemID
            
            Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
            Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
            
            Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
            Ei.CANCEL_FLAG = BD.CANCEL_FLAG
            Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
            
            '--------------------------------------------------------------------------------------------------------------------------------------------
            For Each DocItemLink In Di.LotItemLinkCollection
               Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
               If LotItemLink Is Nothing Then
                  Set LotItemLink = New CLotItemLink
                  LotItemLink.Flag = "A"
                  LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                  LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                  LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                  Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
               Else
                  LotItemLink.Flag = "E"
                  LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
               End If
            Next DocItemLink
            '--------------------------------------------------------------------------------------------------------------------------------------------
            
            If PartItemID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
               Call Ivd.ImportExportItems.add(Ei)
            End If
            Set Ei = Nothing
         ElseIf Di.Flag = "E" Then
            Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
            If Ei Is Nothing Then
               'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
               If PartItemID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
                  Set Ei = New CLotItem
                  If Area = 1 Then
                     Ei.TX_TYPE = "E"
                     Ei.MULTIPLIER = -1
                     If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                        Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                        Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                        Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                     End If
                  ElseIf Area = 2 Then
                     Ei.TX_TYPE = "I"
                     Ei.MULTIPLIER = 1
                     Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                     Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
                     Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
                  End If
                  Ei.Flag = "E"
                  
                  Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
            
                  Ei.PART_ITEM_ID = PartItemID
                  Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  
                  Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
                  Ei.CANCEL_FLAG = BD.CANCEL_FLAG
                  Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  For Each DocItemLink In Di.LotItemLinkCollection
                     Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
                     If LotItemLink Is Nothing Then
                        Set LotItemLink = New CLotItemLink
                        LotItemLink.Flag = "A"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                        LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                        LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                        Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
                     Else
                        LotItemLink.Flag = "E"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                     End If
                  Next DocItemLink
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  Call Ivd.ImportExportItems.add(Ei)
                  Set Ei = Nothing
               End If
            Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
               If PartItemID > 0 Then
                  'แก้ไขปกติ
                  Ei.Flag = Di.Flag
                  If Area = 1 Then
                     Ei.TX_TYPE = "E"
                     Ei.MULTIPLIER = -1
                     If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                        Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                        Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                        Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                     End If
                  ElseIf Area = 2 Then
                     Ei.TX_TYPE = "I"
                     Ei.MULTIPLIER = 1
                     Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                     Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
                     Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
                  End If
                  
                  Ei.PART_ITEM_ID = PartItemID
                  
                  Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                  
                  Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  
                  Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  
                  Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
                  Ei.CANCEL_FLAG = BD.CANCEL_FLAG
                  Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  For Each DocItemLink In Di.LotItemLinkCollection
                     Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
                     If LotItemLink Is Nothing Then
                        Set LotItemLink = New CLotItemLink
                        LotItemLink.Flag = "A"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                        LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                        LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                        Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
                     Else
                        LotItemLink.Flag = "E"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                     End If
                  Next DocItemLink
                  '--------------------------------------------------------------------------------------------------------------------------------------------
               Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
                  Ei.Flag = "D"
               End If
            End If
         ElseIf Di.Flag = "D" Then
            Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
            If Not (Ei Is Nothing) Then
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.GetFieldValue("SELL_TYPE_OLD") = 1 And Di.GetFieldValue("SELL_TYPE") <> 1 Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function Rt2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
Dim Ei As CLotItem
Dim PartItemID As Long
Dim DocItemLink As CDocItemLink
Dim LotItemLink As CLotItemLink
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   Rt2InventoryDoc = True
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Call Ivd.SetFieldValue("INVENTORY_DOC_ID", BD.INVENTORY_DOC_ID)
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
            
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Call Ivd.SetFieldValue("DOCUMENT_DATE", BD.DOCUMENT_DATE)
   Call Ivd.SetFieldValue("DOCUMENT_NO", BD.DOCUMENT_NO)
   Call Ivd.SetFieldValue("COMMIT_FLAG", BD.COMMIT_FLAG)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call Ivd.SetFieldValue("APAR_MAS_ID", BD.APAR_MAS_ID)
   
   Call Ivd.SetFieldValue("DELIVERY_FEE", 0)
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("DEPARTMENT_ID", BD.DEPARTMENT_ID)
   Call Ivd.SetFieldValue("CANCEL_FLAG", BD.CANCEL_FLAG)
   
   For Each Di In BD.DocItems
      If Di.GetFieldValue("SELL_TYPE") = 1 Then
         PartItemID = Di.GetFieldValue("PART_ITEM_RETURN_ID")
         If Di.Flag = "A" Or Di.GetFieldValue("SELL_TYPE_OLD") <> 1 Then
            Set Ei = New CLotItem
            
            If Area = 2 Then     ' ส่งคืนสินค้า
               Ei.TX_TYPE = "E"
               Ei.MULTIPLIER = -1
               If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                  If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                     Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                     Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                     Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                  Else
                     Ei.TOTAL_INCLUDE_PRICE = 0
                     Ei.AVG_PRICE = 0
                     Ei.NEW_AVG_PRICE = 0
                  End If
               End If
            ElseIf Area = 1 Then   ' รับคืนสินค้า
               Ei.TX_TYPE = "I"
               Ei.MULTIPLIER = 1
               If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE_EX") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
                  Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
               Else
                  Ei.TOTAL_INCLUDE_PRICE = 0
                  Ei.AVG_PRICE = 0
                  Ei.NEW_AVG_PRICE = 0
               End If
            End If
            
            Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
            Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
            
            Ei.Flag = "A"
            
            Ei.PART_ITEM_ID = PartItemID
            
            Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
            If Di.GetFieldValue("LOT_ITEM_FLAG") = "N" Then
               Ei.TX_AMOUNT = 0
            Else
               Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
            End If
            
            Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
            Ei.CANCEL_FLAG = BD.CANCEL_FLAG
            Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
            
            '--------------------------------------------------------------------------------------------------------------------------------------------
            For Each DocItemLink In Di.LotItemLinkCollection
               Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
               If LotItemLink Is Nothing Then
                  Set LotItemLink = New CLotItemLink
                  LotItemLink.Flag = "A"
                  LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                  LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                  LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                  Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
               Else
                  LotItemLink.Flag = "E"
                  LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
               End If
            Next DocItemLink
            '--------------------------------------------------------------------------------------------------------------------------------------------
            
            If PartItemID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
               Call Ivd.ImportExportItems.add(Ei)
            End If
            Set Ei = Nothing
         ElseIf Di.Flag = "E" Then
            Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
            If Ei Is Nothing Then
               'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
               If PartItemID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
                  Set Ei = New CLotItem
                  If Area = 2 Then
                     Ei.TX_TYPE = "E"
                     Ei.MULTIPLIER = -1
                     If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                        If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                           Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                           Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                           Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                        Else
                           Ei.TOTAL_INCLUDE_PRICE = 0
                           Ei.AVG_PRICE = 0
                           Ei.NEW_AVG_PRICE = 0
                        End If
                     End If
                  ElseIf Area = 1 Then
                     Ei.TX_TYPE = "I"
                     Ei.MULTIPLIER = 1
                     If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                        Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE_EX") * Di.GetFieldValue("ITEM_AMOUNT")
                        Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
                        Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
                     Else
                        Ei.TOTAL_INCLUDE_PRICE = 0
                        Ei.AVG_PRICE = 0
                        Ei.NEW_AVG_PRICE = 0
                     End If
                  End If
                  Ei.Flag = "E"
                  
                  Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
            
                  Ei.PART_ITEM_ID = PartItemID
                  Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  If Di.GetFieldValue("LOT_ITEM_FLAG") = "N" Then
                     Ei.TX_AMOUNT = 0
                  Else
                     Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  End If
                  Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
                  Ei.CANCEL_FLAG = BD.CANCEL_FLAG
                  Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  For Each DocItemLink In Di.LotItemLinkCollection
                     Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
                     If LotItemLink Is Nothing Then
                        Set LotItemLink = New CLotItemLink
                        LotItemLink.Flag = "A"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                        LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                        LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                        Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
                     Else
                        LotItemLink.Flag = "E"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                     End If
                  Next DocItemLink
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  Call Ivd.ImportExportItems.add(Ei)
                  Set Ei = Nothing
               End If
            Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
               If PartItemID > 0 Then
                  'แก้ไขปกติ
                  Ei.Flag = Di.Flag
                  If Area = 2 Then
                     Ei.TX_TYPE = "E"
                     Ei.MULTIPLIER = -1
                     If Di.GetFieldValue("CAPITAL_AMOUNT") > 0 Then
                        If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                           Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT") * Di.GetFieldValue("ITEM_AMOUNT")
                           Ei.AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                           Ei.NEW_AVG_PRICE = Di.GetFieldValue("CAPITAL_AMOUNT")
                        Else
                           Ei.TOTAL_INCLUDE_PRICE = 0
                           Ei.AVG_PRICE = 0
                           Ei.NEW_AVG_PRICE = 0
                        End If
                     End If
                  ElseIf Area = 1 Then
                     Ei.TX_TYPE = "I"
                     Ei.MULTIPLIER = 1
                     If Di.GetFieldValue("LOT_ITEM_FLAG") = "Y" Then
                        Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE_EX") * Di.GetFieldValue("ITEM_AMOUNT")
                        Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
                        Ei.NEW_AVG_PRICE = Di.GetFieldValue("AVG_PRICE_EX")
                     Else
                        Ei.TOTAL_INCLUDE_PRICE = 0
                        Ei.AVG_PRICE = 0
                        Ei.NEW_AVG_PRICE = 0
                     End If
                  End If
                  
                  Ei.PART_ITEM_ID = PartItemID
                  
                  Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                  
                  Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  
                  If Di.GetFieldValue("LOT_ITEM_FLAG") = "N" Then
                     Ei.TX_AMOUNT = 0
                  Else
                     Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  End If
                  
                  Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
                  Ei.CANCEL_FLAG = BD.CANCEL_FLAG
                  Ei.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  '--------------------------------------------------------------------------------------------------------------------------------------------
                  For Each DocItemLink In Di.LotItemLinkCollection
                     Set LotItemLink = GetObject("CLotItemLink", Ei.LotItemLink, Trim(Str(DocItemLink.IMPORT_LOT_ITEM_ID)), False)
                     If LotItemLink Is Nothing Then
                        Set LotItemLink = New CLotItemLink
                        LotItemLink.Flag = "A"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                        LotItemLink.IMPORT_LOT_ITEM_ID = DocItemLink.IMPORT_LOT_ITEM_ID
                        LotItemLink.MAIN_IMPORT_LOT_ITEM_ID = DocItemLink.MAIN_IMPORT_LOT_ITEM_ID
                        Call Ei.LotItemLink.add(LotItemLink, Trim(Str(LotItemLink.IMPORT_LOT_ITEM_ID)))
                     Else
                        LotItemLink.Flag = "E"
                        LotItemLink.IMPORT_AMOUNT = DocItemLink.IMPORT_AMOUNT
                     End If
                  Next DocItemLink
                  '--------------------------------------------------------------------------------------------------------------------------------------------
               Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
                  Ei.Flag = "D"
               End If
            End If
         ElseIf Di.Flag = "D" Then
            Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
            If Not (Ei Is Nothing) Then
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.GetFieldValue("SELL_TYPE_OLD") = 1 And Di.GetFieldValue("SELL_TYPE") <> 1 Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function

Public Function RO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Call Ivd.SetFieldValue("INVENTORY_DOC_ID", BD.INVENTORY_DOC_ID)
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
         
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
      
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
      
   Call Ivd.SetFieldValue("DOCUMENT_DATE", BD.DOCUMENT_DATE)
   Call Ivd.SetFieldValue("DOCUMENT_NO", BD.DOCUMENT_NO)
   Call Ivd.SetFieldValue("COMMIT_FLAG", BD.COMMIT_FLAG)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call Ivd.SetFieldValue("CANCEL_FLAG", BD.CANCEL_FLAG)
   
   Call Ivd.SetFieldValue("DELIVERY_FEE", 0)
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("DEPARTMENT_ID", BD.DEPARTMENT_ID)
   
   Dim II As CLotItem
   
   For Each Di In BD.DocItems
      If Di.GetFieldValue("DOC_ITEM_TYPE") = 1 Then
         If Di.Flag = "A" Or Di.GetFieldValue("DOC_ITEM_TYPE_OLD") = 2 Then
            Set II = New CLotItem
      
            II.Flag = "A"
            
            II.TX_TYPE = "I"
            II.MULTIPLIER = 1
            
            II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
            II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
            II.NEW_AVG_PRICE = II.AVG_PRICE
            
            II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
            II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
            
            II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
            II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
            II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
            II.LINK_ID = Di.GetFieldValue("LINK_ID")
            II.CANCEL_FLAG = BD.CANCEL_FLAG
            II.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
            
            II.WEIGHT_AMOUNT = Di.GetFieldValue("WEIGHT_AMOUNT")
            
            If Di.GetFieldValue("PART_ITEM_ID") > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
               Call Ivd.ImportExportItems.add(II)
            End If
         ElseIf Di.Flag = "E" Then
            Set II = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"), "I")
            If II Is Nothing Then
               'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
               If Di.GetFieldValue("PART_ITEM_ID") > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
                  
                  II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                  
                  II.TX_TYPE = "I"
                  II.MULTIPLIER = 1
                  
                  II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
                  II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
                  II.NEW_AVG_PRICE = II.AVG_PRICE
                  
                  II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
                  II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  II.LINK_ID = Di.GetFieldValue("LINK_ID")
                  II.CANCEL_FLAG = BD.CANCEL_FLAG
                  II.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  II.WEIGHT_AMOUNT = Di.GetFieldValue("WEIGHT_AMOUNT")
                  
               End If
            Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
               If Di.GetFieldValue("PART_ITEM_ID") > 0 Then
                  'แก้ไขปกติ
                  II.Flag = Di.Flag
                                    
                  II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
                  II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                  
                  II.TX_TYPE = "I"
                  II.MULTIPLIER = 1
                  
                  II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
                  II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
                  II.NEW_AVG_PRICE = II.AVG_PRICE
                  
                  II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
                  II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
                  II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
                  'II.LINK_ID = Di.GetFieldValue("LINK_ID"))  ควรจะเป็นเลขเดิม
                  II.CANCEL_FLAG = BD.CANCEL_FLAG
                  II.FREE_FLAG = Di.GetFieldValue("FREE_FLAG")
                  
                  II.WEIGHT_AMOUNT = Di.GetFieldValue("WEIGHT_AMOUNT")
               Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
                  II.Flag = "D"
               End If
            End If
         ElseIf Di.Flag = "D" Then
            Set II = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"), "I")
            If Not (II Is Nothing) Then
               II.Flag = "D"
            End If
         End If
      ElseIf Di.GetFieldValue("DOC_ITEM_TYPE_OLD") = 1 And Di.GetFieldValue("DOC_ITEM_TYPE") = 2 Then
         Set II = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"), "I")
         If Not (II Is Nothing) Then
            II.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function AddEditMasterFromTo(Ua As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditMasterFromTo"
   AddEditMasterFromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CMasterFromToDetail
   For Each Gr In Ua.Details
      Call Gr.SetFieldValue("MASTER_FROMTO_ID", Ua.GetFieldValue("MASTER_FROMTO_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr

   Dim Gx As CMasterFromToEx
   For Each Gx In Ua.CommissionExs
      Gx.MASTER_FROMTO_ID = Ua.GetFieldValue("MASTER_FROMTO_ID")
      If Gx.Flag = "A" Then
         Gx.AddEditMode = SHOW_ADD
         Call Gx.AddEditData
      ElseIf Gx.Flag = "E" Then
         Gx.AddEditMode = SHOW_EDIT
         Call Gx.AddEditData
      ElseIf Gx.Flag = "D" Then
         Call Gx.DeleteData
      End If
      
   Next Gx
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterFromTo = False
End Function

Public Function QueryMasterFromTo(Ua As CMasterFromTo, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMasterFromTo"
   QueryMasterFromTo = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CMasterFromToDetail
      Set Gr = New CMasterFromToDetail
      Call Gr.SetFieldValue("MASTER_FROMTO_ID", Ua.GetFieldValue("MASTER_FROMTO_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.Details = Nothing
      Set Ua.Details = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CMasterFromToDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.Details.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Gx As CMasterFromToEx
      Set Gx = New CMasterFromToEx
      Gx.MASTER_FROMTO_ID = Ua.GetFieldValue("MASTER_FROMTO_ID")
      Call Gx.QueryData(1, m_Rs1, iCount)
      Set Gx = Nothing

      Set Ua.CommissionExs = Nothing
      Set Ua.CommissionExs = New Collection

      While Not m_Rs1.EOF
         Set Gx = New CMasterFromToEx
         Call Gx.PopulateFromRS(1, m_Rs1)
         
         Gx.Flag = "I"
         Call Ua.CommissionExs.add(Gx)
         Set Gx = Nothing

         m_Rs1.MoveNext
      Wend
         
   End If
   
   QueryMasterFromTo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryMasterFromTo = False
End Function

Public Function DeleteMasterFromTo(Ua As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteMasterFromTo"
   DeleteMasterFromTo = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterFromTo = True
End Function
Public Function AddEditCommissionChart(Ua As CCommissionChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionChart"
   AddEditCommissionChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditCommissionChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCommissionChart = False
End Function
Public Function QueryCommissionChart(Ua As CCommissionChart, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCommissionChart"
   QueryCommissionChart = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCommissionChart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCommissionChart = False
End Function
Public Function DeleteCommissionChart(Ua As CCommissionChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCommissionChart"
   DeleteCommissionChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCommissionChart = True
End Function
Public Function CopyMasterFromTo(Ug As CMasterFromTo, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyMasterFromTo"
   CopyMasterFromTo = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryMasterFromTo(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   If Ug.GetFieldValue("MASTER_FROMTO_TYPE") = 3 Then
      Dim Cm As CCommissionChart
      Dim m_CommissionCharts As Collection
      Set Cm = New CCommissionChart
      Set m_CommissionCharts = New Collection
      Call LoadCommissionChartEx(Nothing, m_CommissionCharts, Ug.GetFieldValue("MASTER_FROMTO_ID"))
   End If

   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("MASTER_FROMTO_ID", -1)
   Call Ug.SetFieldValue("MASTER_FROMTO_NO", "***" & Ug.GetFieldValue("MASTER_FROMTO_NO"))
   
   Dim Det As CMasterFromToDetail
   For Each Det In Ug.Details
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditMasterFromTo(Ug, IsOK, False, ErrorObj)
      
   If Ug.GetFieldValue("MASTER_FROMTO_TYPE") = 3 Then
      Dim CmC As CCommissionChart
      Dim TempSearch As CCommissionChart
      
      For Each CmC In m_CommissionCharts
         CmC.ShowMode = SHOW_ADD
         Call CmC.SetFieldValue("MASTER_FROMTO_ID", Ug.GetFieldValue("MASTER_FROMTO_ID"))
         If CmC.GetFieldValue("PARENT_ID") <= 0 Then
            Call CmC.AddEditData
         Else
            Set TempSearch = GetObject("CCommissionChart", m_CommissionCharts, CmC.GetFieldValue("PARENT_ID"))
            Call CmC.SetFieldValue("PARENT_ID", TempSearch.GetFieldValue("COMMISSION_CHART_ID"))
            Call CmC.AddEditData
         End If
      Next CmC
      
   End If
   
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyMasterFromTo = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyMasterFromTo = False
End Function
Public Function AddEditPackage(Ua As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditPackage"
   AddEditPackage = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CPackageDetail
   For Each Gr In Ua.PackageDeTails
      Call Gr.SetFieldValue("PACKAGE_ID", Ua.GetFieldValue("PACKAGE_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr

   Dim Package As CPackageDetail
   Set Package = New CPackageDetail
   Call Package.SetFieldValue("PACKAGE_DETAIL_ID", 1)
   Call LoadPackageDetail(Package, Nothing, LoadPackageColl)
      
   Set Package = Nothing
   
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPackage = False
End Function

Public Function QueryPackage(Ua As CPackage, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPackage"
   QueryPackage = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CPackageDetail
      Set Gr = New CPackageDetail
      Call Gr.SetFieldValue("PACKAGE_ID", Ua.GetFieldValue("PACKAGE_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.PackageDeTails = Nothing
      Set Ua.PackageDeTails = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CPackageDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.PackageDeTails.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If

   QueryPackage = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPackage = False
End Function

Public Function DeletePackage(Ua As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeletePackage"
   DeletePackage = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePackage = True
End Function

Public Function CopyPackage(Ug As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyPackage"
   CopyPackage = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryPackage(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("PACKAGE_ID", -1)
   Call Ug.SetFieldValue("PACKAGE_NO", "***" & Ug.GetFieldValue("PACKAGE_NO"))
   Call Ug.SetFieldValue("PACKAGE_MASTER_FLAG", "N")
   
   Dim Det As CPackageDetail
   For Each Det In Ug.PackageDeTails
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditPackage(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyPackage = False
End Function
Public Function QueryTaget(Ua As CTaget, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryTaget"
   QueryTaget = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Gr As CTagetDetail
      Set Gr = New CTagetDetail
      Call Gr.SetFieldValue("TAGET_ID", Ua.GetFieldValue("TAGET_ID"))
      Call Gr.QueryData(1, m_Rs1, ItemCount)
      Set Gr = Nothing
      
      Set Ua.TagetDetails = Nothing
      Set Ua.TagetDetails = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CTagetDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.TagetDetails.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryTaget = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTaget = False
End Function

Public Function DeleteTaget(Ua As CTaget, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteTaget"
   DeleteTaget = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteTaget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteTaget = True
End Function

Public Function AddEditTaget(Ua As CTaget, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditTaget"
   AddEditTaget = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CTagetDetail
   For Each Gr In Ua.TagetDetails
      Call Gr.SetFieldValue("TAGET_ID", Ua.GetFieldValue("TAGET_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditTaget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditTaget = False
End Function
Public Function CopyTaget(Ug As CTaget, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyTaget"
   CopyTaget = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryTaget(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("TAGET_ID", -1)
   Call Ug.SetFieldValue("TAGET_DESC", "COPY จากเดือน " & Ug.GetFieldValue("MONTH_ID") & "ปี " & Ug.GetFieldValue("YEAR_NO") + 543 & " กรุณาแก้ไขเดือนและปีให้ถูกต้อง")
   
   Dim Det As CTagetDetail
   For Each Det In Ug.TagetDetails
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditTaget(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyTaget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyTaget = False
End Function
Public Sub CreateBillingTransferItems(BD As CBillingDoc)
Dim Ti As CCashTransferItem
Dim TempCt As CCashTran
Dim IIEx As CCashTran
Dim Ei As CCashTran
Dim II As CCashTran
Dim I As Long
Dim j As Long
Dim AddFlag As Boolean
Dim AddFlag1 As Boolean
Dim AddFlag2 As Boolean
Dim AddFlag3 As Boolean
   
   Set BD.TransferItems = Nothing
   Set BD.TransferItems = New Collection
   
   For I = 1 To BD.Payments.Count
      Set TempCt = BD.Payments(I)
      If TempCt.GetFieldValue("STEP_ID") > 1 Then
         If TempCt.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And TempCt.GetFieldValue("TX_TYPE") = "I" Then
            Set IIEx = BD.Payments(I)
            AddFlag1 = True
         ElseIf TempCt.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And TempCt.GetFieldValue("TX_TYPE") = "E" Then
            Set Ei = BD.Payments(I)
            AddFlag2 = True
         ElseIf TempCt.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
            Set II = BD.Payments(I)
            AddFlag3 = True
         End If
      Else
         j = I
         Set IIEx = BD.Payments(I)
         AddFlag = True
      End If
      If AddFlag Or (AddFlag1 And AddFlag2 And AddFlag3) Then
         Set Ti = New CCashTransferItem
         Set Ti.ImportItemEx = IIEx
         If (AddFlag1 And AddFlag2 And AddFlag3) Then
            Set Ti.ExportItem = Ei
            Set Ti.ImportItem = II
         End If
         Ti.Flag = "I"
         Call BD.TransferItems.add(Ti)
         Set Ti = Nothing
         AddFlag = False
         AddFlag1 = False
         AddFlag2 = False
         AddFlag3 = False
      End If
   Next I
End Sub
Public Function DeleteCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCashDoc"
   DeleteCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCashDoc = True
End Function
Public Function QueryCashDoc(Ua As CCashDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCashDoc"
   QueryCashDoc = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.CashTranItems = Nothing
      Set Ua.CashTranItems = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.SetFieldValue("TEMP_FEE_AMOUNT", Ct.GetFieldValue("FEE_AMOUNT"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.CashTranItems.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim CDP As CCashDocPost
      Set CDP = New CCashDocPost
      Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing

      Set Ua.PostItems = Nothing
      Set Ua.PostItems = New Collection

      While Not m_Rs1.EOF
         Set CDP = New CCashDocPost
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         Call Ua.PostItems.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryCashDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCashDoc = False
End Function
Public Function AddEditCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim DocType As CASH_DOC_TYPE

   RName = "AddEditCashDoc"
   AddEditCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Ct As CCashTran
   Dim PaymentType As Long
   Dim OldPaymentType As Long
   
   DocType = Ua.GetFieldValue("DOCUMENT_TYPE")
   
   For Each Ct In Ua.CashTranItems
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.SetFieldValue("TX_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Ct.SetFieldValue("TX_DATE", Ua.GetFieldValue("DOCUMENT_DATE"))
      Call Ct.SetFieldValue("APAR_MAS_ID", Ua.GetFieldValue("APAR_MAS_ID"))
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If DocType = CASH_DEPOSIT Then 'ฝากเช็คที่มีอยู่ในระบบแล้ว
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "Y")
            Call Ct.Cheque.SetFieldValue("POST_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Then 'เช็ค
               Ct.Cheque.ShowMode = SHOW_ADD
               Call Ct.Cheque.AddEditData
            End If
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'เมื่อเป็นการนำฝากจะไม่มีการแก้ไขเกิดขึ้น จะต้องลบแล้วสร้างใหม่
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = CHEQUE_HAND_PMT) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If DocType = CASH_DEPOSIT Then
            'ถ้าเป็นการลบ item สำหรับการนำเช็คในมือไปเข้า ธนาคาร ไม่ต้องลบเช็คนั้นทิ้งเพียงแต่ set BANK_FLAG ให้เป็น 'N'
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("OLD_CHEQUE_ID") > 0 Then 'เช็ค
               Call Ct.Cheque.DeleteData
            End If
         End If
      End If
   Next Ct
   
   For Each Ct In Ua.CashTranItems
        If (DocType = CASH_DEPOSIT) Then
          If Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      End If
   Next Ct
   
   Dim CDP As CCashDocPost
   Dim Ch As CCheque
   Set Ch = New CCheque
   For Each CDP In Ua.PostItems
      If CDP.Flag = "A" Then
         Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
         CDP.ShowMode = SHOW_ADD
         Call CDP.AddEditData
            
         Call Ch.SetFieldValue("CHEQUE_ID", CDP.GetFieldValue("CHEQUE_ID"))
         If CDP.GetFieldValue("POST_TYPE") = POST_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         ElseIf CDP.GetFieldValue("POST_TYPE") = WAITING_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "N")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         ElseIf CDP.GetFieldValue("POST_TYPE") = PASSED_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         End If
         
         Call Ch.UpdateBankFlag
         
      ElseIf CDP.Flag = "E" Then
         CDP.ShowMode = SHOW_EDIT
         Call CDP.AddEditData
      ElseIf CDP.Flag = "D" Then
         Call CDP.DeleteData
      End If
   Next CDP
   


   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCashDoc = False
End Function
Public Sub CreateCashTransferItems(Ivd As CCashDoc)
Dim Ti As CCashTransferItem
Dim O As Object
Dim Ei As CCashTran
Dim II As CCashTran
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.CashTranItems.Count \ 2
   Count2 = Ivd.CashTranItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.CashTranItems(j)
         If (O.GetFieldValue("TX_TYPE") = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.GetFieldValue("TX_TYPE") = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CCashTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function DeleteJob(Ua As CJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim BD As CJob
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long
   
   RName = "DeleteJob"
   DeleteJob = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Set BD = New CJob
   Set m_Rs = New ADODB.Recordset
   
   BD.JOB_ID = Ua.JOB_ID
   BD.QueryFlag = 1
   Call QueryJob(BD, m_Rs, ItemCount, IsOK, glbErrorLog)
   
   Call BD.PopulateFromRS(1, m_Rs)
   
   Ua.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
   Call Ua.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteJob = True
End Function
Public Function QueryJob(Ua As CJob, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryJob"
   QueryJob = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      
      Dim Ji As CJobItem
      Set Ji = New CJobItem
      Ji.JOB_ID = Ua.JOB_ID
      Call Ji.QueryData(1, m_Rs1, iCount, True)
      Set Ji = Nothing
      
      Set Ua.JobInItems = Nothing
      Set Ua.JobInItems = New Collection
      Set Ua.JobOutItems = Nothing
      Set Ua.JobOutItems = New Collection
      
      While Not m_Rs1.EOF
         Set Ji = New CJobItem
         Call Ji.PopulateFromRS(1, m_Rs1)
         
         If Ji.TX_TYPE = "I" Then      'ด้าน Job เข้าแสดงว่าต้องเบิกของออกมาจากคลัง
            Call Ua.JobInItems.add(Ji)
         Else
            Call Ua.JobOutItems.add(Ji)
         End If
         
         Set Ji = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryJob = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryJob = False
End Function
Public Function AddEditJob(Ua As CJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditJob"
   AddEditJob = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Ji As CJobItem
   
   For Each Ji In Ua.JobInItems
      If Ji.Flag = "A" Then
         Ji.JOB_ID = Ua.JOB_ID
         Ji.AddEditMode = SHOW_ADD
         Call Ji.AddEditData
      ElseIf Ji.Flag = "E" Then
         Ji.AddEditMode = SHOW_EDIT
         Call Ji.AddEditData
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji
   
   For Each Ji In Ua.JobOutItems
      If Ji.Flag = "A" Then
         Ji.JOB_ID = Ua.JOB_ID
         Ji.AddEditMode = SHOW_ADD
         Call Ji.AddEditData
         ' ต้อง update Export_Lot_Item_ID ให้ Link ไปยัง LotItem ได้เลยเพื่อให้การออกรายงานเป็นไปได้อย่างรวดเร็ว
      ElseIf Ji.Flag = "E" Then
         Ji.AddEditMode = SHOW_EDIT
         Call Ji.AddEditData
         ' ต้อง update Export_Lot_Item_ID ให้ Link ไปยัง LotItem ได้เลยเพื่อให้การออกรายงานเป็นไปได้อย่างรวดเร็ว
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditJob = False
End Function
Public Function DeleteFormula(Ua As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteFormula"
   DeleteFormula = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteFormula = True
End Function
Public Function QueryFormula(Ua As CFormula, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryFormula"
   QueryFormula = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
      
   If Ua.QueryFlag = 1 And ItemCount > 0 Then
      
      Dim Ji As CFormulaItem
      Set Ji = New CFormulaItem
      
      Call Ua.PopulateFromRS(1, Rs)
      Ji.FORMULA_ID = Ua.FORMULA_ID
      Call Ji.QueryData(1, m_Rs1, iCount)
      Set Ji = Nothing
      
      Set Ua.FormulaItems = Nothing
      Set Ua.FormulaItems = New Collection
      
      While Not m_Rs1.EOF
         Set Ji = New CFormulaItem
         Call Ji.PopulateFromRS(1, m_Rs1)
         
         Call Ua.FormulaItems.add(Ji)
         Set Ji = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryFormula = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryFormula = False
End Function
Public Function AddEditFormula(Ua As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFormula"
   AddEditFormula = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Ji As CFormulaItem
   
   For Each Ji In Ua.FormulaItems
      If Ji.Flag = "A" Then
         Ji.FORMULA_ID = Ua.FORMULA_ID
         Ji.AddEditMode = SHOW_ADD
         Call Ji.AddEditData
      ElseIf Ji.Flag = "E" Then
         Ji.AddEditMode = SHOW_EDIT
         Call Ji.AddEditData
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditFormula = False
End Function
Public Function CopyFormula(Ug As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyFormula"
   CopyFormula = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryFormula(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.FORMULA_ID = -1
   Ug.FORMULA_NO = "***" & Ug.FORMULA_NO
   
   Dim Det As CFormulaItem
   For Each Det In Ug.FormulaItems
      Det.Flag = "A"
      Det.AddEditMode = SHOW_ADD
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditFormula(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyFormula = False
End Function
Public Function Job2InventoryDoc(Jb As CJob, Ivd As CInventoryDoc, Optional DocumentType As Long = -1, Optional TempColl As Collection, Optional StockColl As Collection) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Ji As CJobItem
Dim Ei As CLotItem
Dim Lk As CLotItemLink
Dim TempLk As CLotItemLink
Dim TempSc As CStockCode
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   If Jb.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Call Ivd.SetFieldValue("INVENTORY_DOC_ID", Jb.INVENTORY_DOC_ID)
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Call Ivd.SetFieldValue("DOCUMENT_DATE", Jb.JOB_DATE)
   Call Ivd.SetFieldValue("DOCUMENT_NO", Jb.JOB_NO)
   Call Ivd.SetFieldValue("COMMIT_FLAG", Jb.COMMIT_FLAG)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("CANCEL_FLAG", "N")
   
   For Each Ji In Jb.JobInItems ' In ของ Job ก็ คืนการเบิกออกจากคลัง
      If Ji.Flag = "A" Then
         Set Ei = New CLotItem
               
         Ei.TX_TYPE = "E"
         Ei.MULTIPLIER = -1
            
         Ei.UNIT_TRAN_ID = Ji.UNIT_ID
         Ei.UNIT_MULTIPLE = 1
         
         Ei.Flag = "A"
         
         Ei.FREE_FLAG = "N"
         Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
         Ei.LOCATION_ID = Ji.LOCATION_ID
         Ei.TX_AMOUNT = Ji.TX_AMOUNT
         
         Ei.LINK_ID = Ji.LINK_ID
         Ei.CANCEL_FLAG = "N"
                        
         For Each Lk In TempColl
            If Lk.PART_ITEM_ID = Ji.PART_ITEM_ID Then
               Lk.Flag = "A"
               Call Ei.LotItemLink.add(Lk)
            End If
         Next Lk
         Call Ivd.ImportExportItems.add(Ei)
         Set Ei = Nothing
      End If
   Next Ji
   
   For Each Ji In Jb.JobOutItems ' Out ของ Job ก็ คืนการเก็บสินค้าที่ได้เข้าคลังหลังจากผลิตเสร็จ
      If Ji.LOCATION_ID > 0 Then
         If Ji.Flag = "A" Then
            Set Ei = New CLotItem
            
            Ei.TX_TYPE = "I"
            Ei.MULTIPLIER = 1
            
            Ei.UNIT_TRAN_ID = Ji.UNIT_ID
            Ei.UNIT_MULTIPLE = 1
            
            Ei.Flag = "A"
            Ei.FREE_FLAG = "N"
            Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
            
            Ei.LOCATION_ID = Ji.LOCATION_ID
            Ei.TX_AMOUNT = Ji.TX_AMOUNT
            Set TempSc = GetObject("CStockCode", StockColl, Trim(Str(Ji.PART_ITEM_ID)), False)
            If Not (TempSc Is Nothing) Then
               Ei.AVG_PRICE = TempSc.COST_PER_AMOUNT
               Ei.TOTAL_INCLUDE_PRICE = Ei.AVG_PRICE * Ei.TX_AMOUNT
            End If
            
            Ei.LINK_ID = Ji.LINK_ID
            Ei.CANCEL_FLAG = "N"
            
            If Ji.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
               Call Ivd.ImportExportItems.add(Ei)
            End If
            Set Ei = Nothing
         ElseIf Ji.Flag = "E" Then
            Set Ei = GetExportItem(Ivd, Ji.LINK_ID)
            If Ji.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Ji.Flag
               Ei.TX_AMOUNT = Ji.TX_AMOUNT
               Ei.CANCEL_FLAG = "N"
            End If
         End If
      End If
   Next Ji
   
End Function
Public Function GetNextLotItemID(LotItemID As Long, IvdID As Long, PartItemID As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim Lt As CLotItem
Dim ItemCount  As Long
Dim TempValue  As Long
Dim TempMainValue As Long

   GetNextLotItemID = False
   MasterInd = "8"
   Set Lt = New CLotItem
   MasterInd = "1"
   
   Set TempRs = New ADODB.Recordset
   Lt.INVENTORY_DOC_ID = IvdID
   Lt.PART_ITEM_ID = PartItemID
   
   Call Lt.QueryData(8, TempRs, ItemCount, False)
   
   TempValue = 0
   TempMainValue = 0
   While Not TempRs.EOF
      Call Lt.PopulateFromRS(8, TempRs)
      
      If (Lt.LOT_ITEM_ID > TempValue) And (Lt.LOT_ITEM_ID < LotItemID) Then
         TempValue = Lt.LOT_ITEM_ID
         TempMainValue = Lt.MAIN_IMPORT_LOT_ITEM_ID
      End If
      
      GetNextLotItemID = True
      TempRs.MoveNext
   Wend
   
   LotItemID = TempMainValue
   
   Set TempRs = Nothing
   Set Lt = Nothing
End Function
Public Function AddEditBalanceVerify(Ua As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditBalanceVerify"
   AddEditBalanceVerify = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CBalanceVerifyDeTail
   For Each Gr In Ua.BalanceVerifyDetail
      Gr.BALANCE_VERIFY_ID = Ua.BALANCE_VERIFY_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBalanceVerify = False
End Function

Public Function QueryBalanceVerify(Ua As CBalanceVerify, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBalanceVerify"
   QueryBalanceVerify = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CBalanceVerifyDeTail
      Set Gr = New CBalanceVerifyDeTail
      Gr.BALANCE_VERIFY_ID = Ua.BALANCE_VERIFY_ID
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.BalanceVerifyDetail = Nothing
      Set Ua.BalanceVerifyDetail = New Collection
      
      While Not m_Rs1.EOF
         Set Gr = New CBalanceVerifyDeTail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.BalanceVerifyDetail.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If

   QueryBalanceVerify = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBalanceVerify = False
End Function

Public Function DeleteBalanceVerify(Ua As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteBalanceVerify"
   DeleteBalanceVerify = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBalanceVerify = True
End Function
Public Function CopyBalanceVerify(Ug As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyBalanceVerify"
   CopyBalanceVerify = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryBalanceVerify(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.BALANCE_VERIFY_ID = -1
   Ug.BALANCE_VERIFY_NO = "***" & Ug.BALANCE_VERIFY_NO
   
   Dim Det As CBalanceVerifyDeTail
   For Each Det In Ug.BalanceVerifyDetail
      Det.Flag = "A"
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditBalanceVerify(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBalanceVerify = False
End Function
Public Function CopyBalanceVerifyToAdjust(Ug As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim Ivd As CInventoryDoc
Dim Lt As CLotItem
Dim InventoryBals2 As Collection
Dim TempLt As CLotItem
Dim Pt As CStockCode
Dim PartItemColl As Collection
   
   RName = "CopyBalanceVerifyToAdjust"
   CopyBalanceVerifyToAdjust = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryBalanceVerify(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Set InventoryBals2 = New Collection
   Set PartItemColl = New Collection
   
   Call LoadLeftAmountLotItem(InventoryBals2, -1, Ug.BALANCE_VERIFY_DATE)
   Call LoadStockCode(Nothing, PartItemColl)
      
   Set Ivd = New CInventoryDoc
   Ivd.ShowMode = SHOW_ADD
   Call Ivd.SetFieldValue("INVENTORY_DOC_ID", -1)
   Call Ivd.SetFieldValue("DOCUMENT_NO", Ug.BALANCE_VERIFY_NO & "/" & Now)
   Call Ivd.SetFieldValue("DOCUMENT_DATE", Ug.BALANCE_VERIFY_DATE)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", ADJUST_DOCTYPE)
   
   
   Call Ivd.SetFieldValue("COMMIT_FLAG", "N")
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("DOCUMENT_DESC", "มากจากการตั้งค่าของหน้าจอตรวจยอดสต็อค" & Now)
   Call Ivd.SetFieldValue("CANCEL_FLAG", "N")
   
   Dim Det As CBalanceVerifyDeTail
   For Each Det In Ug.BalanceVerifyDetail
      Set TempLt = GetObject("CLotItem", InventoryBals2, Trim(Det.LOCATION_ID & "-" & Det.PART_ITEM_ID))
      If Not (TempLt.SUM_AMOUNT = Det.PART_ITEM_AMOUNT) Then
         Set Lt = New CLotItem
         Lt.Flag = "A"
         If TempLt.SUM_AMOUNT > Det.PART_ITEM_AMOUNT Then    ' ต้องปรับ ออก
            Lt.TX_AMOUNT = TempLt.SUM_AMOUNT - Det.PART_ITEM_AMOUNT
            Lt.TX_TYPE = "E"
            Lt.MULTIPLIER = -1
         Else
            Lt.TX_AMOUNT = Det.PART_ITEM_AMOUNT - TempLt.SUM_AMOUNT
            Lt.TX_TYPE = "I"
            Lt.MULTIPLIER = 1
         End If
                           
         Lt.PART_ITEM_ID = Det.PART_ITEM_ID
         Lt.LOCATION_ID = Det.LOCATION_ID
         
         Set Pt = GetObject("CStockCode", PartItemColl, Trim(Str(Det.PART_ITEM_ID)))
         Lt.UNIT_TRAN_ID = Pt.UNIT_CHANGE_ID
         Lt.UNIT_MULTIPLE = 1
         
         Call Ivd.ImportExportItems.add(Lt)
         Set Lt = Nothing
      End If
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   Set InventoryBals2 = Nothing
   Set Ivd = Nothing
   Set PartItemColl = Nothing
   
   HasBegin = False
   IsOK = True

   CopyBalanceVerifyToAdjust = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBalanceVerifyToAdjust = False
   Set InventoryBals2 = Nothing
   Set Ivd = Nothing
End Function

Public Function AddEditKeyAccount(Ua As CKeyAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditKeyAccount"
   AddEditKeyAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CKeyAccountDetail
   For Each Gr In Ua.KeyAccountDetail
      Gr.KEY_ACCOUNT_ID = Ua.KEY_ACCOUNT_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditKeyAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditKeyAccount = False
End Function

Public Function QueryKeyAccount(Ua As CKeyAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryKeyAccount"
   QueryKeyAccount = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Call Ua.PopulateFromRS(1, Rs)
      If Ua.KEY_ACCOUNT_ID > 0 Then
         Dim Gr As CKeyAccountDetail
         Set Gr = New CKeyAccountDetail
         Gr.KEY_ACCOUNT_ID = Ua.KEY_ACCOUNT_ID
         Call Gr.QueryData(1, m_Rs1, iCount)
         Set Gr = Nothing
         
         Set Ua.KeyAccountDetail = Nothing
         Set Ua.KeyAccountDetail = New Collection
         
         While Not m_Rs1.EOF
            Set Gr = New CKeyAccountDetail
            Call Gr.PopulateFromRS(1, m_Rs1)
            
            Gr.Flag = "I"
            Call Ua.KeyAccountDetail.add(Gr)
            Set Gr = Nothing
   
            m_Rs1.MoveNext
         Wend
      End If
   End If

   QueryKeyAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryKeyAccount = False
End Function

Public Function DeleteKeyAccount(Ua As CKeyAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteKeyAccount"
   DeleteKeyAccount = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteKeyAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteKeyAccount = True
End Function
Public Function AddEditTagetJob(Ua As CTagetJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditTagetJob"
   AddEditTagetJob = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CTagetJobDetail
   For Each Gr In Ua.TagetJobDetail
      Gr.TAGET_JOB_ID = Ua.TAGET_JOB_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditTagetJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditTagetJob = False
End Function

Public Function QueryTagetJob(Ua As CTagetJob, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryTagetJob"
   QueryTagetJob = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CTagetJobDetail
      Set Gr = New CTagetJobDetail
      Gr.TAGET_JOB_ID = Ua.TAGET_JOB_ID
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.TagetJobDetail = Nothing
      Set Ua.TagetJobDetail = New Collection
      
      While Not m_Rs1.EOF
         Set Gr = New CTagetJobDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.TagetJobDetail.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If

   QueryTagetJob = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTagetJob = False
End Function
Public Function DeleteTagetJob(Ua As CTagetJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteTagetJob"
   DeleteTagetJob = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteTagetJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteTagetJob = True
End Function
Public Function QueryLockDate(Ua As CLockDate, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryLockDate"
   QueryLockDate = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryLockDate = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLockDate = False
End Function
Public Function AddEditLockDate(Ua As CLockDate, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLockDate"
   AddEditLockDate = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditLockDate = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLockDate = False
End Function

Public Function DeleteSumBill(Ua As CSumBill, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteSumBill"
   DeleteSumBill = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ua.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSumBill = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteSumBill = True
End Function
Public Function QuerySumBill(Ua As CSumBill, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySumBill"
   QuerySumBill = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      Dim Rcp As CBillDetail
      Set Rcp = New CBillDetail
      Rcp.SUM_BILL_ID = Ua.SUM_BILL_ID
      Call Rcp.QueryData(1, m_Rs1, iCount)
      Set Rcp = Nothing
      
      Set Ua.RcpCnDnItems = Nothing
      Set Ua.RcpCnDnItems = New Collection

      While Not m_Rs1.EOF
         Set Rcp = New CBillDetail
         Call Rcp.PopulateFromRS(1, m_Rs1)
         Rcp.Flag = "I"
         
         Call Ua.RcpCnDnItems.add(Rcp, Trim(Str(Rcp.BILLING_DOC_ID)))
         
         Set Rcp = Nothing

         m_Rs1.MoveNext
      Wend
      Set Rcp = Nothing
      
   End If
   
   QuerySumBill = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySumBill = False
End Function
Public Function AddEditSumBill(Ua As CSumBill, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PaymentType As Long
Dim OldPaymentType As Long
Dim I As Long
   RName = "AddEditSumBill"
   AddEditSumBill = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Rc As CBillDetail
   For Each Rc In Ua.RcpCnDnItems
      Rc.SUM_BILL_ID = Ua.SUM_BILL_ID
      If Rc.Flag = "A" Then
         Rc.ShowMode = SHOW_ADD
         Call Rc.AddEditData
      ElseIf Rc.Flag = "E" Then
         Rc.ShowMode = SHOW_EDIT
         Call Rc.AddEditData
      ElseIf Rc.Flag = "D" Then
         Call Rc.DeleteData
      End If
   Next Rc
   Set Rc = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditSumBill = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSumBill = False
End Function
Public Function AddEditSaleChart(Ua As CSaleChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSaleChart"
   AddEditSaleChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditSaleChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSaleChart = False
End Function
Public Function QuerySaleChart(Ua As CSaleChart, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QuerySaleChart"
   QuerySaleChart = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   QuerySaleChart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySaleChart = False
End Function
Public Function DeleteSaleChart(Ua As CSaleChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteSaleChart"
   DeleteSaleChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSaleChart = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSaleChart = True
End Function
Public Function CopyStockCode(Ug As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyStockCode"
   CopyStockCode = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryStockCode(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Ug.STOCK_CODE_ID = -1
   Ug.STOCK_NO = Ug.STOCK_NO & "*"
         
   Dim Det As CStockCodeChange
   For Each Det In Ug.StockCodeChange
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det
   
   Dim JiFt As CStockCodeChangeFt
   For Each JiFt In Ug.StockCodeChangeFt
      JiFt.Flag = "A"
      JiFt.ShowMode = SHOW_ADD
   Next JiFt
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditStockCode(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyStockCode = False
End Function

