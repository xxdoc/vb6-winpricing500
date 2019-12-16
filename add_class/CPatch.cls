VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection
Private m_Conn2 As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set m_Conn2 = glbDatabaseMngr.DB2Connection
End Sub
Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim itemcount As Long
Dim Result As Boolean
   
   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, itemcount, TempRs, glbErrorLog)
   
   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function
Public Function IsPatch2(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim itemcount As Long
Dim Result As Boolean
   
   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   
   Call glbDatabaseMngr.GetRs2(SQL, "", False, itemcount, TempRs, glbErrorLog)

   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch2 = Result
   Exit Function
   
ErrorHandler:
   IsPatch2 = False
End Function
Public Function Patch_2008_01_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 43 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 66 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 68 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_10_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_10_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_10_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 39 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 40 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 67 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM APP_VIEW WHERE KEY_ID = 400 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 400 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_10_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_10_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_10_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 6 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM APAR_ADDRESS WHERE ADDRESS_ID IN (SELECT UG.ADDRESS_ID FROM ADDRESS UG WHERE  UG.COUNTRY_ID = 5) "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM ADDRESS WHERE COUNTRY_ID = 5 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM APP_VIEW WHERE KEY_ID = 5 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 5 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 58 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 59 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 60 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 61 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 62 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 63 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 64 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 65 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_10_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_12_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_12_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD STOCK_TYPE_SUB ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_12_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_12_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_12_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE STOCK_CODE ADD CONSTRAINT STOCK_CODE_TYPE_SUB_FK FOREIGN KEY (STOCK_TYPE_SUB) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_12_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_01_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DROP TABLE MASTER_REF_DETAIL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_20_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_09_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_09_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE LOT_ITEM LT SET LT.UNIT_TRAN_ID = (SELECT SC.UNIT_CHANGE_ID FROM STOCK_CODE SC WHERE LT.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE LT.INVENTORY_DOC_ID IN (SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD WHERE IVD.DOCUMENT_TYPE = 1000) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_09_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM DROP PREVIOUS_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP PREVIOUS_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP NEW_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP NEW_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD STD_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_19_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_19_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_19_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DOC_ITEM ADD STD_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_19_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE LOT_ITEM LT SET LT.STD_COST = (SELECT SC.COST_PER_AMOUNT FROM STOCK_CODE SC WHERE LT.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DOC_ITEM DI SET DI.STD_COST = (SELECT SC.COST_PER_AMOUNT FROM STOCK_CODE SC WHERE DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_20_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD TAGET_VALUE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_25_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD NO_X_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD AMOUNT_X_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD VALUE_X_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_26_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD AMOUNT_P_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD VALUE_P_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_27_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_27_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_27_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO DROP MASTER_FROMTO_VALUE1 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO DROP MASTER_FROMTO_VALUE2 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO DROP MASTER_FROMTO_VALUE3 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_FROMTO DROP MASTER_FROMTO_VALUE4 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_27_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_27_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_02_27_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO DROP MULTIPLE_VALUE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_27_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET SALE_BY = 22 WHERE DOCUMENT_TYPE = 3 AND APAR_MAS_ID = 102 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_02_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_02_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_02_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET SALE_BY = 16 WHERE DOCUMENT_TYPE = 3 AND APAR_MAS_ID = 102 AND  (SUBSTR(DOCUMENT_NO,1,3) = 'IVV') "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_02_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_02_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_02_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  MASTER_FROMTO_EX  "
   SQL = SQL & "( "
   SQL = SQL & "MASTER_FROMTO_EX_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_FROMTO_ID                    ID_TYPE , "
   SQL = SQL & "CUSTOMER_ID                    ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "EMP_ID                    ID_TYPE , "
   SQL = SQL & "PERCENT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE MASTER_FROMTO_EX  ADD CONSTRAINT MASTER_FROMTO_EX_PK  PRIMARY KEY (MASTER_FROMTO_EX_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2008_03_02_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_02_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_02_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE MASTER_FROMTO_EX ADD CONSTRAINT MASTER_FROMTO_EX_CUS_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES APAR_MAS "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MASTER_FROMTO_EX ADD CONSTRAINT MASTER_FROMTO_EX_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MASTER_FROMTO_EX ADD CONSTRAINT MASTER_FROMTO_EX_EMP_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2008_03_02_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE GENERATOR MASTER_FROMTO_EX_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_04_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB ADD PRODUCT_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_13_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_13_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_13_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE JOB ADD CONSTRAINT JOB_PRODUCT_ID_FK FOREIGN KEY (PRODUCT_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_13_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_13_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_13_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB ADD PRODUCT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_13_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_13_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_13_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE JOB J SET J.PRODUCT_ID =  (SELECT MAX(JI.PART_ITEM_ID) FROM JOB_ITEM JI WHERE (JI.TX_TYPE = 'I') AND (J.JOB_ID = JI.JOB_ID)) "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE JOB J SET J.PRODUCT_AMOUNT =  (SELECT SUM(LK.IMPORT_AMOUNT) FROM JOB_ITEM JI LEFT OUTER JOIN LOT_ITEM_LINK LK ON (JI.EXPORT_LOT_ITEM_ID = LK.EXPORT_LOT_ITEM_ID) WHERE (JI.TX_TYPE = 'I') AND (J.JOB_ID = JI.JOB_ID)) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_13_4_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB ADD PRODUCT_LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_14_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_14_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_14_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE JOB ADD CONSTRAINT JOB_LOCATION_ID_FK FOREIGN KEY (PRODUCT_LOCATION_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_14_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_14_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_14_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE JOB J SET J.PRODUCT_LOCATION_ID =  (SELECT MAX(JI.LOCATION_ID) FROM JOB_ITEM JI WHERE (JI.TX_TYPE = 'I') AND (J.JOB_ID = JI.JOB_ID)) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_14_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_14_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_03_14_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB_ITEM DROP LOT_ITEM_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_14_4_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB_ITEM ADD ORDER_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_18_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID = 31159 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID = 31159 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_27_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE STOCK_CODE SET LOT_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_30_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_31_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_31_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER LOT_ITEM_LINK_BI FOR LOT_ITEM_LINK"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.LOT_ITEM_LINK_ID  = GEN_ID (LOT_ITEM_LINK_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   SQL = "                 CREATE TRIGGER DOC_ITEM_LINK_BI FOR DOC_ITEM_LINK"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.DOC_ITEM_LINK_ID  = GEN_ID (DOC_ITEM_LINK_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_31_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_03_31_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_03_31_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD EXCEPTION_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE STOCK_CODE SET EXCEPTION_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_03_31_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DRIVER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD TRANSPORTOR_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CAR_LICENSE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_04_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_04_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_04_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BD_DRIVER_ID_FK FOREIGN KEY (DRIVER_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BD_TRANSPORTOR_ID_FK FOREIGN KEY (TRANSPORTOR_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BD_CAR_LICENSE_ID_FK FOREIGN KEY (CAR_LICENSE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_04_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_07_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_07_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  STOCK_CODE_CHANGE  "
   SQL = SQL & "( "
   SQL = SQL & "STOCK_CODE_CHANGE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "STOCK_CODE_ID                    ID_TYPE , "
   SQL = SQL & "UNIT_CHANGE_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE STOCK_CODE_CHANGE  ADD CONSTRAINT STOCK_CODE_CHANGE_ID_PK  PRIMARY KEY (STOCK_CODE_CHANGE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR STOCK_CODE_CHANGE_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2008_04_07_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_07_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_04_07_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE STOCK_CODE_CHANGE ADD CONSTRAINT STOCK_CODE_CHANGE_ID_FK FOREIGN KEY (STOCK_CODE_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_07_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_07_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_04_07_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER STOCK_CODE_CHANGE_BI FOR STOCK_CODE_CHANGE"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.STOCK_CODE_CHANGE_ID  = GEN_ID (STOCK_CODE_CHANGE_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_07_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_08_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_08_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE_CHANGE ADD UNIT_CHANGE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_08_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_08_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_08_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE STOCK_CODE_CHANGE ADD CONSTRAINT STOCK_CODE_CHANGE_ID_UNIT_FK FOREIGN KEY (UNIT_CHANGE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_08_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_08_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_08_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " DROP TABLE PART_CHILD "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_08_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_08_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_08_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " DROP GENERATOR PART_CHILD_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_08_4_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DROP TABLE APP_VIEW "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_10_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_10_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_10_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 33 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   
   SQL = "UPDATE STOCK_CODE SET UNIT_ID = 240 WHERE UNIT_ID = 45 "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE STOCK_CODE SET UNIT_CHANGE_ID  = 240 WHERE UNIT_CHANGE_ID = 45 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 45 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 32 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 34 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "UPDATE STOCK_CODE SET UNIT_ID = 117 WHERE UNIT_ID = 31 "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE STOCK_CODE SET UNIT_CHANGE_ID  = 117 WHERE UNIT_CHANGE_ID = 31 "
   m_Conn.Execute (SQL)
   
    SQL = "UPDATE DOC_ITEM SET UNIT_TRAN_ID  = 117 WHERE UNIT_TRAN_ID = 31 "
   m_Conn.Execute (SQL)
   
    SQL = "UPDATE LOT_ITEM SET UNIT_TRAN_ID  = 117 WHERE UNIT_TRAN_ID = 31 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 31 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 48 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 46 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   SQL = "UPDATE STOCK_CODE SET UNIT_ID = 798 WHERE UNIT_ID = 49 "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE STOCK_CODE SET UNIT_CHANGE_ID  = 798 WHERE UNIT_CHANGE_ID = 49 "
   m_Conn.Execute (SQL)
   
    SQL = "UPDATE DOC_ITEM SET UNIT_TRAN_ID  = 798 WHERE UNIT_TRAN_ID = 49 "
   m_Conn.Execute (SQL)
   
    SQL = "UPDATE LOT_ITEM SET UNIT_TRAN_ID  = 798 WHERE UNIT_TRAN_ID = 49 "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM MASTER_REF WHERE KEY_ID = 49 "
   m_Conn.Execute (SQL)
   '----------------------------------------------------------------------------------------------------------------------------------------
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_10_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_10_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_10_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE_CHANGE ADD CUSTOMER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_10_3_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_10_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_10_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE STOCK_CODE_CHANGE ADD CONSTRAINT STOCK_CODE_CHANGE_ID_CUS_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES APAR_MAS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_10_4_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_16_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  TRANSPORT_DETAIL  "
   SQL = SQL & "( "
   SQL = SQL & "TRANSPORT_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DRIVER_ID                    ID_TYPE , "
   SQL = SQL & "CAR_LICENSE_ID                    ID_TYPE , "
   SQL = SQL & "TRANSPORTOR_ID                    ID_TYPE , "
   SQL = SQL & "TRANSPORT_DETAIL_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE TRANSPORT_DETAIL  ADD CONSTRAINT TRANSPORT_DETAIL_ID_PK  PRIMARY KEY (TRANSPORT_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TRANSPORT_DETAIL_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_16_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_16_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_04_16_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE TRANSPORT_DETAIL ADD CONSTRAINT TRANSPORT_DETAIL_DRIVER_ID_FK FOREIGN KEY (DRIVER_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE TRANSPORT_DETAIL ADD CONSTRAINT TRANSPORT_DETAIL_TSPTOR_ID_FK FOREIGN KEY (TRANSPORTOR_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE TRANSPORT_DETAIL ADD CONSTRAINT TRANSPORT_DETAIL_CAR_ID_FK FOREIGN KEY (CAR_LICENSE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_16_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_16_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_04_16_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER TRANSPORT_DETAIL_BI FOR TRANSPORT_DETAIL"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.TRANSPORT_DETAIL_ID  = GEN_ID (TRANSPORT_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_17_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM CASH_TRAN WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM CHEQUE WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM RCPCNDN_ITEM WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM BILLING_SUBTRACT WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM BILLING_ADDITION WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_PACK = " & "18704"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_18_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_18_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_18_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC BD SET BD.PAID_AMOUNT "
   SQL = SQL & "= (SELECT SUM(RI.PAID_AMOUNT) FROM RCPCNDN_ITEM RI "
   SQL = SQL & "LEFT OUTER JOIN BILLING_DOC UG ON (RI.DOC_ID =UG.BILLING_DOC_ID) "
   SQL = SQL & "WHERE (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID ) AND (UG.DOCUMENT_TYPE = 3) ) "
   SQL = SQL & "WHERE BD.DOCUMENT_TYPE = 5 "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILLING_DOC BD SET BD.CREDIT_AMOUNT "
   SQL = SQL & "= (SELECT SUM(RI.PAID_AMOUNT) FROM RCPCNDN_ITEM RI "
   SQL = SQL & "LEFT OUTER JOIN BILLING_DOC UG ON (RI.DOC_ID =UG.BILLING_DOC_ID) "
   SQL = SQL & "WHERE (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID ) AND (UG.DOCUMENT_TYPE = 6) ) "
   SQL = SQL & "WHERE BD.DOCUMENT_TYPE = 5 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_18_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE NAME NM2 SET NM2.SHORT_NAME = NM2.LONG_NAME WHERE NAME_ID IN (SELECT APN.NAME_ID FROM APAR_NAME APN WHERE APN.APAR_MAS_ID IS NOT NULL) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_21_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET ENTERPRISE_ADDRESS_ID = 2 WHERE DOCUMENT_DATE >= '2008-04-01 00:00:00' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_25_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE NAME NM2 SET NM2.SHORT_NAME = 'โลตัส' WHERE NAME_ID IN (SELECT APN.NAME_ID FROM APAR_NAME APN LEFT OUTER JOIN APAR_MAS APM ON (APM.APAR_MAS_ID = APN.APAR_MAS_ID) WHERE (SUBSTR(APM.APAR_CODE,1,3) = '2-อ')  ) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_26_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE NAME NM2 SET NM2.SHORT_NAME = 'แม็คโคร' WHERE NAME_ID IN (SELECT APN.NAME_ID FROM APAR_NAME APN LEFT OUTER JOIN APAR_MAS APM ON (APM.APAR_MAS_ID = APN.APAR_MAS_ID) WHERE (SUBSTR(APM.APAR_CODE,1,3) = '2-ส')  ) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_26_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_04_26_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_04_26_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE NAME NM2 SET NM2.SHORT_NAME = 'คาร์ฟู' WHERE NAME_ID IN (SELECT APN.NAME_ID FROM APAR_NAME APN LEFT OUTER JOIN APAR_MAS APM ON (APM.APAR_MAS_ID = APN.APAR_MAS_ID) WHERE (SUBSTR(APM.APAR_CODE,1,3) = '2-ซ')  ) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_04_26_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_05_17_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_05_17_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BD_BILLING_DOC_PACK_FK FOREIGN KEY (BILLING_DOC_PACK) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_05_17_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_05_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_05_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET ENTERPRISE_ADDRESS_ID = 2 WHERE DOCUMENT_DATE >= '2008-04-01 00:00:00' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_05_30_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_03_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_03_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP ENTERPRISE_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP CODE1 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP YEAR_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP CODE2 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP MONTH_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP CODE3 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP MM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP UPDATE_MONTH_FLAG "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC DROP UPDATE_YEAR_FLAG "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_03_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE COMMISSION_CHART ADD ORDER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_04_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  KEY_ACCOUNT  "
   SQL = SQL & "( "
   SQL = SQL & "KEY_ACCOUNT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "SALE_ID                    ID_TYPE , "
   SQL = SQL & "KEY_ACCOUNT_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE KEY_ACCOUNT  ADD CONSTRAINT KEY_ACCOUNT_ID_PK  PRIMARY KEY (KEY_ACCOUNT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR KEY_ACCOUNT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_10_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_10_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_10_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  KEY_ACCOUNT_DETAIL  "
   SQL = SQL & "( "
   SQL = SQL & "KEY_ACCOUNT_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "KEY_ACCOUNT_ID                    ID_TYPE , "
   SQL = SQL & "CUSTOMER_ID                    ID_TYPE , "
   SQL = SQL & "BRANCH_ID                    ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE KEY_ACCOUNT_DETAIL  ADD CONSTRAINT KEY_ACCOUNT_DETAIL_ID_PK  PRIMARY KEY (KEY_ACCOUNT_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR KEY_ACCOUNT_DETAIL_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_10_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_10_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_10_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE KEY_ACCOUNT ADD CONSTRAINT KEY_ACCOUNT_SALE_FK FOREIGN KEY (SALE_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE KEY_ACCOUNT_DETAIL ADD CONSTRAINT KEY_ACCOUNT_DE_CUS_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES APAR_MAS "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE KEY_ACCOUNT_DETAIL ADD CONSTRAINT KEY_ACCOUNT_DE_BB_FK FOREIGN KEY (BRANCH_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_10_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_10_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_06_10_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER KEY_ACCOUNT_DETAIL_BA FOR KEY_ACCOUNT_DETAIL"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.KEY_ACCOUNT_DETAIL_ID  = GEN_ID (KEY_ACCOUNT_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_10_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Patch_2008_06_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  TAGET_JOB  "
   SQL = SQL & "( "
   SQL = SQL & "TAGET_JOB_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MONTH_ID                    ID_TYPE , "
   SQL = SQL & "YEAR_NO                    CODE_TYPE , "
   SQL = SQL & "INPUT_ID                    ID_TYPE , "
   SQL = SQL & "INPUT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "TAGET_JOB_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE TAGET_JOB ADD CONSTRAINT TAGET_JOB_ID_PK  PRIMARY KEY (TAGET_JOB_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TAGET_JOB_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_26_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  TAGET_JOB_DETAIL  "
   SQL = SQL & "( "
   SQL = SQL & "TAGET_JOB_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "TAGET_JOB_ID                    ID_TYPE , "
   SQL = SQL & "OUTPUT_TYPE_ID                    ID_TYPE , "
   SQL = SQL & "OUTPUT_AMOUNT                    ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE TAGET_JOB_DETAIL  ADD CONSTRAINT TAGET_JOB_DETAIL_ID_PK  PRIMARY KEY (TAGET_JOB_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TAGET_JOB_DETAIL_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_26_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_26_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_26_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE TAGET_JOB ADD CONSTRAINT TAGET_JOB_INPUT_FK FOREIGN KEY (INPUT_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE TAGET_JOB_DETAIL ADD CONSTRAINT TAGET_JOB_DETAIL_OUTPUT_FK FOREIGN KEY (OUTPUT_TYPE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_26_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_26_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_06_26_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER TAGET_JOB_DETAIL_BA FOR TAGET_JOB_DETAIL"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.TAGET_JOB_DETAIL_ID  = GEN_ID (TAGET_JOB_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_26_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_06_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_06_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE REPORT_CONFIG DROP FONT_NAME "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE REPORT_CONFIG ADD FONT_NAME NAME_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_06_27_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_07_05_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_07_05_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE TAGET_JOB ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_07_05_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_07_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_07_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE APAR_MAS  DROP RESPONSE_BY "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_07_14_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_07_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_07_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE TAGET_JOB DROP YYYYMM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TAGET_JOB_DETAIL ADD BATCH_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_07_21_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_08_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_08_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE TRANSPORT_DETAIL ADD TRANSPORT_PATH DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TRANSPORT_DETAIL ADD CAR_TYPE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TRANSPORT_DETAIL ADD COST_PER_ROUND MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_08_25_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_08_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_08_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD TRANSPORT_CYCLE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_08_26_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_08_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_08_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_CYCLE_FK FOREIGN KEY (TRANSPORT_CYCLE) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_08_26_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  STOCK_CODE_CHANGE_FT  "
   SQL = SQL & "( "
   SQL = SQL & "STOCK_CODE_CHANGE_FT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "STOCK_CODE_ID                    ID_TYPE , "
   SQL = SQL & "BOX                    MONEY_TYPE , "
   SQL = SQL & "TRAY                    MONEY_TYPE , "
   SQL = SQL & "PACK                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE STOCK_CODE_CHANGE_FT ADD CONSTRAINT STOCK_CODE_CHANGE_FT_PK  PRIMARY KEY (STOCK_CODE_CHANGE_FT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR STOCK_CODE_CHANGE_FT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_18_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_18_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_18_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE STOCK_CODE_CHANGE_FT ADD CONSTRAINT STOCK_CODE_CHANGE_FT_FK FOREIGN KEY (STOCK_CODE_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_18_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_18_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_09_18_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER STOCK_CODE_CHANGE_FT_BA FOR STOCK_CODE_CHANGE_FT"
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.STOCK_CODE_CHANGE_FT_ID  = GEN_ID (STOCK_CODE_CHANGE_FT_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_18_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
     
   VersionStr = "2008_09_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'Y' "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'N' WHERE SC.STOCK_TYPE = 444 " 'พัสดุหีบห่อ
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'N' WHERE SC.STOCK_TYPE = 881 " 'ถาดพลาสติก
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'N' WHERE SC.STOCK_TYPE = 794 " 'พัสดุและของใช้สินเปลือง
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'N' WHERE SC.STOCK_TYPE = 917 " 'จุลินทรีย์
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE STOCK_CODE SC SET CHK_STD_COST = 'N' WHERE SC.STOCK_TYPE = 820 " 'เครื่องเขียนและแบบพิมพ์
'   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_28_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_28_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_28_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  CAPITAL_MOVEMENT  "
   SQL = SQL & "( "
   SQL = SQL & "CAPITAL_MOVEMENT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_DATE                    DATE_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "LOCATION_ID                    ID_TYPE , "
   SQL = SQL & "CAPITAL_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_PK  PRIMARY KEY (CAPITAL_MOVEMENT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CAPITAL_MOVEMENT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_28_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_28_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_28_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES STOCK_CODE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_LC_FK FOREIGN KEY (LOCATION_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_28_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_28_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2008_09_28_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER CAPITAL_MOVEMENT_BA FOR CAPITAL_MOVEMENT "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.CAPITAL_MOVEMENT_ID  = GEN_ID (CAPITAL_MOVEMENT_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_28_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM ADD AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE LOT_ITEM SET AVG_PRICE = ACTUAL_UNIT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP ACTUAL_UNIT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP INCLUDE_UNIT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP TOTAL_ACTUAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP DISCOUNT_AMT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP ITEM_DESC "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP SETTING_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM DROP SETTING_PRICE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_29_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_29_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_29_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM ADD NEW_AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_29_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DOC_ITEM ADD TOTAL_INCLUDE_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DOC_ITEM ADD CAPITAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DOC_ITEM SET CAPITAL_AMOUNT = STD_COST "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DOC_ITEM DROP STD_COST "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_30_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_09_30_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_09_30_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   Dim Di As CDocItem
   Dim m_Rs As ADODB.Recordset
   Dim itemcount As Long
   Set Di = New CDocItem
   Set m_Rs = New ADODB.Recordset
   
   Di.DOC_ITEM_ID = -1
   Call Di.QueryData(1, m_Rs, itemcount)
   
   While Not m_Rs.EOF
      Call Di.PopulateFromRS(1, m_Rs)
      Call Di.SetFieldValue("TOTAL_INCLUDE_PRICE", Di.GetFieldValue("ITEM_AMOUNT") * Di.GetFieldValue("CAPITAL_AMOUNT"))
      If Di.GetFieldValue("TOTAL_INCLUDE_PRICE") > 0 Then
         Di.DOC_ITEM_ID = Di.GetFieldValue("DOC_ITEM_ID")
         Di.CAPITAL_AMOUNT = Di.GetFieldValue("CAPITAL_AMOUNT")
         Di.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("TOTAL_INCLUDE_PRICE")
         Di.ShowMode = SHOW_EDIT
         Call Di.UpdateCapitalSell
      End If
      m_Rs.MoveNext
   Wend
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_09_30_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_10_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_10_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RIGHT_ITEM ADD ADDEDIT CODE_TYPE NOT NULL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_10_02_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_10_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_10_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DOC_ITEM ADD LOT_ITEM_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_10_13_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_10_13_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_10_13_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE DOC_ITEM SET LOT_ITEM_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_10_13_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_10_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_10_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_FROMTO_DETAIL ADD GROUP_COM_ID ID_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = " ALTER TABLE MASTER_FROMTO_DETAIL ADD CONSTRAINT MASTER_FT_DT_GC_FK FOREIGN KEY (GROUP_COM_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_10_28_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_10_28_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2008_10_28_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD GROUP_COM_ID ID_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = " ALTER TABLE STOCK_CODE ADD CONSTRAINT STOCK_CODE_GC_FK FOREIGN KEY (GROUP_COM_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_10_28_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_05_16_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_05_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD BARCODE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_05_16_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_05_22_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_05_22_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD BARCODE_JOB_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_05_22_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_05_22_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_05_22_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE INVENTORY_DOC SET BARCODE_JOB_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_05_22_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_08_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_08_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  LOCK_DATE  "
   SQL = SQL & "( "
   SQL = SQL & "LOCK_DATE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "FROM_DATE                    DATE_TYPE , "
   SQL = SQL & "TO_DATE                    DATE_TYPE , "
   SQL = SQL & "LOCK_TYPE                    ID_TYPE , "          ' 1 IS GLOBLE
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOCK_DATE  ADD CONSTRAINT LOCK_DATE_PK  PRIMARY KEY (LOCK_DATE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR LOCK_DATE_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2012_08_20_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_08_20_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_08_20_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER LOCK_DATE_BA FOR LOCK_DATE "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.LOCK_DATE_ID  = GEN_ID (LOCK_DATE_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_08_20_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_08_20_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_08_20_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 INSERT INTO LOCK_DATE (FROM_DATE,TO_DATE,LOCK_TYPE,CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY) "
   SQL = SQL & "  VALUES "
   SQL = SQL & "  ('2012-08-01 00:00:00','2012-08-31 00:00:00',1,'2012-08-20 15:50:00',1,'2012-08-20 15:50:00',1) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_08_20_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_08_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_08_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD TEMP_SALE_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_08_27_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_08_27_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_08_27_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE INVENTORY_DOC SET TEMP_SALE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_08_27_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_09_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_09_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSIGNMENT_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_09_26_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_09_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_09_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET CONSIGNMENT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_09_26_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_09_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_09_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOCK_DATE ADD FROM_INVENTORY_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCK_DATE ADD TO_INVENTORY_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCK_DATE ADD FROM_INVOICE_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCK_DATE ADD TO_INVOICE_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCK_DATE ADD FROM_RECEIPT_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCK_DATE ADD TO_RECEIPT_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_09_28_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_09_28_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_09_28_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE LOCK_DATE SET FROM_INVENTORY_DATE = '2012-09-22 00:00:00',TO_INVENTORY_DATE = '2012-12-31 00:00:00',FROM_INVOICE_DATE = '2012-09-22 00:00:00',TO_INVOICE_DATE = '2012-12-31 00:00:00',FROM_RECEIPT_DATE = '2012-08-01 00:00:00',TO_RECEIPT_DATE = '2012-12-31 00:00:00' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_09_28_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_11_1_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_11_1_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE APAR_MAS ADD FLAG_EDIT FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_11_1_1_yong = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_11_1_2_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_11_1_2_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE APAR_MAS SET FLAG_EDIT = 'N'"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_11_1_2_yong = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2012_11_28_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "Patch_2012_11_28_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD OUTLAY_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_11_28_1_ging = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2012_11_28_2_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_11_28_2_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE STOCK_CODE SET OUTLAY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_11_28_2_ging = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2012_12_6_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE APAR_MAS ADD APAR_MAS_GROUP_ID ID_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_1_yong = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_2_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_6_2_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE APAR_MAS ADD CONSTRAINT FK_APAR_MAS FOREIGN KEY (APAR_MAS_GROUP_ID) REFERENCES MASTER_REF(KEY_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_2_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_3_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_6_3_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE APAR_MAS DROP CONSTRAINT FK_APAR_MAS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_3_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_4_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_6_4_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE APAR_MAS DROP APAR_MAS_GROUP_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_4_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_5_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_6_5_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE APAR_MAS ADD APAR_MAS_GROUP_CODE CODE_TYPE COLLATE UNICODE_FSS "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE APAR_MAS ADD APAR_MAS_GROUP_NAME NAME_TYPE COLLATE UNICODE_FSS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_5_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_6_6_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_6_6_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = " UPDATE APAR_MAS AP SET AP.APAR_MAS_GROUP_CODE = AP.APAR_CODE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_6_6_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2012_12_26_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_26_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD PRO_FROM_DATE DATE_TYPE"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD PRO_TO_DATE DATE_TYPE"
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE PACKAGE_DETAIL ADD PRO_ITEM_COST  MONEY_TYPE"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_26_1_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2012_12_28_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2012_12_28_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSIGNMENT_NO CODE_TYPE COLLATE UNICODE_FSS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2012_12_28_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_01_03_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_01_03_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = " ALTER TABLE BILLING_DOC ADD CONSIGNMENT_ID ID_TYPE "
   
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_01_03_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_01_03_2_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_01_03_2_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT FK_CONSIGNMENT_ID FOREIGN KEY (CONSIGNMENT_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) "
   
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_01_03_2_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_01_15_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
Dim ID As Long

   VersionStr = "2013_01_15_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   Call glbDatabaseMngr.GetSeqID("MASTER_REF_SEQ", ID, glbErrorLog)

   SQL = " INSERT INTO MASTER_REF (KEY_ID,KEY_CODE,KEY_NAME,MASTER_AREA,PARENT_ID,PARENT_EX_ID,PARENT_EX_ID2,PARENT_EX_ID3,PARENT_EX_ID4,PARENT_EX_ID5,MASTER_FLAG,SHORT_CODE,INDEX_LINK,CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY) "
   SQL = SQL & "  VALUES "
   SQL = SQL & "  (" & ID & ",'A-01',Null,999,Null,Null,Null,Null,Null,Null,'N',Null,0,'" & ServerDateTime & "',4,'" & ServerDateTime & "',4) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_01_15_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_03_07_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_03_07_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE PRINT_LABEL ADD SALE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_03_07_1_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_03_07_2_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_03_07_2_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = " ALTER TABLE PRINT_LABEL ADD CONSTRAINT FK_SALE_ID FOREIGN KEY (SALE_ID) REFERENCES EMPLOYEE(EMP_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_03_07_2_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_03_07_3_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_03_07_3_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE PRINT_LABEL PL SET PL.SALE_ID =  (SELECT MR.PARENT_EX_ID FROM MASTER_REF MR WHERE PL.BRANCH_ID = MR.KEY_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_03_07_3_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_03_12_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_03_12_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE EMPLOYEE ADD MAINSALE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_03_12_1_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_03_12_2_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_03_12_2_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE EMPLOYEE EM SET EM.MAINSALE_FLAG =  'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_03_12_2_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_04_18_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_04_18_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD JOINT_CODE CODE_TYPE COLLATE UNICODE_FSS"
   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)
   
   m_Conn.BeginTrans
   m_Conn2.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   m_Conn2.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_04_18_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_04_23_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_04_23_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE APAR_MAS ADD CANCEL_OUT_DOCUMENT FLAG_TYPE NOT NULL COLLATE UNICODE_FSS"
   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)
   
   SQL = "UPDATE APAR_MAS AP SET AP.CANCEL_OUT_DOCUMENT =  'N'"
   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)
   
   m_Conn.BeginTrans
   m_Conn2.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   m_Conn2.CommitTrans
   HasBegin = False
     
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_04_23_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_04_25_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_04_25_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EMPLOYEE ADD JOINT_CODE CODE_TYPE COLLATE UNICODE_FSS"
   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)
   
   m_Conn.BeginTrans
   m_Conn2.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   m_Conn2.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_04_25_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_05_14_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_05_14_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EMPLOYEE ADD NOT_SHOW_RETURN FLAG_TYPE"

   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)

   SQL = "UPDATE EMPLOYEE SET NOT_SHOW_RETURN = 'N'"

   m_Conn.Execute (SQL)
   m_Conn2.Execute (SQL)
   
   m_Conn.BeginTrans
   m_Conn2.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   m_Conn2.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_14_1_yong = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_10_2_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_2_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE APAR_MAS ADD CONSIGNMENT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = " SET GENERATOR PATCH_DELIVERY_SEQ TO 500 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_2_1_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_10_2_2_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_2_2_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE APAR_MAS AM SET AM.CONSIGNMENT_FLAG =  'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_2_2_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_03_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2013_10_03_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  SUM_BILL  "
   SQL = SQL & "( "
   SQL = SQL & "SUM_BILL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_DATE                    DATE_TYPE , "
   SQL = SQL & "SUM_BILL_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUM_BILL ADD CONSTRAINT SUM_BILL_PK  PRIMARY KEY (SUM_BILL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SUM_BILL_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_03_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_03_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_03_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER SUM_BILL_BA FOR SUM_BILL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.SUM_BILL_ID  = GEN_ID (SUM_BILL_SEQ, 1); "
   SQL = SQL & "  END; "
   
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_03_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_03_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2013_10_03_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  BILL_DETAIL  "
   SQL = SQL & "( "
   SQL = SQL & "BILL_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "SUM_BILL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "BILLING_DOC_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "SUM_BILL_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILL_DETAIL ADD CONSTRAINT BILL_DETAIL_PK  PRIMARY KEY (BILL_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BILL_DETAIL_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_03_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_10_03_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_03_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER BILL_DETAIL_BA FOR BILL_DETAIL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.BILL_DETAIL_ID  = GEN_ID (BILL_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_03_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_10_03_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_03_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILL_DETAIL ADD CONSTRAINT BILL_DETAIL_FK FOREIGN KEY (BILL_DETAIL_ID) REFERENCES SUM_BILL "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILL_DETAIL ADD CONSTRAINT BILL_DETAIL_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_03_5_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_DETAIL DROP CONSTRAINT BILL_DETAIL_FK"
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILL_DETAIL ADD CONSTRAINT BILL_DETAIL_FK FOREIGN KEY (SUM_BILL_ID) REFERENCES SUM_BILL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_14_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_14_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_14_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_DETAIL DROP SUM_BILL_DESC"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_14_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_10_14_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_10_14_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SUM_BILL ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUM_BILL ADD CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUM_BILL ADD DEBIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_10_14_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_11_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2013_11_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  SALE_CHART "
   SQL = SQL & "( "
   SQL = SQL & "SALE_CHART_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "ORDER_ID             ID_TYPE , "
   SQL = SQL & "MASTER_FROMTO_ID             ID_TYPE , "
   SQL = SQL & "PARENT_ID             ID_TYPE , "
   SQL = SQL & "EMP_ID             ID_TYPE , "
   SQL = SQL & "EMP_PERCENT                    MONEY_TYPE , "
   SQL = SQL & "EMP_PARENT_PERCENT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_CHART ADD CONSTRAINT SALE_CHART_PK  PRIMARY KEY (SALE_CHART_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SALE_CHART_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_11_20_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_11_20_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_11_20_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER SALE_CHART_BA FOR SALE_CHART "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.SALE_CHART_ID  = GEN_ID (SALE_CHART_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_11_20_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_11_20_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2013_11_20_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE SALE_CHART ADD CONSTRAINT SALE_CHART_MASTER_FK FOREIGN KEY (MASTER_FROMTO_ID) REFERENCES MASTER_FROMTO "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_CHART ADD CONSTRAINT SALE_CHART_EMP_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_CHART ADD CONSTRAINT SALE_CHART_PARENT_FK FOREIGN KEY (PARENT_ID) REFERENCES SALE_CHART "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_11_20_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_12_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_12_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE NAME ADD BILL_NAME LNAME_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_12_27_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_12_27_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2013_12_27_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE NAME NM SET NM.BILL_NAME = NM.LONG_NAME WHERE NM.NAME_ID IN (SELECT APN.NAME_ID FROM APAR_NAME APN) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_12_27_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_03_03_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_03_03_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD TRIP_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CUSTOMER_BRANCH ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD DRIVER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CAR_LICENSE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD TRANSPORTOR_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_03_03_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_03_03_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_03_03_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE INVENTORY_DOC SET TRIP_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_03_03_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_03_03_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_03_03_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = " ALTER TABLE INVENTORY_DOC ADD CONSTRAINT IVTRD_CUSTOMER_BRANCH_FK FOREIGN KEY (CUSTOMER_BRANCH) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE INVENTORY_DOC ADD CONSTRAINT IVTRD_DRIVER_ID_FK FOREIGN KEY (DRIVER_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE INVENTORY_DOC ADD CONSTRAINT IVTRD_CAR_LICENSE_ID_FK FOREIGN KEY (CAR_LICENSE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE INVENTORY_DOC ADD CONSTRAINT IVTRD_TRANSPORTOR_ID_FK FOREIGN KEY (TRANSPORTOR_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_03_03_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_03_03_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_03_03_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = " ALTER TABLE INVENTORY_DOC ALTER COLUMN TRIP_FLAG SET DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_03_03_4_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_04_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_04_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE APAR_MAS ADD BASKET_FIX_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_04_21_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EMPLOYEE ADD DEALER_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_14_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_15_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_15_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_REF ADD DEALER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_15_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_15_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_15_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE MASTER_REF ADD CONSTRAINT MASTER_DEALER_FK FOREIGN KEY (DEALER_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_15_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_16_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DEALER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_16_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_16_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_16_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_DEALER_FK FOREIGN KEY (DEALER_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_16_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_05_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  EMPLOYEE_DEALER "
   SQL = SQL & "( "
   SQL = SQL & "EMPLOYEE_DEALER_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "EMP_ID             ID_TYPE , "
   SQL = SQL & "YYYYMM             CODE_TYPE , "
   SQL = SQL & "DEALER_TYPE             ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EMPLOYEE_DEALER ADD CONSTRAINT EMPLOYEE_DEALER_PK  PRIMARY KEY (EMPLOYEE_DEALER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EMPLOYEE_DEALER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_21_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_21_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2014_05_21_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER EMPLOYEE_DEALER_BA FOR EMPLOYEE_DEALER "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.EMPLOYEE_DEALER_ID  = GEN_ID (EMPLOYEE_DEALER_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_21_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_21_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_05_21_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE EMPLOYEE_DEALER ADD CONSTRAINT EMPLOYEE_DEALER_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_21_3_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_05_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE STOCK_CODE ADD REBATE_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_26_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_05_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE STOCK_CODE SET REBATE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_26_2_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_05_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_05_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE EMPLOYEE SET DEALER_TYPE = 10 WHERE SUBSTRING(EMP_CODE FROM 1 FOR 7) = '06-001-' "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_05_30_1_jill = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function



Public Function Patch_2014_09_22_1_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_09_22_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD HOLD_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_09_22_1_pui = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_09_22_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2014_09_22_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE PACKAGE_DETAIL SET HOLD_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_09_22_2_pui = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_01_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2017_01_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD DISCOUNT_TYPE CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_DOC ADD DISCOUNT_VALUE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_01_05_1_lek = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_02_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery

   VersionStr = "2017_02_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE USER_ACCOUNT ADD USER_ACCESS DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_02_03_1_lek = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_22_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2017_11_22_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_BY ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_11_22_1_jill = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_22_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2017_11_22_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_PAYMENT_FK FOREIGN KEY (PAYMENT_BY) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_11_22_2_jill = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_24_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2017_11_24_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD GET_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_11_24_1_jill = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_10_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2019_07_10_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_REF ADD CASH_DELIVERY_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  MASTER_REF SET CASH_DELIVERY_FLAG ='N'"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_07_10_1_lek = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2019_07_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  EXTRA_DISCOUNT  "
   SQL = SQL & "( "
   SQL = SQL & "EXTRA_DISCOUNT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "BILLING_DOC_ID                    ID_TYPE , "
   SQL = SQL & "EXTRA_DISCOUNT_NO                    CODE_TYPE , "
   SQL = SQL & "EXTRA_DISCOUNT_DESC                    DESC_TYPE , "
   SQL = SQL & "EXTRA_DISCOUNT_VALUE                   MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EXTRA_DISCOUNT  ADD CONSTRAINT EXTRA_DISCOUNT_ID_PK  PRIMARY KEY (EXTRA_DISCOUNT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD CONSTRAINT BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR EXTRA_DISCOUNT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_07_12_1_lek = True
   Exit Function

ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_18_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2019_07_18_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD ITEM ID_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE EXTRA_DISCOUNT ADD UNIT_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_07_18_1_lek = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_19_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2019_07_19_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT DROP EXTRA_DISCOUNT_NO "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT DROP EXTRA_DISCOUNT_DESC  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD DISCOUNT_TYPE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD CONSTRAINT DISCOUNT_TYPE_ID_FK FOREIGN KEY (DISCOUNT_TYPE_ID) REFERENCES MASTER_REF(KEY_ID) "
   m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_07_19_1_lek = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_23_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim PD As CPatchDelivery
Set PD = New CPatchDelivery
   
   VersionStr = "2019_07_23_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   SQL = "ALTER TABLE BILLING_DOC ADD IS_CALCULATE FLAG_TYPE "
'   m_Conn.Execute (SQL)
   
'   SQL = "UPDATE  BILLING_DOC SET IS_CALCULATE ='N'"
'   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD EXTRA_DISCOUNT_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRA_DISCOUNT ADD EXTRA_DISCOUNT_DESC DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
   Call PD.AddEditData(SHOW_ADD)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_07_23_1_lek = True
   Exit Function
   
ErrorHandler:

      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_08_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_08_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE VERSION_PROGRAM "
   SQL = SQL & "( "
   SQL = SQL & "VERSION_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "VERSION_NAME          CODE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY  ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE VERSION_PROGRAM ADD CONSTRAINT VERSION_ID_PK PRIMARY KEY (VERSION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO VERSION_PROGRAM (VERSION_ID,VERSION_NAME,CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY) "
   SQL = SQL & "  VALUES "
   SQL = SQL & "  (1,'0.0.0.0','2013-08-05 15:50:00',1,'2013-08-05 15:50:00',1) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_08_05_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
'Public Function Patch_2017_01_05_1_lek() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2017_01_05_1_lek"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE BILLING_DOC ADD DISCOUNT_TYPE CODE_TYPE  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE BILLING_DOC ADD DISCOUNT_VALUE MONEY_TYPE  "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2017_01_05_1_lek = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
