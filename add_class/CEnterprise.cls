VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnterprise"
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

Public Flag As String
Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Public EnpNames As Collection
Public EnpAddresses As Collection
Public EnpPersons As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set EnpNames = New Collection
   Set EnpAddresses = New Collection
   Set EnpPersons = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   
   Set EnpNames = Nothing
   Set EnpAddresses = Nothing
   Set EnpPersons = Nothing
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
   
   SQL = "DELETE FROM ENTERPRISE_ADDRESS WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM ENTERPRISE_NAME WHERE " & PKField & " = " & Tf.GetValue
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
Dim SubLen As Long
Dim NewStr As String

Dim ENTERPRISE_ID As Long
Dim BUSINESS_TYPE As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim SHORT_NAME As String
Dim ENTERPRISE_NAME As String

   ENTERPRISE_ID = GetFieldValue("ENTERPRISE_ID")
   BUSINESS_TYPE = Val(GetFieldValue("BUSINESS_TYPE"))
   SHORT_NAME = GetFieldValue("SHORT_NAME")
   ENTERPRISE_NAME = GetFieldValue("ENTERPRISE_NAME")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT ENP.*, NM.LONG_NAME, NM.SHORT_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " ENP "
      FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (ENP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   
      WhereStr = ""
   End If
   
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (ENP.ENTERPRISE_ID = " & ENTERPRISE_ID & ") " & vbCrLf
   End If
   
   If BUSINESS_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (ENP.BUSINESS_TYPE = " & BUSINESS_TYPE & ") " & vbCrLf
   End If
   
   If Len(Trim(SHORT_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(SHORT_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(SHORT_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "')"
      End If
   End If
      
   If Len(Trim(ENTERPRISE_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(ENTERPRISE_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (LONG_NAME = '" & ChangeQuote(Trim(ENTERPRISE_NAME)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.SHORT_NAME "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.LONG_NAME "
      Else
         OrderByStr = "ORDER BY NM.SHORT_NAME "
      End If
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

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
     Call SetFieldValue("ENTERPRISE_ID", NVLI(Rs("ENTERPRISE_ID"), -1))
      Call SetFieldValue("BUSINESS_TYPE", NVLI(Rs("BUSINESS_TYPE"), -1))
      Call SetFieldValue("EMAIL", NVLS(Rs("EMAIL"), ""))
      Call SetFieldValue("WEBSITE", NVLS(Rs("WEBSITE"), ""))
      Call SetFieldValue("TAX_ID", NVLS(Rs("TAX_ID"), ""))
      Call SetFieldValue("SETUP_DATE", InternalDateToDate(NVLS(Rs("SETUP_DATE"), "")))
      Call SetFieldValue("POLICY", NVLS(Rs("POLICY"), ""))
      Call SetFieldValue("ENTERPRISE_TYPE", NVLI(Rs("ENTERPRISE_TYPE"), -1))
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      
      Call SetFieldValue("ENTERPRISE_NAME", NVLS(Rs("LONG_NAME"), ""))
      Call SetFieldValue("SHORT_NAME", NVLS(Rs("SHORT_NAME"), ""))
   End If
End Sub

Private Sub RegisterFields()
   TableName = "ENTERPRISE"
   SequenceName = "ENTERPRISE_SEQ"
   PKField = "ENTERPRISE_ID"
   
   Call AddField("ENTERPRISE_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BUSINESS_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("EMAIL", STRING_TYPE, "", DATA_CAT)
   Call AddField("WEBSITE", STRING_TYPE, "", DATA_CAT)
   Call AddField("TAX_ID", STRING_TYPE, "", DATA_CAT)
   Call AddField("SETUP_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("POLICY", STRING_TYPE, "", DATA_CAT)
   Call AddField("ENTERPRISE_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("BRANCH_CODE", STRING_TYPE, "", DATA_CAT)
   Call AddField("BRANCH_NAME", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ENTERPRISE_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("SHORT_NAME", STRING_TYPE, "", TEMP_CAT)

   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("QUERY_FLAG", INT_TYPE, "", TEMP_CAT)
End Sub

