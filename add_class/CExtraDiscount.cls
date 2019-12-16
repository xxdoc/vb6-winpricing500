VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtraDiscount"
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

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
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

Dim EXTRA_DISCOUNT_ID As Long
Dim BILLING_DOC_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   EXTRA_DISCOUNT_ID = Val(GetFieldValue("EXTRA_DISCOUNT_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*,MR.KEY_CODE,MR.KEY_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (NM.DISCOUNT_TYPE_ID = MR.KEY_ID) "
      
      WhereStr = ""
   End If
   
   If EXTRA_DISCOUNT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.EXTRA_DISCOUNT_ID = " & EXTRA_DISCOUNT_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.ITEM "
      Else
         OrderByStr = "ORDER BY MR.KEY_CODE "
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
      Call SetFieldValue("EXTRA_DISCOUNT_ID", NVLI(Rs("EXTRA_DISCOUNT_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("DISCOUNT_TYPE_ID", NVLI(Rs("DISCOUNT_TYPE_ID"), -1))
      Call SetFieldValue("EXTRA_DISCOUNT_VALUE", NVLD(Rs("EXTRA_DISCOUNT_VALUE"), 0))
      Call SetFieldValue("ITEM", NVLI(Rs("ITEM"), -1))
      Call SetFieldValue("UNIT_TYPE", NVLI(Rs("UNIT_TYPE"), -1))
      Call SetFieldValue("KEY_CODE", NVLS(Rs("KEY_CODE"), ""))
      Call SetFieldValue("KEY_NAME", NVLS(Rs("KEY_NAME"), ""))
      Call SetFieldValue("EXTRA_DISCOUNT_NO", NVLS(Rs("EXTRA_DISCOUNT_NO"), ""))
      Call SetFieldValue("EXTRA_DISCOUNT_DESC", NVLS(Rs("EXTRA_DISCOUNT_DESC"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "EXTRA_DISCOUNT"
   SequenceName = "EXTRA_DISCOUNT_SEQ"
   PKField = "EXTRA_DISCOUNT_ID"
   
   Call AddField("EXTRA_DISCOUNT_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BILLING_DOC_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("EXTRA_DISCOUNT_VALUE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("ITEM", INT_TYPE, -1, DATA_CAT)
   Call AddField("UNIT_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("DISCOUNT_TYPE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("EXTRA_DISCOUNT_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("EXTRA_DISCOUNT_DESC", STRING_TYPE, "", DATA_CAT)
   
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   
   Call AddField("KEY_CODE", STRING_TYPE, -1, TEMP_CAT)
   Call AddField("KEY_NAME", STRING_TYPE, -1, TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

