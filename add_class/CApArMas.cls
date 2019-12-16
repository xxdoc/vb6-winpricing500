VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAPARMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public Flag As String
Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Public CstNames As Collection
Public CstAddresses As Collection
Public CstPersons As Collection
Public CstBranchs As Collection

Public APAR_MAS_ID As Long
Public ADDRESS_ID As Long
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public APAR_CODE As String
Public APAR_NAME As String
Public APAR_IND As Long
Public CREDIT  As Double
Public APAR_GRADE As Long
Public APAR_TYPE As Long
Public APAR_GROUP As Long
Public TAX_ID  As String
Public BIRTH_DATE As Date
Public PHONE As String
Public Address As String
Public BUSINESS_TYPE  As Long
Public BUSINESS_DESC As String
Public EMAIL  As String
Public WEBSITE As String
Public PASSWD As String
Public NORMAL_DISCOUNT  As Double
Public PACKAGE_ID  As Long
Public LABEL_FLAG As String
Public CONSIGNMENT_FLAG As String
Public ADD_BRANCH_NAME As String
Public BASKET_FIX_AMOUNT As Double

Public BILL_NAME As String

Public APAR_SHORT_NAME As String
Public APAR_TYPE_NAME As String
Public APAR_GROUP_NAME  As String
Public APAR_GRADE_NAME  As String

Public BRANCH_CODE As String
Public BRANCH_NAME As String
Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LNAME  As String
Public POSITION_NAME  As String

Public APAR_MAS_GROUP_CODE  As String
Public APAR_MAS_GROUP_NAME  As String
'------------------------------------------------------------------------------------------------------------------------------------------------
Public KEY_ID As Long
Public KEY_LOOKUP As String
Public FROM_APAR_CODE As String
Public TO_APAR_CODE As String

Public FLAG_EDIT As String
Public CANCEL_OUT_DOCUMENT As String
'------------------------------------------------------------------------------------------------------------------------------------------------
Public HOME As String
Public SOI As String
Public MOO As String
Public VILLAGE As String
Public ROAD As String
Public DISTRICT As String
Public AMPHUR As String
Public PROVINCE As String
Public ZIPCODE As String
Public BANGKOK_FLAG As String
'------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set CstNames = New Collection
   Set CstAddresses = New Collection
   Set CstPersons = New Collection
   Set CstBranchs = New Collection
   
   LABEL_FLAG = "N"
   ADD_BRANCH_NAME = "N"
End Sub

Private Sub Class_Terminate()
   Set CstNames = Nothing
   Set CstAddresses = Nothing
   Set CstPersons = Nothing
   Set CstBranchs = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM MASTER_REF WHERE PARENT_EX_ID2  = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM APAR_ADDRESS WHERE APAR_MAS_ID = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM APAR_NAME WHERE APAR_MAS_ID = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM APAR_MAS WHERE APAR_MAS_ID = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long, CountFlag As Boolean)
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

   
   If Ind = 1 Then
      SelectStr1 = "SELECT APM.*, NM.LONG_NAME,NM.BILL_NAME, NM.SHORT_NAME, MR1.KEY_NAME APAR_GRADE_NAME, MR2.KEY_NAME APAR_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (APM.APAR_GRADE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT APM.APAR_MAS_ID,APM.APAR_CODE, NM.LONG_NAME,NM.BILL_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then     ' CReportMain001
       SelectStr1 = "SELECT MR.*,"
       SelectStr1 = SelectStr1 & " APM.APAR_CODE,"
       SelectStr1 = SelectStr1 & " APM.BIRTH_DATE,"
       SelectStr1 = SelectStr1 & " APA.*,"
       SelectStr1 = SelectStr1 & "EM.EMP_CODE,"
       SelectStr1 = SelectStr1 & "NM4.LONG_NAME CUS_LONG_NAME,"
       SelectStr1 = SelectStr1 & "NM4.BILL_NAME CUS_BILL_NAME,"
       SelectStr1 = SelectStr1 & "NM4.SHORT_NAME CUS_SHORT_NAME,"
       SelectStr1 = SelectStr1 & "NM.LONG_NAME EMP_LONG_NAME,"
       SelectStr1 = SelectStr1 & "NM.SHORT_NAME EMP_SHORT_NAME,"
       SelectStr1 = SelectStr1 & "NM.BILL_NAME EMP_BILL_NAME,"
       SelectStr1 = SelectStr1 & "EM2.EMP_CODE DEALER_CODE,"
       SelectStr1 = SelectStr1 & "NM2.LONG_NAME DEALER_NAME,"
       SelectStr1 = SelectStr1 & "NM2.LAST_NAME DEALER_LNAME,"
      SelectStr1 = SelectStr1 & "APT.KEY_NAME APAR_TYPE_NAME,"
      SelectStr1 = SelectStr1 & "APG.KEY_NAME APAR_GROUP_NAME,"
      SelectStr1 = SelectStr1 & "ADR.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = " FROM MASTER_REF MR "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN APAR_ADDRESS APA ON (APM.APAR_MAS_ID = APA.APAR_MAS_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APT ON (APM.APAR_TYPE = APT.KEY_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APG ON (APG.KEY_ID = APT.PARENT_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID)"
      FromStr = FromStr & " LEFT OUTER JOIN NAME NM4 ON (EN.NAME_ID = NM4.NAME_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR.PARENT_EX_ID = EM.EMP_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EM.EMP_ID = EN2.EMP_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN2.NAME_ID = NM.NAME_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM2 ON (MR.DEALER_ID = EM2.EMP_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN3 ON (EM2.EMP_ID = EN3.EMP_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN3.NAME_ID = NM2.NAME_ID)"
      
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADR ON (ADR.ADDRESS_ID = APA.ADDRESS_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (MR.MASTER_AREA = 21) " & vbCrLf
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT APM.APAR_CODE,NM.BILL_NAME "
      SelectStr1 = SelectStr1 & ", NM.LONG_NAME,NM.SHORT_NAME  "
      SelectStr1 = SelectStr1 & ",APT.KEY_NAME  APAR_TYPE_NAME  "
      SelectStr1 = SelectStr1 & ",APG.KEY_NAME  APAR_GROUP_NAME  "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE ,MR.KEY_NAME BRANCH_NAME "    '™◊ËÕ “¢“"
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      SelectStr1 = SelectStr1 & ",NM2.LONG_NAME EMP_NAME ,  NM2.LAST_NAME EMP_LNAME "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME POSITION_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APT ON (APM.APAR_TYPE = APT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APG ON (APG.KEY_ID = APT.PARENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EMN ON (EM.EMP_ID = EMN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EMN.NAME_ID = NM2.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EM.CURRENT_POSITION = MR1.KEY_ID) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT APM.APAR_CODE , APM.CONSIGNMENT_FLAG "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      
      WhereStr = ""
   ElseIf Ind = 6 Then        'LoadApArAddress
      SelectStr1 = "SELECT APM.APAR_MAS_ID,"
      SelectStr1 = SelectStr1 & "ADR.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = " FROM APAR_MAS APM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_ADDRESS APA ON (APM.APAR_MAS_ID = APA.APAR_MAS_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADR ON (ADR.ADDRESS_ID = APA.ADDRESS_ID) "
   End If
   
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_MAS_ID = " & APAR_MAS_ID & ") " & vbCrLf
   End If
   
   If APAR_IND > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_IND = " & APAR_IND & ") " & vbCrLf
   End If
   
   If Len(Trim(APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(EMP_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EM.EMP_CODE ,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (EM.EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
      End If
   End If
   
   
   
   If Len(Trim(FROM_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE >= '" & ChangeQuote(Trim(FROM_APAR_CODE)) & "')"
   End If
   
   If Len(Trim(TO_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE <= '" & ChangeQuote(Trim(TO_APAR_CODE)) & "')"
   End If
   
'   If Len(Trim(APAR_NAME)) > 0 Then
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      If WildCard(APAR_NAME, SubLen, NewStr) Then
'         WhereStr = WhereStr & TempStr & " (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'      Else
'         WhereStr = WhereStr & TempStr & " (LONG_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
'      End If
'   End If
   
   If Trim(APAR_NAME) <> "" Then
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "LONG_NAME", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "LONG_NAME", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
         End If
      End If
   End If
   
   If APAR_GRADE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_GRADE = " & APAR_GRADE & ") " & vbCrLf
   End If
   
   If APAR_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_TYPE = " & APAR_TYPE & ") " & vbCrLf
   End If
   
   If APAR_GROUP > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APT.PARENT_ID = " & APAR_GROUP & ") " & vbCrLf
   End If
   
    If CONSIGNMENT_FLAG <> "" Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.CONSIGNMENT_FLAG = '" & CONSIGNMENT_FLAG & "') " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.LONG_NAME "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE "
      End If
   ElseIf Ind = 2 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.LONG_NAME "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE "
      End If
   ElseIf Ind = 3 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,NM.LONG_NAME "
      Else
'         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
         OrderByStr = "ORDER BY APM.APAR_CODE "
         ORDER_TYPE = 1
      End If
   ElseIf Ind = 4 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
          If ORDER_TYPE = 1 Then
            OrderByStr = OrderByStr & " ASC "
         ElseIf ORDER_TYPE = 2 Then
            OrderByStr = OrderByStr & " DESC "
         Else
            OrderByStr = OrderByStr & " ASC "
         End If
          OrderByStr = OrderByStr & " ,MR.KEY_CODE ASC, EM.EMP_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,NM.LONG_NAME "
         If ORDER_TYPE = 1 Then
            OrderByStr = OrderByStr & " ASC "
         ElseIf ORDER_TYPE = 2 Then
            OrderByStr = OrderByStr & " DESC "
         Else
            OrderByStr = OrderByStr & " ASC "
         End If
          OrderByStr = OrderByStr & " ,MR.KEY_CODE ASC, EM.EMP_CODE "
       Else
            OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
            If ORDER_TYPE = 1 Then
               OrderByStr = OrderByStr & " ASC "
            ElseIf ORDER_TYPE = 2 Then
               OrderByStr = OrderByStr & " DESC "
            Else
               OrderByStr = OrderByStr & " ASC "
            End If
             OrderByStr = OrderByStr & " ,MR.KEY_CODE ASC, EM.EMP_CODE "
      End If
      ORDER_TYPE = 1
  ElseIf Ind = 5 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
         ORDER_TYPE = 1
   End If
   
   If Len(OrderByStr) > 0 Then
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & " ASC "
      Else
         OrderByStr = OrderByStr & " DESC "
      End If
   End If
      
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
'   Debug.Print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountFlag, itemcount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_GRADE = NVLI(Rs("APAR_GRADE"), -1)
      APAR_TYPE = NVLI(Rs("APAR_TYPE"), -1)
      CREDIT = NVLD(Rs("CREDIT"), 0)
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
      BUSINESS_TYPE = NVLI(Rs("BUSINESS_TYPE"), -1)
      BUSINESS_DESC = NVLS(Rs("BUSINESS_DESC"), "")
      EMAIL = NVLS(Rs("EMAIL"), "")
      WEBSITE = NVLS(Rs("WEBSITE"), "")
      PASSWD = NVLS(Rs("PASSWD"), "")
      NORMAL_DISCOUNT = NVLD(Rs("NORMAL_DISCOUNT"), 0)
      APAR_IND = NVLD(Rs("APAR_IND"), -1)
      PACKAGE_ID = NVLD(Rs("PACKAGE_ID"), -1)
      
      LABEL_FLAG = NVLS(Rs("LABEL_FLAG"), "N")
      ADD_BRANCH_NAME = NVLS(Rs("ADD_BRANCH_NAME"), "N")
      
      KEY_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("LONG_NAME"), "")
      APAR_GRADE_NAME = NVLS(Rs("APAR_GRADE_NAME"), "")
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      BILL_NAME = NVLS(Rs("BILL_NAME"), "")
      
      BASKET_FIX_AMOUNT = NVLD(Rs("BASKET_FIX_AMOUNT"), 0)
      
      APAR_SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
      
      FLAG_EDIT = NVLS(Rs("FLAG_EDIT"), "")
      CANCEL_OUT_DOCUMENT = NVLS(Rs("CANCEL_OUT_DOCUMENT"), "")
      APAR_MAS_GROUP_CODE = NVLS(Rs("APAR_MAS_GROUP_CODE"), "")
      APAR_MAS_GROUP_NAME = NVLS(Rs("APAR_MAS_GROUP_NAME"), "")
      CONSIGNMENT_FLAG = NVLS(Rs("CONSIGNMENT_FLAG"), "")
   ElseIf Ind = 2 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("LONG_NAME"), "")
      BILL_NAME = NVLS(Rs("BILL_NAME"), "")
   ElseIf Ind = 3 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_MAS_ID = NVLI(Rs("APAR_CODE"), -1)
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      APAR_NAME = NVLS(Rs("CUS_LONG_NAME"), "")
      APAR_SHORT_NAME = NVLS(Rs("CUS_SHORT_NAME"), "")
      BILL_NAME = NVLS(Rs("CUS_BILL_NAME"), "")
      BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
      
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_LONG_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_SHORT_NAME"), "")
      
      
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "")
      
      PHONE = NVLS(Rs("PHONE1"), "") & " " & NVLS(Rs("PHONE2"), "")
      
   ElseIf Ind = 4 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("LONG_NAME"), "")
      APAR_SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
      BILL_NAME = NVLS(Rs("BILL_NAME"), "")
      
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
            
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      POSITION_NAME = NVLS(Rs("POSITION_NAME"), "")
   ElseIf Ind = 5 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      CONSIGNMENT_FLAG = NVLS(Rs("CONSIGNMENT_FLAG"), "")
   ElseIf Ind = 6 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "")
      
      PHONE = NVLS(Rs("PHONE1"), "") & " " & NVLS(Rs("PHONE2"), "")
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = APAR_GRADE
   If APAR_GRADE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = APAR_TYPE
   If APAR_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
      
   TempID3 = BUSINESS_TYPE
   If BUSINESS_TYPE <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PACKAGE_ID
   If PACKAGE_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("APAR_MAS_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO APAR_MAS " & _
                     "( " & _
                     "   APAR_MAS_ID, " & _
                     "   APAR_CODE, " & _
                     "   APAR_GRADE, " & _
                     "   APAR_TYPE, " & _
                     "   CREDIT, " & _
                     "   TAX_ID, " & _
                     "   BIRTH_DATE, " & _
                     "   EMAIL, " & _
                     "   WEBSITE, " & _
                     "   PASSWD, " & _
                     "   BUSINESS_TYPE, " & _
                     "   BUSINESS_DESC, " & _
                     "   NORMAL_DISCOUNT, " & _
                     "   APAR_IND, " & _
                     "   PACKAGE_ID, " & _
                     "   LABEL_FLAG, " & _
                     "   ADD_BRANCH_NAME, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, "
         SQL1 = SQL1 & _
                    "    FLAG_EDIT, " & "    CONSIGNMENT_FLAG, " & "    CANCEL_OUT_DOCUMENT, " & "    APAR_MAS_GROUP_CODE, " & "    APAR_MAS_GROUP_NAME " & _
                    "    ,BASKET_FIX_AMOUNT " & _
                     ") "    '
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        ID & "," & _
                        "'" & ChangeQuote(Trim(APAR_CODE)) & "'," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        CREDIT & "," & _
                        "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                        "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                        "'" & ChangeQuote(Trim(PASSWD)) & "'," & _
                        TempID3 & "," & _
                        "'" & ChangeQuote(Trim(BUSINESS_DESC)) & "'," & _
                        NORMAL_DISCOUNT & "," & _
                        APAR_IND & "," & _
                        TempID4 & "," & _
                        "'" & ChangeQuote(Trim(LABEL_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(ADD_BRANCH_NAME)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & "'" & ChangeQuote(Trim(FLAG_EDIT)) & "'," & "'" & ChangeQuote(Trim(CONSIGNMENT_FLAG)) & "'," & "'" & ChangeQuote(Trim(CANCEL_OUT_DOCUMENT)) & "'," & "'" & APAR_MAS_GROUP_CODE & "'," & "'" & APAR_MAS_GROUP_NAME & "',"
            SQL1 = SQL1 & _
                        BASKET_FIX_AMOUNT & "" & _
                     ") "    '
         APAR_MAS_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE APAR_MAS SET " & _
                     "    APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "', " & _
                     "    APAR_GRADE = " & TempID1 & ", " & _
                     "    APAR_TYPE = " & TempID2 & ", " & _
                     "    CREDIT = " & CREDIT & ", " & _
                     "    TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   BIRTH_DATE= '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "    EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "    WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "    PASSWD = '" & ChangeQuote(Trim(PASSWD)) & "', " & _
                     "    BUSINESS_TYPE = " & TempID3 & ", " & _
                     "    BUSINESS_DESC = '" & ChangeQuote(Trim(BUSINESS_DESC)) & "', " & _
                     "    NORMAL_DISCOUNT = " & NORMAL_DISCOUNT & ", " & _
                     "    APAR_IND = " & APAR_IND & ", " & _
                     "    PACKAGE_ID = " & TempID4 & ", " & _
                     "    LABEL_FLAG = '" & ChangeQuote(Trim(LABEL_FLAG)) & "', " & _
                     "    ADD_BRANCH_NAME = '" & ChangeQuote(Trim(ADD_BRANCH_NAME)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & ", " & _
                      "    FLAG_EDIT = '" & ChangeQuote(Trim(FLAG_EDIT)) & "', " & _
                      "    CONSIGNMENT_FLAG = '" & ChangeQuote(Trim(CONSIGNMENT_FLAG)) & "', " & _
                      "    CANCEL_OUT_DOCUMENT = '" & ChangeQuote(Trim(CANCEL_OUT_DOCUMENT)) & "', " & _
                      "    APAR_MAS_GROUP_CODE = '" & APAR_MAS_GROUP_CODE & "', " & _
                      "    APAR_MAS_GROUP_NAME = '" & APAR_MAS_GROUP_NAME & "',"
         SQL1 = SQL1 & _
                     "    BASKET_FIX_AMOUNT = " & BASKET_FIX_AMOUNT & " " & _
                     "WHERE APAR_MAS_ID = " & APAR_MAS_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function

Private Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    
    GenerateSearchLike = WhereStr
End Function

Public Function PackAddress() As String
Dim AddressStr As String
   AddressStr = ""
   
   If Trim(HOME) <> "" Then
      AddressStr = AddressStr & HOME & " "
   End If

   If Trim(SOI) <> "" Then
      AddressStr = AddressStr & "´." & SOI & " "
   End If

   If Trim(MOO) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË") & MOO & " "
   End If

   If Trim(VILLAGE) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË∫È“π") & VILLAGE & " "
   End If

   If Trim(ROAD) <> "" Then
      AddressStr = AddressStr & MapText("∂.") & ROAD & " "
   End If

   If Trim(DISTRICT) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("·¢«ß") & DISTRICT & " "
      Else
         AddressStr = AddressStr & MapText("µ”∫≈") & DISTRICT & " "
      End If
   End If

   If Trim(AMPHUR) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("‡¢µ") & AMPHUR & " "
      Else
         AddressStr = AddressStr & MapText("Õ”‡¿Õ") & AMPHUR & " "
      End If
   End If

   If Trim(PROVINCE) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & "" & PROVINCE & " "
      Else
         AddressStr = AddressStr & MapText("®.") & PROVINCE & " "
      End If
   End If

   If Trim(ZIPCODE) <> "" Then
      AddressStr = AddressStr & " " & ZIPCODE & " "
   End If

   PackAddress = AddressStr
End Function

