VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBillingD002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBillingD002"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<48|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double

Private CreditBalances As Collection
Private TotalMoves As Collection

Private m_PaidAmounts As Collection

Private m_CheckBillingMovement As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   
   Set CreditBalances = New Collection
   Set TotalMoves = New Collection
   
   Set m_PaidAmounts = New Collection
   
   Set m_CheckBillingMovement = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set CreditBalances = Nothing
   Set TotalMoves = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_CheckBillingMovement = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================
Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Amt As Double
Dim iCount As Long

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(4, "^", "�������١���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�١���", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "�ѹ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�͡���#", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��ѡ�ҹ���", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ʹ㹺��", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "����˹��", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Ŵ˹��/�Ѻ�׹", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ʹ����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ����ҧ", ">")
   Call m_Headers1.AddField(Cf)
   
   
   Set Cf = Nothing
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim ToRcp As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("FROM_BILL_DATE") < 1 Then
      FromDate = ""
      ToDate = "� �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   Else
      FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
      ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   End If
   
   If mcolParam("TO_RCP_DATE") < 1 Then
      ToRcp = "�Դ�ѹ�Ѻ���з�����"
   Else
      ToRcp = "�֧�ѹ����Ѻ���� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_RCP_DATE")), "N/A")
   End If
   Call GetPaidAmountByDocID(m_PaidAmounts, mcolParam("FROM_BILL_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), RECEIPT2_DOCTYPE)
   
   If mcolParam("SHOW_ONLY_MOVE") = 1 Then
      Call GetDistinctCheckBillingMovement(m_CheckBillingMovement, mcolParam("FROM_MOVE_DATE"), mcolParam("TO_MOVE_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"))
   End If
   
   MODULE_DESC = vbCrLf & glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " " & ToRcp & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & vbCrLf & _
                                        "�ҡ��ѡ�ҹ��� " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " �֧��ѡ�ҹ��� " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
   
   If mcolParam("SHOW_ONLY_MOVE") = 1 Then
      MODULE_DESC = MODULE_DESC & " �ʴ�੾������͹��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_MOVE_DATE")), "N/A") & "-" & EmptyToString(DateToStringExtEx2(mcolParam("TO_MOVE_DATE")), "N/A")
   End If
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim TempBd As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(10) As Double
Dim Total2(10) As Double
Dim Total3(10) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim Rcp As CRcpCnDn_Item
Dim ShowHeader As Boolean
Dim M As Long
Dim TempRcp As Double
Dim TempDebit As Double
Dim TempCreDit As Double
Dim ShortCode As String
Dim CheckMove As Boolean
Dim tempCheckMove As CBillingDoc

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_TO_EXCEL") = 1 Then
      Call printHeader
   End If
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   
   m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
   m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.APAR_TYPE = mcolParam("APAR_TYPE")
   'APAR_TYPE
   If mcolParam("SHORT_CODE") = 0 Then
      ShortCode = ""
   Else
      ShortCode = mcolParam("SHORT_CODE")
   End If
   m_BillingDoc.SHORT_CODE = ShortCode
   m_BillingDoc.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
   m_BillingDoc.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
   Call m_BillingDoc.QueryData(13, Rs, iCount)
   
   I = 0
   k = 0
   L = 0
   M = 0
   TempCreDit = 0
   TempDebit = 0
   TempRcp = 0
   
   CheckMove = False
   While Not Rs.EOF
      Call m_Details1.ClearField
      Call m_BillingDoc.PopulateFromRS(13, Rs)
      
'      If m_BillingDoc.APAR_CODE = "0-�001" Then
'         Debug.Print
'      End If
      If mcolParam("SHOW_ONLY_MOVE") = 0 Then
         CheckMove = True
      ElseIf mcolParam("SHOW_ONLY_MOVE") = 1 Then
         Set tempCheckMove = GetObject("CBillingDoc", m_CheckBillingMovement, Trim(m_BillingDoc.APAR_CODE), False)
         If tempCheckMove Is Nothing Then
            CheckMove = False
         Else
            CheckMove = True
         End If
      End If
      
      If CheckMove Then
         Set Rcp = GetObject("CRcpCnDn_Item", m_PaidAmounts, Trim(Str(m_BillingDoc.BILLING_DOC_ID)), False)
         If Not (Rcp Is Nothing) Then
            Rcp.Flag = "Y"
         Else
            Set Rcp = New CRcpCnDn_Item
            Rcp.Flag = "N"
         End If
   '      If Round(Rcp.PAID_AMOUNT, 2) <> 0 Then
   '         Call m_PaidAmounts.Remove(Trim(Str(m_BillingDoc.BILLING_DOC_ID)))
   '      End If
         
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT - Rcp.PAID_AMOUNT
            TempRcp = TempRcp + Rcp.PAID_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT - Rcp.PAID_AMOUNT
            TempCreDit = TempCreDit + Rcp.PAID_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            Amt = m_BillingDoc.PAY_AMOUNT + m_BillingDoc.VAT_AMOUNT - Rcp.PAID_AMOUNT
         End If
         
         If Round(Amt, 2) <> 0 Or mcolParam("SHOW_DETAIL_ZERO") = 1 Then
            If PrevKey1 <> m_BillingDoc.APAR_TYPE_CODE And M > 0 Then
               
               If Round(Total1(8), 2) <> 0 Then
                  If mcolParam("SHOW_DETAIL") = 1 Then
                     Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
                  Else
                     Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
                  End If
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  If mcolParam("SHOW_DETAIL") = 1 Then
                     Vsp.FontBold = True
                  End If
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  ShowHeader = False
               End If
               
               If Round(Total2(8), 2) <> 0 Then
                  Call GenerateFooter("���������", " " & PrevKey2, k & " �", Total2, m_Details1, 0)            '�������١���
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
               
               k = 0
               I = 0
               
               Vsp = "�������١���  :  " & m_BillingDoc.APAR_TYPE_NAME & " ( " & m_BillingDoc.APAR_TYPE_CODE & " ) "
               
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
                  Total2(j) = 0
               Next j
            ElseIf PrevKey3 <> m_BillingDoc.APAR_CODE And M > 0 Then
               If Round(Total1(8), 2) <> 0 Then
                  If mcolParam("SHOW_DETAIL") = 1 Then
                     Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
                  Else
                     Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
                  End If
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  If mcolParam("SHOW_DETAIL") = 1 Then
                     Vsp.FontBold = True
                  End If
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  ShowHeader = False
               End If
               
               I = 0
               
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            ElseIf M = 0 Then
               ShowHeader = False
               Vsp = "�������١���  :  " & m_BillingDoc.APAR_TYPE_NAME & " ( " & m_BillingDoc.APAR_TYPE_CODE & " ) "
            End If
            
            PrevKey1 = m_BillingDoc.APAR_TYPE_CODE
            PrevKey2 = m_BillingDoc.APAR_TYPE_NAME
            PrevKey3 = m_BillingDoc.APAR_CODE
            PrevKey4 = m_BillingDoc.APAR_LONG_NAME
   
            I = I + 1
            k = k + 1
            L = L + 1
            M = M + 1
            SumRow = 0
            
            Call m_Details1.ClearField
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = "" & DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = m_BillingDoc.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = m_BillingDoc.SALE_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(4)
            Amt = 0
            If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
               Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               TempStr = FormatNumberToNull(Amt)
               Total1(4) = Total1(4) + Amt
               Total2(4) = Total2(4) + Amt
               Total3(4) = Total3(4) + Amt
               SumRow = Amt
            Else
               TempStr = ""
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(5)
            If m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
               Amt = m_BillingDoc.PAY_AMOUNT + m_BillingDoc.VAT_AMOUNT
               TempStr = FormatNumberToNull(Amt)
               Total1(5) = Total1(5) + Amt
               Total2(5) = Total2(5) + Amt
               Total3(5) = Total3(5) + Amt
               SumRow = Amt
            Else
               TempStr = ""
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
               Amt = -(m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT)
               TempStr = FormatNumberToNull(Amt)
               Total1(6) = Total1(6) + Amt
               Total2(6) = Total2(6) + Amt
               Total3(6) = Total3(6) + Amt
               SumRow = Amt
            ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
               Amt = -(m_BillingDoc.PAY_AMOUNT + m_BillingDoc.VAT_AMOUNT)
               TempStr = FormatNumberToNull(Amt)
               Total1(6) = Total1(6) + Amt
               Total2(6) = Total2(6) + Amt
               Total3(6) = Total3(6) + Amt
               SumRow = Amt
            Else
               TempStr = ""
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(7)
            If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
               Amt = Rcp.PAID_AMOUNT
            ElseIf m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
               Amt = Rcp.PAID_AMOUNT
            ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
               Amt = -Rcp.PAID_AMOUNT
            ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
               Amt = -Rcp.PAID_AMOUNT
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt)
            Total1(7) = Total1(7) + Amt
            Total2(7) = Total2(7) + Amt
            Total3(7) = Total3(7) + Amt
            SumRow = SumRow - Amt
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(8)
            TempStr = FormatNumberToNull(SumRow)
            Total1(8) = Total1(8) + SumRow
            Total2(8) = Total2(8) + SumRow
            Total3(8) = Total3(8) + SumRow
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Round(SumRow, 2) <> 0 Or mcolParam("SHOW_DETAIL_ZERO") = 1 Then
               If Not (ShowHeader) Then
                  If mcolParam("SHOW_DETAIL") = 1 Then
                     Vsp = "     �١���  :  " & m_BillingDoc.APAR_LONG_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
                  End If
                  ShowHeader = True
               End If
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
            Else
               I = I - 1
               k = k - 1
               L = L - 1
            End If
         End If
      End If
      Rs.MoveNext
   Wend
   
   If Round(Total1(8), 2) <> 0 Then
      If Not (ShowHeader) Then
         Vsp = "     �١���  :  " & m_BillingDoc.APAR_LONG_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
      End If
               
      If mcolParam("SHOW_DETAIL") = 1 Then
         Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
      Else
         Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      If mcolParam("SHOW_DETAIL") = 1 Then
         Vsp.FontBold = True
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If
   
   If Round(Total2(8), 2) <> 0 Then
      Call GenerateFooter("���������", " " & PrevKey2, k & " �", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If
   
   If Round(Total3(8), 2) <> 0 Then
      Call GenerateFooter("", "�ط��", L & " �", Total3, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If

   'debug.print (TempRcp)
   'debug.print (TempCreDit)
'   Debug.Print m_PaidAmounts.Count
'   For Each Rcp In m_PaidAmounts
'      If Rcp.Flag <> "Y" Then
'         Debug.Print
'      End If
'   Next Rcp
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function
Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_TO_EXCEL") = 1 Then
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = txt1
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = txt2
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Else
      Set HeadCf = m_Headers1.Fields(1)
      TempWidth = TempWidth + HeadCf.Width
      Set HeadCf = m_Headers1.Fields(2)
      TempWidth = TempWidth + HeadCf.Width
      TempStr = txt1 & " " & txt2
      Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)

   End If
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j >= 3 Then
         If j = 3 Then
            TempStr = Txt3
         Else
            TempStr = FormatNumberToNull(Tot(j))
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.GetFieldValue("TAX_ID"))
   Vsp.CurrentY = dY(0) - 100
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   
   Vsp.FontSize = iSize
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub
Private Sub GenerateHeader(txt1 As String, txt2 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
'Private Sub GenerateFooterDebitCredit(txt1 As String, txt2 As String, Txt3 As String, DebitAmount As Double, CreditAmount As Double, Details As CFieldList)
'Dim HeadCf As CReportField
'Dim BodyCf As CReportField
'Dim j As Long
'Dim TempStr As String
'Dim TempWidth As Double
'
'   Call Details.ClearField
'   Set BodyCf = New CReportField
'
'   Set HeadCf = m_Headers1.Fields(1)
'   TempWidth = TempWidth + HeadCf.Width
'   Set HeadCf = m_Headers1.Fields(2)
'   TempWidth = TempWidth + HeadCf.Width
'   TempStr = txt1 & " " & txt2
'   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, HeadCf.BodyAlign)
'   Call Details.AddField(BodyCf)
'
'   j = 0
'   For Each HeadCf In m_Headers1.Fields
'      j = j + 1
'      If j >= 3 Then
'         If j = 3 Then
'            TempStr = Txt3
'         ElseIf j = 5 Then
'            TempStr = FormatNumberToNull(DebitAmount)
'         ElseIf j = 6 Then
'            TempStr = FormatNumberToNull(-CreditAmount)
'         ElseIf j = 7 Then
'            TempStr = ""
'         ElseIf j = 8 Then
'            TempStr = FormatNumberToNull(DebitAmount - CreditAmount)
'         End If
'
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
'         Call Details.AddField(BodyCf)
'      End If
'   Next HeadCf
'
'   Set BodyCf = Nothing
'End Sub