VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPrintLabel006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CReportInterface

Private Const MODULE_NAME = "CReportLabelDelivery006"
Private MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Details1 As CFieldList
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<52|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig

Private m_cData As CBillingDoc

Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private m_BillingDocIDColl As Collection
Private DriverIdCOll  As Collection
Private TranSportorIdCOll  As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_ReportConfig = New CReportConfig
   Set m_BillingDocIDColl = New Collection
   Set DriverIdCOll = New Collection
   Set TranSportorIdCOll = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList

End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_ReportConfig = Nothing
   Set m_BillingDocIDColl = Nothing
   Set DriverIdCOll = Nothing
   Set TranSportorIdCOll = Nothing
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
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
'   Vsp.FontSize = HEADER_SIZE
   Vsp.FontSize = 18
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Set Cf = New CReportField
   Call Cf.SetFieldValue(10, "<", "", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "<", "", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(10, "^", "", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "", "<")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
            
'   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Function initDoc() As Boolean
Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   
   Call LoadDisTinctBillingDocID(m_BillingDocIDColl, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), PatchWildCard(mcolParam("DOCUMENT_NO_SEARCH")), , , , mcolParam("ORDER_BY"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")", mcolParam("DRIVER_ID"), mcolParam("TRANSPORTOR_ID"), , , PatchWildCard(mcolParam("DOCUMENT_NO_FROM")), PatchWildCard(mcolParam("DOCUMENT_NO_TO")))
   Call LoadMaster(Nothing, DriverIdCOll, , , MASTER_DRIVER)
   Call LoadMaster(Nothing, TranSportorIdCOll, , , MASTER_TRANSPORTOR)

   Vsp.PaperSize = pprA4
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 0
   Vsp.MarginFooter = 0
   Vsp.MarginHeader = 0
   Vsp.MarginLeft = 0
   Vsp.MarginRight = 0
   Vsp.MarginTop = 0
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   MODULE_DESC = ""

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
   
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
Dim iCount As Long
Dim I As Long
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim BodyCf1 As CReportField
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CDocItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim j As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim AskMessage As VbMsgBoxResult
Dim TempString As String
Dim TempBillingDocID  As CBillingDoc
Dim TempBd As CBillingDoc
Dim CountRow As Long
Dim CountCol As Long
Dim TempLine As Double
Dim TempStr As String
Dim NewPage As Boolean
Dim Total As Double


Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
   RName = "genDoc"
   
   TempString = "ต้องการพิมพ์ทันทีหรือไม่? (YES พิมพ์ทันที   NO แสดงผลรายงานก่อน  CANCEL ออก) "
   AskMessage = MsgBox(TempString, vbYesNoCancel, glbErrorLog.MsgBoxTitle)
   If AskMessage = vbCancel Then
      frmReport.AutoPrintMode = True
      Exit Function
   End If
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function
   
   If m_BillingDocIDColl.Count <= 0 Then
         mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
         Exit Function
   End If
   
   If AskMessage = vbNo Then
      Vsp.Preview = True
   Else
      Vsp.Preview = False
   End If
   
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   Call StartExportFile(Vsp)
      
   mblnHeader = True
   Vsp.StartDoc
   I = 0
   CountRow = CountRow + 1
   NewPage = True
   For Each TempBillingDocID In m_BillingDocIDColl
      I = I + 1

      Set Rs = New ADODB.Recordset
      Set m_cData = New CBillingDoc
      m_cData.BILLING_DOC_ID = TempBillingDocID.BILLING_DOC_ID
      m_cData.QueryFlag = 1
      If Not glbDaily.QueryBillingDocOnlyDo(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
         mstrErrMsg = "Query Error."
         Exit Function
      End If

      Call m_cData.PopulateFromRS2(172, Rs)
      Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
      Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
      Call mcolParam.add(m_cData.APAR_CODE, "CUSTOMER_CODE")
      Call mcolParam.add(m_cData.APAR_NAME, "CUSTOMER_NAME")
      Call mcolParam.add(m_cData.APAR_BILL_NAME, "CUSTOMER_BILL_NAME")
      Call mcolParam.add(m_cData.CREDIT, "CREDIT")
      Call mcolParam.add(m_cData.Due_Date, "DUE_DATE")
      Call mcolParam.add(m_cData.TAX_ID, "TAX_ID")
      
      Call mcolParam.add(m_cData.SALE_CODE, "SALE_CODE")
      Call mcolParam.add(m_cData.SALE_LONG_NAME & " " & m_cData.SALE_LAST_NAME, "SALE_NAME")
      Call mcolParam.add(m_cData.CUS_PO, "CUS_PO")
      Call mcolParam.add(m_cData.CUSTOMER_BRANCH_NAME, "CUSTOMER_BRANCH_NAME")
      
      Call mcolParam.add(m_cData.TRANSPORTOR_ID2, "TRANSPORTOR_ID2")
      Call mcolParam.add(m_cData.TRANSPORTOR_NAME, "TRANSPORTOR_NAME")
      Call mcolParam.add(m_cData.CASH_DELIVERY_FLAG, "CASH_DELIVERY_FLAG")
      Call mcolParam.add(m_cData.INDEX_LINK, "INDEX_LINK")
      Call mcolParam.add(m_cData.SHORT_CODE, "SHORT_CODE")
      
      If m_cData.BILLING_ADDRESS_ID > 0 Then
         Set Addr = New CAddress
         Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BILLING_ADDRESS_ID)
         Call Addr.QueryData(1, Rs, iCount)
         If Not Rs.EOF Then
            Call Addr.PopulateFromRS(1, Rs)
            Call mcolParam.add(Addr.PackAddress2, "CUSTOMER_ADDRESS")
            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "PHONE")
            Call mcolParam.add(Addr.GetFieldValue("FAX1"), "FAX")
         End If
         Set Addr = Nothing
      Else
         Call mcolParam.add("", "CUSTOMER_ADDRESS")
         Call mcolParam.add("", "PHONE")
         Call mcolParam.add("", "FAX")
      End If
   
      If m_cData.BRANCH_ADDRESS > 0 Then
         Set Addr = New CAddress
         Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BRANCH_ADDRESS)
         Call Addr.QueryData(1, Rs, iCount)
         If Not Rs.EOF Then
            Call Addr.PopulateFromRS(1, Rs)
            Call mcolParam.add(Addr.PackAddress2, "BRANCH_ADDRESS")
            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "BRANCH_ADDRESS_PHONE")
            Call mcolParam.add(Addr.GetFieldValue("FAX1"), "BRANCH_ADDRESS_FAX")
            Call mcolParam.add(Addr.GetFieldValue("SHOW_LOCATION_FLAG"), "SHOW_LOCATION_FLAG")
         End If
         Set Addr = Nothing
      Else
         Call mcolParam.add("", "BRANCH_ADDRESS")
         Call mcolParam.add("", "BRANCH_ADDRESS_PHONE")
         Call mcolParam.add("", "BRANCH_ADDRESS_FAX")
         Call mcolParam.add("", "SHOW_LOCATION_FLAG")
      End If

   If mcolParam("TRANSPORTOR_ID2") > 0 And mcolParam("INDEX_LINK") = 1 Then
      CountCol = CountCol + 1
      If CountCol Mod 4 = 0 Then 'เมื่อครบ 3 แถวก็ทำการพิมพ์ลงไป 1 บรรทัด
          CountCol = 1
          CountRow = CountRow + 1
         If CountRow = 8 Then 'ครบ 7 บรรทัดให้ขึ้นหน้าใหม่s
            Vsp.TableBorder = tbNone
            Vsp.EndTable
         
            Vsp.NewPage
            NewPage = True
            CountRow = 1
            CountCol = 1
         End If
      End If
         
      If NewPage Then
         NewPage = False
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 3
         Vsp.TableCell(tcRows) = 1
          Vsp.TableCell(tcRowHeight) = (1.6 * 567) + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
          Vsp.TableCell(tcColWidth, , 1) = "9.9cm"
          Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
          Vsp.TableCell(tcColWidth, , 3) = "9.9cm"
          Vsp.TableBorder = tbNone
          Vsp.EndTable
   
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 3
         Vsp.TableCell(tcRows) = 7
          Vsp.TableCell(tcRowHeight) = "3.80cm"
          Vsp.TableCell(tcColWidth, , 1) = "9.9cm"
          Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
          Vsp.TableCell(tcColWidth, , 3) = "9.9cm"
         
         Vsp.FontSize = 15
      End If
      
         TempStr = "ชื่อและที่อยู่ผู้รับ : " & mcolParam("CUSTOMER_NAME") _
                              & vbCrLf & mcolParam("BRANCH_ADDRESS")  '& "  " & mcolParam("DOCUMENT_NO")
                                    
        
        If mcolParam("CASH_DELIVERY_FLAG") = "Y" Then
            Total = m_cData.TOTAL_PRICE - m_cData.DISCOUNT_AMOUNT - m_cData.EXT_DISCOUNT_AMOUNT + m_cData.VAT_AMOUNT
            If mcolParam("INDEX_LINK") = 1 Then
                TempStr = TempStr & "(" & mcolParam("SHORT_CODE") & " " & FormatNumber(RoundNumber(Total, 0.5), 0) & " บ.)" '& "->" & Total
             Else
                 TempStr = TempStr & "(" & mcolParam("TRANSPORTOR_NAME") & " " & FormatNumber(RoundNumber(Total, 0.5), 0) & " บ.)" '& "->" & Total
            End If
        ElseIf mcolParam("INDEX_LINK") = 1 Then
            TempStr = TempStr & "(" & mcolParam("SHORT_CODE") & ")"
        End If
               
         Vsp.TableCell(tcText, CountRow, CountCol) = TempStr
         Vsp.TableCell(tcAlign, CountRow, CountCol) = taLeftTop
         
         If CountCol + 1 = 2 Then 'เขียนช่องว่าง
            CountCol = CountCol + 1
            
            Vsp.TableCell(tcText, CountRow, CountCol + 1) = ""
            Vsp.TableCell(tcAlign, CountRow, CountCol + 1) = taLeftTop
         End If
   End If

'
      '--------------      Clear mColparam ----------------'
      Call mcolParam.Remove("DOCUMENT_NO")
      Call mcolParam.Remove("DOCUMENT_DATE")
      Call mcolParam.Remove("CUSTOMER_CODE")
      Call mcolParam.Remove("CUSTOMER_NAME")
      Call mcolParam.Remove("CUSTOMER_BILL_NAME")
      Call mcolParam.Remove("CREDIT")
      Call mcolParam.Remove("DUE_DATE")
      Call mcolParam.Remove("TRANSPORTOR_ID2")
      Call mcolParam.Remove("TRANSPORTOR_NAME")
      Call mcolParam.Remove("CASH_DELIVERY_FLAG")
      Call mcolParam.Remove("INDEX_LINK")
      Call mcolParam.Remove("SHORT_CODE")

      Call mcolParam.Remove("SALE_CODE")
      Call mcolParam.Remove("SALE_NAME")
      Call mcolParam.Remove("CUS_PO")
      Call mcolParam.Remove("TAX_ID")
      Call mcolParam.Remove("CUSTOMER_BRANCH_NAME")

      Call mcolParam.Remove("CUSTOMER_ADDRESS")
      Call mcolParam.Remove("PHONE")
      Call mcolParam.Remove("FAX")

      Call mcolParam.Remove("BRANCH_ADDRESS")
      Call mcolParam.Remove("BRANCH_ADDRESS_PHONE")
      Call mcolParam.Remove("BRANCH_ADDRESS_FAX")
      Call mcolParam.Remove("SHOW_LOCATION_FLAG")
      
      '--------------      Clear mColparam ----------------'
'
      Set Rs = Nothing
      Set m_cData = Nothing
   Next TempBillingDocID
   
   Vsp.TableBorder = tbNone
   Vsp.EndTable
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   
   If AskMessage = vbYes Then
      Vsp.PrintDoc (True)
      frmReport.AutoPrintMode = True
   End If
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function

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
   
  Vsp.MarginLeft = m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
  'Vsp.MarginTop = m_ReportConfig.GetFieldValue("HEAD_OFFSET")
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
      
   
   If mblnHeader Then
      Call printHeader
   End If
   
   'mdY = Vsp.CurrentY
   
   'Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
   
   
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub



