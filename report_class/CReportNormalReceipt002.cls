VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalRcp002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalRcp002"
Private MODULE_DESC  As String

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
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private Const HEADER_RECEIPT1 = "^1|^7|^2|^2|^2;"
Private Const HEADER_RECEIPT2 = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_BILLS = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_CNDN = "^1|^4|^4|^4;"

Private Const DETAIL_RECEIPT1 = "^0|^1|<7|>2|>2|>2|^0|^0;"
Private Const DETAIL_RECEIPT2 = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_BILLS = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_CNDN = "^1|^4|^4|>4;"

Private Const FOOTER_RECEIPT1 = "<10|^2|>2;"
Private Const FOOTER_RECEIPT1_1 = "<10|>2|>2;"
Private Const FOOTER_RECEIPT1_2 = "<8|>4|>2;"
Private Const FOOTER_RECEIPT1_3 = ">12|>2;"
Private Const FOOTER_RECEIPT2 = "<13|^3|>3;"
Private Const FOOTER_BILLS = "<13|^3|>3;"
Private Const FOOTER_CNDN = "<5|>4|>4;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As cName
Private m_Addr As CAddress
Private m_cData As CBillingDoc
Private m_ChildData As CBillingDoc
Private m_OwnerAddress As CAddress

Private Sp As CSystemParam
Private X1 As Long
Private X2  As Long
Private Y1  As Long
Private Y2  As Long
Private AllEditEndPage As Boolean
Private m_Cd  As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New cName
   Set m_Addr = New CAddress
   Set m_OwnerAddress = New CAddress
   Set m_Cd = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_Addr = Nothing
   Set m_OwnerAddress = Nothing
   Set m_Cd = Nothing
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
   If mcolParam("PREVIEW_TYPE") = 2 Then
      CReportInterface_Preview = genDoc(False)
   Else
      CReportInterface_Preview = genDoc(True)
   End If
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "ผู้จัดการ|.....................................|ชื่อผู้รับเงิน|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "วันที่|........./........./.........|วันที่|........./........./.........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

Private Sub PrintBillHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempAddr As String
Dim Doc As CDocItem
Dim iSize As Integer
   
   If I = 0 Then
      I = 1
      Set TempRs = New ADODB.Recordset
      
      Call m_Enterprise.SetFieldValue("ENTERPRISE_ID", -1)
      Call m_Enterprise.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_Enterprise.PopulateFromRS(1, TempRs)
      End If
      
      Call m_EnpAddress.SetFieldValue("ADDRESS_ID", m_cData.ENTERPRISE_ADDRESS_ID)
      Call m_EnpAddress.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_EnpAddress.PopulateFromRS(1, TempRs)
      End If
      
      Call m_CustAddress.SetFieldValue("ADDRESS_ID", m_cData.BILLING_ADDRESS_ID)
      Call m_CustAddress.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_CustAddress.PopulateFromRS(1, TempRs)
      End If
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = m_Enterprise.GetFieldValue("ENTERPRISE_NAME") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   Y1 = Vsp.CurrentY
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   iSize = Vsp.FontSize
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|ใบเสร็จรับเงิน" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.GetFieldValue("TAX_ID") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontBold = False
   Y2 = Vsp.CurrentY
   
   '--------------------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.CurrentY = Y1
   Vsp.FontSize = 24
   strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
   strPrint = "|ใบเสร็จรับเงิน" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   '--------------------------------------------------------------------------------------------------------------------------------------------------------
   Vsp.FontSize = iSize
   Vsp.CurrentY = Y2
   
   'เลขที่ Run จาก AUTO ของใบเสร็จ แนบใบส่งของ
   '------------------------------------------------------------------------------------------------------------------------------
   Dim HeadNo As String
   Dim RunningNo  As Long
   Dim ConFigDocType As Long
   Dim FormatAmount As String
   
   strFormat = VSP_CalTable(">10|^2;", mdblWid, alngX)
   strPrint = "เลขที่|" & GetDocumentNo(-1, -1, m_cData.DOCUMENT_DATE, HeadNo, RunningNo, ConFigDocType, FormatAmount) & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   '------------------------------------------------------------------------------------------------------------------------------
   Dim Cd  As CConfigDoc
   Set Cd = New CConfigDoc
   Call Cd.SetFieldValue("RUNNING_NO", RunningNo)
   Call Cd.SetFieldValue("LAST_NO", HeadNo & Format(Trim(Str(RunningNo)), FormatAmount))
   Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ConFigDocType)
   Call Cd.UpdateRunningNo
   Set Cd = Nothing
   
   strFormat = VSP_CalTable(">10|^2;", mdblWid, alngX)
   strPrint = "วันที่|" & DateToStringExtEx2((m_cData.DOCUMENT_DATE)) & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<2|<8;", mdblWid, alngX)
   strPrint = "ได้รับเงินจาก | " & m_cData.APAR_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strFormat = VSP_CalTable("<2|<8;", mdblWid, alngX)
   strPrint = "ที่อยู่ | " & mcolParam("CUSTOMER_ADDRESS") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<2|<8|<2|<8;", mdblWid, alngX)
   strPrint = "ชำระค่า |สินค้า|ตามบิล|" & m_cData.DOCUMENT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub
Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdt As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call PrintBillHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   'Vsp.FontBold = True
   'Vsp.TableBorder = tbAll
   
      
'   strFormat = VSP_CalTable(HEADER_RECEIPT1, mdblWid, alngX)
'   strPrint = "No.|รายการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'   X1 = Vsp.PageWidth - Vsp.MarginRight
'   Y1 = Vsp.CurrentY
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Function initDoc() As Boolean
   Dim strDate As String
   Dim TempPa As PaperSizeSettings
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   
   Vsp.TableBorder = tbNone
   Vsp.PaperSize = pprLetter
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 100
   Vsp.MarginFooter = 100
   Vsp.MarginHeader = 100
   Vsp.MarginLeft = 100
   Vsp.MarginRight = 100
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
'   For TempPa = 1 To 500
'      If Vsp.PaperSizes(TempPa) Then
'
'         Vsp.PaperSize = TempPa
'         If Vsp.PageHeight <= 10000 Then
'            'debug.print (TempPa)
'            'debug.print (Vsp.PageHeight)
'            'debug.print (Vsp.PageWidth)
'            'debug.print
'         End If
'      End If
'   Next TempPa
'   Vsp.PaperSize = pprA5
'   Vsp.Orientation = orLandscape
'   Vsp.PaperSize = pprEnvDL
'   'debug.print (Vsp.PageHeight)
'   'debug.print (Vsp.PageWidth)
   
   Call LoadConfigDoc(Nothing, m_Cd)
   
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
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Bi As CDocItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim DiscountAmount As Double
Dim VatAmount As Double
Dim IncludeDiscountAmount As Double
Dim IncludeVatAmount As Double
Dim TempVatAmount As Double
Dim TempDeposit As Double
Dim TempDiscount As Double
Dim TempCol As Collection
Dim SpaceCount As Long
Dim Ri As CRcpCnDn_Item
Dim Temp1 As Double
Dim Temp2 As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Call m_cData.PopulateFromRS(1, Rs)
   
   Dim Addr As CAddress
   If m_cData.BILLING_ADDRESS_ID > 0 Then
      Set Addr = New CAddress
      Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BILLING_ADDRESS_ID)
      Call Addr.QueryData(1, Rs, iCount)
      If Not Rs.EOF Then
         Call Addr.PopulateFromRS(1, Rs)
          Call mcolParam.add(Addr.PackAddress, "CUSTOMER_ADDRESS")
          Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "PHONE")
          Call mcolParam.add(Addr.GetFieldValue("FAX1"), "FAX")
      End If
      Set Addr = Nothing
   Else
      Call mcolParam.add("", "CUSTOMER_ADDRESS")
      Call mcolParam.add("", "PHONE")
      Call mcolParam.add("", "FAX")
   End If
   
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   Call StartExportFile(Vsp)
   
   Vsp.PenStyle = psDash
   
   Vsp.StartDoc
      
      'Vsp.TableBorder = tbAll
      I = 0
      AllEditEndPage = True
      TempDiscount = 0
      TempVatAmount = 0
   
'      strFormat = VSP_CalTable(DETAIL_RECEIPT1, mdblWid, alngX)
'      Set TempCol = m_cData.DocItems
'      For Each Bi In TempCol
'         I = I + 1
'         strPrint = "|"
'         strPrint = strPrint & I & "|"
'         strPrint = strPrint & Bi.ShowDescCode & " " & Bi.ShowDescText & "|"
'         strPrint = strPrint & FormatNumber(MyDiff(Bi.GetFieldValue("ITEM_AMOUNT"), Bi.GetFieldValue("UNIT_MULTIPLE"))) & " " & Bi.GetFieldValue("UNIT_TRAN_NAME") & "|"
'         strPrint = strPrint & FormatNumber(Bi.GetFieldValue("AVG_PRICE") * Bi.GetFieldValue("UNIT_MULTIPLE")) & "|"
'         strPrint = strPrint & FormatNumber(Bi.GetFieldValue("TOTAL_PRICE")) & "|12|12;"
'
'         Sum3 = Sum3 + Bi.GetFieldValue("TOTAL_PRICE")
'         TempDiscount = TempDiscount + Bi.GetFieldValue("DISCOUNT_AMOUNT")
'
'         If I >= TempCol.Count Then
'            Call AdjustPage(Vsp, strFormat, strPrint, 5000)
'         End If
'         Call Vsp.AddTable(strFormat, "", strPrint)
'      Next Bi
'
'      Call PatchTable(Vsp, strFormat, "||||;", 5000) '10
'
'      X2 = Vsp.PageWidth - Vsp.MarginRight
'      Y2 = Vsp.CurrentY
'
'      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
'      Call Vsp.DrawLine(X1, Y1, X2, Y2)
'      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
'      Vsp.TableBorder = tbNone
'      Y1 = Vsp.CurrentY
'      AllEditEndPage = False
'
'      strFormat = VSP_CalTable(FOOTER_RECEIPT1, mdblWid, alngX)
'      Vsp.FontBold = True
'      strPrint = "หมายเหตุ    " & m_cData.NOTE & "|รวม|" & FormatNumber(Sum3) & ";"
'      Call Vsp.AddTable(strFormat, "", strPrint)
'
'      strPrint = "มูลค่าของสินค้าหรือบริการที่ถูกต้อง|" & FormatNumber(Temp2 - Temp1) & ";"
'      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "ชำระเป็น       (  ) เงินสด"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "                     (  ) โอนเงิน"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "                     (  ) เช็คธนาคาร ______________________ เลขที่ ______________________ ลงวันที่ ______________________ ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      
      Sum3 = m_cData.TOTAL_PRICE + m_cData.VAT_AMOUNT - m_cData.DISCOUNT_AMOUNT - m_cData.EXT_DISCOUNT_AMOUNT
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "                     จำนวนเงิน   " & FormatNumber(Sum3) & "   บาท  (" & ThaiBaht(Sum3) & ");"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      
      Vsp.FontBold = False
      Call PrintBillFooter(alngX)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   
   If mcolParam("PREVIEW_TYPE") = 2 Then
      Vsp.PrintDoc (True)
   End If
   
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
End Function
'Private Sub VSP_EndDoc()
''This event occur when VSPrinter.EndDoc is used
'End Sub

'Private Sub VSP_EndPage()
'   If AllEditEndPage Then
'      X2 = Vsp.PageWidth - Vsp.MarginRight
'      Y2 = Vsp.CurrentY
         
'      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
'      Call Vsp.DrawLine(X1, Y1, X2, Y2)
'      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
'   End If
   
'   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'End Sub

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
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   'Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   'Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
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

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Private Function GetDocumentNo(DocumentType As Long, DocumentSubType As Long, DocumentDate As Date, HeadNo As String, RunningNo As Long, ConFigDocType As Long, TempStr As String) As String
Dim ID As Long
Dim Cd As CConfigDoc
Dim I As Long
   
   GetDocumentNo = ""
   
   ID = ConvertDocToConfigNo(3, DocumentType, DocumentSubType)
   If ID <= 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถดำเนินการต่อได้ เนื่องจากระบบจำเป็นที่จะต้องตั้งหมายเลขเอกสารอัตโนมัติไว้ก่อน"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         Dim TempCd As CConfigDoc
         ''''''''''''''
         
         GetDocumentNo = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         HeadNo = GetDocumentNo
         GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1, TempStr)
         RunningNo = Cd.GetFieldValue("RUNNING_NO") + 1
         ConFigDocType = ID
      ElseIf Cd Is Nothing Then
         glbErrorLog.LocalErrorMsg = "ไม่สามารถดำเนินการต่อได้ เนื่องจากระบบจำเป็นที่จะต้องตั้งหมายเลขเอกสารอัตโนมัติไว้ก่อน"
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
End Function


