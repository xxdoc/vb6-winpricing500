VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBillingPo001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBillingPo001"
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

Private Const TITLE_FORMAT = "<52|<3|<5;"

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
Private m_YearSeqs As Collection
Private m_InvoidByPo As Collection

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
   Set m_YearSeqs = New Collection
   Set m_InvoidByPo = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_InvoidByPo = Nothing
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
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(5, "^", "�Ţ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ѹ���", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(5, "^", "�����١���", "<")
   Call m_Headers1.AddField(Cf)
   
    If mcolParam("SHOW_PO_DETAIL") = 0 And mcolParam("SHOW_INV_DETAIL") = 0 Then
      Call Cf.SetFieldValue(0, "^", "��¡��", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "��Ť��", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(5, "^", "��¡��", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "��Ť��", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(3, "^", "�ӹǹ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "�Ҥ����", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_PO_DETAIL") = 0 And mcolParam("SHOW_INV_DETAIL") = 0 Then
      Call Cf.SetFieldValue(0, "^", "˹���", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(3, "^", "˹���", "^")
      Call m_Headers1.AddField(Cf)
   End If

   
   Set Cf = Nothing
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadInvoidByPo(m_InvoidByPo, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
   
    FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & " " & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A")
                                       
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
Dim II As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String

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

   For j = 1 To UBound(Total2)
      Total2(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField

      Set m_BillingDoc = New CBillingDoc
      
      m_BillingDoc.BILLING_DOC_ID = -1
      m_BillingDoc.DOCUMENT_NO = mcolParam("DOCUMENT_NO_PO")
      m_BillingDoc.DOCUMENT_TYPE = PO_DOCTYPE
      m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
      m_BillingDoc.APAR_IND = 1
      m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
      m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
      m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
      m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
'      m_BillingDoc.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
      Call m_BillingDoc.QueryData(176, Rs, iCount)
         
      I = 0

      While Not Rs.EOF
         I = I + 1
         Call m_BillingDoc.PopulateFromRS2(176, Rs)
         
         If PrevKey1 <> Trim(m_BillingDoc.DOCUMENT_NO) And I > 1 Then
            If mcolParam("SHOW_PO_DETAIL") = 0 Then
               Call GenerateFooter2(Str(I - 1), PrevKey1, PrevKey2, PrevKey3, "", Total1, m_Details1, 0)
            Else
              Call GenerateFooter("", "������觫���", PrevKey1, Total1, m_Details1, 0)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            
         PrevKey4 = ""
            If mcolParam("SHOW_INV_DETAIL") = 1 Then
               Call GenerateFooterInvoid(PrevKey1, m_Details1, I - 1)
            End If
         End If
         

         
         j = 0
         Call m_Details1.ClearField
         
         If PrevKey1 <> Trim(m_BillingDoc.DOCUMENT_NO) Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_BillingDoc.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE) 'm_BillingDoc.STOCK_DESC & " ( " & m_BillingDoc.STOCK_NO & " ) "
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
          j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_BillingDoc.APAR_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Else
               I = I - 1
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = ""
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
         End If
         
         PrevKey1 = Trim(m_BillingDoc.DOCUMENT_NO)
         PrevKey2 = Trim(m_BillingDoc.DOCUMENT_DATE)
         PrevKey3 = Trim(m_BillingDoc.APAR_NAME)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_BillingDoc.STOCK_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumber(m_BillingDoc.AVG_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + m_BillingDoc.AVG_PRICE
         Total2(j) = Total2(j) + m_BillingDoc.AVG_PRICE
         
         
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumber(m_BillingDoc.ITEM_AMOUNT) 'MyDiff(m_BillingDoc.TOTAL_AMOUNT, m_BillingDoc.UNIT_MULTIPLE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + m_BillingDoc.ITEM_AMOUNT
         Total2(j) = Total2(j) + m_BillingDoc.ITEM_AMOUNT
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumber(m_BillingDoc.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + m_BillingDoc.TOTAL_PRICE
         Total2(j) = Total2(j) + m_BillingDoc.TOTAL_PRICE
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_BillingDoc.UNIT_TRAN_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("SHOW_PO_DETAIL") = 1 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         Rs.MoveNext
      Wend
      
            If mcolParam("SHOW_PO_DETAIL") = 0 Then
               Call GenerateFooter2(Str(I - 1), PrevKey1, PrevKey2, PrevKey3, "", Total1, m_Details1, 0)
            Else
              Call GenerateFooter("", "���", PrevKey1, Total1, m_Details1, 0)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            If mcolParam("SHOW_INV_DETAIL") = 1 Then
               Call GenerateFooterInvoid(PrevKey1, m_Details1, I - 1)
            End If

'       Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
'
      Call GenerateFooter("�ط��", "���觫��ͷ�����", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
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

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 6 Then
         TempStr = ""
      Else
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(txt1 As String, txt2 As String, Txt3 As String, Txt4 As String, Txt5 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 4 Then
         TempStr = Txt4
      ElseIf j = 5 Then
         TempStr = Txt5
       ElseIf j = 6 Then
         TempStr = ""
      Else
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
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
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
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
Private Sub GenerateEndDoc()
''   If mcolParam("INCLUDE_FREE") > 0 Then
''      Vsp.Paragraph = "*�����¡�âͧ��"
''   End If
End Sub
Private Sub GenerateFooterInvoid(Key As String, Details As CFieldList, iCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim I As Long
Dim j As Long
Dim k As Long
Dim TempStr As String
Dim StartCollumn As Long
Dim StartRow As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
'Dim Details As CFieldList
Dim TempBilling As CBillingDoc
Dim TempBillingItem As CBillingDoc
Dim PrevKey1 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
      
   j = 0
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Set Details = New CFieldList
   Set BodyCf = New CReportField
   
       Set TempBilling = GetObject("CBillingDoc", m_InvoidByPo, Trim(Key))
       If Not TempBilling Is Nothing Then
               For Each TempBillingItem In TempBilling.collBillSub
                        j = 0
                        Call m_Details1.ClearField
                         
                        If PrevKey1 <> Trim(TempBillingItem.DOCUMENT_NO) Then
                     
                          I = I + 1
'                           If I > 1 Then
'                              Call GenerateFooter("", "���", PrevKey1, Total1, m_Details1, 0)
'                              Call m_Details1.GetString(1, TempStr1, TempStr2)
'                              strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'                              Vsp.FontBold = True
'                              Call Vsp.AddTable(strFormat, "", TempStr2)
'                              Vsp.FontBold = False
'                           End If
'
'                           For j = 1 To UBound(Total1)
'                               Total1(j) = 0
'                           Next j
                           
                          Call m_Details1.ClearField
                          j = 0
                          
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = iCount & "." & I
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                  
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = TempBillingItem.DOCUMENT_NO
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                           
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = DateToStringExtEx2(TempBillingItem.DOCUMENT_DATE)
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                           
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = TempBillingItem.APAR_NAME
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                        Else
                             j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = ""
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                  
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = ""
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                           
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = ""
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                           
                           j = j + 1
                           Set HeadCf = m_Headers1.Fields(j)
                           TempStr = ""
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                           Call m_Details1.AddField(BodyCf)
                        End If
                        
                        PrevKey1 = Trim(TempBillingItem.DOCUMENT_NO)
                        
                        j = j + 1
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = TempBillingItem.STOCK_DESC
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        
                        j = j + 1
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = FormatNumber(TempBillingItem.AVG_PRICE)
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        Total1(j) = Total1(j) + TempBillingItem.AVG_PRICE
                        Total2(j) = Total2(j) + TempBillingItem.AVG_PRICE
                        
                        j = j + 1
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = FormatNumber(TempBillingItem.ITEM_AMOUNT) 'MyDiff(TempBillingItem.TOTAL_AMOUNT, TempBillingItem.UNIT_MULTIPLE)
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        Total1(j) = Total1(j) + TempBillingItem.ITEM_AMOUNT
                        Total2(j) = Total2(j) + TempBillingItem.ITEM_AMOUNT
                        
                        j = j + 1
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = FormatNumber(TempBillingItem.TOTAL_PRICE)
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        Total1(j) = Total1(j) + TempBillingItem.TOTAL_PRICE
                        Total2(j) = Total2(j) + TempBillingItem.TOTAL_PRICE
                        
                        j = j + 1
                        Set HeadCf = m_Headers1.Fields(j)
                        TempStr = TempBillingItem.UNIT_TRAN_NAME
                        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                        Call m_Details1.AddField(BodyCf)
                        
               
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     
               Next TempBillingItem
               
'                Call GenerateFooter("", "���", PrevKey1, Total1, m_Details1, 0)
'               Call m_Details1.GetString(1, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
               
               Call GenerateFooter("", "�����Ţ��", "", Total2, m_Details1, 0)
                Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If

End Sub