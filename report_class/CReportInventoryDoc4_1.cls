VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryDoc4_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryDoc4"
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
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList

Private MovementPartType As Collection
Private MovementPartTypeBack As Collection
Private DistinctPartType As Collection
Private MasterID As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set Rs = New ADODB.Recordset
   Set MovementPartType = New Collection
   Set MovementPartTypeBack = New Collection
   Set DistinctPartType = New Collection
   Set MasterID = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set MovementPartType = Nothing
   Set MovementPartTypeBack = Nothing
   Set DistinctPartType = Nothing
   Set MasterID = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
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
Dim TempStr2 As String
Dim Cf As CReportField
Dim TempLt  As CLotItem
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   
   Call Cf.SetFieldValue(6, "^", "��������´", "<")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "�͹�", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�͹��Ѻ", ">", , True)
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "DIFF", ">", , True)
   Call m_Headers1.AddField(Cf)
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
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
Dim FromLocationName As String
Dim FromLocationName2 As String
Dim ToLocationName As String
Dim ToLocationName2 As String
Dim TempMaster As CMasterRef
   
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
   
   FromDate = " �ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = " �֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   Call LoadMasterID(Nothing, MasterID, 14)
   If mcolParam("FROM_LOCATION_ID") <> "" Then
      Set TempMaster = GetObject("CMasterRef", MasterID, Trim(Str(mcolParam("FROM_LOCATION_ID"))))
      FromLocationName = TempMaster.KEY_NAME
   Else
      FromLocationName = "N/A"
   End If
   If mcolParam("FROM_LOCATION_ID2") <> "" Then
      Set TempMaster = GetObject("CMasterRef", MasterID, Trim(Str(mcolParam("FROM_LOCATION_ID2"))))
      FromLocationName2 = TempMaster.KEY_NAME
   Else
      FromLocationName2 = "N/A"
   End If
   If mcolParam("TO_LOCATION_ID") <> "" Then
   Set TempMaster = GetObject("CMasterRef", MasterID, Trim(Str(mcolParam("TO_LOCATION_ID"))))
   ToLocationName = TempMaster.KEY_NAME
   Else
      ToLocationName = "N/A"
   End If
   If mcolParam("TO_LOCATION_ID2") <> "" Then
      Set TempMaster = GetObject("CMasterRef", MasterID, Trim(Str(mcolParam("TO_LOCATION_ID2"))))
      ToLocationName2 = TempMaster.KEY_NAME
   Else
      ToLocationName2 = "N/A"
   End If
   
   MODULE_DESC = glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & " " & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & FromDate & ToDate & vbCrLf & _
                                       " �ҡ�Թ��� " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " �֧�Թ��� " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " �ҡ��ѧ " & FromLocationName & "-" & FromLocationName2 & "   ��ѧ��ѧ " & ToLocationName & "-" & ToLocationName2
                                       
   Call GetDistinctTransferPartItemConsignment(DistinctPartType, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_LOCATION_ID"), mcolParam("FROM_LOCATION_ID2"), mcolParam("TO_LOCATION_ID"), mcolParam("TO_LOCATION_ID2"), TRANSFER_DOCTYPE, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("CONSIGNMENT"), mcolParam("SHOW_OUTLAY"))
   Call GetTransferPartItemDocDateConsignment(MovementPartType, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_LOCATION_ID"), mcolParam("FROM_LOCATION_ID"), mcolParam("TO_LOCATION_ID"), mcolParam("TO_LOCATION_ID2"), TRANSFER_DOCTYPE, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("CONSIGNMENT"), mcolParam("SHOW_OUTLAY"))
   Call GetTransferPartItemDocDateConsignment(MovementPartTypeBack, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("TO_LOCATION_ID"), mcolParam("TO_LOCATION_ID2"), mcolParam("FROM_LOCATION_ID"), mcolParam("FROM_LOCATION_ID2"), TRANSFER_DOCTYPE, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("CONSIGNMENT"), mcolParam("SHOW_OUTLAY"))
   
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
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim Amt2 As Double

Dim TempLi As CLotItem
Dim TempLiMovement As CLotItem
Dim SumRow As Double
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim LFromDate As Date
Dim LToDate As Date

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   '=== ��¡��
   For Each TempLi In DistinctPartType
      Call m_Details1.ClearField
      I = I + 1
      
      If PrevKey1 <> TempLi.PART_NO And I <> 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Call GenerateFooter("��� " & PrevKey1, "", "", Total2, m_Details1)
               Vsp.FontBold = True
            Else
               Call GenerateFooter(" " & PrevKey2 & "/" & PrevKey1, "", "", Total2, m_Details1)
            End If
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total2)
               Total2(j) = 0
            Next j
            
            If mcolParam("SUMMARY") = 0 Then
               Vsp = "�Թ��� : " & TempLi.PART_DESC & "/" & TempLi.PART_NO
            End If
      ElseIf I = 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Vsp = "�Թ��� : " & TempLi.PART_DESC & "/" & TempLi.PART_NO
            End If
      End If
      PrevKey1 = TempLi.PART_NO
      PrevKey2 = TempLi.PART_DESC
      
      Call m_Details1.ClearField
      j = 0
      Amt = 0
      Amt2 = 0
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = "      �ѹ��� " & DateToStringExtEx2(TempLi.DOCUMENT_DATE) & "      �Ţ��� " & TempLi.DOCUMENT_NO
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      Set TempLiMovement = GetObject("CLotItem", MovementPartType, Trim(TempLi.DOCUMENT_DATE & "-" & TempLi.PART_ITEM_ID & "-" & TempLi.DOCUMENT_NO))
      Amt = TempLiMovement.TX_AMOUNT
      Total1(j) = Total1(j) + Amt
      Total2(j) = Total2(j) + Amt
      TempStr = FormatNumberToNull(Amt)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      Set TempLiMovement = GetObject("CLotItem", MovementPartTypeBack, Trim(TempLi.DOCUMENT_DATE & "-" & TempLi.PART_ITEM_ID & "-" & TempLi.DOCUMENT_NO))
      Amt2 = TempLiMovement.TX_AMOUNT
      Total1(j) = Total1(j) + Amt2
      Total2(j) = Total2(j) + Amt2
      TempStr = FormatNumberToNull(Amt2)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      Total1(j) = Total1(j) + Amt - Amt2
      Total2(j) = Total2(j) + Amt - Amt2
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      If mcolParam("SUMMARY") = 0 Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
   Next TempLi
'   '=== ��¡��
   If mcolParam("SUMMARY") = 0 Then
      Call GenerateFooter("��� " & PrevKey1, "", "", Total2, m_Details1)
      Vsp.FontBold = True
   Else
      Call GenerateFooter(" " & PrevKey2 & "/" & PrevKey1, "", "", Total2, m_Details1)
   End If
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   Call GenerateFooter("�ط��", "", "", Total1, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
End Function
Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
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
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
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
