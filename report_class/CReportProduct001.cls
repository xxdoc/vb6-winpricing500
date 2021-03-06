VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportJobItemSumItem"
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
Private m_DistinctInput As Collection
Private m_SumInput As Collection
Private m_SumInputOutPutLost As Collection

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
   Set m_DistinctInput = New Collection
   Set m_SumInput = New Collection
   Set m_SumInputOutPutLost = New Collection
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
   Set m_DistinctInput = Nothing
   Set m_SumInput = Nothing
   Set m_SumInputOutPutLost = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim Ji As CJobItem

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(2, "^", "�����Թ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�����Թ���", "<")
   Call m_Headers1.AddField(Cf)
   
   For Each Ji In m_DistinctInput
      Call Cf.SetFieldValue(2, "^", Ji.PART_NO & " " & Ji.PART_DESC, ">")
      Call m_Headers1.AddField(Cf)
   Next Ji
   
   Call Cf.SetFieldValue(2, "^", "���", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
    FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = mcolParam("REPORT_NAME") & " " & FromDate & " " & ToDate & " �ҡ�ѵ�شԺ " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " �֧ " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " ʶҹ����Ե " & EmptyToString(mcolParam("PRODUCTION_LOCATION_NAME"), "������") & " BATCH : " & EmptyToString(mcolParam("BATCH_NO_SET"), "������")
   
   Call GetDistinctJobInItem(m_DistinctInput, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), GetBatchNoSet())
   Call GetSumJobInItem(m_SumInput, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), GetBatchNoSet())
   Call GetSumJobInOutLostItem(m_SumInputOutPutLost, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), GetBatchNoSet())
                                                
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
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_JobItem As CJobItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Ji  As CJobItem
Dim TempJi  As CJobItem
Dim HaveFlag As Boolean
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

   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Set m_JobItem = New CJobItem
      
      m_JobItem.JOB_ID = -1
      m_JobItem.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_JobItem.TO_DATE = mcolParam("TO_BILL_DATE")
      m_JobItem.TX_TYPE = "E"
      m_JobItem.PRD_LOCATION_ID = mcolParam("PRODUCTION_LOCATION")
      m_JobItem.BATCH_NO_SET = GetBatchNoSet
      Call m_JobItem.QueryData(2, Rs, iCount)
      
      If mcolParam("PRINT_TO_FILE") = 1 Then
         Call printHeader
      End If
      
      I = 0
      
      Call m_Details1.ClearField
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      k = 2
      SumRow = 0
      HaveFlag = False
      For Each Ji In m_DistinctInput
         k = k + 1
         Set TempJi = GetObject("CJobItem", m_SumInput, Trim(Str(Ji.PART_ITEM_ID)), True)
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TempJi.TX_AMOUNT, 0)
         SumRow = SumRow + TempJi.TX_AMOUNT
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         If TempJi.TX_AMOUNT <> 0 Then
            HaveFlag = True
         End If
      Next
      
      k = k + 1
      Set HeadCf = m_Headers1.Fields(k)
      TempStr = FormatNumberToNull(SumRow, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      If HaveFlag Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      While Not Rs.EOF
         I = I + 1
         Call m_Details1.ClearField
         Call m_JobItem.PopulateFromRS(2, Rs)
         
         Set HeadCf = m_Headers1.Fields(1)
         If m_JobItem.PART_ITEM_ID > 0 Then
            TempStr = m_JobItem.PART_NO
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         If m_JobItem.PART_ITEM_ID > 0 Then
            TempStr = m_JobItem.PART_DESC
         Else
            TempStr = m_JobItem.PROBLEM_DESC
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = 2
         SumRow = 0
         HaveFlag = False
         For Each Ji In m_DistinctInput
            k = k + 1
            Set TempJi = GetObject("CJobItem", m_SumInputOutPutLost, Trim(m_JobItem.BATCH_NO & "-" & Ji.PART_ITEM_ID & "-" & m_JobItem.PART_ITEM_ID & "-" & m_JobItem.LOST_ID), True)
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(TempJi.TX_AMOUNT, 0)
            SumRow = SumRow + TempJi.TX_AMOUNT
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If TempJi.TX_AMOUNT <> 0 Then
               HaveFlag = True
            End If
         Next
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(SumRow, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If HaveFlag Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         Rs.MoveNext
      Wend
      
      Vsp.TableBorder = tbNone
      Call PatchTable(Vsp, strFormat, ";", 1000) '10
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "................................................|................................................;"
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "���Ѵ��|����Ǩ�ͺ;"
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
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
      ElseIf j = 5 Then
         TempStr = FormatNumberToNull(Tot(j))
      ElseIf j = 6 Then
         TempStr = ""
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
Dim strFormat   As String
Dim strPrint  As String
Dim alngX() As Long
Dim tbdt As TableBorderSettings
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
   'Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.GetFieldValue("TAX_ID"))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taLeftMiddle
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
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
Private Function GetBatchNoSet()
Dim TempStr As String
   GetBatchNoSet = ""
   If Len(mcolParam("BATCH_NO_SET")) > 0 Then
      TempStr = mcolParam("BATCH_NO_SET")
      GetBatchNoSet = "("
      While Len(TempStr) > 0
         If InStr(1, TempStr, ",") > 0 Then
            GetBatchNoSet = GetBatchNoSet & Left(TempStr, InStr(1, TempStr, ",") - 1) & ","
            TempStr = Mid(TempStr, InStr(1, TempStr, ",") + 1, Len(TempStr) - InStr(1, TempStr, ","))
         Else
            GetBatchNoSet = GetBatchNoSet & TempStr & ")"
            TempStr = ""
         End If
      Wend
   End If
End Function
