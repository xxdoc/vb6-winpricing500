VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportMN03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportMaster001"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

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
Private m_BranchColl As Collection
Private CollApArAddess As Collection

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
   Set m_BranchColl = New Collection
   Set CollApArAddess = New Collection
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
   Set m_BranchColl = Nothing
   Set CollApArAddess = Nothing
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
Dim ImportItems As Collection
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
   
   Call Cf.SetFieldValue(1, "^", "NO", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "���ʼ����", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "���ͼ����", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "�������", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(5, "^", "�����͡���", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "�����������", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(7, "^", "�Ң�", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(8, "^", "����������", "<")
   Call m_Headers1.AddField(Cf)
   
   If Val(mcolParam("SHOW_TAX")) = 1 Then
         Call Cf.SetFieldValue(4, "^", "�Ţ�������������", "<")
        Call m_Headers1.AddField(Cf)
   End If
   
   If Val(mcolParam("SHOW_ADDRESS")) = 1 Then
         Call Cf.SetFieldValue(10, "^", "�������", "<")
        Call m_Headers1.AddField(Cf)
   End If
   
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
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
  Call LoadApArAddress(CollApArAddess, 2)
   
   MODULE_DESC = glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & " " & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME")
                                       
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
Dim m_Customer As CAPARMas
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Mr As CMasterRef
Dim ck As Boolean
Dim TempApar As CAPARMas

   
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
      
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      Set BodyCf = New CReportField
      
      MasterInd = "3"
      Set m_Customer = New CAPARMas
      MasterInd = "1"
      
      m_Customer.APAR_MAS_ID = -1
      m_Customer.APAR_IND = 2
      m_Customer.APAR_CODE = PatchWildCard(mcolParam("APAR_CODE"))
      m_Customer.APAR_NAME = PatchWildCard(mcolParam("APAR_NAME"))
      m_Customer.APAR_GRADE = mcolParam("MASTER_SUPGRADE")
      m_Customer.APAR_TYPE = mcolParam("MASTER_SUPTYPE")
      m_Customer.ORDER_BY = mcolParam("ORDER_BY")
      m_Customer.ORDER_TYPE = mcolParam("ORDER_TYPE")
      
      
      
      Call m_Customer.QueryData(1, Rs, iCount, False)
      
      Call LoadMaster(Nothing, m_BranchColl, , , MASTER_APARMAS_BRANCH)
      I = 0

      While Not Rs.EOF
         I = I + 1
         j = 0
         ck = False
         Call m_Customer.PopulateFromRS(1, Rs)
         
         Call m_Details1.ClearField
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.APAR_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.APAR_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.APAR_SHORT_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.BILL_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.APAR_TYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
       
        For Each Mr In m_BranchColl
            If Mr.PARENT_EX_ID2 = m_Customer.APAR_MAS_ID Then
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = Mr.KEY_CODE & " " & Mr.KEY_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               ck = True
               Exit For
            End If
        Next Mr
       If ck = False Then
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
          TempStr = ""
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
       End If
       
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = m_Customer.APAR_MAS_GROUP_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If Val(mcolParam("SHOW_TAX")) = 1 Then
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = m_Customer.TAX_ID
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         If Val(mcolParam("SHOW_ADDRESS")) = 1 Then
            
            Set TempApar = GetObject("CAPARMas", CollApArAddess, Trim(Str(m_Customer.APAR_MAS_ID)))
           
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = TempApar.PackAddress()
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
   
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend
       
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
         TempStr = FormatNumberToNull(Tot(j))
      ElseIf j = 4 Then
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
   Vsp = ""
   
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

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub