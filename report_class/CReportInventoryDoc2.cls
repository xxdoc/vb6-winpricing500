VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryDoc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryDoc2"
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

Private InventoryBals2 As Collection
Private DistinctPartItem As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set Rs = New ADODB.Recordset
   Set InventoryBals2 = New Collection
   Set DistinctPartItem = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set InventoryBals2 = Nothing
   Set DistinctPartItem = Nothing
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
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(7, "^", "��������´�͡���", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�Ѻ���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�����͡", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�������", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "***", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ����͡���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�Ţ����͡���", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ѹ������ҧ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ������", ">")
   Call m_Headers1.AddField(Cf)
   Set Cf = Nothing
   
  Call m_Headers0.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
   
   FromDate = " �ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = " �֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & " " & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & FromDate & ToDate & vbCrLf & _
                                       " �ҡ�Թ��� " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " �֧�Թ��� " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " ��ѧ " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
                                          
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
Dim NewDate As Date

Dim Li As CLotItem
Dim Ba As CLotItem

Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String

Dim SumNewAmount As Double
Dim SumNewPrice As Double
Dim InFlag As Boolean

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'
Dim LocationId As Long
Dim LocationName As String
Call getLocationId(mcolParam("LOCATION_NO"), LocationId, LocationName)
Call mcolParam.add(LocationId, "LOCATION_ID")
Call mcolParam.Remove("LOCATION_NAME")
Call mcolParam.add(LocationName, "LOCATION_NAME")

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

   Set Li = New CLotItem
   
   Set InventoryBals2 = New Collection
   NewDate = DateAdd("D", -1, mcolParam("FROM_BILL_DATE"))
   If mcolParam("FROM_BILL_DATE") > 0 Then
      Call LoadLeftAmount(InventoryBals2, -1, NewDate, mcolParam("LOCATION_ID"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   End If
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   '=== ��¡��
   MasterInd = "5"
   Set Li = New CLotItem
   MasterInd = "1"
   Li.LOT_ITEM_ID = -1
   Li.FROM_DOC_DATE = mcolParam("FROM_BILL_DATE")
   Li.TO_DOC_DATE = mcolParam("TO_BILL_DATE")
   Li.LOCATION_ID = mcolParam("LOCATION_ID")
   Li.ORDER_BY = mcolParam("ORDER_BY")
   Li.ORDER_TYPE = mcolParam("ORDER_TYPE")
   Li.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   Li.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   If mcolParam("SHOW_OUTLAY") = 0 Then
      Li.OUTLAY_FLAG = "N"
   End If
   Call Li.QueryData(5, Rs, iCount, False)
   
   While (Not Rs.EOF)
      
      Call Li.PopulateFromRS(5, Rs)
      
      If PrevKey1 <> Li.PART_TYPE_CODE And I > 0 Then
         Total2(10) = Total2(10) + Total1(10)
         Total3(10) = Total3(10) + Total1(10)
         Total2(12) = Total2(12) + Total1(12)
         Total3(12) = Total3(12) + Total1(12)
         
         Call GenerateFooter("���", PrevKey2, "", Total1, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
            
         Call GenerateFooter("���", PrevKey1, "", Total2, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         For j = 1 To UBound(Total1)    ' TOTAL1 FOR  PARTITEM
            Total1(j) = 0
            Total2(j) = 0
         Next j
             
         Call m_Details1.ClearField
         Call BodyCf.SetFieldValue(7, "^", Li.PART_TYPE_NAME & " (" & Li.PART_TYPE_CODE & ")")
         Call m_Details1.AddField(BodyCf)
         Call BodyCf.SetFieldValue(22, "^", Li.PART_DESC & " (" & Li.PART_NO & ")" & "    ˹��¹Ѻ�� " & Li.UNIT_CHANGE_NAME & " �鹷ع STD =  " & Li.CHK_STD_COST)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
            
         Call GenerateBalance(Li.PART_ITEM_ID, Total1, Total2)
         
      ElseIf PrevKey2 <> Li.PART_NO And I > 0 Then
         Total2(10) = Total2(10) + Total1(10)
         Total3(10) = Total3(10) + Total1(10)
         Total2(12) = Total2(12) + Total1(12)
         Total3(12) = Total3(12) + Total1(12)
         Call GenerateFooter("���", PrevKey2, "", Total1, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         For j = 1 To UBound(Total1)    ' TOTAL1 FOR  PARTITEM
            Total1(j) = 0
          Next j
          
         Call m_Details1.ClearField
         Call BodyCf.SetFieldValue(7, "^", Li.PART_TYPE_NAME & " (" & Li.PART_TYPE_CODE & ")")
         Call m_Details1.AddField(BodyCf)
         Call BodyCf.SetFieldValue(22, "^", Li.PART_DESC & " (" & Li.PART_NO & ")" & "    ˹��¹Ѻ�� " & Li.UNIT_CHANGE_NAME & " �鹷ع STD =  " & Li.CHK_STD_COST)
         Call m_Details1.AddField(BodyCf)

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
            
         Call GenerateBalance(Li.PART_ITEM_ID, Total1, Total2)
      ElseIf I = 0 Then
         Call m_Details1.ClearField
         Call BodyCf.SetFieldValue(7, "^", Li.PART_TYPE_NAME & " (" & Li.PART_TYPE_CODE & ")")
         Call m_Details1.AddField(BodyCf)
         Call BodyCf.SetFieldValue(22, "^", Li.PART_DESC & " (" & Li.PART_NO & ")" & "    ˹��¹Ѻ�� " & Li.UNIT_CHANGE_NAME & " �鹷ع STD =  " & Li.CHK_STD_COST)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Call GenerateBalance(Li.PART_ITEM_ID, Total1, Total2)
      End If
      
      InFlag = False
      
      PrevKey1 = Li.PART_TYPE_CODE
      PrevKey2 = Li.PART_NO
         
      Call m_Details1.ClearField
      I = I + 1
         
      For j = 1 To m_Headers1.Fields.Count
         Set HeadCf = m_Headers1.Fields(j)
            
         If j = 1 Then
            TempStr = I
         ElseIf j = 2 Then
            TempStr = DateToStringExtEx2(Li.DOCUMENT_DATE)
         ElseIf j = 3 Then
            TempStr = Li.DOCUMENT_NO
         ElseIf j = 4 Then
            If Li.TX_TYPE = "I" Then
               Amt = Li.TX_AMOUNT
               Total1(10) = Total1(10) + Li.TX_AMOUNT
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            
            If Amt > 0 Then
               InFlag = True
            End If
         ElseIf j = 5 Then
            If Li.TX_TYPE = "I" Then
               Amt = Li.AVG_PRICE
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            
            If Amt > 0 Then
               InFlag = True
            End If
         ElseIf j = 6 Then
            If Li.TX_TYPE = "I" Then
               Amt = Li.TOTAL_INCLUDE_PRICE
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            
         ElseIf j = 7 Then
            If Li.TX_TYPE = "E" Then
               Amt = Li.TX_AMOUNT
               Total1(10) = Total1(10) - Li.TX_AMOUNT
            Else
               Amt = 0
            End If
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            
            If Amt > 0 Then
               InFlag = True
            End If
         ElseIf j = 8 Then
            If Li.TX_TYPE = "E" Then
               Amt = Li.AVG_PRICE
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            If Amt > 0 Then
               InFlag = True
            End If
         ElseIf j = 9 Then
            If Li.TX_TYPE = "E" Then
               Amt = Li.TX_AMOUNT * Li.AVG_PRICE
            Else
               Amt = 0
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt
            If Amt > 0 Then
               InFlag = True
            End If
         ElseIf j = 10 Then
            Amt = Total1(10) 'Li.NEW_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Amt
            SumNewAmount = Amt
            
         ElseIf j = 11 Then
            Amt = Li.NEW_AVG_PRICE
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Amt
         ElseIf j = 12 Then
            Amt = Total1(10) * Li.NEW_AVG_PRICE
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Amt
            SumNewPrice = Amt
         ElseIf j = 13 Then
            TempStr = DateToStringExtEx2(Li.CREATE_DATE)
         ElseIf j = 14 Then
            TempStr = DateToStringExtEx2(Li.MODIFY_DATE)
         Else
            TempStr = ""
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
      Next j
      
      If InFlag Then
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      
      Rs.MoveNext
   Wend
   '=== ��¡��
   
   Total2(10) = Total2(10) + Total1(10)
   Total3(10) = Total3(10) + Total1(10)
   Total2(12) = Total2(12) + Total1(12)
   Total3(12) = Total3(12) + Total1(12)
   
   Call GenerateFooter("���", PrevKey2, "", Total1, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("���", PrevKey1, "", Total2, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   Call GenerateFooter("�ط��", "", "", Total3, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set InventoryBals2 = Nothing
   Set Li = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
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
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
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
Public Sub GenerateBalance(Item As Long, Total1() As Double, Total2() As Double)
Dim j As Long
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr As String
Dim Ba As CLotItem
Dim Amt As Double
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim strFormat As String
      
      Set BodyCf = New CReportField
      
      '=== �ʹ¡��
      Call m_Details1.ClearField
      j = 0
      For Each HeadCf In m_Headers1.Fields
         j = j + 1
         If j = 2 Then
            TempStr = "¡��"
         ElseIf j = 10 Then
            Set Ba = GetObject("CLotItem", InventoryBals2, Trim(Str(Item)))
            If mcolParam("FROM_BILL_DATE") > 0 Then
               Amt = Ba.SUM_AMOUNT
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
         ElseIf j = 11 Then
            If mcolParam("FROM_BILL_DATE") > 0 Then
               Amt = 0 'BA.AVG_PRICE
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
         ElseIf j = 12 Then
            If mcolParam("FROM_BILL_DATE") > 0 Then
               Amt = 0 'BA.TOTAL_INCLUDE_PRICE
            End If
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
         Else
            TempStr = ""
         End If

         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
      Next HeadCf

      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      '=== �ʹ¡��
      
End Sub

