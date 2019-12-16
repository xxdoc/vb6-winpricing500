Attribute VB_Name = "modLoadBalance"
Option Explicit
Public Sub LoadLeftAmountLocation(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long, Optional FromStockNo As String, Optional ToStockNo As String, Optional PartItem As Long, Optional OUTLAY_FLAG As Long = 1, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   MasterInd = "31"
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationId
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   D.PART_ITEM_ID = PartItem
   D.LOCATION_GROUP_NO = LocationGroupNo
   If OUTLAY_FLAG = 0 Then
      D.OUTLAY_FLAG = "N"
   End If
   Call D.QueryData(32, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(32, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   MasterInd = "1"
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadLeftAmountLotItem(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   MasterInd = "20"
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationId
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   D.COUNT_AMOUNT = ""
   Call D.QueryData(20, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(20, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   MasterInd = "1"
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Sub LoadLeftAmount(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   MasterInd = "30"
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationId
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   D.COUNT_AMOUNT = ""
   D.LOCATION_GROUP_NO = LocationGroupNo
   Call D.QueryData(30, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(30, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   MasterInd = "1"
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Sub
Public Function LoadCheckBalance(CompareUseAmount As Double, LocationId As Long, PartItemID As Long, PartNo As String, Optional ExCludeLotID As Long) As Boolean
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim I As Long
   
   LoadCheckBalance = False
   MasterInd = "29"
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationId
   D.PART_ITEM_ID = PartItemID
   D.COUNT_AMOUNT = ""
   D.LOT_ITEM_ID = ExCludeLotID
   Call D.QueryData(29, Rs, itemcount, False)
   
   Call D.PopulateFromRS(29, Rs)
         
   If D.SUM_AMOUNT >= CompareUseAmount Then
      LoadCheckBalance = True
   Else
      glbErrorLog.LocalErrorMsg = "มียอด " & PartNo & " ไม่เพียงพอสำหรับเบิก   ( มียอดคงเหลือเพียง " & D.SUM_AMOUNT & " )"
      glbErrorLog.ShowUserError
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   MasterInd = "1"
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   MasterInd = "1"
End Function
Public Sub LoadCapitalFromLotItem(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationId As Long, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationGroupNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationId
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   D.LOCATION_GROUP_NO = LocationGroupNo
   Call D.QueryData(33, Rs, itemcount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(33, Rs)
         
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CLotItem", Cl, Trim(Str(TempData.PART_ITEM_ID)), False)
         If D Is Nothing Then
            Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

