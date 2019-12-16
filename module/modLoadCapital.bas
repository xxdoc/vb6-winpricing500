Attribute VB_Name = "modLoadCapital"
Option Explicit
Public Sub LoadCapitalMovementLocation(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PartItem As Long = -1, Optional LocationId As Long, Optional FromStockNo As String, Optional ToStockNo As String, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_ITEM_ID = PartItem
   D.LOCATION_ID = LocationId
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID))
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
Public Sub LoadCapitalMovementLocationDocDate(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PartItem As Long = -1, Optional LocationId As Long)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_ITEM_ID = PartItem
   D.LOCATION_ID = LocationId
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
'         If TempData.PART_ITEM_ID = 913 And TempData.LOCATION_ID = 258 And TempData.DOCUMENT_DATE = "30/06/2557" Then
'            Debug.Print
'         End If
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE))
         'Debug.Print Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description & " " & Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE)
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
