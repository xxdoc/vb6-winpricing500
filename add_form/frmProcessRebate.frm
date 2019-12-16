VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessRebate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmProcessRebate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   5371
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   1755
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   1560
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Xivess.uctlTextBox txtYear 
         Height          =   435
         Left            =   5160
         TabIndex        =   1
         Top             =   960
         Width           =   1095
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   4
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessRebate.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   11
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblMonth 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   5
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcessRebate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private CollSaleAmounts As Collection
Private SaleChartColl  As Collection
Private OrderSaleChartColl  As Collection
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitThaiMonth(cboMonth)
      
   End If
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If

End Sub

Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitNormalLabel(lblMonth, "เดือน")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblYear, "ปี")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   Call txtYear.SetTextLenType(TEXT_INTEGER, glbSetting.YEAR_TYPE)
   
   Call InitCombo(cboMonth)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
  ' cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
  ' Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
      
   'Set m_Products = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim PartItemID As Long
Dim FromDate As Date
Dim ToDate As Date
   
   If Not VerifyCombo(lblMonth, cboMonth, False) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblYear, txtYear, False) Then
      Exit Sub
   End If
   
   FromDate = DateSerial(Val(txtYear.Text) - 543, Val(cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex))), 1)
   Call GetFirstLastDate(FromDate, FromDate, ToDate)
   
   Call glbDaily.StartTransaction
   
   Me.Enabled = False
   
   Status = UpdateEmpDealerType(FromDate, ToDate, Me.prgProgress, Me.txtPercent)
   
   Me.Enabled = True
   
   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   Call cmdOK_Click
   Exit Sub
   
End Sub

Public Function UpdateEmpDealerType(Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional Parent As Object = Nothing, Optional ParentEx As Object = Nothing) As Boolean
On Error GoTo ErrorHandler
Dim m_Rs As ADODB.Recordset
Dim Emp As CEmployee
Dim CollEmpDealerType As Collection
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim PERCENT As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean
Dim TempEmpDealer As CEmployeeDealer
Dim TempBd As CBillingDoc
Dim Amt As Double
Dim m_SaleChart As CSaleChart
Dim TotalSale As CTotalSale
Dim SumTotalChart  As Collection
   
   Set CollSaleAmounts = New Collection
   Set CollEmpDealerType = New Collection
   Set OrderSaleChartColl = New Collection
   Set SumTotalChart = New Collection
   
   If Not (Parent Is Nothing) Then
      Parent.Max = 100
      Parent.Min = 0
   End If
      
   Set m_SaleChart = New CSaleChart
   Set TotalSale = New CTotalSale
   Set m_Rs = New ADODB.Recordset
   Set SaleChartColl = New Collection
   
   m_SaleChart.SALE_CHART_ID = -1
   m_SaleChart.FROM_DATE = FromDate
   m_SaleChart.TO_DATE = ToDate
   Call m_SaleChart.QueryData(3, m_Rs, iCount)

   If iCount <= 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่พบข้อมูลการตั้งแผนภูมิ"
      glbErrorLog.ShowUserError
      Exit Function
   End If

   While Not m_Rs.EOF
      Set m_SaleChart = New CSaleChart
      Call m_SaleChart.PopulateFromRS(3, m_Rs)
      Call SaleChartColl.add(m_SaleChart, Trim(Str(m_SaleChart.SALE_CHART_ID)))
      m_Rs.MoveNext
   Wend
   
   ' Query Employee เฉพาะที่มีการเซ็ตประเภท Dealer_type ไว้ เท่านั้น
   ' Query ประเภทขอมูล Dealer_Type ตามเดือนที่รัน -1
   Call GetEmpDealerTypeYYYYMM(CollEmpDealerType, Year(DateAdd("m", -1, FromDate)) & Format(Month(DateAdd("m", -1, FromDate)), "00"))
   
   Call GetSaleAmountDealerDocType(CollSaleAmounts, FromDate, ToDate)           'Query ยอดขายของเดือนที่รัน
   
   HasBegin = True
   
   Set TempEmpDealer = New CEmployeeDealer
   TempEmpDealer.YYYYMM = Year(FromDate) & Format(Month(FromDate), "00")
   Call TempEmpDealer.DeleteDataYYYYMM
         
   Call GenerateOrderSaleChart(SaleChartColl, -1, 0)
   
   Call SumChart(OrderSaleChartColl, SumTotalChart)
      
   I = 0
   
   For Each m_SaleChart In OrderSaleChartColl
      I = I + 1
      If Not (Parent Is Nothing) Then
         Parent.Value = 25 + MyDiff(I, OrderSaleChartColl.Count) * 75
         ParentEx.Text = FormatNumber(Parent.Value)
         Parent.Refresh
         ParentEx.Refresh
         DoEvents
      End If
      
      'คำนวณยอดของตัองและลูกข่าย 2 ชั้น
      
      Set TempEmpDealer = GetObject("CEmployeeDealer", CollEmpDealerType, Trim(Str(m_SaleChart.EMP_ID)), False)
      ' เช็คเดือนที่แล้ว
      If TempEmpDealer Is Nothing Then
         Set TempEmpDealer = New CEmployeeDealer
         TempEmpDealer.DEALER_TYPE = SILVER 'ถ้ายังไม่ได้เซ็ตตั้งเป็นค่าเริ่มต้นที่ SILVER
      End If
      ' ยอดซื้อและประเภทตัวแทนไปคำนวณ
      Set TotalSale = GetObject("CTotalSale", SumTotalChart, Trim(Str(m_SaleChart.EMP_ID)))
      Call CheckNewEmpDealerStatus(TotalSale.TOTAL_PRICE, TempEmpDealer)
      
      TempEmpDealer.AddEditMode = SHOW_ADD
      TempEmpDealer.EMP_ID = m_SaleChart.EMP_ID
      TempEmpDealer.YYYYMM = Year(FromDate) & Format(Month(FromDate), "00")
      Call TempEmpDealer.AddEditData
      
      Set Emp = New CEmployee
      Emp.ShowMode = SHOW_EDIT
      Emp.EMP_ID = m_SaleChart.EMP_ID
      Emp.DEALER_TYPE = TempEmpDealer.DEALER_TYPE
      Call Emp.UpdateDealerType
      
      Set Emp = Nothing
      
   Next m_SaleChart
   
   If Not (Parent Is Nothing) Then
      Parent.Value = 100
      ParentEx.Text = 100
   End If
   
   Set Emp = Nothing
   
   Set CollSaleAmounts = Nothing
   Set CollEmpDealerType = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   UpdateEmpDealerType = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
   End If
   
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.RoutineName = "UpdateCapitalMovement"
   glbErrorLog.ModuleName = "FrmProcessCommit"
   glbErrorLog.LocalErrorMsg = "Eror"
   glbErrorLog.ShowErrorLog (LOG_MSGBOX)
   
   Set Emp = Nothing
   
   Set CollSaleAmounts = Nothing
   Set CollEmpDealerType = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   UpdateEmpDealerType = False
End Function
Private Sub CheckNewEmpDealerStatus(Amt As Double, TempEmpDealer As CEmployeeDealer)
Dim NewDealerType As DEALER_TYPE_AREA
   If TempEmpDealer.DEALER_TYPE = HEADER_GROUP Then
      Exit Sub
   End If
   NewDealerType = TempEmpDealer.DEALER_TYPE
   If TempEmpDealer.DEALER_TYPE = SILVER Then
      If Amt >= 35000 Then
         NewDealerType = SILVER_PLUS
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = SILVER_PLUS Then
      If Amt >= 35000 Then
         NewDealerType = SILVER_PLUS_PLUS
      Else
         NewDealerType = SILVER
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = SILVER_PLUS_PLUS Then
      If Amt >= 35000 Then
         NewDealerType = GOLD
      Else
         NewDealerType = SILVER
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = GOLD_MUNUS Then
      If Amt >= 35000 Then
         NewDealerType = GOLD
      Else
         NewDealerType = SILVER
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = GOLD Then
      If Amt >= 100000 Then
         NewDealerType = GOLD_PLUS
      ElseIf Amt >= 35000 Then
         NewDealerType = GOLD
      Else
         NewDealerType = GOLD_MUNUS
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = GOLD_PLUS Then
      If Amt >= 100000 Then
         NewDealerType = GOLD_PLUS_PLUS
      ElseIf Amt >= 35000 Then
         NewDealerType = GOLD
      Else
         NewDealerType = GOLD_MUNUS
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = GOLD_PLUS_PLUS Then
      If Amt >= 100000 Then
         NewDealerType = PLATINUM
      ElseIf Amt >= 35000 Then
         NewDealerType = GOLD
      Else
         NewDealerType = GOLD_MUNUS
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = PLATINUM Then
      If Amt >= 100000 Then
         NewDealerType = PLATINUM
      Else
         NewDealerType = PLATINUM_MUNUS
      End If
   ElseIf TempEmpDealer.DEALER_TYPE = PLATINUM_MUNUS Then
      If Amt >= 100000 Then
         NewDealerType = PLATINUM
      Else
         NewDealerType = GOLD
      End If
   End If
   
   TempEmpDealer.DEALER_TYPE = NewDealerType
End Sub
Private Sub SumChart(Coll As Collection, SumTotal As Collection)
Dim Cm As CSaleChart
   For Each Cm In Coll
      Call Recuresive(Cm, SumTotal, GetParent(Cm.SALE_CHART_ID), GetEmp(Cm.SALE_CHART_ID), 0)
   Next Cm
End Sub
Public Sub Recuresive(Cm As CSaleChart, SumTotal As Collection, ParentID As Long, OwnId As Long, Level As Long)
On Error Resume Next
Dim Amt As Double
Dim P1 As CTotalSale
Dim P2 As CTotalSale
Set P1 = New CTotalSale
Dim Old As Double
Dim TempBd As CBillingDoc
Dim Tg As CTagetDetail

   
   P1.EMP_ID = OwnId
   P1.SALE_NAME = Cm.SALE_NAME & " (" & Cm.SALE_CODE & ")"
   
   Amt = 0
         
   Set TempBd = GetObject("CBillingDoc", CollSaleAmounts, Cm.SALE_CODE & "-" & INVOICE_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   Set TempBd = GetObject("CBillingDoc", CollSaleAmounts, Cm.SALE_CODE & "-" & RECEIPT1_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   Set TempBd = GetObject("CBillingDoc", CollSaleAmounts, Cm.SALE_CODE & "-" & RETURN_DOCTYPE)
   Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)

   
   P1.TOTAL_PRICE = Amt
   If Level = 0 Then
      P1.TOTAL_SELF_PRICE = Amt
   End If
   
   If SumTotal.Count = 0 Then
      Call SumTotal.add(P1, Trim(P1.Getkey))
   Else
      Set P2 = SumTotal(Trim(P1.Getkey))
      If P2 Is Nothing Then
         Call SumTotal.add(P1, Trim(P1.Getkey))
      Else
         P2.TOTAL_PRICE = P2.TOTAL_PRICE + P1.TOTAL_PRICE
         P2.TOTAL_SELF_PRICE = P2.TOTAL_SELF_PRICE + P1.TOTAL_SELF_PRICE
      End If
   End If
   
   If ParentID > 0 And Level < 2 Then
      Call Recuresive(Cm, SumTotal, GetParent(ParentID), GetEmp(ParentID), Level + 1)
   End If
End Sub
Private Function GetParent(ID As Long) As Long
Dim Cm As CSaleChart
   Set Cm = GetObject("CSaleChart", SaleChartColl, Trim(Str(ID)))
   GetParent = Cm.PARENT_ID
End Function
Private Function GetEmp(ID As Long) As Long
Dim Cm As CSaleChart
   Set Cm = GetObject("CSaleChart", SaleChartColl, Trim(Str(ID)))
   GetEmp = Cm.EMP_ID
End Function
Private Sub GenerateOrderSaleChart(TempColl As Collection, PID As Long, Level As Long)
Dim O As CSaleChart

   For Each O In TempColl
      If O.PARENT_ID = PID Then
         O.SALE_NAME = Space(Level * 2) & (Level + 1) & " " & O.SALE_NAME
         O.Level = Level
         Call OrderSaleChartColl.add(O, Trim(Str(O.SALE_CHART_ID)))
         Call GenerateOrderSaleChart(TempColl, O.SALE_CHART_ID, Level + 1)
      End If
   Next O
End Sub

