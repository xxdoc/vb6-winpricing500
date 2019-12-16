VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPartItem 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   Icon            =   "frmAddEditPartItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   14340
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtJointCode 
         Height          =   435
         Left            =   12600
         TabIndex        =   41
         Top             =   1020
         Width           =   1575
         _extentx        =   2778
         _extenty        =   767
      End
      Begin VB.ComboBox cboGroupCom 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2880
         Width           =   2205
      End
      Begin VB.ComboBox cboPartTypeSub 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   2205
      End
      Begin VB.ComboBox cboUnitChange 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   1305
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   1275
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   2205
      End
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   4485
         _extentx        =   13309
         _extenty        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2955
         _extentx        =   5212
         _extenty        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3765
         Left            =   150
         TabIndex        =   16
         Top             =   3960
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   6641
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditPartItem.frx":27A2
         Column(2)       =   "frmAddEditPartItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPartItem.frx":290E
         FormatStyle(2)  =   "frmAddEditPartItem.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPartItem.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPartItem.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPartItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPartItem.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   15
         Top             =   3420
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Xivess.uctlTextBox txtBillCode 
         Height          =   435
         Left            =   6600
         TabIndex        =   1
         Top             =   1020
         Width           =   1455
         _extentx        =   5212
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtBillDesc 
         Height          =   435
         Left            =   7800
         TabIndex        =   3
         Top             =   1470
         Width           =   3945
         _extentx        =   5212
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtUnitAmount 
         Height          =   435
         Left            =   3120
         TabIndex        =   9
         Top             =   2400
         Width           =   675
         _extentx        =   1191
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtUnitPerBasket 
         Height          =   435
         Left            =   7800
         TabIndex        =   6
         Top             =   1920
         Width           =   675
         _extentx        =   1191
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtReportPriority 
         Height          =   435
         Left            =   11080
         TabIndex        =   7
         Top             =   1920
         Width           =   675
         _extentx        =   1191
         _extenty        =   767
      End
      Begin Xivess.uctlTextLookup uctlPartItemReturn 
         Height          =   435
         Left            =   1860
         TabIndex        =   14
         Top             =   2835
         Width           =   4425
         _extentx        =   7805
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtCostPerAmount 
         Height          =   435
         Left            =   7800
         TabIndex        =   12
         Top             =   2370
         Width           =   2175
         _extentx        =   3836
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtBarCode 
         Height          =   435
         Left            =   9000
         TabIndex        =   37
         Top             =   1020
         Width           =   1455
         _extentx        =   2566
         _extenty        =   767
      End
      Begin Threed.SSCheck ChkRebateFlag 
         Height          =   375
         Left            =   11760
         TabIndex        =   42
         Top             =   2800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJointCode 
         Caption         =   "Label1"
         Height          =   255
         Left            =   10560
         TabIndex        =   40
         Top             =   1080
         Width           =   1935
      End
      Begin Threed.SSCheck ChkOutlayFlag 
         Height          =   375
         Left            =   11760
         TabIndex        =   39
         Top             =   2400
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblBarCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   8160
         TabIndex        =   38
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblGroupCom 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6360
         TabIndex        =   35
         Top             =   3000
         Width           =   1335
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   34
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   33
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   32
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck ChkExceptionFlag 
         Height          =   435
         Left            =   10200
         TabIndex        =   31
         Top             =   2800
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkLotFlag 
         Height          =   435
         Left            =   10200
         TabIndex        =   13
         Top             =   2400
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck ChkStandardCost 
         Height          =   435
         Left            =   5640
         TabIndex        =   11
         Top             =   2370
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPartItemReturn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   2940
         Width           =   1395
      End
      Begin VB.Label lblReportPriority 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   9720
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblUnitPerBasket 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6360
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   8520
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblBillDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   6390
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblBillCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   345
         Left            =   5160
         TabIndex        =   25
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   23
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   22
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   21
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12600
         TabIndex        =   18
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10920
         TabIndex        =   17
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CStockCode

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public PartGroupID As Long
Private m_Mr As CMasterRef

Private m_PartItemColl As Collection
Private BalanceAmountByLot As Collection
Private Sub cboGroupCom_Click()
   m_HasModify = True
End Sub

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub
Private Sub cboPartTypeSub_Click()
   m_HasModify = True
End Sub
Private Sub cboUnit_Click()
   m_HasModify = True
   txtUnitAmount.Text = "1"
   cboUnitChange.ListIndex = cboUnit.ListIndex
End Sub
Private Sub cboUnitChange_Click()
   m_HasModify = True
End Sub

Private Sub ChkExceptionFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub ChkExceptionFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkLotFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkLotFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkOutlayFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkOutlayFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkRebateFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkRebateFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkStandardCost_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkStandardCost_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วันที่นำเข้า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 5300
   Col.Caption = MapText("สถานที่จัดเก็บ")
    
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนคงคลัง")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาเฉลี่ย")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
   
End Sub
Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("หน่วย")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = ScaleWidth - 5500
   Col.Caption = MapText("ชื่อลูกค้า")
End Sub
Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("กล่อง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ถาด")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("แพ็ค")
   Col.TextAlignment = jgexAlignRight
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PartItem.STOCK_CODE_ID = ID
      m_PartItem.QueryFlag = 1
      If Not glbDaily.QueryStockCode(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartItem.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_PartItem.STOCK_DESC
      txtPartNo.Text = m_PartItem.STOCK_NO
      cboPartType.ListIndex = IDToListIndex(cboPartType, m_PartItem.STOCK_TYPE)
      cboPartTypeSub.ListIndex = IDToListIndex(cboPartTypeSub, m_PartItem.STOCK_TYPE_SUB)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_PartItem.UNIT_ID)
      cboUnitChange.ListIndex = IDToListIndex(cboUnitChange, m_PartItem.UNIT_CHANGE_ID)
      uctlPartItemReturn.MyCombo.ListIndex = IDToListIndex(uctlPartItemReturn.MyCombo, m_PartItem.PART_ITEM_RETURN_ID)
      txtBarCode.Text = m_PartItem.BARCODE
      txtBillCode.Text = m_PartItem.BILL_CODE
      txtBillDesc.Text = m_PartItem.BILL_DESC
      txtUnitAmount.Text = m_PartItem.UNIT_AMOUNT
      txtUnitPerBasket.Text = m_PartItem.UNIT_PER_BASKET
      txtReportPriority.Text = m_PartItem.REPORT_PRIORITY
      ChkStandardCost.Value = FlagToCheck(m_PartItem.CHK_STD_COST)
      txtCostPerAmount.Text = m_PartItem.COST_PER_AMOUNT
      chkLotFlag.Value = FlagToCheck(m_PartItem.LOT_FLAG)
      ChkOutlayFlag.Value = FlagToCheck(m_PartItem.OUTLAY_FLAG)
      ChkExceptionFlag.Value = FlagToCheck(m_PartItem.EXCEPTION_FLAG)
      cboGroupCom.ListIndex = IDToListIndex(cboGroupCom, m_PartItem.GROUP_COM_ID)
      txtJointCode.Text = m_PartItem.JOINT_CODE
      ChkRebateFlag.Value = FlagToCheck(m_PartItem.REBATE_FLAG)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Mr = Nothing
   Set m_PartItemColl = New Collection
   Set BalanceAmountByLot = New Collection
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Tag = "BALANCE_AMOUNT" Then
      If BalanceAmountByLot Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Lt As CLotItem
      If BalanceAmountByLot.Count <= 0 Then
         Exit Sub
      End If
      Set Lt = GetItem(BalanceAmountByLot, RowIndex, RealIndex)
      If Lt Is Nothing Then
         Exit Sub
      End If

      Values(1) = Lt.LOT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = DateToStringExtEx2(Lt.DOCUMENT_DATE)
      Values(4) = Lt.LOCATION_NO
      Values(5) = FormatNumber(Lt.LOT_ITEM_AMOUNT - Lt.TX_AMOUNT)
      Values(6) = FormatNumber(Lt.AVG_PRICE)
      Values(7) = FormatNumber((Lt.LOT_ITEM_AMOUNT - Lt.TX_AMOUNT) * Lt.AVG_PRICE)
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE" Then
      If m_PartItem.StockCodeChange Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Scc As CStockCodeChange
      If m_PartItem.StockCodeChange.Count <= 0 Then
         Exit Sub
      End If
      Set Scc = GetItem(m_PartItem.StockCodeChange, RowIndex, RealIndex)
      If Scc Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = Scc.STOCK_CODE_CHANGE_ID
      Values(2) = RealIndex
      Values(3) = Scc.UNIT_CHANGE_NAME
      Values(4) = Scc.UNIT_CHANGE_AMOUNT
      Values(5) = Scc.CUSTOMER_CODE
      Values(6) = Scc.CUSTOMER_NAME
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE_VERIFY" Then
      If m_PartItem.StockCodeChangeFt Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ft As CStockCodeChangeFt
      If m_PartItem.StockCodeChangeFt.Count <= 0 Then
         Exit Sub
      End If
      Set Ft = GetItem(m_PartItem.StockCodeChangeFt, RowIndex, RealIndex)
      If Ft Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = Ft.STOCK_CODE_CHANGE_FT_ID
      Values(2) = RealIndex
      Values(3) = Ft.BOX
      Values(4) = Ft.TRAY
      Values(5) = Ft.PACK
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboPartType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
'   If Not VerifyCombo(lblParcelType, cboParcelType, False) Then
'      Exit Function
'   End If
   If Not VerifyTextControl(lblUnitPerBasket, txtUnitPerBasket, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblReportPriority, txtReportPriority, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   If (Val(txtBarCode.Text) > 0) Then
      If Not CheckUniqueNs(BARCODE_UNIQUE, Val(txtBarCode.Text), ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtBarCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PartItem.ShowMode = ShowMode
   m_PartItem.STOCK_CODE_ID = ID
   m_PartItem.STOCK_NO = txtPartNo.Text
   m_PartItem.STOCK_DESC = txtName.Text
   m_PartItem.STOCK_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   m_PartItem.STOCK_TYPE_SUB = cboPartTypeSub.ItemData(Minus2Zero(cboPartTypeSub.ListIndex))
   m_PartItem.UNIT_ID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_PartItem.UNIT_CHANGE_ID = cboUnitChange.ItemData(Minus2Zero(cboUnitChange.ListIndex))
   m_PartItem.PART_ITEM_RETURN_ID = uctlPartItemReturn.MyCombo.ItemData(Minus2Zero(uctlPartItemReturn.MyCombo.ListIndex))
   m_PartItem.BARCODE = Val(txtBarCode.Text)
   m_PartItem.BILL_CODE = txtBillCode.Text
   m_PartItem.BILL_DESC = txtBillDesc.Text
   m_PartItem.STOCK_AREA = STOCK_INV
   m_PartItem.UNIT_AMOUNT = Val(txtUnitAmount.Text)
   m_PartItem.UNIT_PER_BASKET = Val(txtUnitPerBasket.Text)
   m_PartItem.REPORT_PRIORITY = Val(txtReportPriority.Text)
   m_PartItem.CHK_STD_COST = Check2Flag(ChkStandardCost.Value)
   m_PartItem.COST_PER_AMOUNT = Val(txtCostPerAmount.Text)
   m_PartItem.LOT_FLAG = Check2Flag(chkLotFlag.Value)
   m_PartItem.OUTLAY_FLAG = Check2Flag(ChkOutlayFlag.Value)
   m_PartItem.EXCEPTION_FLAG = Check2Flag(ChkExceptionFlag.Value)
   m_PartItem.GROUP_COM_ID = cboGroupCom.ItemData(Minus2Zero(cboGroupCom.ListIndex))
   m_PartItem.JOINT_CODE = txtJointCode.Text
   m_PartItem.REBATE_FLAG = Check2Flag(ChkRebateFlag.Value)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditStockCode(m_PartItem, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboPartType, , , , MASTER_STOCKTYPE, , PartGroupID)
      
      Call LoadMaster(cboUnit, , , , MASTER_UNIT)
      
      Call LoadMaster(cboUnitChange, , , , MASTER_UNIT)
      
      Call LoadMaster(cboPartTypeSub, , , , MASTER_STOCKTYPE_SUB)
      
      Call LoadStockCode(uctlPartItemReturn.MyCombo, m_PartItemColl)
      Set uctlPartItemReturn.MyCollection = m_PartItemColl
      
      Call LoadMaster(cboGroupCom, , , , MASTER_GROUP_COM)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         chkLotFlag.Value = ssCBChecked
      End If
      m_HasModify = False
      
      Call TabStrip1_Click
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
   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub InitFormLayout()
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblName, MapText("รายการ"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสคลัง"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทรายการ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblBillCode, MapText("รหัสย่อ"))
   Call InitNormalLabel(lblBarCode, MapText("Barcode"))
   Call InitNormalLabel(lblBillDesc, MapText("ชื่อย่อ"))
   Call InitNormalLabel(lblUnitPerBasket, MapText("จำนวน"))
   Call InitNormalLabel(Label1, MapText("ต่อตะกร้า"))
   Call InitNormalLabel(lblReportPriority, MapText("ลำดับรายงาน"))
   Call InitNormalLabel(lblPartItemReturn, MapText("รับคืน"))
   Call InitNormalLabel(lblGroupCom, MapText("กลุ่มคอม"))
   Call InitNormalLabel(lblJointCode, MapText("รหัสสินค้าอีกบริษัท"))
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBarCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtBillDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtUnitAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtUnitPerBasket.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtReportPriority.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtJointCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboPartTypeSub)
   Call InitCombo(cboUnit)
   Call InitCombo(cboUnitChange)
   Call InitCombo(cboGroupCom)
   
   Call InitCheckBox(ChkStandardCost, "คิดต้นทุนมาตรฐาน")
   Call InitCheckBox(chkLotFlag, "เบิกตัด LOT")
   Call InitCheckBox(ChkExceptionFlag, "ยกเลิก")
   Call InitCheckBox(ChkOutlayFlag, "เป็นค่าใช้จ่าย")
   Call InitCheckBox(ChkRebateFlag, "คิด Rebate")
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("เปลี่ยนแปลงหน่วย")
   T.Tag = "UNIT_CHANGE"
      
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("ยอดคงเหลือ")
   T.Tag = "BALANCE_AMOUNT"
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("เเปลงหน่วยสำหรับตรวจนับ")
   T.Tag = "UNIT_CHANGE_VERIFY"
   
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Set m_PartItem = New CStockCode
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Set m_Mr = New CMasterRef
   Set m_PartItemColl = New Collection
   Set BalanceAmountByLot = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Tag = "UNIT_CHANGE" Then
      Call InitGrid2
      
      Call RefreshGrid(TabStrip1.SelectedItem.Tag)
   ElseIf TabStrip1.SelectedItem.Tag = "BALANCE_AMOUNT" Then
      Call InitGrid1
      
      If BalanceAmountByLot.Count <= 0 Then
         Call GetBalanceByLotItemLinkDate(BalanceAmountByLot, -1, -1, -1, txtPartNo.Text, txtPartNo.Text)
      End If
      GridEX1.ItemCount = CountItem(BalanceAmountByLot)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE_VERIFY" Then
      Call InitGrid3
      
      Call RefreshGrid(TabStrip1.SelectedItem.Tag)
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub
Private Sub txtJointCode_Change()
   m_HasModify = True
End Sub
Private Sub txtBarcode_Change()
   m_HasModify = True
End Sub
Private Sub txtBillCode_Change()
   m_HasModify = True
End Sub
Private Sub txtBillDesc_Change()
   m_HasModify = True
End Sub
Private Sub txtCostPerAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub
Private Sub txtName_Change()
   m_HasModify = True
End Sub
Private Sub txtReportPriority_Change()
   m_HasModify = True
End Sub
Private Sub txtUnitAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtUnitPerBasket_Change()
   m_HasModify = True
End Sub
Private Sub uctlPartItemReturn_Change()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   TabStrip1.Width = GridEX1.Width
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   
End Sub
Private Sub txtPartNo_LostFocus()
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      txtPartNo.SetFocus
      Exit Sub
   End If
End Sub
Public Sub RefreshGrid(Key As String)
   If Key = "UNIT_CHANGE" Then
      GridEX1.ItemCount = CountItem(m_PartItem.StockCodeChange)
      GridEX1.Rebind
   ElseIf Key = "UNIT_CHANGE_VERIFY" Then
      GridEX1.ItemCount = CountItem(m_PartItem.StockCodeChangeFt)
      GridEX1.Rebind
   End If
   
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = "UNIT_CHANGE" Then
      Set frmAddEditPartItemChange.ParentForm = Me
      Set frmAddEditPartItemChange.TempCollection = m_PartItem.StockCodeChange
      frmAddEditPartItemChange.ShowMode = SHOW_ADD
      frmAddEditPartItemChange.HeaderText = MapText("เพิ่มหน่วยนับ")
      Load frmAddEditPartItemChange
      frmAddEditPartItemChange.Show 1
      
      OKClick = frmAddEditPartItemChange.OKClick
      
      Unload frmAddEditPartItemChange
      Set frmAddEditPartItemChange = Nothing
         
      If OKClick Then
         Call RefreshGrid(TabStrip1.SelectedItem.Tag)
      End If
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE_VERIFY" Then
      If m_PartItem.StockCodeChangeFt.Count > 0 Then
         Exit Sub
      End If
      Set frmAddEditPartItemChangeFt.ParentForm = Me
      Set frmAddEditPartItemChangeFt.TempCollection = m_PartItem.StockCodeChangeFt
      frmAddEditPartItemChangeFt.ShowMode = SHOW_ADD
      frmAddEditPartItemChangeFt.HeaderText = MapText("เพิ่มหน่วยนับ")
      Load frmAddEditPartItemChangeFt
      frmAddEditPartItemChangeFt.Show 1
      
      OKClick = frmAddEditPartItemChangeFt.OKClick
      
      Unload frmAddEditPartItemChangeFt
      Set frmAddEditPartItemChangeFt = Nothing
         
      If OKClick Then
         Call RefreshGrid(TabStrip1.SelectedItem.Tag)
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Tag = "UNIT_CHANGE" Then
      Set frmAddEditPartItemChange.ParentForm = Me
      frmAddEditPartItemChange.ID = ID
      Set frmAddEditPartItemChange.TempCollection = m_PartItem.StockCodeChange
      frmAddEditPartItemChange.HeaderText = MapText("แก้ไขหน่วย")
      frmAddEditPartItemChange.ShowMode = SHOW_EDIT
      Load frmAddEditPartItemChange
      frmAddEditPartItemChange.Show 1

      OKClick = frmAddEditPartItemChange.OKClick

      Unload frmAddEditPartItemChange
      Set frmAddEditPartItemChange = Nothing
      
      If OKClick Then
         Call RefreshGrid(TabStrip1.SelectedItem.Tag)
      End If
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE_VERIFY" Then
      Set frmAddEditPartItemChangeFt.ParentForm = Me
      frmAddEditPartItemChangeFt.ID = ID
      Set frmAddEditPartItemChangeFt.TempCollection = m_PartItem.StockCodeChangeFt
      frmAddEditPartItemChangeFt.HeaderText = MapText("แก้ไขหน่วย")
      frmAddEditPartItemChangeFt.ShowMode = SHOW_EDIT
      Load frmAddEditPartItemChangeFt
      frmAddEditPartItemChangeFt.Show 1

      OKClick = frmAddEditPartItemChangeFt.OKClick

      Unload frmAddEditPartItemChangeFt
      Set frmAddEditPartItemChangeFt = Nothing
      
      If OKClick Then
         Call RefreshGrid(TabStrip1.SelectedItem.Tag)
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Tag = "UNIT_CHANGE" Then
      If ID1 <= 0 Then
         m_PartItem.StockCodeChange.Remove (ID2)
      Else
         m_PartItem.StockCodeChange.Item(ID2).Flag = "D"
      End If
   ElseIf TabStrip1.SelectedItem.Tag = "UNIT_CHANGE_VERIFY" Then
      If ID1 <= 0 Then
         m_PartItem.StockCodeChangeFt.Remove (ID2)
      Else
         m_PartItem.StockCodeChangeFt.Item(ID2).Flag = "D"
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   
   Call RefreshGrid(TabStrip1.SelectedItem.Tag)
   m_HasModify = True

End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

