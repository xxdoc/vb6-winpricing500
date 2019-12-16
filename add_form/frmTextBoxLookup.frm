VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTextBoxLookup 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmTextBoxLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12091
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4965
         Left            =   60
         TabIndex        =   1
         Top             =   1830
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   8758
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
         Column(1)       =   "frmTextBoxLookup.frx":27A2
         Column(2)       =   "frmTextBoxLookup.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmTextBoxLookup.frx":290E
         FormatStyle(2)  =   "frmTextBoxLookup.frx":2A6A
         FormatStyle(3)  =   "frmTextBoxLookup.frx":2B1A
         FormatStyle(4)  =   "frmTextBoxLookup.frx":2BCE
         FormatStyle(5)  =   "frmTextBoxLookup.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmTextBoxLookup.frx":2D5E
      End
      Begin Xivess.uctlTextBox txtSearchText 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2865
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtSearchName 
         Height          =   435
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   4545
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblSearchName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label lblSearchText 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   900
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmTextBoxLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Rs As ADODB.Recordset
Public KEYWORD As String
Public KeySearch As String

Public OKClick As Boolean
Public HeaderText As String
Private m_Customer As CAPARMas
Private m_StockNo As CStockCode
Private m_MasterRef As CMasterRef
Private m_Employee As CEmployee
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call RunQuery
   End If
End Sub
Private Sub RunQuery()
   If KeySearch = "CUSTOMER_CODE" Then
      Call QueryCustomer(1)
   ElseIf KeySearch = "SUPPLIER_CODE" Then
      Call QueryCustomer(2)
   ElseIf KeySearch = "STOCK_NO" Or KeySearch = "NOT_STOCK_NO" Then
      Call QueryStockNo
   ElseIf KeySearch = "LOCATION_NO" Then
      Call QueryLocationNo
   ElseIf KeySearch = "SALE_CODE" Then
      Call QuerySaleCode
   ElseIf KeySearch = "EMP_CODE" Then
      Call QuerySaleCode
   ElseIf KeySearch = "LOCATION_GROUP_NO" Then
      Call QueryLocationGroupNo
   End If
End Sub
Private Sub QueryCustomer(Area As Long)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_Customer As CAPARMas
   MasterInd = "2"
   Set m_Customer = New CAPARMas
   MasterInd = "1"
   
   m_Customer.APAR_MAS_ID = -1
   m_Customer.APAR_IND = Area
   m_Customer.APAR_CODE = PatchWildCard(txtSearchText.Text)
   m_Customer.APAR_NAME = PatchWildCard(txtSearchName.Text)
   m_Customer.ORDER_TYPE = 1
   Call m_Customer.QueryData(2, m_Rs, itemcount, True)
   
   Call InitGrid
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryStockNo()
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_StockNo As CStockCode
   Set m_StockNo = New CStockCode
   
   m_StockNo.STOCK_CODE_ID = -1
   m_StockNo.STOCK_NO = PatchWildCard(txtSearchText.Text)
   m_StockNo.STOCK_DESC = PatchWildCard(txtSearchName.Text)
   m_StockNo.EXCEPTION_FLAG = "N"
   Call m_StockNo.QueryData(2, m_Rs, itemcount, True)
   
   Call InitGrid
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryLocationNo()
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_MasterRef As CMasterRef
   MasterInd = "5"
   Set m_MasterRef = New CMasterRef
   MasterInd = "1"
   
   m_MasterRef.KEY_ID = -1
   m_MasterRef.MASTER_AREA = MASTER_LOCATION
   m_MasterRef.KEY_CODE = PatchWildCard(txtSearchText.Text)
   m_MasterRef.KEY_NAME = PatchWildCard(txtSearchName.Text)
   Call m_MasterRef.QueryData(5, m_Rs, itemcount, True)
   
   Call InitGrid
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QueryLocationGroupNo()
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_MasterRef As CMasterRef
   MasterInd = "5"
   Set m_MasterRef = New CMasterRef
   MasterInd = "1"
   
   m_MasterRef.KEY_ID = -1
   m_MasterRef.MASTER_AREA = MASTER_INVENTORY_SALE_GROUP
   m_MasterRef.KEY_CODE = PatchWildCard(txtSearchText.Text)
   m_MasterRef.KEY_NAME = PatchWildCard(txtSearchName.Text)
   Call m_MasterRef.QueryData(5, m_Rs, itemcount, True)
   
   Call InitGrid
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub QuerySaleCode()
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
         
   Dim m_Employee As CEmployee
   Set m_Employee = New CEmployee
   
   m_Employee.EMP_ID = -1
   m_Employee.EMP_CODE = PatchWildCard(txtSearchText.Text)
   m_Employee.EMP_NAME = PatchWildCard(txtSearchName.Text)
   Call m_Employee.QueryData(2, m_Rs, itemcount, True)
   
   Call InitGrid
   
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("����")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = ScaleWidth - 2300
   Col.Caption = MapText("��������´")
   
   GridEX1.itemcount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblSearchText, MapText("����"))
   Call InitNormalLabel(lblSearchName, MapText("��������´"))
      
   txtSearchText.Text = KEYWORD
   txtSearchText.Enabled = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   MasterInd = "2"
   Set m_Customer = New CAPARMas
   Set m_StockNo = New CStockCode
   Set m_MasterRef = New CMasterRef
   Set m_Employee = New CEmployee
   MasterInd = "1"
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Customer = Nothing
   Set m_StockNo = Nothing
   Set m_MasterRef = Nothing
   Set m_Employee = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   MasterInd = "1"
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call ReturnKeyWord
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   
   If KeySearch = "CUSTOMER_CODE" Or KeySearch = "SUPPLIER_CODE" Then
      Call m_Customer.PopulateFromRS(2, m_Rs)
      Values(1) = m_Customer.APAR_MAS_ID
      Values(2) = m_Customer.APAR_CODE
      Values(3) = m_Customer.APAR_NAME
   ElseIf KeySearch = "STOCK_NO" Or KeySearch = "NOT_STOCK_NO" Then
      Call m_StockNo.PopulateFromRS(2, m_Rs)
      Values(1) = m_StockNo.STOCK_CODE_ID
      Values(2) = m_StockNo.STOCK_NO
      Values(3) = m_StockNo.STOCK_DESC
   ElseIf KeySearch = "LOCATION_NO" Then
      Call m_MasterRef.PopulateFromRS(5, m_Rs)
      Values(1) = m_MasterRef.KEY_ID
      Values(2) = m_MasterRef.KEY_CODE
      Values(3) = m_MasterRef.KEY_NAME
   ElseIf KeySearch = "SALE_CODE" Then
      Call m_Employee.PopulateFromRS(2, m_Rs)
      Values(1) = m_Employee.EMP_ID
      Values(2) = m_Employee.EMP_CODE
      Values(3) = m_Employee.EMP_NAME & " " & m_Employee.EMP_LNAME
   ElseIf KeySearch = "EMP_CODE" Then
      Call m_Employee.PopulateFromRS(2, m_Rs)
      Values(1) = m_Employee.EMP_ID
      Values(2) = m_Employee.EMP_CODE
      Values(3) = m_Employee.EMP_NAME & " " & m_Employee.EMP_LNAME
  ElseIf KeySearch = "LOCATION_GROUP_NO" Then
      Call m_MasterRef.PopulateFromRS(5, m_Rs)
      Values(1) = m_MasterRef.KEY_ID
      Values(2) = m_MasterRef.KEY_CODE
      Values(3) = m_MasterRef.KEY_NAME
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      KeyCode = 0
      Unload Me
   ElseIf KeyCode = 13 Or KeyCode = 32 Then
      Call ReturnKeyWord
   End If
End Sub
Private Sub ReturnKeyWord()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
      
   KEYWORD = GridEX1.Value(2)
   Unload Me
End Sub
Private Sub txtSearchName_Change()
   Call RunQuery
End Sub
