VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmInventoryDoc 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmInventoryDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDepartment 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1770
         Width           =   2655
      End
      Begin Xivess.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2220
         Width           =   2655
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2220
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8916
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
         Column(1)       =   "frmInventoryDoc.frx":27A2
         Column(2)       =   "frmInventoryDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmInventoryDoc.frx":290E
         FormatStyle(2)  =   "frmInventoryDoc.frx":2A6A
         FormatStyle(3)  =   "frmInventoryDoc.frx":2B1A
         FormatStyle(4)  =   "frmInventoryDoc.frx":2BCE
         FormatStyle(5)  =   "frmInventoryDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmInventoryDoc.frx":2D5E
      End
      Begin Xivess.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1680
         TabIndex        =   3
         Top             =   1770
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Xivess.uctlDate uctlToDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   24
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtInventoryRefNoMain 
         Height          =   435
         Left            =   5970
         TabIndex        =   25
         Top             =   840
         Width           =   2625
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblInventoryRefNoMain 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin Threed.SSCheck ChkCancelFlag 
         Height          =   465
         Left            =   8700
         TabIndex        =   23
         Top             =   2160
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   820
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   465
         Left            =   8700
         TabIndex        =   5
         Top             =   1740
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   820
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   22
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4800
         TabIndex        =   19
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   9
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   15
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryDoc.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_InventoryDoc As CInventoryDoc
Private m_TempInventoryDoc As CInventoryDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Private m_Mr As CMasterRef

Public OKClick As Boolean
Public DocumentType As INVENTORY_DOCTYPE
Public TempInventorySubTypeColl As Collection
Public InventorySubType As Long
Public HeaderText As String
Private Function DocumentType2Text(TempID As INVENTORY_DOCTYPE) As String
   If TempID = IMPORT_DOCTYPE Then
      DocumentType2Text = "ข้อมูลการนำเข้า"
   ElseIf TempID = EXPORT_DOCTYPE Then
      DocumentType2Text = "ข้อมูลการเบิกจ่าย"
   ElseIf TempID = TRANSFER_DOCTYPE Then
      DocumentType2Text = "ข้อมูลการโอน"
   ElseIf TempID = ADJUST_DOCTYPE Then
      DocumentType2Text = "ข้อมูลการปรับยอด"
   ElseIf TempID = 1000 Then
      DocumentType2Text = "ข้อมูลการผลิต"
   End If
End Function

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditInventoryDoc.InventorySubType = InventorySubType
   frmAddEditInventoryDoc.DocumentType = DocumentType
   frmAddEditInventoryDoc.HeaderText = MapText("เพิ่ม" & DocumentType2Text(DocumentType))
   frmAddEditInventoryDoc.ShowMode = SHOW_ADD
   Load frmAddEditInventoryDoc
   frmAddEditInventoryDoc.Show 1

   OKClick = frmAddEditInventoryDoc.OKClick

   Unload frmAddEditInventoryDoc
   Set frmAddEditInventoryDoc = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtPartNo.Text = ""
   uctlDocumentDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not VerifyLockDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If Not VerifyLockInventoryDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   If Not CheckUniqueNs(CONSIGNMENT_NO, GridEX1.Value(2), 0) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถลบ") & " " & GridEX1.Value(2) & " " & MapText("ได้ เนื่องจากมีการอ้างอิงร่วมกับเอกสารใบสั่งซื้อ")
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   Call EnableForm(Me, False)
   Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", ID)
   If Not glbDaily.DeleteInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
      Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", -1)
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   Call EnableForm(Me, True)
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

   ID = Val(GridEX1.Value(1))
   
   frmAddEditInventoryDoc.InventorySubType = InventorySubType
   frmAddEditInventoryDoc.DocumentType = DocumentType
   frmAddEditInventoryDoc.ID = ID
   frmAddEditInventoryDoc.HeaderText = MapText("แก้ไข" & DocumentType2Text(DocumentType))
   frmAddEditInventoryDoc.ShowMode = SHOW_EDIT
   Load frmAddEditInventoryDoc
   frmAddEditInventoryDoc.Show 1

   OKClick = frmAddEditInventoryDoc.OKClick

   Unload frmAddEditInventoryDoc
   Set frmAddEditInventoryDoc = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadMaster(cboDepartment, , , , MASTER_DEPARTMENT)
      
      Call InitInventoryDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      uctlDocumentDate.ShowDate = Now
      uctlToDate.ShowDate = Now
      
      Call SeparateInventorySubTypeColl(TempInventorySubTypeColl)
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", -1)
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_NO", PatchWildCard(txtDocumentNo.Text))
      Call m_InventoryDoc.SetFieldValue("STOCK_CODE_NO", PatchWildCard(txtPartNo.Text))
      Call m_InventoryDoc.SetFieldValue("FROM_DATE", uctlDocumentDate.ShowDate)
      Call m_InventoryDoc.SetFieldValue("TO_DATE", uctlToDate.ShowDate)
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
      Call m_InventoryDoc.SetFieldValue("INVENTORY_SUB_TYPE", InventorySubType)
      Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", Check2Flag(chkCommit.Value))
      Call m_InventoryDoc.SetFieldValue("ORDER_BY", cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex)))
      Call m_InventoryDoc.SetFieldValue("ORDER_TYPE", cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)))
      Call m_InventoryDoc.SetFieldValue("CANCEL_FLAG", Check2Flag(ChkCancelFlag.Value))
      m_InventoryDoc.INVENTORY_REF_NO_MAIN = txtInventoryRefNoMain.Text
      
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
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
   Col.Width = 2115
   Col.Caption = MapText("เลขที่บิลรับของ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2055
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("แผนก")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 5305
   Col.Caption = MapText("หมายเหตุ")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("COMMIT_FLAG")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1000
   Col.Caption = MapText("เอกสารย่อย")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสสต็อค"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblDepartment, MapText("แผนก"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call InitNormalLabel(lblInventoryRefNoMain, MapText("นำเข้า NO"))
   
   Call txtPartNo.SetKeySearch("STOCK_NO")
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   Call InitCombo(cboDepartment)
   Call InitCheckBox(chkCommit, "ห้ามแก้ไข")
   Call InitCheckBox(ChkCancelFlag, "CANCEL")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Enabled = False
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_InventoryDoc = New CInventoryDoc
   Set m_TempInventoryDoc = New CInventoryDoc
   Set m_Rs = New ADODB.Recordset
   Set m_Mr = New CMasterRef
   Set TempInventorySubTypeColl = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Mr = Nothing
   Set TempInventorySubTypeColl = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(6)
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
   Call m_TempInventoryDoc.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempInventoryDoc.GetFieldValue("INVENTORY_DOC_ID")
   Values(2) = m_TempInventoryDoc.GetFieldValue("DOCUMENT_NO")
   Values(3) = DateToStringExtEx2(m_TempInventoryDoc.GetFieldValue("DOCUMENT_DATE"))
   Values(4) = m_TempInventoryDoc.GetFieldValue("DEPARTMENT_NAME")
   Values(5) = m_TempInventoryDoc.GetFieldValue("DOCUMENT_DESC")
   Values(6) = m_TempInventoryDoc.GetFieldValue("COMMIT_FLAG")
   Values(7) = m_TempInventoryDoc.GetFieldValue("INVENTORY_SUB_TYPE")
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As CPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim TempID2 As Long
Dim BD As CInventoryDoc
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New CPopupMenu
      If chkCommit.Value = ssCBChecked Then
         lMenuChosen = oMenu.Popup("ยกเลิกการคำนวณ ")
      Else
         lMenuChosen = oMenu.Popup("เปลี่ยนประเภทเอกสารย่อย")
         lMenuChosen = lMenuChosen + 1000 'บวก 1000เนื่องจากว่าจะได้ แยกกับส่วนของที่ COMMIT แล้ว
      End If
      If lMenuChosen = 0 Or lMenuChosen = 1000 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Call glbDaily.StartTransaction
      Set BD = New CInventoryDoc
      Call BD.SetFieldValue("INVENTORY_DOC_ID", TempID1)
      Call BD.SetFieldValue("COMMIT_FLAG", "N")
     
      Call BD.UndoCommit
      
      Call glbDaily.CommitTransaction
      
      Call QueryData(True)
      Set BD = Nothing
   ElseIf lMenuChosen = 1001 Then
      Call EnableForm(Me, True)
      Set oMenu = Nothing
      Set oMenu = New CPopupMenu
      lMenuChosen = oMenu.AddMenu(TempInventorySubTypeColl)
      Call EnableForm(Me, False)
      If lMenuChosen > 0 Then
         Set BD = New CInventoryDoc
         BD.INVENTORY_DOC_ID = TempID1
         BD.INVENTORY_SUB_TYPE = lMenuChosen
         Call BD.UpdateInventorySubType
         Set BD = Nothing
         Call QueryData(True)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub SeparateInventorySubTypeColl(TempCollection As Collection)
Dim MenuSelected As Long
Dim oMenu As CPopupMenu
Dim Mr As CMasterRef
Dim MI As CMenuItem
   
   Set TempCollection = Nothing
   Set TempCollection = New Collection
   For Each Mr In InventorySubTypecoll
      If Mr.INDEX_LINK = DocumentType And Mr.KEY_ID <> InventorySubType Then
         Set MI = New CMenuItem
         MI.KEY_ID = Mr.KEY_ID
         MI.KEYWORD = Mr.KEY_NAME
         Call TempCollection.add(MI)
         Set MI = Nothing
      End If
   Next Mr
   
End Sub
