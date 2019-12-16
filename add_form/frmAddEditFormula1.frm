VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditFormula1 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditFormula1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboProductionLocation 
         Height          =   315
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   1755
      End
      Begin Xivess.uctlDate uctlFormulaDate 
         Height          =   405
         Left            =   5580
         TabIndex        =   1
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   5
         Top             =   2355
         Width           =   11595
         _ExtentX        =   20452
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
      Begin Xivess.uctlTextBox txtFormulaNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   900
         Width           =   2385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4815
         Left            =   150
         TabIndex        =   6
         Top             =   2910
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8493
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
         Column(1)       =   "frmAddEditFormula1.frx":27A2
         Column(2)       =   "frmAddEditFormula1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditFormula1.frx":290E
         FormatStyle(2)  =   "frmAddEditFormula1.frx":2A6A
         FormatStyle(3)  =   "frmAddEditFormula1.frx":2B1A
         FormatStyle(4)  =   "frmAddEditFormula1.frx":2BCE
         FormatStyle(5)  =   "frmAddEditFormula1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditFormula1.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextLookup uctlPartItem 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1800
         Width           =   6405
         _ExtentX        =   14684
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtFormulaDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1350
         Width           =   7575
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin VB.Label lblProductionLocation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8400
         TabIndex        =   19
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label lblFormulaDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   1665
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormula1.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label lblFormulaDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4290
         TabIndex        =   16
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormula1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormula1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormula1.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblFormulaNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditFormula1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Formula As CFormula

Private m_PartItemColls  As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Formula.FORMULA_ID = ID
      m_Formula.QueryFlag = 1
      If Not glbDaily.QueryFormula(m_Formula, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Formula.PopulateFromRS(1, m_Rs)
      
      uctlFormulaDate.ShowDate = m_Formula.FORMULA_DATE
      txtFormulaNo.Text = m_Formula.FORMULA_NO
      txtFormulaDesc.Text = m_Formula.FORMULA_DESC
      uctlPartItem.MyCombo.ListIndex = IDToListIndex(uctlPartItem.MyCombo, m_Formula.PART_ITEM_ID)
      cboProductionLocation.ListIndex = IDToListIndex(cboProductionLocation, m_Formula.PRD_LOCATION_ID)
      
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblFormulaNo, txtFormulaNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFormulaDate, uctlFormulaDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartItem, uctlPartItem.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblProductionLocation, cboProductionLocation, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(EXPORT_UNIQUE, txtFormulano.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtFormulano.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Formula.AddEditMode = ShowMode
   m_Formula.FORMULA_ID = ID
   m_Formula.FORMULA_NO = txtFormulaNo.Text
   m_Formula.FORMULA_DESC = txtFormulaDesc.Text
   m_Formula.FORMULA_DATE = uctlFormulaDate.ShowDate
   m_Formula.PART_ITEM_ID = uctlPartItem.MyCombo.ItemData(Minus2Zero(uctlPartItem.MyCombo.ListIndex))
   m_Formula.PRD_LOCATION_ID = cboProductionLocation.ItemData(Minus2Zero(cboProductionLocation.ListIndex))
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditFormula(m_Formula, IsOK, True, glbErrorLog) Then
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
Private Sub cboProductionLocation_Click()
   m_HasModify = True
End Sub
Private Sub cboProductionLocation_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu As CPopupMenu
Dim lMenuChoosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = "DESC" Then
      Set frmAddEditFormulaItem1.ParentForm = Me
      Set frmAddEditFormulaItem1.TempCollection = m_Formula.FormulaItems
      frmAddEditFormulaItem1.ShowMode = SHOW_ADD
      frmAddEditFormulaItem1.HeaderText = MapText("เพิ่ม" & "รายการสูตร")
      Load frmAddEditFormulaItem1
      frmAddEditFormulaItem1.Show 1

      OKClick = frmAddEditFormulaItem1.OKClick

      Unload frmAddEditFormulaItem1
      Set frmAddEditFormulaItem1 = Nothing
   End If

   If OKClick Then
      Call RefreshGrid
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
   
   If TabStrip1.SelectedItem.Tag = "DESC" Then
      If ID1 <= 0 Then
         m_Formula.FormulaItems.Remove (ID2)
      Else
         m_Formula.FormulaItems.Item(ID2).Flag = "D"
      End If
     Call RefreshGrid
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
   
   If TabStrip1.SelectedItem.Tag = "DESC" Then
      Set frmAddEditFormulaItem1.ParentForm = Me
      frmAddEditFormulaItem1.ID = ID
      Set frmAddEditFormulaItem1.TempCollection = m_Formula.FormulaItems
      frmAddEditFormulaItem1.HeaderText = MapText("แก้ไข" & "รายการสูตร")
      frmAddEditFormulaItem1.ShowMode = SHOW_EDIT
      Load frmAddEditFormulaItem1
      frmAddEditFormulaItem1.Show 1
         
      OKClick = frmAddEditFormulaItem1.OKClick
      
      Unload frmAddEditFormulaItem1
      Set frmAddEditFormulaItem1 = Nothing
      
      If OKClick Then
         Call RefreshGrid
      End If
   End If
End Sub
Private Sub cmdOK_Click()
Dim oMenu As CPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_Formula.FORMULA_ID
      m_Formula.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadStockCode(uctlPartItem.MyCombo, m_PartItemColls)
      Set uctlPartItem.MyCollection = m_PartItemColls
      
      Call LoadMaster(cboProductionLocation, , , , MASTER_PRODUCTION_LOCATION)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Formula.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlFormulaDate.ShowDate = Now
         m_Formula.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
'      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Formula = Nothing
   Set m_PartItemColls = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
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
   Col.Caption = MapText("BATCH/ลำดับ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("รหัสสินค้า")

   Set Col = GridEX1.Columns.add '5
   Col.Width = ScaleWidth - 8000
   Col.Caption = MapText("สินค้า")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("ประเภทผลิต")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน %")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.Caption = MapText("รหัสคลัง")

   
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblFormulaNo, MapText("หมายเลขสูตร"))
   Call InitNormalLabel(lblFormulaDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFormulaDate, MapText("วันที่สูตร"))
   Call InitNormalLabel(lblPartItem, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblProductionLocation, MapText("สถานที่ผลิต"))
   
   Call txtFormulaNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
    
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   
   Call InitCombo(cboProductionLocation)
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16

   Dim T As Object
   TabStrip1.Tabs.Clear

   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("รายละเอียด")
   T.Tag = "DESC"
   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Formula = New CFormula
   
   Set m_PartItemColls = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim FI As CFormulaItem

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Tag = "DESC" Then
      If m_Formula.FormulaItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Set FI = GetItem(m_Formula.FormulaItems, RowIndex, RealIndex)
      If FI Is Nothing Then
         Exit Sub
      End If

      Values(1) = FI.FORMULA_ITEM_ID
      Values(2) = RealIndex
      Values(3) = FI.BATCH_NO & "/" & FI.ORDER_NO
      Values(4) = FI.PART_NO
      If Len(FI.PART_DESC) > 0 Then
         Values(5) = FI.PART_DESC
      Else
         Values(5) = FI.PROBLEM_DESC
      End If
      Values(6) = FI.PRODUCTION_TYPE_NAME
      If Len(FI.PART_DESC) > 0 Then
         Values(7) = FormatNumberToNull(FI.FORMULA_ITEM_PERCENT)
      Else
         Values(7) = "<= " & FormatNumberToNull(FI.PROBLEM_LIMIT_PERCENT)
      End If
      If Len(FI.LOCATION_NAME) > 0 Then
         Values(8) = FI.LOCATION_NAME
      ElseIf FI.NEXT_FLAG = "Y" Then
         Values(8) = "------>"
      Else
         Values(8) = ""
      End If
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   Call InitGrid1
   Call RefreshGrid
End Sub

Private Sub txtFormulaDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtFormulano_Change()
   m_HasModify = True
End Sub
Public Sub RefreshGrid()
   m_HasModify = True
   GridEX1.ItemCount = CountItem(m_Formula.FormulaItems)
   GridEX1.Rebind
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

Private Sub uctlFormulaDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPartItem_Change()
   m_HasModify = True
End Sub
