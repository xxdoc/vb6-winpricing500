VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBarcodeTransfer 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmBarcodeTransfer.frx":0000
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
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   5340
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   10
         Top             =   3480
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
      Begin Xivess.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2250
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   1785
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
         Height          =   3765
         Left            =   150
         TabIndex        =   11
         Top             =   3960
         Width           =   11595
         _ExtentX        =   20452
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
         Column(1)       =   "frmBarcodeTransfer.frx":27A2
         Column(2)       =   "frmBarcodeTransfer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBarcodeTransfer.frx":290E
         FormatStyle(2)  =   "frmBarcodeTransfer.frx":2A6A
         FormatStyle(3)  =   "frmBarcodeTransfer.frx":2B1A
         FormatStyle(4)  =   "frmBarcodeTransfer.frx":2BCE
         FormatStyle(5)  =   "frmBarcodeTransfer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBarcodeTransfer.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   10260
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   870
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   2730
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   2
         Top             =   2160
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlPartSubLookup 
         Height          =   435
         Left            =   7080
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlLocationToLookup 
         Height          =   435
         Left            =   7080
         TabIndex        =   1
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblLocationLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -240
         TabIndex        =   25
         Top             =   1620
         Width           =   1845
      End
      Begin Threed.SSCommand cmdSaveBarcode 
         Height          =   525
         Left            =   10200
         TabIndex        =   5
         Top             =   2715
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeTransfer.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2790
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2220
         Width           =   1485
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   3810
         TabIndex        =   22
         Top             =   2790
         Width           =   2565
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         TabIndex        =   21
         Top             =   900
         Width           =   735
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1740
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeTransfer.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeTransfer.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   20
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4050
         TabIndex        =   19
         Top             =   900
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeTransfer.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeTransfer.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   900
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmBarcodeTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As INVENTORY_DOCTYPE

Private FileName As String
Private m_SumUnit As Double

Private m_Cd As Collection
Private DocAdd As Long
Private m_Parts As Collection
Private m_Locations As Collection

'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'------------------------------------------------------
Private Sub CalculateSumPrice()
Dim Li As CLotItem
Dim Sum2 As Double
Dim Ti As CTransferItem
   
   Sum2 = 0
      
   If DocumentType = TRANSFER_DOCTYPE Then
      For Each Ti In m_InventoryDoc.TransferItems
         If Ti.Flag <> "D" Then
            Sum2 = Sum2 + Ti.ImportItem.TX_AMOUNT
         End If
      Next Ti
   Else
      For Each Li In m_InventoryDoc.ImportExportItems
         If Li.Flag <> "D" Then
            Sum2 = Sum2 + Li.TX_AMOUNT
         End If
      Next Li
   End If
   
   
   txtTotalAmount.Text = Format(Sum2, "0.00")
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(INVENTORY_DOC_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.ShowMode = SHOW_ADD
   Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", ID)
   If uctlDocumentDate.ShowDate <= 0 Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", Now)
   Else
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", uctlDocumentDate.ShowDate)
   End If
    
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   If (DocumentType = ADJUST_DOCTYPE) Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", 1000) '��Ե
   Else
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   End If
   
   
   Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("EXCEPTION_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("SALE_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("ADJUST_FLAG", "N")
   If DocumentType = IMPORT_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "�Ѻ��Ҩҡ BARCODE")
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "�ԡ�͡�ҡ BARCODE")
   ElseIf DocumentType = ADJUST_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "��Ե")
   End If
   Call m_InventoryDoc.SetFieldValue("DEPARTMENT_ID", -1)
   Call m_InventoryDoc.SetFieldValue("CANCEL_FLAG", "N")
   If DocumentType = ADJUST_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("BARCODE_JOB_FLAG", "Y")
   End If
   
   
   
   Call EnableForm(Me, False)
   
   If DocumentType = TRANSFER_DOCTYPE Then
      Call CreateImportExportItems
      Call PopulateGuiID(m_InventoryDoc)
   End If
   
   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
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
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   If Len(txtDocumentNo.Text) > 0 Then
      SendKeys ("{TAB}")
      Exit Sub
   End If
   
   ID = ConvertDocToConfigNo(2, DocumentType, -1)
   
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         Call m_InventoryDoc.SetFieldValue("RUNNING_NO", Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd)
         Call m_InventoryDoc.SetFieldValue("CONFIG_DOC_TYPE", ID)
         
         Call txtDocumentNo.SetSelectText(Len(txtDocumentNo.Text) - Cd.GetFieldValue("DIGIT_AMOUNT"), Cd.GetFieldValue("DIGIT_AMOUNT"))
      Else
         txtDocumentNo.Text = ""
      End If
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
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (DocumentType = IMPORT_DOCTYPE) Or (DocumentType = EXPORT_DOCTYPE) Or (DocumentType = ADJUST_DOCTYPE) Then
         If ID1 <= 0 Then
            m_InventoryDoc.ImportExportItems.Remove (ID2)
         Else
            m_InventoryDoc.ImportExportItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = TRANSFER_DOCTYPE Then
         If ID1 <= 0 Then
            m_InventoryDoc.TransferItems.Remove (ID2)
         Else
            m_InventoryDoc.TransferItems.Item(ID2).Flag = "D"
         End If
      End If
      
      Call RefreshGrid(DocumentType, True)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSaveBarcode_Click()
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Mr As CMasterRef
Dim LocationName As String
Dim Pi As CStockCode
   
   If Not VerifyCombo(lblLocationLookup, uctlLocationLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblLocationLookup, uctlLocationToLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not ((uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)) > 0) Or (uctlPartSubLookup.MyCombo.ItemData(Minus2Zero(uctlPartSubLookup.MyCombo.ListIndex)) > 0)) Then
      Call MsgBox("��سҡ�͡������ " & " �Թ��� " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Sub
   End If
     
   Dim EnpAddress As CTransferItem
   Dim Ei As CLotItem
   Dim II As CLotItem
   
   Set Ei = New CLotItem
   Set II = New CLotItem
   Set EnpAddress = New CTransferItem
      
   Ei.Flag = "A"
   II.Flag = "A"
   EnpAddress.Flag = "A"
      
   Set EnpAddress.ExportItem = Ei
   Set EnpAddress.ImportItem = II
   
   Call m_InventoryDoc.TransferItems.add(EnpAddress)
   
   If (uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)) > 0) Then
      EnpAddress.ExportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      EnpAddress.ExportItem.PART_NO = uctlPartLookup.MyTextBox.Text
      EnpAddress.ExportItem.PART_DESC = uctlPartLookup.MyCombo.Text
      
      EnpAddress.ImportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      EnpAddress.ImportItem.PART_NO = uctlPartLookup.MyTextBox.Text
      EnpAddress.ImportItem.PART_DESC = uctlPartLookup.MyCombo.Text
   Else
      EnpAddress.ExportItem.PART_ITEM_ID = uctlPartSubLookup.MyCombo.ItemData(Minus2Zero(uctlPartSubLookup.MyCombo.ListIndex))
      EnpAddress.ExportItem.PART_NO = uctlPartSubLookup.MyTextBox.Text
      EnpAddress.ExportItem.PART_DESC = uctlPartSubLookup.MyCombo.Text
      
      EnpAddress.ImportItem.PART_ITEM_ID = uctlPartSubLookup.MyCombo.ItemData(Minus2Zero(uctlPartSubLookup.MyCombo.ListIndex))
      EnpAddress.ImportItem.PART_NO = uctlPartSubLookup.MyTextBox.Text
      EnpAddress.ImportItem.PART_DESC = uctlPartSubLookup.MyCombo.Text
   End If
   
   EnpAddress.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   
   EnpAddress.ImportItem.LOCATION_ID = uctlLocationToLookup.MyCombo.ItemData(Minus2Zero(uctlLocationToLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.LOCATION_NAME = uctlLocationToLookup.MyCombo.Text
   
   EnpAddress.ExportItem.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.ExportItem.TOTAL_INCLUDE_PRICE = 0
   EnpAddress.ExportItem.AVG_PRICE = 0
   
   EnpAddress.ImportItem.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = 0
   EnpAddress.ImportItem.AVG_PRICE = 0
   
   Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(EnpAddress.ExportItem.PART_ITEM_ID)), False)
   If Not (Pi Is Nothing) Then
      EnpAddress.ExportItem.AVG_PRICE = Pi.COST_PER_AMOUNT
      EnpAddress.ExportItem.TOTAL_INCLUDE_PRICE = EnpAddress.ExportItem.AVG_PRICE * EnpAddress.ExportItem.TX_AMOUNT
      EnpAddress.ExportItem.NEW_AVG_PRICE = EnpAddress.ExportItem.AVG_PRICE
      
      EnpAddress.ImportItem.AVG_PRICE = Pi.COST_PER_AMOUNT
      EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = EnpAddress.ImportItem.AVG_PRICE * EnpAddress.ImportItem.TX_AMOUNT
      EnpAddress.ImportItem.NEW_AVG_PRICE = EnpAddress.ImportItem.AVG_PRICE
   End If
   
   EnpAddress.ExportItem.UNIT_TRAN_ID = UnitID
   EnpAddress.ExportItem.UNIT_MULTIPLE = Multiple
   EnpAddress.ExportItem.UNIT_TRAN_NAME = UnitName
   
   EnpAddress.ImportItem.UNIT_TRAN_ID = UnitID
   EnpAddress.ImportItem.UNIT_MULTIPLE = Multiple
   EnpAddress.ImportItem.UNIT_TRAN_NAME = UnitName
   
   EnpAddress.ExportItem.TX_TYPE = "E"
   EnpAddress.ExportItem.MULTIPLIER = -1
   
   EnpAddress.ImportItem.TX_TYPE = "I"
   EnpAddress.ImportItem.MULTIPLIER = 1
   
   Set EnpAddress = Nothing
   
   Call RefreshGrid(DocumentType, False)
   
   txtQuantity.Text = ""
   uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, -1)
   uctlPartLookup.SetFocus
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadStockBarcode(uctlPartLookup.MyCombo, m_Parts)
      Set uctlPartLookup.MyCollection = m_Parts
      
      Call LoadStockBarcode(uctlPartSubLookup.MyCombo, m_Parts)
      Set uctlPartSubLookup.MyCollection = m_Parts
            
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadMaster(uctlLocationToLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationToLookup.MyCollection = m_Locations
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      m_InventoryDoc.ShowMode = SHOW_ADD
      
      uctlDocumentDate.ShowDate = Now
      Call cmdAuto_Click
      
      Call EnableForm(Me, True)
      m_HasModify = False
      
      uctlLocationLookup.SetFocus
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
      'Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      'Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Cd = Nothing
    Set m_Parts = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As Long)
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
   
   If Ind = TRANSFER_DOCTYPE Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("����ʵ�ͤ")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("��¡��")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("����ҳ")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("˹���")
   
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("�ҡʶҹ���Ѵ��")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 1995
      Col.Caption = MapText("�ʶҹ���Ѵ��")
      
   End If
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("�Ţ����͡���"))
   Call InitNormalLabel(lblTotalAmount, MapText("���"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(lblPart, MapText("�Թ���/�ѵ�شԺ"))
   Call InitNormalLabel(lblQuantity, MapText("�ӹǹ"))
   Call InitNormalLabel(lblLocationLookup, MapText("��ѧ"))
   
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   Call InitNormalLabel(lblUnit, MapText(""))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtQuantity.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   uctlLocationLookup.MyCombo.TabStop = False
   uctlLocationToLookup.MyCombo.TabStop = False
   uctlPartLookup.MyCombo.TabStop = False
   uctlPartSubLookup.MyCombo.TabStop = False
   
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSaveBarcode.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdPrint, MapText("�����"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdSaveBarcode, MapText("�ѹ�֡����"))
   
   
   Call InitGrid1(DocumentType)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   If DocumentType = IMPORT_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = TRANSFER_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = ADJUST_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   End If
End Sub

Private Function Doctype2Text(TempID As INVENTORY_DOCTYPE) As String
   If TempID = IMPORT_DOCTYPE Then
      Doctype2Text = "��¡�ù����"
   ElseIf TempID = EXPORT_DOCTYPE Then
      Doctype2Text = "��¡���ԡ����"
   ElseIf TempID = TRANSFER_DOCTYPE Then
      Doctype2Text = "��¡���͹ʵ�ͤ"
   ElseIf TempID = ADJUST_DOCTYPE Then
      Doctype2Text = "��¡�ü�Ե/��Ѻ�ʹ"
   End If
End Function

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
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Cd = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
End Sub


Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If m_InventoryDoc.ImportExportItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If DocumentType = TRANSFER_DOCTYPE Then
         Dim TR As CTransferItem
         If m_InventoryDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set TR = GetItem(m_InventoryDoc.TransferItems, RowIndex, RealIndex)
         If TR Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = TR.ImportItem.LOT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = TR.ImportItem.PART_NO
         Values(4) = TR.ImportItem.PART_DESC
         Values(5) = FormatNumber(MyDiff(TR.ExportItem.TX_AMOUNT, TR.ExportItem.UNIT_MULTIPLE))
         Values(6) = TR.ExportItem.UNIT_TRAN_NAME
         Values(7) = TR.ExportItem.LOCATION_NAME
         Values(8) = TR.ImportItem.LOCATION_NAME
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid(Ind As INVENTORY_DOCTYPE, Flag As Boolean)
   If (Ind = IMPORT_DOCTYPE) Or (Ind = EXPORT_DOCTYPE) Or (Ind = ADJUST_DOCTYPE) Then
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExportItems)
      GridEX1.Rebind
   ElseIf Ind = TRANSFER_DOCTYPE Then
      GridEX1.ItemCount = CountItem(m_InventoryDoc.TransferItems)
      GridEX1.Rebind
   End If

   Call CalculateSumPrice
   If Flag Then
      m_HasModify = Flag
   End If
End Sub
Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      Call InitGrid1(DocumentType)
      Call RefreshGrid(DocumentType, False)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(INVENTORY_DOC_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub
Private Sub uctlDocumentDate_HasChange()
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
   cmdDelete.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
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

Private Sub uctlPartLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode
   
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
      UnitID = Pi.UNIT_ID
      Multiple = Pi.UNIT_AMOUNT
      UnitName = Pi.UNIT_NAME
      UnitMName = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
      
      uctlPartSubLookup.MyCombo.ListIndex = -1
   End If
   m_HasModify = True
End Sub
Private Sub uctlPartSubLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode
   
   PartItemID = uctlPartSubLookup.MyCombo.ItemData(Minus2Zero(uctlPartSubLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
      UnitID = Pi.UNIT_CHANGE_ID
      Multiple = 1
      UnitName = Pi.UNIT_CHANGE_NAME
      UnitMName = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
            
      uctlPartLookup.MyCombo.ListIndex = -1
   End If
   m_HasModify = True
End Sub
Private Sub CreateImportExportItems()
Dim Ti As CTransferItem
Dim Ei As CLotItem
Dim II As CLotItem

   Set m_InventoryDoc.ImportExportItems = Nothing
   Set m_InventoryDoc.ImportExportItems = New Collection
   
   For Each Ti In m_InventoryDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem
      
      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag
      
      Call m_InventoryDoc.ImportExportItems.add(Ei)
      Call m_InventoryDoc.ImportExportItems.add(II)
   Next Ti
End Sub
Private Sub PopulateGuiID(BD As CInventoryDoc)
Dim Di As CLotItem
Dim I As Long
Dim TempID As Long
   I = 0
   For Each Di In BD.ImportExportItems
      If Di.Flag = "A" Then
         I = I + 1
         If (I Mod 2) = 1 Then
            Di.LINK_ID = GetNextGuiID(BD)
            TempID = Di.LINK_ID
         Else
            Di.LINK_ID = TempID
         End If
         
      End If
   Next Di
End Sub
Private Function GetNextGuiID(BD As CInventoryDoc) As Long
Dim Di As CLotItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.ImportExportItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di

   GetNextGuiID = MaxId + 1
End Function

