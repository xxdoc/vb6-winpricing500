VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditInventoryDoc 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frmAddEditInventoryDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDriver 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3120
         Width           =   2385
      End
      Begin VB.ComboBox cboTranSportor 
         Height          =   315
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3120
         Width           =   2385
      End
      Begin VB.ComboBox cboCarLicense 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3120
         Width           =   2385
      End
      Begin VB.ComboBox cboLocationSale 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2640
         Width           =   1755
      End
      Begin VB.ComboBox cboDepartment 
         Height          =   315
         Left            =   9180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1740
         Width           =   1815
      End
      Begin Xivess.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1740
         TabIndex        =   3
         Top             =   1290
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Xivess.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   2
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   10
         Top             =   3600
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
         TabIndex        =   0
         Top             =   840
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
         Height          =   3645
         Left            =   130
         TabIndex        =   11
         Top             =   4080
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6429
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
         Column(1)       =   "frmAddEditInventoryDoc.frx":27A2
         Column(2)       =   "frmAddEditInventoryDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDoc.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtDeliveryFee 
         Height          =   435
         Left            =   1740
         TabIndex        =   4
         Top             =   1740
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtNote 
         Height          =   435
         Left            =   1740
         TabIndex        =   7
         Top             =   2190
         Width           =   5325
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   5580
         TabIndex        =   5
         Top             =   1740
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlAparMasLookup 
         Height          =   435
         Left            =   1740
         TabIndex        =   31
         Top             =   2640
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblDriver 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblTranSportor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         TabIndex        =   38
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblCarLicense 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   37
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblAparMas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   2700
         Width           =   1395
      End
      Begin VB.Label lblLocationSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         TabIndex        =   32
         Top             =   2730
         Width           =   1095
      End
      Begin Threed.SSCheck chkTripFlag 
         Height          =   435
         Left            =   9480
         TabIndex        =   29
         Top             =   1320
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkTempSaleFlag 
         Height          =   435
         Left            =   7680
         TabIndex        =   28
         Top             =   1320
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin Threed.SSCheck ChkCancelFlag 
         Height          =   435
         Left            =   9480
         TabIndex        =   9
         Top             =   2250
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   7260
         TabIndex        =   26
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         TabIndex        =   25
         Top             =   2250
         Width           =   1365
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1740
         TabIndex        =   1
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   7680
         TabIndex        =   8
         Top             =   2250
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblEmployeeNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   210
         TabIndex        =   24
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3330
         TabIndex        =   23
         Top             =   1830
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4890
         TabIndex        =   22
         Top             =   900
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryFee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -90
         TabIndex        =   20
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDoc"
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
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As INVENTORY_DOCTYPE
Public InventorySubType As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Employee As CEmployee
Private m_Mr As CMasterRef
Private m_Apm  As CAPARMas
Private m_APARCancelOutDocument As Collection
Private m_LocationSales As Collection
Private m_Cd As Collection
Private DocAdd As Long
Private typeDocNum As Long


Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", ID)
      Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", "")
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_InventoryDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_InventoryDoc.GetFieldValue("DOCUMENT_DATE")
      txtDocumentNo.Text = m_InventoryDoc.GetFieldValue("DOCUMENT_NO")
      txtDeliveryFee.Text = Format(m_InventoryDoc.GetFieldValue("DELIVERY_FEE"))
      uctlEmployeeLookUp.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookUp.MyCombo, m_InventoryDoc.GetFieldValue("EMP_ID"))
      chkCommit.Value = FlagToCheck(m_InventoryDoc.GetFieldValue("COMMIT_FLAG"))
      txtNote.Text = m_InventoryDoc.GetFieldValue("DOCUMENT_DESC")
      cboDepartment.ListIndex = IDToListIndex(cboDepartment, m_InventoryDoc.GetFieldValue("DEPARTMENT_ID"))
      ChkCancelFlag.Value = FlagToCheck(m_InventoryDoc.GetFieldValue("CANCEL_FLAG"))
      chkTempSaleFlag.Value = FlagToCheck(m_InventoryDoc.GetFieldValue("TEMP_SALE_FLAG"))
      chkTripFlag.Value = FlagToCheck(m_InventoryDoc.GetFieldValue("TRIP_FLAG"))
      
      uctlAparMasLookup.MyCombo.ListIndex = IDToListIndex(uctlAparMasLookup.MyCombo, m_InventoryDoc.GetFieldValue("APAR_MAS_ID"))
      cboLocationSale.ListIndex = IDToListIndex(cboLocationSale, m_InventoryDoc.GetFieldValue("CUSTOMER_BRANCH"))
      cboDriver.ListIndex = IDToListIndex(cboDriver, m_InventoryDoc.GetFieldValue("DRIVER_ID"))
      cboCarLicense.ListIndex = IDToListIndex(cboCarLicense, m_InventoryDoc.GetFieldValue("CAR_LICENSE_ID"))
      cboTranSportor.ListIndex = IDToListIndex(cboTranSportor, m_InventoryDoc.GetFieldValue("TRANSPORTOR_ID"))
      
      cmdAdd.Enabled = (m_InventoryDoc.GetFieldValue("COMMIT_FLAG") = "N")
      cmdDelete.Enabled = (m_InventoryDoc.GetFieldValue("COMMIT_FLAG") = "N")
      chkCommit.Enabled = (m_InventoryDoc.GetFieldValue("COMMIT_FLAG") = "N")
      
      If DocumentType = TRANSFER_DOCTYPE Then
         Call glbDaily.CreateTransferItems(m_InventoryDoc)
      End If
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Sub CalculateSumPrice()
Dim Li As CLotItem
Dim Ti As CTransferItem
Dim Sum As Double
Dim Sum2 As Double

   Sum = 0
   Sum2 = 0
   If DocumentType = TRANSFER_DOCTYPE Then
      For Each Ti In m_InventoryDoc.TransferItems
         If Ti.Flag <> "D" Then
            Sum = Sum + Ti.ImportItem.TOTAL_INCLUDE_PRICE
            Sum2 = Sum2 + Ti.ImportItem.TX_AMOUNT
         End If
      Next Ti
   Else
      For Each Li In m_InventoryDoc.ImportExportItems
         If Li.Flag <> "D" Then
            Sum = Sum + Li.TOTAL_INCLUDE_PRICE
            Sum2 = Sum2 + Li.TX_AMOUNT
         End If
      Next Li
   End If
   
   txtDeliveryFee.Text = Format(Sum, "0.00")
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
   If Not VerifyCombo(lblEmployeeNo, uctlEmployeeLookUp.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDeliveryFee, txtDeliveryFee, True) Then
      Exit Function
   End If
   
    If Not VerifyLockDate(uctlDocumentDate.ShowDate, m_InventoryDoc.GetFieldValue("DOCUMENT_DATE")) Then
      glbErrorLog.LocalErrorMsg = MapText("�������ö����¹�ŧ�͡��õ���ѹ����͡��÷�����͡�� ��سҵԴ��ͼ������к� ���ͼ�����Է����˹��ѹ����͡�����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not VerifyLockInventoryDate(uctlDocumentDate.ShowDate, m_InventoryDoc.GetFieldValue("DOCUMENT_DATE")) Then
      glbErrorLog.LocalErrorMsg = MapText("�������ö����¹�ŧ�͡��õ���ѹ����͡��÷�����͡�� ��سҵԴ��ͼ������к� ���ͼ�����Է����˹��ѹ����͡�����")
      glbErrorLog.ShowUserError
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
   
   m_InventoryDoc.ShowMode = ShowMode
   Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", ID)
    Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", uctlDocumentDate.ShowDate)
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   Call m_InventoryDoc.SetFieldValue("DELIVERY_FEE", 0)
   Call m_InventoryDoc.SetFieldValue("EMP_ID", uctlEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookUp.MyCombo.ListIndex)))
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", Check2Flag(chkCommit.Value))
   Call m_InventoryDoc.SetFieldValue("EXCEPTION_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("SALE_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("ADJUST_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", txtNote.Text)
   Call m_InventoryDoc.SetFieldValue("DEPARTMENT_ID", cboDepartment.ItemData(Minus2Zero(cboDepartment.ListIndex)))
   If ShowMode = SHOW_ADD Then     '੾��������ҹ�鹷��� SET Sub Type
      Call m_InventoryDoc.SetFieldValue("INVENTORY_SUB_TYPE", InventorySubType)
   End If
   Call m_InventoryDoc.SetFieldValue("CANCEL_FLAG", Check2Flag(ChkCancelFlag.Value))
   Call m_InventoryDoc.SetFieldValue("TEMP_SALE_FLAG", Check2Flag(chkTempSaleFlag.Value))
      
   Call m_InventoryDoc.SetFieldValue("TRIP_FLAG", Check2Flag(chkTripFlag.Value))
   
   Call m_InventoryDoc.SetFieldValue("APAR_MAS_ID", uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex)))
   If cboLocationSale.ListCount > 0 Then
      Call m_InventoryDoc.SetFieldValue("CUSTOMER_BRANCH", cboLocationSale.ItemData(Minus2Zero(cboLocationSale.ListIndex)))
   End If
   Call m_InventoryDoc.SetFieldValue("DRIVER_ID", cboDriver.ItemData(Minus2Zero(cboDriver.ListIndex)))
   Call m_InventoryDoc.SetFieldValue("CAR_LICENSE_ID", cboCarLicense.ItemData(Minus2Zero(cboCarLicense.ListIndex)))
   Call m_InventoryDoc.SetFieldValue("TRANSPORTOR_ID", cboTranSportor.ItemData(Minus2Zero(cboTranSportor.ListIndex)))
      
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
Private Sub cboCarLicense_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub
Private Sub cboDepartment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboInventorySubType_Click()
   m_HasModify = True
End Sub
Private Sub cboInventorySubType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboDriver_Click()
   m_HasModify = True
End Sub

Private Sub cboLocationSale_Click()
   m_HasModify = True
End Sub
Private Sub cboTransportor_Click()
   m_HasModify = True
End Sub

Private Sub ChkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkTempSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkTripFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If DocumentType = TRANSFER_DOCTYPE Then
         Set frmAddEditTransferItem.ParentForm = Me
         frmAddEditTransferItem.DocumentType = DocumentType
         frmAddEditTransferItem.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditTransferItem.TempCollection = m_InventoryDoc.TransferItems
         frmAddEditTransferItem.ShowMode = SHOW_ADD
         frmAddEditTransferItem.HeaderText = MapText("����" & Doctype2Text(DocumentType))
         Load frmAddEditTransferItem
         frmAddEditTransferItem.Show 1
   
         OKClick = frmAddEditTransferItem.OKClick
         
         Unload frmAddEditTransferItem
         Set frmAddEditTransferItem = Nothing
      ElseIf DocumentType = ADJUST_DOCTYPE Or DocumentType = 1000 Then
         Set frmAddEditAdjustItem.ParentForm = Me
         frmAddEditAdjustItem.DocumentType = DocumentType
         frmAddEditAdjustItem.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditAdjustItem.TempCollection = m_InventoryDoc.ImportExportItems
         frmAddEditAdjustItem.ParentShowMode = ShowMode
         frmAddEditAdjustItem.ShowMode = SHOW_ADD
         frmAddEditAdjustItem.HeaderText = MapText("����" & Doctype2Text(DocumentType))
         Load frmAddEditAdjustItem
         frmAddEditAdjustItem.Show 1
   
         OKClick = frmAddEditAdjustItem.OKClick
   
         Unload frmAddEditAdjustItem
         Set frmAddEditAdjustItem = Nothing
      Else
         Set frmAddEditImportItemEx.ParentForm = Me
         frmAddEditImportItemEx.DocumentType = DocumentType
         frmAddEditImportItemEx.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditImportItemEx.TempCollection = m_InventoryDoc.ImportExportItems
         frmAddEditImportItemEx.ParentShowMode = ShowMode
         frmAddEditImportItemEx.ShowMode = SHOW_ADD
         frmAddEditImportItemEx.HeaderText = MapText("����" & Doctype2Text(DocumentType))
         Load frmAddEditImportItemEx
         frmAddEditImportItemEx.Show 1
   
         OKClick = frmAddEditImportItemEx.OKClick
   
         Unload frmAddEditImportItemEx
         Set frmAddEditImportItemEx = Nothing
      End If
      
      If OKClick Then
         Call RefreshGrid(DocumentType, False)
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
      If (DocumentType = IMPORT_DOCTYPE) Or (DocumentType = EXPORT_DOCTYPE) Or (DocumentType = ADJUST_DOCTYPE) Or (DocumentType = 1000) Then
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
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (DocumentType = IMPORT_DOCTYPE) Or (DocumentType = EXPORT_DOCTYPE) Then
         frmAddEditImportItemEx.DocumentType = DocumentType
         Set frmAddEditImportItemEx.ParentForm = Me
         frmAddEditImportItemEx.ID = ID
         frmAddEditImportItemEx.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditImportItemEx.TempCollection = m_InventoryDoc.ImportExportItems
         frmAddEditImportItemEx.HeaderText = MapText("���" & Doctype2Text(DocumentType))
         frmAddEditImportItemEx.ParentShowMode = ShowMode
         frmAddEditImportItemEx.ShowMode = SHOW_EDIT
         Load frmAddEditImportItemEx
         frmAddEditImportItemEx.Show 1
   
         OKClick = frmAddEditImportItemEx.OKClick
   
         Unload frmAddEditImportItemEx
         Set frmAddEditImportItemEx = Nothing
      ElseIf DocumentType = TRANSFER_DOCTYPE Then
         frmAddEditTransferItem.DocumentType = DocumentType
         Set frmAddEditTransferItem.ParentForm = Me
         frmAddEditTransferItem.ID = ID
         frmAddEditTransferItem.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditTransferItem.TempCollection = m_InventoryDoc.TransferItems
         frmAddEditTransferItem.ShowMode = SHOW_EDIT
         frmAddEditTransferItem.HeaderText = MapText("���" & Doctype2Text(DocumentType))
         Load frmAddEditTransferItem
         frmAddEditTransferItem.Show 1
         
         OKClick = frmAddEditTransferItem.OKClick
         
         Unload frmAddEditTransferItem
         Set frmAddEditTransferItem = Nothing
      ElseIf DocumentType = ADJUST_DOCTYPE Or DocumentType = 1000 Then
         frmAddEditAdjustItem.DocumentType = DocumentType
         Set frmAddEditAdjustItem.ParentForm = Me
         frmAddEditAdjustItem.ID = ID
         frmAddEditAdjustItem.COMMIT_FLAG = m_InventoryDoc.GetFieldValue("COMMIT_FLAG")
         Set frmAddEditAdjustItem.TempCollection = m_InventoryDoc.ImportExportItems
         frmAddEditAdjustItem.ParentShowMode = ShowMode
         frmAddEditAdjustItem.ShowMode = SHOW_EDIT
         frmAddEditAdjustItem.HeaderText = MapText("���" & Doctype2Text(DocumentType))
         Load frmAddEditAdjustItem
         frmAddEditAdjustItem.Show 1
   
         OKClick = frmAddEditAdjustItem.OKClick
   
         Unload frmAddEditAdjustItem
         Set frmAddEditAdjustItem = Nothing
      End If
      
      If OKClick Then
         Call RefreshGrid(DocumentType, True)
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

Private Sub cmdOK_Click()
Dim oMenu As CPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("�ѹ�֡", "-", "�ѹ�֡����͡�ҡ˹�Ҩ�")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_InventoryDoc.GetFieldValue("INVENTORY_DOC_ID")
      m_InventoryDoc.QueryFlag = 1
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

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim OKClick As Boolean
Dim iCount As Long
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   
   If m_HasModify Or (m_InventoryDoc.GetFieldValue("INVENTORY_DOC_ID") <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("��سҷӡ�úѹ�֡������������º���¡�͹")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New CPopupMenu
'   If DocumentType = IMPORT_DOCTYPE Then
'===========================================��ǹ�������¹���ͷ� Report ========================================================
'-------------------����  typeDocNum  ���� �繪��� popup ������� ������͡ ���ç�Ѻ �͡��÷��о����----------------
'-------------------����  lMenuChosen  ������͡ �ҡ����� popup ����� �ҡ��----------------
'-------------------------------1. CReportNormalExportPromotion    ��͡�þ���� ��ԡ   �� return ��� 1 ,3
'-------------------------------2. CReportNormalTransferStock   ��͡�þ���� ��͹ʵ�͡   �� return ��� 1 ,3
'------------------------------*** ���������������ѹ�Ѻ�ѹ ��ͧ������ҵ�ҧ�ѹ �����¤�� +100���� �ѧ��� ��ҷ�� return 㹢�� 2 �֧��  101 ,103 ****----
   If typeDocNum = 1 Then
      lMenuChosen = oMenu.Popup("���������������â��", "-", "��Ѻ���˹�ҡ�д��")
   ElseIf typeDocNum = 2 Then
         lMenuChosen = oMenu.Popup("�������͹ʵ�͡", "-", "��Ѻ���˹�ҡ�д��")
         lMenuChosen = lMenuChosen + 100
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   ReportMode = 1
   ReportFlag = False
   
   

   If lMenuChosen = 1 Then
       ReportKey = "CReportNormalExportPromotion"
       Set Report = New CReportNormalExportPromotion
       ReportFlag = True
   ElseIf lMenuChosen = 3 Then
      ReportKey = "CReportNormalExportPromotion"
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)

      HeaderText = MapText("��Ѻ���˹�ҡ�д��")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
      
 ElseIf lMenuChosen = 101 Then
      ReportKey = "CReportNormalTransferStock"
      Set Report = New CReportNormalTransferStock
      ReportFlag = True
                  
ElseIf lMenuChosen = 103 Then
      ReportMode = 1
      ReportKey = "CReportNormalTransferStock"
              
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      
       HeaderText = MapText("��͹��ѧ")
  
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
      
   Else
      Exit Sub
   End If
'========================================================================================================
   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_InventoryDoc.GetFieldValue("INVENTORY_DOC_ID"), "INVENTORY_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   Call EnableForm(Me, False)
   If ReportFlag Then
      frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.GetFieldValue("REPORT_CONFIG_ID")
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1

      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = -1
      Call LoadEmployee(m_Employee, uctlEmployeeLookUp.MyCombo)
      Set uctlEmployeeLookUp.MyCollection = m_EmployeeColl
            
      Call LoadMaster(cboDepartment, , , , MASTER_DEPARTMENT)
            
      Call LoadApArMas(m_Apm, uctlAparMasLookup.MyCombo, m_APARCancelOutDocument)
      Set uctlAparMasLookup.MyCollection = m_CustomerColl
               
      Call LoadMaster(cboDriver, , , , MASTER_DRIVER)
      Call LoadMaster(cboCarLicense, , , , MASTER_CAR_LICENSE)
      Call LoadMaster(cboTranSportor, , , , MASTER_TRANSPORTOR)
         
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         Call cmdAuto_Click
         m_InventoryDoc.QueryFlag = 0
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
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Employees = Nothing
   Set m_Employee = Nothing
   Set m_Cd = Nothing
   Set m_Apm = Nothing
   Set m_APARCancelOutDocument = Nothing
   Set m_LocationSales = Nothing
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
   
   If Ind = IMPORT_DOCTYPE Then
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
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ�")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ����")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("ʶҹ���Ѵ��")
            
   ElseIf Ind = EXPORT_DOCTYPE Then
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
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ�")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ����")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("ʶҹ���Ѵ��")
            
   ElseIf Ind = TRANSFER_DOCTYPE Then
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
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ�")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ����")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("�ҡʶҹ���Ѵ��")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 1995
      Col.Caption = MapText("�ʶҹ���Ѵ��")
      
   ElseIf Ind = ADJUST_DOCTYPE Or DocumentType = 1000 Then
      
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1000
      Col.Caption = MapText("������")
      
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
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ�")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("�Ҥ����")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("ʶҹ���Ѵ��")
   End If
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("�Ţ����͡���"))
   Call InitNormalLabel(lblDeliveryFee, MapText("��Ť�����"))
   Call InitNormalLabel(lblTotalAmount, MapText("�ӹǹ���"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(lblEmployeeNo, MapText("����Ѻ�Դ�ͺ"))
   Call InitNormalLabel(lblNote, MapText("�����˵�"))
   Call InitNormalLabel(lblDepartment, MapText("Ἱ�"))
   
   Call InitNormalLabel(lblLocationSale, MapText("�Ң�"))
   Call InitNormalLabel(lblDriver, MapText("���Ѻ"))
   Call InitNormalLabel(lblCarLicense, MapText("����¹"))
   Call InitNormalLabel(lblTranSportor, MapText("����"))
   Call InitNormalLabel(lblAparMas, MapText("�����١���"))
   Call InitCombo(cboDriver)
   Call InitCombo(cboCarLicense)
   Call InitCombo(cboTranSportor)
   Call InitCombo(cboLocationSale)
         
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDeliveryFee.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDeliveryFee.Enabled = False
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
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
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdPrint, MapText("�����"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitCheckBox(chkCommit, MapText("�������"))
   Call InitCheckBox(ChkCancelFlag, "CANCEL")
   Call InitCheckBox(chkTempSaleFlag, "�ҡ���")
   Call InitCheckBox(chkTripFlag, "�Ѵ��Ի")
   
  '================================��ǹ�������¹���ͷ� Report ========================================
   '---------------------------------������ print �¡�� visible ���Ѻ���������ҵ�ͧ�������ա�á��������  ��� DocumentType
   '---------------------------------㹷��������ö������� Report ��  DocumentType 2 �ѹ ���  EXPORT_DOCTYPE  , TRANSFER_DOCTYPE
   '---------------------------------�͡��� invisible ���� visible =false ���
   If DocumentType = EXPORT_DOCTYPE Then      '  visible
      cmdPrint.Visible = True
      typeDocNum = 1
 ElseIf DocumentType = TRANSFER_DOCTYPE Then     '  visible
     cmdPrint.Visible = True
       typeDocNum = 2
   Else                             '  invisible
      cmdPrint.Visible = False
   End If
'==============================================================================================
   Call InitCombo(cboDepartment)
   
   Call InitGrid1(DocumentType)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   chkTempSaleFlag.Visible = False
   
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
      
      chkTempSaleFlag.Visible = True
   ElseIf DocumentType = ADJUST_DOCTYPE Or (DocumentType = 1000) Then
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
      Doctype2Text = "��¡�û�Ѻ�ʹʵ�ͤ"
   ElseIf TempID = 1000 Then
      Doctype2Text = "��¡�ü�Ե"
   End If
End Function
'Private Function typeDocument(typeDoc As Long)
'   If typeDoc = IMPORT_DOCTYPE Then
'     typeDocNum = 1
'   ElseIf TempID = EXPORT_DOCTYPE Then
'     typeDocNum = 2
'   ElseIf TempID = TRANSFER_DOCTYPE Then
'      typeDocNum = 3
'   ElseIf TempID = ADJUST_DOCTYPE Then
'      typeDocNum = 4
'
'   End If
'End Function

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
   Set m_Employee = New CEmployee
   Set m_Employees = New Collection
   Set m_Mr = New CMasterRef
   Set m_Cd = New Collection
   Set m_Apm = New CAPARMas
   Set m_APARCancelOutDocument = New Collection
   Set m_LocationSales = New Collection
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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

      If (DocumentType = IMPORT_DOCTYPE) Or (DocumentType = EXPORT_DOCTYPE) Then
         Dim CR As CLotItem
         If m_InventoryDoc.ImportExportItems.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_InventoryDoc.ImportExportItems, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
         
         Values(1) = CR.LOT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.PART_NO
         Values(4) = CR.PART_DESC
         Values(5) = FormatNumber(MyDiff(CR.TX_AMOUNT, CR.UNIT_MULTIPLE))
         Values(6) = FormatNumber(CR.AVG_PRICE * CR.UNIT_MULTIPLE)
         Values(7) = FormatNumber(CR.TOTAL_INCLUDE_PRICE)
         Values(8) = CR.LOCATION_NAME
      ElseIf (DocumentType = ADJUST_DOCTYPE) Or (DocumentType = 1000) Then
         Dim Aj As CLotItem
         If m_InventoryDoc.ImportExportItems.Count <= 0 Then
            Exit Sub
         End If
         Set Aj = GetItem(m_InventoryDoc.ImportExportItems, RowIndex, RealIndex)
         If Aj Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Aj.LOT_ITEM_ID
         Values(2) = RealIndex
         If Aj.TX_TYPE = "E" Then
            Values(3) = "��ѺŴ"
         Else
            Values(3) = "��Ѻ����"
         End If
         Values(4) = Aj.PART_NO
         Values(5) = Aj.PART_DESC
         Values(6) = FormatNumber(MyDiff(Aj.TX_AMOUNT, Aj.UNIT_MULTIPLE))
         Values(7) = FormatNumber(Aj.AVG_PRICE * Aj.UNIT_MULTIPLE)
         Values(8) = FormatNumber(Aj.TOTAL_INCLUDE_PRICE)
         Values(9) = Aj.LOCATION_NAME
      ElseIf DocumentType = TRANSFER_DOCTYPE Then
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
         Values(6) = FormatNumber(TR.ExportItem.AVG_PRICE * TR.ExportItem.UNIT_MULTIPLE)
         Values(7) = FormatNumber(TR.ExportItem.TOTAL_INCLUDE_PRICE)
         Values(8) = TR.ExportItem.LOCATION_NAME
         Values(9) = TR.ImportItem.LOCATION_NAME
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
   If (Ind = IMPORT_DOCTYPE) Or (Ind = EXPORT_DOCTYPE) Or (Ind = ADJUST_DOCTYPE) Or (Ind = 1000) Then
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

Private Sub txtDeliveryFee_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub
Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(INVENTORY_DOC_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub
Private Sub txtNote_Change()
   m_HasModify = True
End Sub
Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlApArMasLookup_Change()
Dim ID As Long
Static OldID As Long
Dim AparMas As CAPARMas
   
   ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
    
   If ID = OldID Then
      Exit Sub
   End If
   If ID > 0 Then
      OldID = ID
      Set AparMas = m_CustomerColl(Trim(Str(ID)))
            
      m_HasModify = True
      
      Call LoadMaster(cboLocationSale, m_LocationSales, , , MASTER_APARMAS_BRANCH, ID)
         
      If m_LocationSales.Count > 0 Then
         Dim BranchID As Long
         Dim Branch As CMasterRef
         BranchID = cboLocationSale.ItemData(Minus2Zero(1))
         If AparMas.LABEL_FLAG = "N" Then
            cboLocationSale.Enabled = True
            cboLocationSale.ListIndex = 1
         Else
            cboLocationSale.Enabled = False
         End If
      End If
   End If
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
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

