VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditDoItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5145
   ClientLeft      =   4335
   ClientTop       =   240
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditDOItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSOption SSOption2 
         Height          =   255
         Left            =   11880
         TabIndex        =   52
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   131073
         Caption         =   "SSOption1"
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   6375
      Left            =   0
      TabIndex        =   29
      Top             =   5160
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   11245
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin GridEX20.GridEX GridEX1 
         Height          =   2835
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   5001
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
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditDOItem.frx":08CA
         Column(2)       =   "frmAddEditDOItem.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditDOItem.frx":0A36
         FormatStyle(2)  =   "frmAddEditDOItem.frx":0B92
         FormatStyle(3)  =   "frmAddEditDOItem.frx":0C42
         FormatStyle(4)  =   "frmAddEditDOItem.frx":0CF6
         FormatStyle(5)  =   "frmAddEditDOItem.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditDOItem.frx":0E86
      End
      Begin Xivess.uctlTextLookup uctlBlock 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3000
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtItemAmount 
         Height          =   435
         Left            =   12120
         TabIndex        =   21
         Top             =   3480
         Width           =   915
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlBranch 
         Height          =   435
         Left            =   1800
         TabIndex        =   20
         Top             =   3480
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   12120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3000
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtSumAmount 
         Height          =   435
         Left            =   12120
         TabIndex        =   27
         Top             =   3960
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlSale 
         Height          =   435
         Left            =   1800
         TabIndex        =   67
         Top             =   3960
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   66
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label lblUnitSum 
         Height          =   375
         Left            =   13200
         TabIndex        =   54
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label lblSumAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11040
         TabIndex        =   53
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   11040
         TabIndex        =   33
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label lblBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext2 
         Height          =   525
         Left            =   5160
         TabIndex        =   23
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   11040
         TabIndex        =   30
         Top             =   3000
         Width           =   1005
      End
      Begin Threed.SSCommand cmdAdd2 
         Height          =   525
         Left            =   120
         TabIndex        =   24
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit2 
         Height          =   525
         Left            =   1800
         TabIndex        =   25
         Top             =   4560
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete2 
         Height          =   525
         Left            =   3480
         TabIndex        =   26
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":1378
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4575
      Left            =   0
      TabIndex        =   34
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   8070
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDocItemType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11880
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cboSellType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   1515
      End
      Begin Xivess.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1350
         TabIndex        =   6
         Top             =   2520
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1350
         TabIndex        =   3
         Top             =   1080
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtAvgPriceEx 
         Height          =   435
         Left            =   12000
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1365
         TabIndex        =   11
         Top             =   2970
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1365
         TabIndex        =   4
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtManual 
         Height          =   465
         Left            =   8520
         TabIndex        =   1
         Top             =   120
         Width           =   5355
         _ExtentX        =   7964
         _ExtentY        =   820
      End
      Begin Xivess.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   11040
         TabIndex        =   13
         Top             =   3000
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtLeft 
         Height          =   435
         Left            =   11040
         TabIndex        =   14
         Top             =   3480
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlProductTypeLookup 
         Height          =   435
         Left            =   1365
         TabIndex        =   2
         Top             =   600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   10440
         TabIndex        =   12
         Top             =   3000
         Width           =   555
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   11040
         TabIndex        =   8
         Top             =   2520
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlProductReturn 
         Height          =   435
         Left            =   1350
         TabIndex        =   5
         Top             =   2040
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtRefNo 
         Height          =   435
         Left            =   11880
         TabIndex        =   58
         Top             =   1080
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtProDuctDetail 
         Height          =   435
         Left            =   10200
         TabIndex        =   61
         Top             =   1560
         Width           =   2280
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtWeightAmount 
         Height          =   435
         Left            =   1365
         TabIndex        =   15
         Top             =   3420
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   1245
      End
      Begin Threed.SSCheck ChkLotitemFlag 
         Height          =   435
         Left            =   6960
         TabIndex        =   65
         Top             =   2040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
         _Version        =   131073
         Enabled         =   0   'False
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdAddLotItem 
         Height          =   405
         Left            =   12600
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":1692
         ButtonStyle     =   3
      End
      Begin VB.Label lblWeightAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   63
         Top             =   3540
         Width           =   1095
      End
      Begin Threed.SSCommand cmdUnit 
         Height          =   435
         Left            =   3360
         TabIndex        =   7
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":19AC
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkFree 
         Height          =   435
         Left            =   3000
         TabIndex        =   62
         Top             =   120
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdBrowse 
         Height          =   405
         Left            =   13440
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   11040
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDocItemType 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   10320
         TabIndex        =   56
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblProductReturn 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   120
         TabIndex        =   55
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4005
         TabIndex        =   50
         Top             =   2640
         Width           =   1965
      End
      Begin VB.Label Label5 
         Height          =   345
         Left            =   13200
         TabIndex        =   49
         Top             =   3480
         Width           =   855
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   8880
         TabIndex        =   16
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblSellType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   105
         TabIndex        =   48
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   45
         TabIndex        =   47
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   13200
         TabIndex        =   46
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9720
         TabIndex        =   45
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Height          =   345
         Left            =   3405
         TabIndex        =   44
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   13200
         TabIndex        =   43
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9600
         TabIndex        =   42
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblManual 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7320
         TabIndex        =   41
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   165
         TabIndex        =   39
         Top             =   3090
         Width           =   1050
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9720
         TabIndex        =   38
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblAvgPriceEx 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   10800
         TabIndex        =   37
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   135
         TabIndex        =   36
         Top             =   1200
         Width           =   1125
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10560
         TabIndex        =   17
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12240
         TabIndex        =   18
         Top             =   3960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   35
         Top             =   2610
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmAddEditDoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ShowMode2 As SHOW_MODE_TYPE

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ParentForm As Form
Public HeaderText As String
Public ID As Long
Public ID2 As Long
Public OKClick As Boolean

Public HoldFlag As String

Public TempCollection As Collection
Public TempCollection2 As Collection
Public LotItemLinkCollection As Collection

Public Area As Long

Private m_ProductTypes As Collection
Private m_PartTypes As Collection
Private m_Products As Collection
Private m_ProductReturns As Collection
Private m_Locations As Collection

Private m_EmployeeColl As Collection
Private TempEmp As CEmployee

Public DocumentDate As Date
Public DocumentType As SELL_BILLING_DOCTYPE

Public CusID As Long
Private UNIT As String
Private m_Mr As CMasterRef
Private m_Branchs  As Collection
Private m_Blocks As Collection
Private m_Sc As CStockCode
'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'--------------------------------------------------

Private Sub cboDocItemType_Click()
   m_HasModify = True
End Sub
Private Sub cboSellType_Click()
Dim TempID As Long

   TempID = cboSellType.ItemData(Minus2Zero(cboSellType.ListIndex))
   If TempID = 1 Then
      
      Call LoadMaster(uctlProductTypeLookup.MyCombo, m_ProductTypes, , , MASTER_STOCKTYPE)
      Set uctlProductTypeLookup.MyCollection = m_ProductTypes
      
      txtManual.Enabled = False
      uctlLocationLookup.Enabled = True
      uctlProductTypeLookup.Enabled = True
      uctlProductLookup.Enabled = True
   ElseIf TempID = 2 Then
      txtManual.Enabled = False
      uctlLocationLookup.Enabled = False
      uctlLocationLookup.MyCombo.ListIndex = -1
      uctlLocationLookup.MyTextBox.Text = ""
      uctlProductTypeLookup.Enabled = True
      uctlProductLookup.Enabled = True
   Else
      txtManual.Enabled = True
      uctlLocationLookup.Enabled = False
      uctlProductTypeLookup.Enabled = False
      uctlProductLookup.Enabled = False
      uctlProductReturn.Enabled = False
      txtAvgPriceEx.Enabled = False
   End If
End Sub

Private Sub cboSellType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkFree_Click(Value As Integer)
   m_HasModify = True
   If Value = 1 Then
      txtAvgPrice.Text = ""
   End If
End Sub
Private Sub ChkLotitemFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAddLotItem_Click()
   If Not cmdAddLotItem.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   frmAddDocLotItem.TempPartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   frmAddDocLotItem.TempLocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   Set frmAddDocLotItem.TempCollection = LotItemLinkCollection
   frmAddDocLotItem.DocumentDate = Now
   frmAddDocLotItem.HeaderText = "เลือก LOT ที่ต้องการ"
   Load frmAddDocLotItem
   frmAddDocLotItem.Show 1
      
   Unload frmAddDocLotItem
   Set frmAddDocLotItem = Nothing
   
   Call CalulateTotatLotAmount
End Sub
Private Sub CalulateTotatLotAmount()
Dim Lk  As CDocItemLink
Dim Sum As Double
   Sum = 0
   For Each Lk In LotItemLinkCollection
      Sum = Sum + Lk.IMPORT_AMOUNT
   Next Lk
   txtQuantity.Text = ""
   txtQuantity.Text = Sum
   If LotItemLinkCollection.Count > 0 Then
      txtQuantity.Enabled = False
   End If
End Sub
Private Sub cmdBrowse_Click()
   frmAddReturnItem.CusID = CusID
   frmAddReturnItem.DocumentType = DocumentType
   frmAddReturnItem.ShowMode = SHOW_VIEW_ONLY
   frmAddReturnItem.HeaderText = MapText("หมายเลขใบส่งของสำหรับรับคืนสินค้า")
   Load frmAddReturnItem
   frmAddReturnItem.Show 1
   
   OKClick = frmAddReturnItem.OKClick
   txtRefNo.Text = frmAddReturnItem.PO_NO
   cmdBrowse.Tag = frmAddReturnItem.PO_ID
   
   Unload frmAddReturnItem
   Set frmAddReturnItem = Nothing
   
   Call uctlProductLookup.MyTextBox.SetFocus
End Sub

Private Sub cmdEdit2_Click()
   ShowMode2 = SHOW_EDIT
   
   If Not cmdEdit2.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID2 = Val(GridEX1.Value(2))
   
   Call QueryData2(True)
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblWeightAmount, MapText("นน."))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblLocation, MapText("ที่จัดเก็บ"))
   Call InitNormalLabel(lblAvgPriceEx, MapText("ตท@"))
   'Call InitNormalLabel(Label1, MapText("ก.ก."))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblAvgPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblProductType, MapText("ประเภท"))
   Call InitNormalLabel(lblProduct, MapText("รายการ"))
   Call InitNormalLabel(lblProductReturn, MapText("รายการคืน"))
   Call InitNormalLabel(lblManual, MapText("กำหนดเอง"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitNormalLabel(Label7, MapText("บาท"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label5, MapText("บาท"))
   Call InitNormalLabel(lblSellType, MapText("สินค้า"))
   Call InitNormalLabel(lblUnit, MapText(""))
   Call InitNormalLabel(lblBlock, MapText("บล็อค"))
   Call InitNormalLabel(lblBranch, MapText("สาขา"))
   Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblItemAmount, MapText("จ.หน่วย"))
   Call InitNormalLabel(lblPackAmount, MapText("จ.ตะกร้า"))
   Call InitNormalLabel(lblSumAmount, MapText("จำนวนในรายการ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วย"))
   Call InitNormalLabel(lblUnitSum, MapText("หน่วย"))
   Call InitNormalLabel(lblDocItemType, MapText("ประเภทรายการ"))
   Call InitNormalLabel(lblRefNo, MapText("DO NO."))
 '  lblWarningUnavailableSale.Visible = False
'   Call InitNormalLabel(lblWarningUnavailableSale, MapText("สินค้างดจำหน่ายชั่วคราว !!"))
    
'       Dim PD   As CPackageDetail
'       Set PD = LoadPackageColl(Trim(Str(ID)))
'
'   If HoldFlag = "N" Then
'        lblWarningUnavailableSale.Visible = False
'     '   Call InitNormalLabel(lblWarningUnavailableSale, MapText("สินค้างดจำหน่ายชั่วคราว !!"))
'   End If
'
   
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtItemAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtSumAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtSumAmount.Enabled = False
   
   Call txtDiscountPercent.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtAvgPriceEx.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtManual.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtRefNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtProDuctDetail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWeightAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   
   txtLeft.Enabled = False
   lblRefNo.Enabled = False
   txtRefNo.Enabled = False
   Call InitCheckBox(chkFree, "ฟรี")
   Call InitCheckBox(ChkLotitemFlag, "ลิงค์คลัง")
   
   If DocumentType = S_PO_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Then
      txtWeightAmount.Enabled = True
   Else
      txtWeightAmount.Enabled = False
   End If
   
   If DocumentType = RETURN_DOCTYPE Then
      txtAvgPriceEx.Visible = True
      lblAvgPriceEx.Visible = True
      lblRefNo.Visible = True
      txtRefNo.Visible = True
      cmdBrowse.Visible = True
      ChkLotitemFlag.Enabled = True
      ChkLotitemFlag.Value = ssCBChecked
   ElseIf DocumentType = S_RETURN_DOCTYPE Then
      lblRefNo.Visible = True
      txtRefNo.Visible = True
      cmdBrowse.Visible = True
      ChkLotitemFlag.Enabled = True
      ChkLotitemFlag.Value = ssCBChecked
   Else
      lblRefNo.Visible = False
      txtRefNo.Visible = False
      cmdBrowse.Visible = False
   End If
   
   lblProductReturn.Enabled = False
   uctlProductReturn.Enabled = False
      
   Call InitCombo(cboSellType)
   Call InitCombo(cboDocItemType)
   Call InitOptionEx(SSOption1, "รายละเอียดหลัก(F9)")
   Call InitOptionEx(SSOption2, "รายละเอียดย่อย(F9)")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdAdd2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdBrowse.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUnit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAddLotItem.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdAdd2, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit2, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete2, MapText("ลบ (F6)"))
   Call InitMainButton(cmdNext2, MapText("ถัดไป"))
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
   Call InitMainButton(cmdBrowse, MapText("B"))
   Call InitMainButton(cmdUnit, MapText("U"))
   Call InitMainButton(cmdAddLotItem, MapText("F5"))
   InitGrid2
End Sub
Private Sub CalculatePrice()
   txtLeft.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim iCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CDocItem
         Set Di = TempCollection.Item(ID)
         cboSellType.ListIndex = IDToListIndex(cboSellType, Di.GetFieldValue("SELL_TYPE"))
         cboDocItemType.ListIndex = IDToListIndex(cboDocItemType, Di.GetFieldValue("DOC_ITEM_TYPE"))
          txtManual.Text = Di.GetFieldValue("ITEM_DESC")
          uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, Di.GetFieldValue("STOCK_TYPE"))
          uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Di.GetFieldValue("PART_ITEM_ID"))
          If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
            uctlProductReturn.MyCombo.ListIndex = IDToListIndex(uctlProductReturn.MyCombo, Di.GetFieldValue("PART_ITEM_RETURN_ID"))
          End If
          uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, Di.GetFieldValue("LOCATION_ID"))
          txtQuantity.Text = MyDiffEx(Di.GetFieldValue("ITEM_AMOUNT"), Di.GetFieldValue("UNIT_MULTIPLE"))
          txtAvgPrice.Text = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("UNIT_MULTIPLE")
          txtAvgPriceEx.Text = Di.GetFieldValue("AVG_PRICE_EX") * Di.GetFieldValue("UNIT_MULTIPLE")
          txtTotalPrice.Text = Di.GetFieldValue("TOTAL_PRICE")
          
         UnitID = Di.GetFieldValue("UNIT_TRAN_ID")
         Multiple = Di.GetFieldValue("UNIT_MULTIPLE")
         UnitName = Di.GetFieldValue("UNIT_TRAN_NAME")
          
         Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
          
          txtDiscountPercent.Text = Di.GetFieldValue("DISCOUNT_PERCENT")
          txtDiscount.Text = Di.GetFieldValue("DISCOUNT_AMOUNT")
          
          txtRefNo.Text = Di.GetFieldValue("PO_NO")
          cmdBrowse.Tag = Di.GetFieldValue("PO_ID")
          txtProDuctDetail.Text = Di.GetFieldValue("PRODUCT_DETAIL")
          chkFree.Value = FlagToCheck(Di.GetFieldValue("FREE_FLAG"))
          ChkLotitemFlag.Value = FlagToCheck(Di.GetFieldValue("LOT_ITEM_FLAG"))
          
          txtWeightAmount.Text = Di.GetFieldValue("WEIGHT_AMOUNT")
          
'          Dim PD   As CPackageDetail
'        Set PD = New CPackageDetail
'      HoldFlag = PD.GetFieldValue("HOLD_FLAG")
      
          
          
          Set TempCollection2 = Di.PrintLabels
          Set LotItemLinkCollection = Di.LotItemLinkCollection
          Set Di = Nothing
      End If
      
   End If
   
   Call GetAmount
      
   GridEX1.itemcount = CountItem(TempCollection2)
   GridEX1.Rebind
      
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long
   If Not cmdNext.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      txtManual.Text = ""
      uctlProductLookup.MyCombo.ListIndex = -1
      txtQuantity.Text = ""
      txtAvgPrice.Text = ""
      txtDiscountPercent.Text = ""
      
      Set TempCollection2 = Nothing
      Set TempCollection2 = New Collection
      
      Set LotItemLinkCollection = Nothing
      Set LotItemLinkCollection = New Collection
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGrid
      
   SSOption1.Value = True
   Call uctlProductLookup.SetFocus
End Sub


Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim TempID As Long
Dim Pl As CPrintLabel
Dim SumAmount As Double
   TempID = cboSellType.ItemData(Minus2Zero(cboSellType.ListIndex))
   
   If Not VerifyCombo(lblSellType, cboSellType, False) Then
      Exit Function
   End If
   
   If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      If Not VerifyCombo(lblProduct, uctlProductReturn.MyCombo, Not (uctlProductReturn.Enabled)) Then
         Exit Function
      End If
      If Not VerifyTextControl(lblRefNo, txtRefNo, Not (txtRefNo.Visible)) Then
         Exit Function
      End If
   End If
   
   If TempID = 1 Or TempID = 2 Then
      If Not VerifyCombo(lblProductType, uctlProductTypeLookup.MyCombo, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, Not (uctlLocationLookup.Enabled)) Then
         Exit Function
      End If

   Else
      If Not VerifyTextControl(lblManual, txtManual, False) Then
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDiscount, txtDiscount, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ShowMode = SHOW_ADD Then
'      If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)), uctlProductLookup.MyTextBox.Text)) Then
'         Exit Function
'      End If
   End If
   
   ''------------------สำหรับตัดแบบ First In First out
'   Dim Part As CStockCode
'   Set Part = GetObject("CStockCode", m_Products, uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
'
'   '-------------------------------------------------------------------------------------------------
'   If Part.LOT_FLAG = "Y" And LotItemLinkCollection.Count <= 0 And (DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE) Then
'      If Not GenerateAutoLotLink() Then
'         glbErrorLog.LocalErrorMsg = "ไม่มีจำนวน " & uctlProductLookup.MyCombo.Text & " เพียงพอสำหรับเบิก"
'         glbErrorLog.ShowUserError
'         'Exit Function
'      End If
'   ElseIf Part.LOT_FLAG = "Y" And LotItemLinkCollection.Count > 0 And (DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE) Then
'      If Not CheckLotItemAmount() Then
'         glbErrorLog.LocalErrorMsg = "จำนวน " & uctlProductLookup.MyCombo.Text & "ตัด LOT กับยอดเบิกไม่เท่ากัน กรุณาแก้ไขจำนวนทั้งคู่ให้เท่ากัน"
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'    End If
''------------------สำหรับตัดแบบ First In First out
'Set Part = Nothing
'-------------------------------------------------------------------------------------------------
   
   Dim D As CAPARMas
   If Area = 2 Then
      Set D = m_SupplierColl(Trim(Str(CusID)))
   Else
      Set D = m_CustomerColl(Trim(Str(CusID)))
   End If
      
   If D.LABEL_FLAG = "Y" Then
      SumAmount = 0
      For Each Pl In TempCollection2
         If Pl.Flag <> "D" Then
            SumAmount = SumAmount + MyDiff(Pl.GetFieldValue("ITEM_AMOUNT"), Multiple)
         End If
      Next Pl
      If FormatNumber(SumAmount, , False) <> FormatNumber(Val(txtQuantity.Text), , False) Then
         Call MsgBox("กรณีใส่จำนวนตามบิลให้ตรงกับจำนวนที่ส่งตามสาขา", vbOKOnly, PROJECT_NAME)
         Exit Function
      End If
   End If
   
'    Dim PD   As CPackageDetail
'    Set PD = LoadPackageColl(Trim(Str(ID)))
'
'   If PD.HOLD_FLAG = "Y" Then
'      lblWarningUnavailableSale.Visible = True
'    Call InitNormalLabel(lblWarningUnavailableSale, MapText("สินค้างดจำหน่ายชั่วคราว !!"))
'     'HoldFlag = "N"
'    ElseIf PD.HOLD_FLAG = "N" Then
'    lblWarningUnavailableSale.Visible = True
'   End If


   
   
   Dim Di As CDocItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CDocItem
      
      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If
   
   If ShowMode = SHOW_EDIT Then
      
   End If
   
   Call Di.SetFieldValue("SELL_TYPE", TempID)
   Call Di.SetFieldValue("DOC_ITEM_TYPE", cboDocItemType.ItemData(cboDocItemType.ListIndex))
   If TempID = 1 Then
      Call Di.SetFieldValue("STOCK_TYPE", uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)))
      Call Di.SetFieldValue("PART_ITEM_ID", uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
      Call Di.SetFieldValue("STOCK_DESC", uctlProductLookup.MyCombo.Text)
      Call Di.SetFieldValue("STOCK_NO", uctlProductLookup.MyTextBox.Text)
      If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
         Call Di.SetFieldValue("PART_ITEM_RETURN_ID", uctlProductReturn.MyCombo.ItemData(Minus2Zero(uctlProductReturn.MyCombo.ListIndex)))
      End If
   ElseIf TempID = 2 Then
      Call Di.SetFieldValue("STOCK_TYPE", uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)))
      Call Di.SetFieldValue("PART_ITEM_ID", uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
      Call Di.SetFieldValue("STOCK_DESC", uctlProductLookup.MyCombo.Text)
      Call Di.SetFieldValue("STOCK_NO", uctlProductLookup.MyTextBox.Text)
      If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
         Call Di.SetFieldValue("PART_ITEM_RETURN_ID", uctlProductReturn.MyCombo.ItemData(Minus2Zero(uctlProductReturn.MyCombo.ListIndex)))
      End If
   Else
      Call Di.SetFieldValue("ITEM_DESC", txtManual.Text)
   End If
   Call Di.SetFieldValue("PO_ID", Val(cmdBrowse.Tag))
   Call Di.SetFieldValue("PO_NO", txtRefNo.Text)
   Call Di.SetFieldValue("PRODUCT_DETAIL", txtProDuctDetail.Text)
   Call Di.SetFieldValue("FREE_FLAG", Check2Flag(chkFree.Value))
   Call Di.SetFieldValue("LOT_ITEM_FLAG", Check2Flag(ChkLotitemFlag.Value))
   
   Call Di.SetFieldValue("LOCATION_ID", uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)))
   Call Di.SetFieldValue("ITEM_AMOUNT", Val(txtQuantity.Text) * Multiple)
   Call Di.SetFieldValue("AVG_PRICE", MyDiffEx(Val(txtAvgPrice.Text), Multiple))
   Call Di.SetFieldValue("AVG_PRICE_EX", MyDiffEx(Val(txtAvgPriceEx.Text), Multiple))
   Call Di.SetFieldValue("DISCOUNT_AMOUNT", Val(txtDiscount.Text))
   Call Di.SetFieldValue("DISCOUNT_PERCENT", Val(txtDiscountPercent.Text))
   Call Di.SetFieldValue("TOTAL_PRICE", FormatNumber(Val(txtQuantity.Text) * Val(txtAvgPrice.Text), , False))
   
   Call Di.SetFieldValue("UNIT_TRAN_ID", UnitID)
   Call Di.SetFieldValue("UNIT_MULTIPLE", Multiple)
   Call Di.SetFieldValue("UNIT_TRAN_NAME", UnitName)

   If m_Sc.CHK_STD_COST = "Y" Then        'ถ้าเป็น Standard แล้วให้นำต้นทุน Standard เป็นต้นทุนขายด้วยทันที
      Call Di.SetFieldValue("CAPITAL_AMOUNT", m_Sc.COST_PER_AMOUNT)
      Call Di.SetFieldValue("TOTAL_INCLUDE_PRICE", m_Sc.COST_PER_AMOUNT * Val(txtQuantity.Text) * Multiple)
   End If
   
   Call Di.SetFieldValue("ITEM_AMOUNT", Val(txtQuantity.Text) * Multiple)
   
   Call Di.SetFieldValue("WEIGHT_AMOUNT", Val(txtWeightAmount.Text))
      
   For Each Pl In TempCollection2
      If Pl.Flag <> "A" And Pl.Flag <> "D" Then
         Pl.Flag = "E"
      End If
      Call Pl.SetFieldValue("UNIT_TRAN_ID", UnitID)
      Call Pl.SetFieldValue("UNIT_MULTIPLE", Multiple)
      Call Pl.SetFieldValue("TOTAL_PRICE", Pl.GetFieldValue("ITEM_AMOUNT") * MyDiffEx(Val(txtAvgPrice.Text), Multiple))
      
   Next Pl
    
   Set Di.PrintLabels = TempCollection2
   
   Set Di.LotItemLinkCollection = LotItemLinkCollection
   
   
   
   Set Di = Nothing
   SaveData = True
End Function
Public Function GenerateAutoLotLink() As Boolean
Dim m_LotItem As CLotItem
Dim Lk As CDocItemLink
Dim CompareAmount  As Double
Dim itemcount As Long
Dim TempID As Long
Dim m_Rs  As ADODB.Recordset
   
   Set m_Rs = New ADODB.Recordset
   GenerateAutoLotLink = False
   CompareAmount = Val(txtQuantity.Text) * Multiple
   MasterInd = "6"
   Set m_LotItem = New CLotItem
   
   m_LotItem.LOT_ITEM_ID = -1
   m_LotItem.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   m_LotItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   m_LotItem.COUNT_AMOUNT = "Y"
   Call m_LotItem.QueryData(6, m_Rs, itemcount, False)
   
   While Not m_Rs.EOF
      If CompareAmount <= 0 Then
         If m_Rs.State = adStateOpen Then
            m_Rs.Close
         End If
         Set m_Rs = Nothing
         GenerateAutoLotLink = True
         MasterInd = "1"
         Set Lk = Nothing
         Exit Function
      End If
      Call m_LotItem.PopulateFromRS(6, m_Rs)
      
      Set Lk = New CDocItemLink
      Lk.Flag = "A"
      Lk.IMPORT_LOT_ITEM_ID = m_LotItem.LOT_ITEM_ID
      If Round(CompareAmount, 2) = Round(m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT, 2) Then
         Lk.IMPORT_AMOUNT = CompareAmount
      ElseIf Round(CompareAmount, 2) > Round(m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT, 2) Then
         Lk.IMPORT_AMOUNT = m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT
         CompareAmount = CompareAmount - (m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT)
      ElseIf Round(CompareAmount, 2) < Round(m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT, 2) Then
         Lk.IMPORT_AMOUNT = CompareAmount
         CompareAmount = 0
      End If
      Lk.MAIN_IMPORT_LOT_ITEM_ID = Lk.IMPORT_LOT_ITEM_ID
      TempID = Lk.IMPORT_LOT_ITEM_ID
      
      Call glbDaily.GetNextLotItemID(TempID, m_LotItem.INVENTORY_DOC_ID, m_LotItem.PART_ITEM_ID)
      
      If TempID > 0 Then
         Lk.MAIN_IMPORT_LOT_ITEM_ID = TempID
      End If
      
      Call LotItemLinkCollection.add(Lk, Trim(m_LotItem.LOT_ITEM_ID & "-" & m_LotItem.PART_ITEM_ID))
   
      Set Lk = Nothing
      m_Rs.MoveNext
   Wend
   
   If CompareAmount > 0 Then
      GenerateAutoLotLink = False
   Else
      GenerateAutoLotLink = True
   End If
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   MasterInd = "1"
End Function
Public Function CheckLotItemAmount() As Boolean
Dim Lk As CDocItemLink
Dim SumAmount As Double
   
   CheckLotItemAmount = True
   SumAmount = 0
   For Each Lk In LotItemLinkCollection
      SumAmount = SumAmount + Lk.IMPORT_AMOUNT
   Next Lk
   If Round(SumAmount, 2) <> Round(Val(txtQuantity.Text) * Multiple, 2) Then
      CheckLotItemAmount = False
   End If
End Function
Private Sub cmdUnit_Click()
   frmChangeUnit.HeaderText = MapText("เปลี่ยนหน่วย")
   frmChangeUnit.UnitID = UnitID
   frmChangeUnit.Multiple = Multiple
   frmChangeUnit.UnitName = UnitName
   frmChangeUnit.UnitMName = UnitMName
   
   Load frmChangeUnit
   frmChangeUnit.Show 1
   
   UnitID = frmChangeUnit.UnitID
   Multiple = frmChangeUnit.Multiple
   UnitName = frmChangeUnit.UnitName
   UnitMName = frmChangeUnit.UnitMName
   
   Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   
   Unload frmChangeUnit
   Set frmChangeUnit = Nothing
   m_HasModify = True
   
   Call txtAvgPrice.SetFocus

End Sub

Private Sub Form_Activate()
Dim D As CAPARMas
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitSellType(cboSellType)
      Call InitDocItemType(cboDocItemType)
      
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      If Area = 2 Then
         Set D = m_SupplierColl(Trim(Str(CusID)))
      Else
         Set D = m_CustomerColl(Trim(Str(CusID)))
      End If
   
      If D.LABEL_FLAG = "Y" Then
         Call LoadMaster(uctlBlock.MyCombo, m_Blocks, , , MASTER_CUSTOMER_BLOCK)
         Set uctlBlock.MyCollection = m_Blocks
         
         SSOption1.Visible = True
         SSOption2.Visible = True
         
         
         Me.Height = 10950
         Me.Top = -200
         
         Me.Refresh
      Else
         SSOption1.Visible = False
         SSOption2.Visible = False
         
         
         
'          Me.Height = 1000
'         Me.Top = -200
          
         
      End If
      
          
      
      If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
         Call LoadStockCode(uctlProductReturn.MyCombo, m_ProductReturns)
         Set uctlProductReturn.MyCollection = m_ProductReturns
      End If
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
         cboSellType.ListIndex = 1
         cboDocItemType.ListIndex = 1
         
         '42 '002
         Dim Pt As CMasterRef
         For Each Pt In m_ProductTypes
            If Pt.KEY_CODE = "42" Then
               uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, Pt.KEY_ID)
            End If
         Next Pt
         Set Pt = Nothing
         For Each Pt In m_Locations
            If Pt.KEY_CODE = "01" Then
               uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, Pt.KEY_ID)
            End If
         Next Pt
         Set Pt = Nothing
         
         If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
            Call cmdBrowse_Click
         End If
      End If
      
      If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Then
         cboDocItemType.Enabled = True
      Else
         cboDocItemType.Enabled = False
      End If
         
      m_HasModify = False
      
      SSOption1.Value = True
      If cboSellType.ListIndex = 1 Then
         Call uctlProductLookup.SetFocus
      End If
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdAddLotItem_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
     
      If SSFrame2.Enabled Then
        
              Call cmdNext_Click
      
      Else
           Call cmdAdd2_Click
      End If
      KeyCode = 0
    
      
      
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete2_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit2_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 120 Then
      Call ExChangeMode
   End If
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Products = New Collection
   Set m_Locations = New Collection
   Set m_PartTypes = New Collection
   Set m_ProductTypes = New Collection
   Set m_ProductReturns = New Collection
   Set m_Mr = New CMasterRef
   Set m_Branchs = New Collection
   Set m_Blocks = New Collection

   Set TempCollection2 = New Collection
   Set LotItemLinkCollection = New Collection
   
   Set m_EmployeeColl = New Collection
   Set TempEmp = New CEmployee
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Products = Nothing
   Set m_Locations = Nothing
   Set m_PartTypes = Nothing
   Set m_ProductTypes = Nothing
   Set m_Mr = Nothing
   Set m_Branchs = Nothing
   Set m_Blocks = Nothing
   Set m_Sc = Nothing
   Set m_ProductReturns = Nothing
   
   Set TempCollection2 = Nothing
   Set LotItemLinkCollection = Nothing
   
   Set uctlLocationLookup.MyCollection = Nothing
   Set uctlProductLookup.MyCollection = Nothing
   Set uctlProductTypeLookup.MyCollection = Nothing
   Set uctlSale.MyCollection = Nothing
   
   Set m_EmployeeColl = Nothing
   Set TempEmp = Nothing
End Sub



Private Sub SSOption1_Click(Value As Integer)
   If SSOption1.Value Then
      SSFrame2.Enabled = True
      SSFrame3.Enabled = False
   Else
      SSFrame2.Enabled = False
      SSFrame3.Enabled = True
   End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
   If SSOption1.Value Then
      SSFrame2.Enabled = True
      SSFrame3.Enabled = False
   Else
      SSFrame2.Enabled = False
      SSFrame3.Enabled = True
   End If
End Sub
Private Sub txtAvgPrice_LostFocus()
   m_HasModify = True
    txtTotalPrice.Text = FormatNumber(Val(txtAvgPrice.Text) * Val(txtQuantity.Text), , False)
End Sub

Private Sub txtAvgPriceEx_Change()
   m_HasModify = True
End Sub
Private Sub txtDiscount_Change()
   m_HasModify = True
   Call CalculatePrice
End Sub
Private Sub txtDiscountPercent_Change()
   m_HasModify = True
   txtDiscount.Text = Val(txtTotalPrice.Text) * Val(txtDiscountPercent.Text) / 100
End Sub
Private Sub txtLeft_Change()
   m_HasModify = True
End Sub
Private Sub txtManual_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtProDuctDetail_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   txtTotalPrice.Text = FormatNumber(Val(txtAvgPrice.Text) * Val(txtQuantity.Text), , False)
End Sub
Private Sub txtTotalPrice_LostFocus()
Dim TotalPrice As Double
Dim Quantity As Double

   TotalPrice = Val(txtTotalPrice.Text)
   Quantity = Val(txtQuantity.Text)
   txtAvgPrice.Text = MyDiff(TotalPrice, Quantity)
   If Not (FormatNumber(TotalPrice - Val(txtDiscount.Text)) = FormatNumber(Val(txtAvgPrice.Text) * Quantity)) Then
      If ((TotalPrice - Val(txtDiscount.Text)) > (Val(txtAvgPrice.Text) * Quantity)) Then
         txtAvgPrice.Text = Val(txtAvgPrice.Text) + 0.01
         Call txtAvgPrice_LostFocus
         txtDiscount.Text = FormatNumber((Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) - TotalPrice, , False)
      ElseIf (Val(txtTotalPrice.Text) - Val(txtDiscount.Text)) < (Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) Then
         Call txtAvgPrice_LostFocus
         txtDiscount.Text = FormatNumber((Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) - TotalPrice, , False)
      End If
   End If
   Call CalculatePrice
End Sub
Private Sub txtRefNo_Change()
   m_HasModify = True
End Sub


Private Sub txtWeightAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlBranch_Change()
Dim ID As Long

   uctlSale.MyCombo.ListIndex = -1
   ID = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex))

   If ID > 0 Then
      Call LoadMasterKeyID(uctlSale.MyCombo, m_EmployeeColl, ID)
      Set uctlSale.MyCollection = m_EmployeeColl
      Set m_Mr = m_EmployeeColl.Item(1)
      uctlSale.MyTextBox.Text = m_Mr.EMP_CODE
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlProductLookup_Change()
On Error Resume Next
Dim ID As Long
Dim D As CAPARMas
Dim PkgDetail As CPackageDetail


   cmdOK.Enabled = True
   cmdNext.Enabled = True

   ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If ID > 0 Then
      Set m_Sc = GetObject("CStockCode", m_Products, Trim(Str(ID)))
      
      UnitID = m_Sc.UNIT_ID
      Multiple = m_Sc.UNIT_AMOUNT
      UnitName = m_Sc.UNIT_NAME
      UnitMName = m_Sc.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
                  
      If m_Sc.LOT_FLAG = "Y" And (DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE) Then
         cmdAddLotItem.Enabled = True
      Else
         cmdAddLotItem.Enabled = False
      End If
                        
      If Area = 1 Then
         Set D = m_CustomerColl(Trim(Str(CusID)))
      ElseIf Area = 2 Then
         Set D = m_SupplierColl(Trim(Str(CusID)))
      End If
         
      
      For Each PkgDetail In LoadPackageColl
         If D.PACKAGE_ID <= 0 Then
            If PkgDetail.GetFieldValue("PACKAGE_MASTER_FLAG") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
               Exit For
            End If
         Else
            If PkgDetail.GetFieldValue("PACKAGE_ID") = D.PACKAGE_ID And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
               Exit For
            End If
         End If
      Next PkgDetail
      
      HoldFlag = PkgDetail.GetFieldValue("HOLD_FLAG")
      
      If Not (PkgDetail Is Nothing) Then
         '------
         If HoldFlag = "Y" Then
'              HoldFlag = "Y"
             glbErrorLog.LocalErrorMsg = "สินค้า งดจำหน่ายชั่วคราว  ควรเปลี่ยนการตั้งราคาสินค้าก่อน"
             glbErrorLog.ShowUserError
            cmdOK.Enabled = False
            cmdNext.Enabled = False
     '
            ' lblWarningUnavailableSale.Visible = True
        
         ElseIf HoldFlag = "N" Then
            ' lblWarningUnavailableSale.Visible = False
'            HoldFlag = "N"
            cmdOK.Enabled = True
            cmdNext.Enabled = True
         '
         End If
         
        If DocumentDate >= PkgDetail.GetFieldValue("PRO_FROM_DATE") And DocumentDate <= PkgDetail.GetFieldValue("PRO_TO_DATE") Then
           txtAvgPrice.Text = PkgDetail.GetFieldValue("PRO_ITEM_COST")
        Else
           txtAvgPrice.Text = PkgDetail.GetFieldValue("PART_ITEM_COST")
        End If
      Else
         txtAvgPrice.Text = 0
      End If
      
   Else
      lblUnit.Caption = ""
   End If
   
   If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      If m_Sc.PART_ITEM_RETURN_ID > 0 Then
         uctlProductReturn.MyCombo.ListIndex = IDToListIndex(uctlProductReturn.MyCombo, m_Sc.PART_ITEM_RETURN_ID)
      Else
         uctlProductReturn.MyCombo.ListIndex = IDToListIndex(uctlProductReturn.MyCombo, m_Sc.STOCK_CODE_ID)
      End If
   End If
'   HoldFlag = PkgDetail.GetFieldValue("HOLD_FLAG ")
   'PkgDetail.GetFieldValue("HOLD_FLAG ") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID
   'PkgDetail.HOLD_FLAG = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID
'
'   If PkgDetail.GetFieldValue("HOLD_FLAG") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
'        lblWarningUnavailableSale.Visible = True
'
'   ElseIf PkgDetail.GetFieldValue("HOLD_FLAG") = "N" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
'         lblWarningUnavailableSale.Visible = False
'  End If
   
   m_HasModify = True

End Sub
Private Sub uctlProductReturn_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductTypeLookup_Change()
Dim PartTypeID As Long
Static OldPartType As Long
   
   PartTypeID = uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex))
   If OldPartType <> PartTypeID Then
      OldPartType = PartTypeID
   Else
      Exit Sub
   End If
   If PartTypeID > 0 Then
      uctlProductLookup.MyTextBox.Text = ""
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products, PartTypeID)
      Set uctlProductLookup.MyCollection = m_Products
   End If
   
   m_HasModify = True
End Sub
Private Sub ExChangeMode()
   If Not (SSOption1.Visible) Then
      Exit Sub
   End If
   
   SSOption1.Value = Not (SSOption1.Value)
   SSOption2.Value = Not (SSOption1.Value)
   
   If Not CheckSave2 Then
      SSOption1.Value = Not (SSOption1.Value)
      SSOption2.Value = Not (SSOption1.Value)
      Call uctlProductLookup.SetFocus
      Exit Sub
   End If
      
   If SSOption2.Value = True Then
      Call uctlBlock.SetFocus
   End If
   If SSOption1.Value = True Then
      Call uctlProductLookup.SetFocus
   End If
End Sub
Private Sub uctlBlock_Change()
Dim ID As Long
   
   ID = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex))
   If ID > 0 Then
      Call LoadMaster(uctlBranch.MyCombo, m_Branchs, , , MASTER_APARMAS_BRANCH, , ID)
      Set uctlBranch.MyCollection = m_Branchs
   End If
   
   m_HasModify = True
End Sub
Private Sub cmdAdd2_Click()
   ShowMode2 = SHOW_ADD
   uctlBlock.MyCombo.ListIndex = -1
   uctlBranch.MyCombo.ListIndex = -1
   uctlSale.MyCombo.ListIndex = -1
   txtItemAmount.Text = ""
   txtPackAmount.Text = ""
   Call uctlBlock.SetFocus
End Sub
Private Sub cmdDelete2_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete2.Enabled Then
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

   If ID1 <= 0 Then
      TempCollection2.Remove (ID2)
   Else
      TempCollection2.Item(ID2).Flag = "D"
   End If

   Call GetAmount
   GridEX1.itemcount = CountItem(TempCollection2)
   GridEX1.Rebind
   m_HasModify = True

End Sub
Private Sub GetAmount()
Dim II As CPrintLabel
Dim Sum1 As Double

   Sum1 = 0
   
   For Each II In TempCollection2
      If II.Flag <> "D" Then
         'debug.print II.Flag
         Sum1 = Sum1 + MyDiffEx(II.GetFieldValue("ITEM_AMOUNT"), Multiple)
      End If
   Next II
   
   txtSumAmount.Text = Sum1
   
End Sub

Private Sub cmdEdit_Click()
   ShowMode = SHOW_EDIT
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   
   Call QueryData(True)
End Sub
Private Sub QueryData2(Flag As Boolean)
Dim IsOK As Boolean

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode2 = SHOW_EDIT Then
         Dim Plb As CPrintLabel
         Set Plb = TempCollection2.Item(ID2)
         
         uctlBlock.MyCombo.ListIndex = IDToListIndex(uctlBlock.MyCombo, Plb.GetFieldValue("BLOCK_ID"))
         uctlBranch.MyCombo.ListIndex = IDToListIndex(uctlBranch.MyCombo, Plb.GetFieldValue("BRANCH_ID"))
         uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, Plb.GetFieldValue("SALE_ID"))
         txtItemAmount.Text = MyDiff(Plb.GetFieldValue("ITEM_AMOUNT"), Multiple)
         txtPackAmount.Text = Plb.GetFieldValue("PACK_AMOUNT")
                
'               HoldFlag = Plb.GetFieldValue("HOLD_FLAG")
                
         Set Plb = Nothing
      End If
   End If
   
   Call GetAmount
   
   GridEX1.itemcount = CountItem(TempCollection2)
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdNext2_Click()
Dim NewID As Long
   
   If ShowMode2 <> SHOW_EDIT Then
      ShowMode2 = SHOW_ADD
   End If
   If Not SaveData2 Then
      Exit Sub
   End If
   
   If ShowMode2 = SHOW_EDIT Then
      NewID = GetNextID(ID2, TempCollection2)
      If ID2 = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
      Else
         ID2 = NewID
      End If
      
      
   ElseIf ShowMode2 = SHOW_ADD Then
      uctlBlock.MyCombo.ListIndex = -1
      uctlBranch.MyCombo.ListIndex = -1
      uctlSale.MyCombo.ListIndex = -1
      txtItemAmount.Text = ""
      txtPackAmount.Text = ""
   End If
   Call QueryData2(True)
   
   Call uctlBlock.MyTextBox.SetFocus
End Sub
Private Function SaveData2() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim TempID As Long

   If Not VerifyCombo(lblBlock, uctlBlock.MyCombo, False) Then
      Exit Function
   End If
      
   If Not VerifyCombo(lblBranch, uctlBranch.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblSale, uctlSale.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblItemAmount, txtItemAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData2 = True
      Exit Function
   End If
   
   TempID = 0
   Dim Plb As CPrintLabel
   For Each Plb In TempCollection2
      If ShowMode = SHOW_ADD Then
         If Plb.GetFieldValue("BLOCK_ID") = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)) And Plb.GetFieldValue("BRANCH_ID") = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)) Then
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลบล็อค ") & Plb.GetFieldValue("BLOCK_NAME") & " และสาขา " & Plb.GetFieldValue("BRANCH_NAME") & " " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      Else
         TempID = TempID + 1
         If Plb.GetFieldValue("BLOCK_ID") = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)) And Plb.GetFieldValue("BRANCH_ID") = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex) And Not (TempID = ID2)) Then     'And Not (TempID = ID) เอาออกวันที่ 17/03/2558 ไม่รู้ว่าใครใส่และใส่ไว้ทำไม และเปลี่ยนเป็น TempID = ID2 ตอน 30/06/2559
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลบล็อค ") & Plb.GetFieldValue("BLOCK_NAME") & " และสาขา " & Plb.GetFieldValue("BRANCH_NAME") & " " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
         Exit Function
      End If
      End If
   Next
   
   If ShowMode2 = SHOW_ADD Then
      Set Plb = New CPrintLabel

      Plb.Flag = "A"
      Call TempCollection2.add(Plb)
   Else
      Set Plb = TempCollection2.Item(ID2)
      If Plb.Flag <> "A" Then
         Plb.Flag = "E"
      End If
   End If
   
   Call Plb.SetFieldValue("BLOCK_ID", uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)))
   Call Plb.SetFieldValue("BLOCK_NAME", uctlBlock.MyCombo.Text)
   Call Plb.SetFieldValue("BRANCH_ID", uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)))
   Call Plb.SetFieldValue("BRANCH_CODE", uctlBranch.MyTextBox.Text)
   Call Plb.SetFieldValue("BRANCH_NAME", uctlBranch.MyCombo.Text)
   Call Plb.SetFieldValue("EMP_ID", uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex)))
   Call Plb.SetFieldValue("SALE_ID", uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex)))
'   If ShowMode2 = SHOW_ADD Or (ShowMode2 = SHOW_EDIT And Plb.Flag = "E") Then
     Call Plb.SetFieldValue("SALE_LONG_NAME", uctlSale.MyCombo.Text)
     Call Plb.SetFieldValue("SALE_LAST_NAME", "")
'   End If
   Call Plb.SetFieldValue("ITEM_AMOUNT", Val(txtItemAmount.Text) * Multiple)
   Call Plb.SetFieldValue("PACK_AMOUNT", Val(txtPackAmount.Text))
   Call Plb.SetFieldValue("TOTAL_PRICE", FormatNumber(Val(txtItemAmount.Text) * Val(txtAvgPrice.Text), , False))
   Call Plb.SetFieldValue("TOTAL_AMOUNT", Val(txtItemAmount.Text) * Multiple)
   
   Call Plb.SetFieldValue("UNIT_TRAN_ID", UnitID)
   Call Plb.SetFieldValue("UNIT_MULTIPLE", Multiple)
   
   SaveData2 = True
End Function
Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 1
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 100
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("บล็อค")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("สาขา")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("สาขา")
   
   Set Col = GridEX1.Columns.add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2000
   Col.Caption = MapText("ตะกร้า")
   
   Set Col = GridEX1.Columns.add '7
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2000
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 4000
   Col.Caption = MapText("พนักงานขาย")               ' ** พนักงานขาย **
   
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit2_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Plb As CPrintLabel
   If TempCollection2.Count <= 0 Then
      Exit Sub
   End If
   Set Plb = GetItem(TempCollection2, RowIndex, RealIndex)
   If Plb Is Nothing Then
      Exit Sub
   End If

   Values(1) = Plb.GetFieldValue("PRINT_LABEL_ID")
   Values(2) = RealIndex
   Values(3) = Plb.GetFieldValue("BLOCK_NAME")
   Values(4) = Plb.GetFieldValue("BRANCH_CODE")
   Values(5) = Plb.GetFieldValue("BRANCH_NAME")
   Values(6) = FormatNumber(Plb.GetFieldValue("PACK_AMOUNT"))
   Values(7) = FormatNumber(MyDiff(Plb.GetFieldValue("ITEM_AMOUNT"), Multiple))
   Values(8) = "  " & Plb.GetFieldValue("SALE_LONG_NAME") & " " & Plb.GetFieldValue("SALE_LAST_NAME")
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub txtItemAmount_Change()
 Dim D As CAPARMas
   
   If Area = 2 Then
      Set D = m_SupplierColl(Trim(Str(CusID)))
   Else
      Set D = m_CustomerColl(Trim(Str(CusID)))
   End If
   If D.BASKET_FIX_AMOUNT > 0 Then
      txtPackAmount.Text = "1"
   Else
      txtPackAmount.Text = MyDiffEx(Val(txtItemAmount.Text), m_Sc.UNIT_PER_BASKET)
   End If
   m_HasModify = True
End Sub
Private Function CheckSave2() As Boolean
Dim TempID As Long

   TempID = cboSellType.ItemData(Minus2Zero(cboSellType.ListIndex))
   
   CheckSave2 = True
   If Not VerifyCombo(lblSellType, cboSellType, False) Then
      CheckSave2 = False
   End If
   
   If TempID = 1 Or TempID = 2 Then
      If Not VerifyCombo(lblProductType, uctlProductTypeLookup.MyCombo, False) Then
         CheckSave2 = False
      End If
      
      If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
         CheckSave2 = False
      End If
      
      If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
         CheckSave2 = False
      End If

   Else
      If Not VerifyTextControl(lblManual, txtManual, False) Then
         CheckSave2 = False
      End If
   End If

   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      CheckSave2 = False
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      CheckSave2 = False
   End If
   If Not VerifyTextControl(lblDiscount, txtDiscount, True) Then
      CheckSave2 = False
   End If
   
End Function

