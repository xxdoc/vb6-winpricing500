VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAddEditApArMas 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   Icon            =   "frmAddEditApArMas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13665
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlDate uctlBirthDate 
         Height          =   495
         Left            =   7920
         TabIndex        =   11
         Top             =   2880
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin Xivess.uctlTextBox txtAparMasGroupName 
         Height          =   435
         Left            =   2920
         TabIndex        =   13
         Top             =   2880
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtAparMasGroupCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
         _ExtentX        =   2355
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboBusinessType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3400
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   4245
         Width           =   13275
         _ExtentX        =   23416
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
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1470
         Width           =   4515
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtShortName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtEmail 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   1920
         Width           =   4515
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtWebSite 
         Height          =   435
         Left            =   1860
         TabIndex        =   10
         Top             =   2370
         Width           =   4515
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBusinessDesc 
         Height          =   450
         Left            =   7920
         TabIndex        =   15
         Top             =   3360
         Width           =   4605
         _ExtentX        =   18627
         _ExtentY        =   794
      End
      Begin Xivess.uctlTextBox txtCredit 
         Height          =   435
         Left            =   4380
         TabIndex        =   1
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Xivess.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   6390
         TabIndex        =   2
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2955
         Left            =   150
         TabIndex        =   18
         Top             =   4800
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   5212
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
         Column(1)       =   "frmAddEditApArMas.frx":27A2
         Column(2)       =   "frmAddEditApArMas.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditApArMas.frx":290E
         FormatStyle(2)  =   "frmAddEditApArMas.frx":2A6A
         FormatStyle(3)  =   "frmAddEditApArMas.frx":2B1A
         FormatStyle(4)  =   "frmAddEditApArMas.frx":2BCE
         FormatStyle(5)  =   "frmAddEditApArMas.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditApArMas.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   34
         Top             =   0
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextLookup uctlPackage 
         Height          =   465
         Left            =   1860
         TabIndex        =   16
         Top             =   3840
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin Xivess.uctlTextBox txtShort 
         Height          =   435
         Left            =   7920
         TabIndex        =   7
         Top             =   1920
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtTaxID 
         Height          =   435
         Left            =   7920
         TabIndex        =   9
         Top             =   2370
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBillName 
         Height          =   435
         Left            =   7920
         TabIndex        =   6
         Top             =   1440
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBasketFixAmount 
         Height          =   435
         Left            =   10440
         TabIndex        =   43
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin VB.Label lblBirthDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   45
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblBasketFixAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9240
         TabIndex        =   44
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblBillName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         TabIndex        =   42
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblTaxID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         TabIndex        =   41
         Top             =   2490
         Width           =   855
      End
      Begin Threed.SSCheck chkConsignmentFlag 
         Height          =   375
         Left            =   13320
         TabIndex        =   40
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkCancelOutDocument 
         Height          =   435
         Left            =   9720
         TabIndex        =   39
         Top             =   3840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAparMasGroup 
         Height          =   375
         Left            =   720
         TabIndex        =   38
         Top             =   2880
         Width           =   975
      End
      Begin Threed.SSCheck chkFlagEdit 
         Height          =   435
         Left            =   7320
         TabIndex        =   37
         Top             =   3840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblShort 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         TabIndex        =   36
         Top             =   2040
         Width           =   855
      End
      Begin Threed.SSCheck chkAddBranchName 
         Height          =   375
         Left            =   11400
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkLabelFlag 
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPackage 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   3840
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9960
         TabIndex        =   22
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditApArMas.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   11760
         TabIndex        =   23
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditApArMas.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditApArMas.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5220
         TabIndex        =   33
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   6570
         TabIndex        =   32
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3630
         TabIndex        =   31
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblBusinessDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   30
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblWebsite 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   27
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblBusinessType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   1485
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditApArMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Customer As CAPARMas
Private m_Employees As Collection
Private m_Employee As CEmployee

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ApArInd As Long

Private ApArText As String
Private FileName As String
Private m_MasterRef As CMasterRef
Private m_Package As CPackage
Private m_Packages As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Customer.APAR_MAS_ID = ID
      If Not glbDaily.QueryCustomer(m_Customer, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Customer.PopulateFromRS(1, m_Rs)
      
      txtEmail.Text = m_Customer.EMAIL
      uctlBirthDate.ShowDate = m_Customer.BIRTH_DATE
      txtWebSite.Text = m_Customer.WEBSITE
      cboBusinessType.ListIndex = IDToListIndex(cboBusinessType, m_Customer.APAR_TYPE)
      txtShortName.Text = m_Customer.APAR_CODE
      txtBusinessDesc.Text = m_Customer.BUSINESS_DESC
      txtCredit.Text = m_Customer.CREDIT
      txtDiscountPercent.Text = m_Customer.NORMAL_DISCOUNT
      txtBasketFixAmount.Text = m_Customer.BASKET_FIX_AMOUNT
      
      uctlPackage.MyCombo.ListIndex = IDToListIndex(uctlPackage.MyCombo, m_Customer.PACKAGE_ID)
      chkLabelFlag.Value = FlagToCheck(m_Customer.LABEL_FLAG)
      chkAddBranchName.Value = FlagToCheck(m_Customer.ADD_BRANCH_NAME)
      chkFlagEdit.Value = FlagToCheck(m_Customer.FLAG_EDIT)
      chkCancelOutDocument.Value = FlagToCheck(m_Customer.CANCEL_OUT_DOCUMENT)
      chkConsignmentFlag.Value = FlagToCheck(m_Customer.CONSIGNMENT_FLAG)
      
      txtAparMasGroupCode.Text = m_Customer.APAR_MAS_GROUP_CODE
      txtAparMasGroupName.Text = m_Customer.APAR_MAS_GROUP_NAME
      txtTaxID.Text = m_Customer.TAX_ID
      
      Dim Name As cName
      Dim CstName As CApArName
      If (Not m_Customer.CstNames Is Nothing) And (m_Customer.CstNames.Count > 0) Then
         Set CstName = m_Customer.CstNames(1)
         Set Name = CstName.Name
         txtName.Text = Name.GetFieldValue("LONG_NAME")
         txtShort.Text = Name.GetFieldValue("SHORT_NAME")
         txtBillName.Text = Name.GetFieldValue("BILL_NAME")
      Else
         txtName.Text = ""
      End If
   Else
      ShowMode = SHOW_ADD
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblShortName, txtShortName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBusinessType, cboBusinessType, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(APARCODE_UNIQUE, txtShortName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtShortName.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If m_Customer.CstBranchs.Count <= 0 And ApArInd = 1 Then
   
      Call AutoGenBranch
      'Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Customer.ShowMode = ShowMode
   m_Customer.BIRTH_DATE = -1
   m_Customer.PASSWD = ""
   m_Customer.EMAIL = txtEmail.Text
   m_Customer.BIRTH_DATE = uctlBirthDate.ShowDate
   m_Customer.WEBSITE = txtWebSite.Text
   m_Customer.APAR_TYPE = cboBusinessType.ItemData(Minus2Zero(cboBusinessType.ListIndex))
   m_Customer.CREDIT = Val(txtCredit.Text)
   m_Customer.APAR_CODE = txtShortName.Text
   m_Customer.BUSINESS_DESC = txtBusinessDesc.Text
   m_Customer.NORMAL_DISCOUNT = Val(txtDiscountPercent.Text)
   m_Customer.PACKAGE_ID = uctlPackage.MyCombo.ItemData(Minus2Zero(uctlPackage.MyCombo.ListIndex))
   m_Customer.APAR_IND = ApArInd
   m_Customer.LABEL_FLAG = Check2Flag(chkLabelFlag.Value)
   m_Customer.ADD_BRANCH_NAME = Check2Flag(chkAddBranchName.Value)
   m_Customer.FLAG_EDIT = Check2Flag(chkFlagEdit.Value)
   m_Customer.CANCEL_OUT_DOCUMENT = Check2Flag(chkCancelOutDocument.Value)
   m_Customer.CONSIGNMENT_FLAG = Check2Flag(chkConsignmentFlag.Value)
   m_Customer.TAX_ID = txtTaxID.Text
   m_Customer.BASKET_FIX_AMOUNT = Val(txtBasketFixAmount.Text)
   
   m_Customer.APAR_MAS_GROUP_CODE = txtAparMasGroupCode.Text
   m_Customer.APAR_MAS_GROUP_NAME = txtAparMasGroupName.Text
   
   Dim CstName As CApArName
   If m_Customer.CstNames.Count <= 0 Then
      Set CstName = New CApArName
      CstName.Flag = "A"
      Call m_Customer.CstNames.add(CstName)
   Else
      Set CstName = m_Customer.CstNames.Item(1)
      CstName.Flag = "E"
   End If
   
   Dim Name As cName
   If m_Customer.CstNames.Count <= 0 Then
      Set Name = CstName.Name
      Call Name.SetFieldValue("LONG_NAME", txtName.Text)
      Call Name.SetFieldValue("SHORT_NAME", txtShort.Text)
      Call Name.SetFieldValue("BILL_NAME", txtBillName.Text)
      Name.Flag = "A"
   Else
      Set Name = CstName.Name
      Call Name.SetFieldValue("LONG_NAME", txtName.Text)
      Call Name.SetFieldValue("SHORT_NAME", txtShort.Text)
      Call Name.SetFieldValue("BILL_NAME", txtBillName.Text)
      Name.Flag = "E"
   End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCustomer(m_Customer, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub
Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub
Private Sub chkAddBranchName_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkAddBranchName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkLabelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkLabelFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkConsignmentFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkConsignmentFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkFlagEdit_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkCancelOutDocument_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkFlagEdit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
 
  If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "EDIT" & "_" & "ADD", "����") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditApArMasAddress.TempCollection = m_Customer.CstAddresses
      frmAddEditApArMasAddress.ShowMode = SHOW_ADD
      frmAddEditApArMasAddress.HeaderText = MapText("�����������")
      Load frmAddEditApArMasAddress
      frmAddEditApArMasAddress.Show 1

      OKClick = frmAddEditApArMasAddress.OKClick

      Unload frmAddEditApArMasAddress
      Set frmAddEditApArMasAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAddresses)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Tag = "CUSTOMER_BRANCH" Then
      Set frmAddEditApArMasBranch.TempCollection = m_Customer.CstBranchs
      frmAddEditApArMasBranch.CustomerID = ID
      Set frmAddEditApArMasBranch.ParentForm = Me
      frmAddEditApArMasBranch.ShowMode = SHOW_ADD
      frmAddEditApArMasBranch.HeaderText = MapText("�����Ң�")
      Load frmAddEditApArMasBranch
      frmAddEditApArMasBranch.Show 1

      OKClick = frmAddEditApArMasBranch.OKClick

      Unload frmAddEditApArMasBranch
      Set frmAddEditApArMasBranch = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstBranchs)
         GridEX1.Rebind
      End If
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
     If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "EDIT" & "_" & "DELETE", "ź") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Customer.CstAddresses.Remove (ID2)
      Else
         m_Customer.CstAddresses.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstAddresses)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Tag = "CUSTOMER_BRANCH" Then
      If ID1 <= 0 Then
         m_Customer.CstBranchs.Remove (ID2)
      Else
         m_Customer.CstBranchs.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Customer.CstBranchs)
      GridEX1.Rebind
      m_HasModify = True

   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID1 As Long
Dim OKClick As Boolean
   
     If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "EDIT" & "_" & "EDIT", "���") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID1 = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditApArMasAddress.ID = ID1
      Set frmAddEditApArMasAddress.TempCollection = m_Customer.CstAddresses
      frmAddEditApArMasAddress.HeaderText = MapText("��䢷������")
      frmAddEditApArMasAddress.ShowMode = SHOW_EDIT
      Load frmAddEditApArMasAddress
      frmAddEditApArMasAddress.Show 1

      OKClick = frmAddEditApArMasAddress.OKClick

      Unload frmAddEditApArMasAddress
      Set frmAddEditApArMasAddress = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstAddresses)
         GridEX1.Rebind
      End If
  ElseIf TabStrip1.SelectedItem.Tag = "CUSTOMER_BRANCH" Then
      Set frmAddEditApArMasBranch.TempCollection = m_Customer.CstBranchs
      frmAddEditApArMasBranch.CustomerID = ID
      frmAddEditApArMasBranch.ID = ID1
      Set frmAddEditApArMasBranch.ParentForm = Me
      frmAddEditApArMasBranch.ShowMode = SHOW_EDIT
      frmAddEditApArMasBranch.HeaderText = MapText("����Ң�")
      Load frmAddEditApArMasBranch
      frmAddEditApArMasBranch.Show 1

      OKClick = frmAddEditApArMasBranch.OKClick

      Unload frmAddEditApArMasBranch
      Set frmAddEditApArMasBranch = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Customer.CstBranchs)
         GridEX1.Rebind
      End If
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
    If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "EDIT" & "_" & "SAVE", "�ѹ�֡") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
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
      ID = m_Customer.APAR_MAS_ID
      m_Customer.QueryFlag = 1
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
      If ApArInd = 1 Then
         Call LoadMaster(cboBusinessType, , , , MASTER_CUSTYPE)
         
         Call LoadPackage(m_Package, uctlPackage.MyCombo, m_Packages)
         Set uctlPackage.MyCollection = m_Packages

      ElseIf ApArInd = 2 Then
         Call LoadMaster(cboBusinessType, , , , MASTER_SUPTYPE)
                  
         Call LoadPackage(m_Package, uctlPackage.MyCombo, m_Packages)
         Set uctlPackage.MyCollection = m_Packages
      End If
'   Call LoadControl("D", uctlBirthDate.Width, True, "", , "BIRTH_DATE")
   'Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Customer.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Customer.QueryFlag = 0
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
   
   Set m_Customer = Nothing
   Set m_Employees = Nothing
   Set m_MasterRef = Nothing
   Set m_Employee = Nothing
   
   Set m_Package = New CPackage
   Set m_Packages = New Collection
   
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
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
   Col.Width = 11550
   Col.Caption = MapText("�������")
End Sub
Private Sub InitGrid3()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

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

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2235
   Col.Caption = "�����Ң�"
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5100
   Col.Caption = "�Ң�"
   
   Set Col = GridEX1.Columns.add '
   Col.Width = 2000
   Col.Caption = "���ʾ�ѡ�ҹ"
         
   GridEX1.ItemCount = 0
   GridEX1.Rebind
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   If ApArInd = 1 Then
      ApArText = MapText("�١���")
   ElseIf ApArInd = 2 Then
      ApArText = MapText("�����")
   End If
   
   Call InitNormalLabel(lblWebsite, MapText("���䫵�"))
   Call InitNormalLabel(lblShortName, MapText("����" & ApArText))
   Call InitNormalLabel(lblAparMasGroup, MapText("�����" & ApArText))
   Call InitNormalLabel(lblName, MapText("����" & ApArText))
   Call InitNormalLabel(lblShort, MapText("�������"))
   Call InitNormalLabel(lblBillName, MapText("�����͡���"))
   Call InitNormalLabel(lblEmail, MapText("�������"))
   Call InitNormalLabel(lblBusinessType, MapText("������" & ApArText))
   Call InitNormalLabel(lblBirthDate, MapText("�ѹ�Դ"))
   Call InitNormalLabel(lblBusinessDesc, MapText("��������´" & ApArText))
   Call InitNormalLabel(lblCredit, MapText("�ôԵ"))
   Call InitNormalLabel(Label2, MapText("�ѹ"))
   Call InitNormalLabel(lblDiscountPercent, MapText("% ��ǹŴ"))
   
   Call InitNormalLabel(lblPackage, MapText("���Ẻ�Ҥ�"))
   Call InitNormalLabel(lblTaxID, MapText("TAX ID"))
   Call InitNormalLabel(lblBasketFixAmount, MapText("FIX �С���"))
   
   Call InitCombo(cboBusinessType)
   If ApArInd <> 1 Then
      lblAparMasGroup.Visible = False
   End If
   
   Call InitCheckBox(chkLabelFlag, "���к����Ң�����")
   Call InitCheckBox(chkAddBranchName, "�ʴ��Ң��١���")
   Call InitCheckBox(chkFlagEdit, "�ʴ������䢢�����")
   Call InitCheckBox(chkCancelOutDocument, "¡��ԡ��â�� Credit (���ʴ��ҹ��)")
   Call InitCheckBox(chkConsignmentFlag, "���½ҡ������§ҹ�����͹")
   
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWebSite.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBusinessDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   
   Call InitGrid1
   
   Dim T As Object
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("�������")
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("�Ң�")
   T.Tag = "CUSTOMER_BRANCH"
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
   Set m_Customer = New CAPARMas
   Set m_Employees = New Collection
   Set m_MasterRef = New CMasterRef
   Set m_Employee = New CEmployee
   
   Set m_Package = New CPackage
   Set m_Packages = New Collection
   
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

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Customer.CstAddresses Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CApArAddress
      Dim Addr As CAddress
      If m_Customer.CstAddresses.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Customer.CstAddresses, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses

      Values(1) = Addr.GetFieldValue("ADDRESS_ID")
      Values(2) = RealIndex
      Values(3) = Addr.PackAddress

   ElseIf TabStrip1.SelectedItem.Tag = "CUSTOMER_BRANCH" Then
      If m_Customer.CstBranchs Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Br As CMasterRef
      If m_Customer.CstBranchs.Count <= 0 Then
         Exit Sub
      End If
      Set Br = GetItem(m_Customer.CstBranchs, RowIndex, RealIndex)
      If Br Is Nothing Then
         Exit Sub
      End If

      Values(1) = Br.KEY_ID
      Values(2) = RealIndex
      Values(3) = Br.KEY_CODE
      Values(4) = Br.KEY_NAME
      Values(5) = Br.EMP_CODE
      
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Customer.CstAddresses)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = "CUSTOMER_BRANCH" Then
      Call InitGrid3
      GridEX1.ItemCount = CountItem(m_Customer.CstBranchs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtAparMasGroupCode_Change()
   m_HasModify = True
End Sub

Private Sub txtAparMasGroupName_Change()
   m_HasModify = True
End Sub

Private Sub txtBasketFixAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtBillName_Change()
   m_HasModify = True
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtShort_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_LostFocus()
   If Not CheckUniqueNs(APARCODE_UNIQUE, txtShortName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtShortName.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub

Private Sub txtTaxID_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub
Private Sub uctlSaleByLookup_Change()
   m_HasModify = True
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
Public Sub RefreshGridBranch()

   GridEX1.ItemCount = CountItem(m_Customer.CstBranchs)
   GridEX1.Rebind
End Sub

Public Function AutoGenBranch()
Dim m_MasterRef As CMasterRef

   If m_EmployeeColl.Count <= 0 Then
      Exit Function
   End If
   
   Set m_MasterRef = New CMasterRef
   
   
   m_MasterRef.ShowMode = ShowMode
   m_MasterRef.KEY_NAME = txtShortName.Text
   m_MasterRef.KEY_CODE = txtShortName.Text
   
   Dim I As Long
   Dim TempName As String
   Dim TempEmp As CEmployee
   For Each TempEmp In m_EmployeeColl
      If TempEmp.MAINSALE_FLAG = "Y" Then
         m_MasterRef.MASTER_AREA = MASTER_APARMAS_BRANCH
         m_MasterRef.EMP_CODE = TempEmp.EMP_CODE  ' ��������
         m_MasterRef.PARENT_EX_ID = TempEmp.EMP_ID   'ID SALE
         TempName = TempEmp.EMP_CODE & " " & TempEmp.EMP_NAME & " " & TempEmp.EMP_LNAME
         Exit For
      End If
   Next TempEmp
   
   If Len(TempName) <= 2 Then
      Exit Function
   End If
   
   m_MasterRef.Flag = "A"
   Call m_Customer.CstBranchs.add(m_MasterRef)
      
   glbErrorLog.LocalErrorMsg = MapText("�к��ѵ��ѵ�����" & vbCrLf & "�����Ң�: " & txtShortName.Text & "  �Ң�: " & txtShortName.Text & "  ��о�ѡ�ҹ��� : " & TempName & "  ����")
   glbErrorLog.ShowUserError
   

'   m_MasterRef.MASTER_AREA = MASTER_APARMAS_BRANCH
'   m_MasterRef.PARENT_ID = cboParent.ItemData(Minus2Zero(cboParent.ListIndex))
'   m_MasterRef.PARENT_EX_ID3 = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
    m_MasterRef.MASTER_FLAG = "N"
    
   Set m_MasterRef = Nothing
   m_HasModify = True
   
End Function

Private Sub uctlBirthDate_HasChange()
    m_HasModify = True
End Sub

Private Sub uctlPackage_Change()
   m_HasModify = True
End Sub
