VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMaster1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMaster1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame Frame1 
      Height          =   3285
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5794
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCustomerAddress 
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   6885
      End
      Begin Xivess.uctlTextLookup uctlSale 
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         Top             =   1150
         Visible         =   0   'False
         Width           =   5380
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboParent 
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
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Visible         =   0   'False
         Width           =   3525
      End
      Begin Xivess.uctlTextBox txtCode 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   210
         Width           =   1845
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   2520
         TabIndex        =   2
         Top             =   660
         Width           =   4425
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   2520
         TabIndex        =   6
         Top             =   1605
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtShortCode 
         Height          =   435
         Left            =   6960
         TabIndex        =   3
         Top             =   660
         Width           =   885
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlDealer 
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   2480
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkMaster2 
         Height          =   555
         Left            =   7920
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   979
         _Version        =   131073
         Caption         =   "chkMaster2"
      End
      Begin VB.Label lblDealer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2480
         Visible         =   0   'False
         Width           =   2325
      End
      Begin Threed.SSCheck chkMaster 
         Height          =   555
         Left            =   7920
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   979
         _Version        =   131073
         Caption         =   "chkMaster"
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2070
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lblParentEx 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1110
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         TabIndex        =   16
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2325
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   705
      Left            =   0
      TabIndex        =   13
      Top             =   3840
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   3510
         TabIndex        =   9
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5145
         TabIndex        =   10
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   14
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMaster1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public MasterKey As String
Public MasterArea As MASTER_TYPE

Private m_MasterRef As CMasterRef

Public KEY_CODE As String
Public KEY_NAME As String
Public MasterMode As Long
Public m_TempMr As CMasterRef
Public m_TempEmp As CEmployee

Private m_Apm  As CAPARMas
Private m_Adr As CAddress

Private m_BankColl As Collection
Private m_BankBranch As Collection

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub
Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboParent_Click()
   m_HasModify = True
End Sub
Private Sub cboParent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkMaster_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkMaster_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkMaster2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkMaster2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, "")
   Call InitNormalLabel(lblName, "")

   Call InitNormalLabel(lblCode, MapText(KEY_CODE))
   Call InitNormalLabel(lblName, MapText(KEY_NAME))
      
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdCancel, MapText("¡��ԡ (ESC)"))
   Call InitCombo(cboParent)
   Call InitCombo(cboCustomerAddress)
   
   If MasterKey = MASTER_LOCATION & "-X" Then
      Call InitCheckBox(chkMaster, MapText("�ͧ"))
   ElseIf MasterKey = MASTER_PRODUCTION_LOST & "-X" Then
      Call InitCheckBox(chkMaster, MapText("���"))
   ElseIf MasterKey = MASTER_PRODUCTION_TYPE & "-X" Then
      Call InitCheckBox(chkMaster, MapText("�ʴ���������´"))
   ElseIf MasterKey = MASTER_PRODUCTION_LOCATION & "-X" Then
      Call InitCheckBox(chkMaster, MapText("ʶҹ����Ե��ѡ"))
   ElseIf MasterKey = MASTER_TRANSPORTOR & "-X" Then
      Call InitCheckBox(chkMaster, MapText("���Թ���·ҧ"))
      Call InitCheckBox(chkMaster2, MapText("�ʴ��������"))
   End If
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      m_MasterRef.KEY_ID = ID
      Call m_MasterRef.QueryData(1, m_Rs, itemcount, True)
      If itemcount > 0 Then
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         txtCode.Text = m_MasterRef.KEY_CODE
         txtName.Text = m_MasterRef.KEY_NAME
         If MasterKey = MASTER_INVENTORY_SUB_TYPE & "-X" Then
            cboParent.ListIndex = IDToListIndex(cboParent, m_MasterRef.INDEX_LINK)
         Else
            cboParent.ListIndex = IDToListIndex(cboParent, m_MasterRef.PARENT_ID)
         End If
         txtShortCode.Text = m_MasterRef.SHORT_CODE
                  
         If MasterKey = "21-X" Then    '�Ң��١���
            uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_MasterRef.PARENT_EX_ID)
            uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_MasterRef.PARENT_EX_ID2)
            cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_MasterRef.PARENT_EX_ID3)
            uctlDealer.MyCombo.ListIndex = IDToListIndex(uctlDealer.MyCombo, m_MasterRef.DEALER_ID)
         ElseIf MasterKey = "26-X" Then   '�ѭ�ո�Ҥ��
            uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_MasterRef.PARENT_EX_ID4)
            uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_MasterRef.PARENT_EX_ID5)
         ElseIf MasterKey = "43-X" Then   '�������ѧ��ѡ�ҹ���
            uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_MasterRef.PARENT_EX_ID4)
         End If
         chkMaster.Value = FlagToCheck(m_MasterRef.MASTER_FLAG)
         
        If MasterKey = MASTER_TRANSPORTOR & "-X" Then
           chkMaster.Value = FlagToCheck(m_MasterRef.CASH_DELIVERY_FLAG)
           chkMaster2.Value = m_MasterRef.INDEX_LINK
        End If
         
      End If
      Call EnableForm(Me, True)
   End If
   
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdSave_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, Not uctlCustomerLookup.Visible) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(MASTER_CODE, txtCode.Text, ID, Val(MasterArea)) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtCode.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
'   If Not CheckUniqueNs(MASTER_NAME, txtName.Text, ID, Val(MasterArea)) Then
'      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtName.Text & " " & MapText("������к�����")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
      
   
   m_MasterRef.ShowMode = ShowMode
   m_MasterRef.KEY_NAME = txtName.Text
   m_MasterRef.KEY_CODE = txtCode.Text
   m_MasterRef.MASTER_AREA = MasterArea
   If MasterKey = MASTER_INVENTORY_SUB_TYPE & "-X" Then
      m_MasterRef.INDEX_LINK = cboParent.ItemData(Minus2Zero(cboParent.ListIndex))
   Else
      m_MasterRef.PARENT_ID = cboParent.ItemData(Minus2Zero(cboParent.ListIndex))
   End If
   m_MasterRef.SHORT_CODE = txtShortCode.Text
   
   If MasterKey = "21-X" Then
      m_MasterRef.PARENT_EX_ID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
      m_MasterRef.PARENT_EX_ID2 = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_MasterRef.PARENT_EX_ID3 = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
      m_MasterRef.DEALER_ID = uctlDealer.MyCombo.ItemData(Minus2Zero(uctlDealer.MyCombo.ListIndex))
   ElseIf MasterKey = "26-X" Then
      m_MasterRef.PARENT_EX_ID4 = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
      m_MasterRef.PARENT_EX_ID5 = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   ElseIf MasterKey = "43-X" Then
      m_MasterRef.PARENT_EX_ID4 = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   End If
   
   m_MasterRef.MASTER_FLAG = Check2Flag(chkMaster.Value)
   If MasterKey = MASTER_TRANSPORTOR & "-X" Then
      m_MasterRef.CASH_DELIVERY_FLAG = Check2Flag(chkMaster.Value)
      m_MasterRef.INDEX_LINK = chkMaster2.Value
   End If
   
   Call glbDaily.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   
   Call EnableForm(Me, True)
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Call EnableForm(Me, True)
   SaveData = False
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If MasterKey = "12-X" Then
         Call LoadMaster(cboParent, , , , MASTER_STOCKGROUP)
         cboParent.Visible = True
         
         Call InitEmptyCombo(uctlSale.MyCombo)
         
         Call InitEmptyCombo(uctlCustomerLookup.MyCombo)
         
         Call InitEmptyCombo(cboCustomerAddress)
      ElseIf MasterKey = "17-X" Then
         Call LoadMaster(cboParent, , , , MASTER_BANK)
         cboParent.Visible = True
         
         Call InitEmptyCombo(uctlSale.MyCombo)
         
         Call InitEmptyCombo(uctlCustomerLookup.MyCombo)
         
         Call InitEmptyCombo(cboCustomerAddress)
      ElseIf MasterKey = "21-X" Then
         Call LoadMaster(cboParent, , , , MASTER_CUSTOMER_BLOCK)
         cboParent.Visible = True
         
         m_TempEmp.EMP_ID = -1
         Call LoadEmployee(m_TempEmp, uctlSale.MyCombo, m_EmployeeColl)
         Set uctlSale.MyCollection = m_EmployeeColl
         uctlSale.Visible = True
         
         Call InitNormalLabel(lblParentEx, MapText("����Ѻ�Դ�ͺ"))
         lblParentEx.Visible = True
         
         m_TempEmp.EMP_ID = -1
         Call LoadEmployee(m_TempEmp, uctlDealer.MyCombo, m_EmployeeColl)
         Set uctlDealer.MyCollection = m_EmployeeColl
         uctlDealer.Visible = True
         
         Call InitNormalLabel(lblDealer, MapText("���᷹"))
         lblDealer.Visible = True
         
         m_Apm.APAR_IND = 1
         Call LoadApArMas(m_Apm, uctlCustomerLookup.MyCombo, m_CustomerColl)
         Set uctlCustomerLookup.MyCollection = m_CustomerColl
         uctlCustomerLookup.Visible = True
         
         Call InitNormalLabel(lblCustomer, MapText("�١���"))
         lblCustomer.Visible = True
         
         cboCustomerAddress.Visible = True
         Call InitNormalLabel(lblCustomerAddress, MapText("��������Ң�"))
         lblCustomerAddress.Visible = True
      ElseIf MasterKey = "26-X" Then
         Call LoadMaster(cboParent, , , , MASTER_BACCOUNT_TYPE)
         cboParent.Visible = True
                  
         Call LoadMaster(uctlSale.MyCombo, m_BankColl, , , MASTER_BANK)
         Set uctlSale.MyCollection = m_BankColl
         uctlSale.Visible = True
         
         Call InitNormalLabel(lblParentEx, MapText("��Ҥ��"))
         lblParentEx.Visible = True
         
         Call LoadMaster(uctlCustomerLookup.MyCombo, m_BankBranch, , , MASTER_BBRANCH)
         Set uctlCustomerLookup.MyCollection = m_BankBranch
         uctlCustomerLookup.Visible = True
         
         Call InitNormalLabel(lblCustomer, MapText("�ҢҸ�Ҥ��"))
         lblCustomer.Visible = True
         
         Call InitEmptyCombo(cboCustomerAddress)
      ElseIf MasterKey = MASTER_CUSTYPE & "-X" Then
         Call LoadMaster(cboParent, , , , MASTER_CUSGROUP)
         cboParent.Visible = True
         
         Call InitEmptyCombo(uctlSale.MyCombo)
         
         Call InitEmptyCombo(uctlCustomerLookup.MyCombo)
         
         Call InitEmptyCombo(cboCustomerAddress)
      ElseIf MasterKey = MASTER_INVENTORY_SUB_TYPE & "-X" Then
         
         Call InitInventoryDocType(cboParent)
         cboParent.Visible = True
         
         Call InitEmptyCombo(uctlSale.MyCombo)
         
         Call InitEmptyCombo(uctlCustomerLookup.MyCombo)
         
         Call InitEmptyCombo(cboCustomerAddress)
    ElseIf MasterKey = "14-X" Then
         Call LoadMaster(cboParent, , , , MASTER_INVENTORY_SALE_GROUP)
         cboParent.Visible = True
'
'         m_TempEmp.EMP_ID = -1
'         Call LoadEmployee(m_TempEmp, uctlSale.MyCombo, m_EmployeeColl)
'         Set uctlSale.MyCollection = m_EmployeeColl
'         uctlSale.Visible = True
'
'         Call InitNormalLabel(lblParentEx, MapText("����Ѻ�Դ�ͺ"))
'         lblParentEx.Visible = True

      Else
         Call InitEmptyCombo(cboParent)
         
         Call InitEmptyCombo(uctlSale.MyCombo)
         
         Call InitEmptyCombo(uctlCustomerLookup.MyCombo)
         
         Call InitEmptyCombo(cboCustomerAddress)
      
         Call InitEmptyCombo(uctlDealer.MyCombo)
      End If
      If MasterKey = MASTER_LOCATION & "-X" Or MasterKey = MASTER_PRODUCTION_LOST & "-X" Or _
      MasterKey = MASTER_PRODUCTION_TYPE & "-X" Or MasterKey = MASTER_PRODUCTION_LOCATION & "-X" Then
         chkMaster.Visible = True
      ElseIf MasterKey = MASTER_TRANSPORTOR & "-X" Then
         chkMaster.Visible = True
         chkMaster2.Visible = True
      End If
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdSave_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()
   
   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TempMr = New CMasterRef
   Set m_MasterRef = New CMasterRef
   Set m_TempEmp = New CEmployee
   Set m_Apm = New CAPARMas
   Set m_Adr = New CAddress
   Set m_BankColl = New Collection
   Set m_BankBranch = New Collection
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_TempMr = Nothing
   Set m_MasterRef = Nothing
   Set m_TempEmp = Nothing
   Set m_BankColl = Nothing
   Set m_BankBranch = Nothing
End Sub


Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtShortCode_Change()
   m_HasModify = True
End Sub

Private Sub uctlDealer_Change()
   m_HasModify = True
End Sub

Private Sub uctlSale_Change()
   If MasterKey = "26-X" Then
      Dim SaleID As Long
      Static OldSaleID As Long
      SaleID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
      If OldSaleID <> SaleID Then
         OldSaleID = SaleID
      Else
         Exit Sub
      End If
      Dim Sale As CMasterRef
      If SaleID > 0 Then
         Set Sale = m_BankColl(Trim(Str(SaleID)))
         Call LoadMaster(uctlCustomerLookup.MyCombo, m_BankBranch, , , MASTER_BBRANCH, , Sale.KEY_ID)
         Set uctlCustomerLookup.MyCollection = m_BankBranch
         uctlCustomerLookup.Visible = True
      End If
      Set Sale = Nothing
   End If
   m_HasModify = True
End Sub
Private Sub uctlCustomerLookup_Change()
   If MasterKey = "21-X" Then
      Dim CustomerID As Long
      Dim Customer As CAPARMas
      CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      If CustomerID > 0 Then
         Set Customer = m_CustomerColl(Trim(Str(CustomerID)))
         Call m_Adr.SetFieldValue("APAR_MAS_ID", CustomerID)
         Call LoadApArMasAddress(m_Adr, cboCustomerAddress, , True)
      End If
   End If
      
   m_HasModify = True
End Sub

