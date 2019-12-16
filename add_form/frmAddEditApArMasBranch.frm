VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditApArMasBranch 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditApArMasBranch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame Frame1 
      Height          =   2565
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4524
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
         TabIndex        =   5
         Top             =   1560
         Width           =   8685
      End
      Begin Xivess.uctlTextLookup uctlSale 
         Height          =   435
         Left            =   2520
         TabIndex        =   4
         Top             =   1150
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
         Width           =   5385
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlDealer 
         Height          =   435
         Left            =   2520
         TabIndex        =   6
         Top             =   1965
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblDealer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2325
      End
      Begin Threed.SSCheck chkMaster 
         Height          =   555
         Left            =   8040
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   979
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1590
         Width           =   2325
      End
      Begin VB.Label lblParentEx 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1110
         Width           =   2325
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         TabIndex        =   15
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2325
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   705
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3360
         TabIndex        =   7
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   4950
         TabIndex        =   8
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
         Left            =   6585
         TabIndex        =   9
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
         TabIndex        =   13
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditApArMasBranch"
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

Public ParentForm As Form
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public MasterKey As String

Private m_MasterRef As CMasterRef

Public KEY_CODE As String
Public KEY_NAME As String
Public MasterMode As Long
Public m_TempMr As CMasterRef
Public m_TempEmp As CEmployee

Private m_Emp As Collection
Private m_Apm  As CAPARMas
Private m_Customers As Collection
Private m_Adr As CAddress

Public TempCollection As Collection
Public CustomerID As Long

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
   
   Call InitNormalLabel(lblCode, "รหัสสาขา")
   Call InitNormalLabel(lblName, "สาขา")
   Call InitNormalLabel(lblParentEx, MapText("ผู้รับผิดชอบ"))
   Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่สาขา"))
   Call InitNormalLabel(lblDealer, MapText("ตัวแทน"))
   
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
   
   Call InitCombo(cboParent)
   Call InitCombo(cboCustomerAddress)
   
   Call InitCheckBox(chkMaster, MapText("จอง"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim iCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim m_MasterRef As CMasterRef
         Set m_MasterRef = TempCollection.Item(ID)
         
         txtCode.Text = m_MasterRef.KEY_CODE
         txtName.Text = m_MasterRef.KEY_NAME
         cboParent.ListIndex = IDToListIndex(cboParent, m_MasterRef.PARENT_ID)
         
         uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_MasterRef.PARENT_EX_ID)
         cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_MasterRef.PARENT_EX_ID3)
         uctlDealer.MyCombo.ListIndex = IDToListIndex(uctlDealer.MyCombo, m_MasterRef.DEALER_ID)
         
         chkMaster.Value = FlagToCheck(m_MasterRef.MASTER_FLAG)
         
          Set m_MasterRef = Nothing
      End If
      
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGridBranch
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      txtCode.Text = ""
      txtName.Text = ""
      cboParent.ListIndex = -1
      uctlSale.MyCombo.ListIndex = -1
      
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGridBranch
   
   Call txtCode.SetFocus
End Sub

Private Sub cmdSave_Click()
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

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   If Not VerifyCombo(lblParentEx, uctlSale.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   Dim m_MasterRef As CMasterRef
   If ShowMode = SHOW_ADD Then
      Set m_MasterRef = New CMasterRef
      m_MasterRef.Flag = "A"
      Call TempCollection.add(m_MasterRef)
   Else
      Set m_MasterRef = TempCollection.Item(ID)
      If m_MasterRef.Flag <> "A" Then
         m_MasterRef.Flag = "E"
      End If
   End If
   
'   If Not CheckUniqueNs(MASTER_CODE, txtCode.Text, m_MasterRef.KEY_ID, MASTER_APARMAS_BRANCH) Then              'เนื่องจากสาขารอบนี้เดิมซ้ำไม่ได้แต่ตอนนี้ อิออน 2 รหัสลูกค้าจะต้องมีสาขาซ้ำกันในการที่ DC จะทำ BLOACK ลาเบล
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   m_MasterRef.ShowMode = ShowMode
   m_MasterRef.KEY_NAME = txtName.Text
   m_MasterRef.KEY_CODE = txtCode.Text
   m_MasterRef.MASTER_AREA = MASTER_APARMAS_BRANCH
   m_MasterRef.PARENT_ID = cboParent.ItemData(Minus2Zero(cboParent.ListIndex))
   
   m_MasterRef.PARENT_EX_ID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   
   m_MasterRef.EMP_CODE = uctlSale.MyTextBox.Text
   
   m_MasterRef.PARENT_EX_ID3 = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_MasterRef.DEALER_ID = uctlDealer.MyCombo.ItemData(Minus2Zero(uctlDealer.MyCombo.ListIndex))
   
   m_MasterRef.MASTER_FLAG = Check2Flag(chkMaster.Value)
   
   Set m_MasterRef = Nothing

   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(cboParent, , , , MASTER_CUSTOMER_BLOCK)
      cboParent.Visible = True
      
      m_TempEmp.EMP_ID = -1
      Call LoadEmployee(m_TempEmp, uctlSale.MyCombo, m_EmployeeColl)
      Set uctlSale.MyCollection = m_EmployeeColl
      uctlSale.Visible = True
      
      m_TempEmp.EMP_ID = -1
      Call LoadEmployee(m_TempEmp, uctlDealer.MyCombo, m_EmployeeColl)
      Set uctlDealer.MyCollection = m_EmployeeColl
      uctlDealer.Visible = True
         
      Call m_Adr.SetFieldValue("APAR_MAS_ID", CustomerID)
      Call LoadAparMasAddress(m_Adr, cboCustomerAddress, , True)
         
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
      Call cmdNext_Click
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
   Set m_Emp = New Collection
   Set m_Apm = New CAPARMas
   Set m_Adr = New CAddress
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_TempMr = Nothing
   Set m_MasterRef = Nothing
   Set m_Emp = Nothing
   Set m_TempEmp = Nothing
End Sub
Private Sub txtCode_Change()
   m_HasModify = True
End Sub
Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub uctlDealer_Change()
   m_HasModify = True
End Sub

Private Sub uctlSale_Change()
   m_HasModify = True
End Sub
