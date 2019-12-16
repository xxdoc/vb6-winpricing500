VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditFormulaItem1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmAddEditFormulaItem1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   7646
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboProductionType 
         Height          =   315
         Left            =   8460
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   990
         Width           =   1755
      End
      Begin Xivess.uctlTextLookup uctlPartItem 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1440
         Width           =   5300
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtOrderNo 
         Height          =   435
         Left            =   5100
         TabIndex        =   1
         Top             =   990
         Width           =   1455
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtFormulaItemPercent 
         Height          =   435
         Left            =   8460
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlLocation 
         Height          =   405
         Left            =   1860
         TabIndex        =   6
         Top             =   2430
         Width           =   5295
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtProblemLimitPercent 
         Height          =   435
         Left            =   8460
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBatchNO 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   1455
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlLostLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   4
         Top             =   1920
         Width           =   5295
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin VB.Label lblProductionType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6600
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
      End
      Begin Threed.SSCheck chkNext 
         Height          =   435
         Left            =   7800
         TabIndex        =   7
         Top             =   2400
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label lblProblemLimitPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7200
         TabIndex        =   20
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label1"
         Height          =   435
         Left            =   10080
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblProblemDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3525
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaItem1.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   10050
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblFormulaItemPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7200
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblOrderNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6825
         TabIndex        =   10
         Top             =   3120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5175
         TabIndex        =   9
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaItem1.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditFormulaItem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object

Public TempCollection As Collection
Private m_PartItemColls As Collection
Private m_LocationColls As Collection
Private m_LostColls As Collection
Private Sub cboProductionType_Click()
   m_HasModify = True
End Sub
Private Sub cboProductionType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkNext_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkNext_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      txtBatchNO.Text = ""
      txtOrderNo.Text = ""
      uctlPartItem.MyCombo.ListIndex = -1
      uctlLocation.MyCombo.ListIndex = -1
      txtFormulaItemPercent.Text = ""
      uctlLostLookup.MyCombo.ListIndex = -1
      txtProblemLimitPercent.Text = ""
   End If
   
   Call txtBatchNO.SetFocus
   Call ParentForm.RefreshGrid
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
      
   OKClick = True
   Unload Me

End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim PaymentType As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      
      Dim FI As CFormulaItem
      Set FI = TempCollection.Item(ID)
      
      txtBatchNO.Text = FI.BATCH_NO
      txtOrderNo.Text = FI.ORDER_NO
      txtFormulaItemPercent.Text = FI.FORMULA_ITEM_PERCENT
      uctlPartItem.MyCombo.ListIndex = IDToListIndex(uctlPartItem.MyCombo, FI.PART_ITEM_ID)
      uctlLocation.MyCombo.ListIndex = IDToListIndex(uctlLocation.MyCombo, FI.LOCATION_ID)
      uctlLostLookup.MyCombo.ListIndex = IDToListIndex(uctlLostLookup.MyCombo, FI.LOST_ID)
      txtProblemLimitPercent.Text = FI.PROBLEM_LIMIT_PERCENT
      chkNext.Value = FlagToCheck(FI.NEXT_FLAG)
      cboProductionType.ListIndex = IDToListIndex(cboProductionType, FI.PRODUCTION_TYPE)
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblBatchNo, txtBatchNO, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblOrderNo, txtOrderNo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   Dim FI As CFormulaItem
   
   If ShowMode = SHOW_ADD Then
      Set FI = New CFormulaItem
      FI.Flag = "A"
      Call TempCollection.add(FI)
   Else
      Set FI = TempCollection.Item(ID)
      If FI.Flag <> "A" Then
         FI.Flag = "E"
      End If
   End If
   
   FI.BATCH_NO = Val(txtBatchNO.Text)
   FI.ORDER_NO = Val(txtOrderNo.Text)
   FI.PART_ITEM_ID = uctlPartItem.MyCombo.ItemData(Minus2Zero(uctlPartItem.MyCombo.ListIndex))
   FI.PART_NO = uctlPartItem.MyTextBox.Text
   FI.PART_DESC = uctlPartItem.MyCombo.Text
   FI.FORMULA_ITEM_PERCENT = Val(txtFormulaItemPercent.Text)
   FI.LOCATION_ID = uctlLocation.MyCombo.ItemData(Minus2Zero(uctlLocation.MyCombo.ListIndex))
   FI.LOCATION_NO = uctlLocation.MyTextBox.Text
   FI.LOCATION_NAME = uctlLocation.MyCombo.Text
   
   FI.LOST_ID = uctlLostLookup.MyCombo.ItemData(Minus2Zero(uctlLostLookup.MyCombo.ListIndex))
   FI.PROBLEM_DESC = uctlLostLookup.MyCombo.Text
   
   FI.PROBLEM_LIMIT_PERCENT = Val(txtProblemLimitPercent.Text)
   
   FI.NEXT_FLAG = Check2Flag(chkNext.Value)
   FI.PRODUCTION_TYPE = cboProductionType.ItemData(Minus2Zero(cboProductionType.ListIndex))
   FI.PRODUCTION_TYPE_NAME = cboProductionType.Text
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadStockCode(uctlPartItem.MyCombo, m_PartItemColls)
      Set uctlPartItem.MyCollection = m_PartItemColls
      
      Call LoadMaster(uctlLocation.MyCombo, m_LocationColls, , , MASTER_LOCATION)
      Set uctlLocation.MyCollection = m_LocationColls
      
      Call LoadMaster(uctlLostLookup.MyCombo, m_LostColls, , , MASTER_PRODUCTION_LOST)
      Set uctlLostLookup.MyCollection = m_LostColls
      
      Call LoadMaster(cboProductionType, , , , MASTER_PRODUCTION_TYPE)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         m_HasModify = False
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblBatchNo, MapText("BATCH_NO"))
   Call InitNormalLabel(lblOrderNo, MapText("ลำดับการแสดง"))
   Call InitNormalLabel(lblPartItem, MapText("สินค้า"))
   Call InitNormalLabel(lblFormulaItemPercent, MapText("มาตรฐาน"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("%"))
   Call InitNormalLabel(Label3, MapText("%"))
   Call InitNormalLabel(lblProblemDesc, MapText("ERROR."))
   Call InitNormalLabel(lblProblemLimitPercent, MapText("ไม่เกิน"))
   Call InitNormalLabel(lblProductionType, MapText("ประเภทสินค้าผลิต"))
   
   Call txtBatchNO.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtOrderNo.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtFormulaItemPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtProblemLimitPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkNext, "นำไปยังโปรเซสถัดไป")
   Call InitCombo(cboProductionType)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartItemColls = New Collection
   Set m_LocationColls = New Collection
   Set m_LostColls = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   
   Set m_PartItemColls = Nothing
   Set m_LocationColls = Nothing
   Set m_LostColls = Nothing
   
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
   
End Sub

Private Sub txtBatchNO_Change()
   m_HasModify = True
End Sub

Private Sub txtFormulaItemPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtOrderNo_Change()
   m_HasModify = True
End Sub

Private Sub txtProblemDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtProblemLimitPercent_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocation_Change()
   m_HasModify = True
End Sub

Private Sub uctlLostLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartItem_Change()
   m_HasModify = True
End Sub
