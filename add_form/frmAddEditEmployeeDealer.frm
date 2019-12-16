VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditEmployeeDealer 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmAddEditEmployeeDealer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   7646
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDealerType 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   3915
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   1755
      End
      Begin Xivess.uctlTextBox txtYear 
         Height          =   435
         Left            =   5640
         TabIndex        =   1
         Top             =   960
         Width           =   1095
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4560
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDealerType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1365
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEmployeeDealer.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblMonth 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4665
         TabIndex        =   5
         Top             =   2280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3015
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEmployeeDealer.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditEmployeeDealer"
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
Private Sub cboDealerType_Click()
   m_HasModify = True
End Sub
Private Sub cboDealerType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboMonth_Click()
   m_HasModify = True
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
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
      cboDealerType.ListIndex = -1
      
   End If
   
   Call cboMonth.SetFocus
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
      
      Dim EmpDl As CEmployeeDealer
      Set EmpDl = TempCollection.Item(ID)
      
      cboMonth.ListIndex = IDToListIndex(cboMonth, Val(Right(EmpDl.YYYYMM, 2)))
      txtYear.Text = Val(Left(EmpDl.YYYYMM, 4)) + 543
      cboDealerType.ListIndex = IDToListIndex(cboDealerType, EmpDl.DEALER_TYPE)
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyCombo(lblMonth, cboMonth, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblYear, txtYear, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblDealerType, cboDealerType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EmpDl As CEmployeeDealer
   
   If ShowMode = SHOW_ADD Then
      Set EmpDl = New CEmployeeDealer
      EmpDl.Flag = "A"
      Call TempCollection.add(EmpDl)
   Else
      Set EmpDl = TempCollection.Item(ID)
      If EmpDl.Flag <> "A" Then
         EmpDl.Flag = "E"
      End If
   End If
   
   EmpDl.YYYYMM = (Val(txtYear.Text) - 543) & Format(cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex)), "00")
   EmpDl.DEALER_TYPE = cboDealerType.ItemData(Minus2Zero(cboDealerType.ListIndex))
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitThaiMonth(cboMonth)
      Call LoadDealerType(cboDealerType)
      
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
   
   Call InitNormalLabel(lblMonth, MapText("เดือน"))
   Call InitNormalLabel(lblYear, MapText("ปี พ.ศ."))
   Call InitNormalLabel(lblDealerType, MapText("ประเภทตัวแทน"))
   
   Call txtYear.SetTextLenType(TEXT_INTEGER, glbSetting.YEAR_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitCombo(cboMonth)
   Call InitCombo(cboDealerType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub txtYear_Change()
   m_HasModify = True
End Sub
