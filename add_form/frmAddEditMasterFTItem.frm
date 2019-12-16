VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMasterFTItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMasterFTItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   11033
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboGroupCom 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   1725
      End
      Begin Xivess.uctlTextBox txtTo 
         Height          =   435
         Left            =   4080
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtFrom 
         Height          =   435
         Left            =   1440
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtValue1 
         Height          =   435
         Left            =   3000
         TabIndex        =   11
         Top             =   3960
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtValue2 
         Height          =   435
         Left            =   3000
         TabIndex        =   12
         Top             =   4440
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtValue3 
         Height          =   435
         Left            =   3000
         TabIndex        =   13
         Top             =   4920
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1720
         _Version        =   131073
         Caption         =   "ประเภทตัวคูณ"
         Begin Threed.SSOption SSOption6 
            Height          =   495
            Left            =   4920
            TabIndex        =   10
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption5 
            Height          =   495
            Left            =   2520
            TabIndex        =   9
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption4 
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2355
         _Version        =   131073
         Caption         =   "ช่วงเปรียบเทียบ"
         Begin Threed.SSOption SSOption8 
            Height          =   495
            Left            =   4920
            TabIndex        =   5
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption7 
            Height          =   495
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption1 
            Height          =   495
            Left            =   2520
            TabIndex        =   1
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption2 
            Height          =   495
            Left            =   4920
            TabIndex        =   2
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
         Begin Threed.SSOption SSOption3 
            Height          =   495
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            _Version        =   131073
            Caption         =   "SSOption1"
         End
      End
      Begin VB.Label lblGroupCom 
         Alignment       =   1  'Right Justify
         Caption         =   "L"
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   120
         Width           =   1635
      End
      Begin VB.Line Line9 
         X1              =   4680
         X2              =   5280
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line8 
         X1              =   5280
         X2              =   5280
         Y1              =   3720
         Y2              =   4200
      End
      Begin VB.Line Line7 
         X1              =   3120
         X2              =   3120
         Y1              =   3600
         Y2              =   3960
      End
      Begin VB.Line Line6 
         X1              =   720
         X2              =   1200
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line5 
         X1              =   720
         X2              =   720
         Y1              =   3720
         Y2              =   4680
      End
      Begin VB.Line Line4 
         X1              =   4920
         X2              =   4920
         Y1              =   2160
         Y2              =   2280
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   2280
         Y1              =   2160
         Y2              =   2280
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4920
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   120
         Y1              =   2040
         Y2              =   2160
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   4680
         TabIndex        =   27
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   4680
         TabIndex        =   26
         Top             =   4560
         Width           =   1755
      End
      Begin VB.Label lblValue3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label lblValue2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   4560
         Width           =   1635
      End
      Begin VB.Label lblValue1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   4080
         Width           =   1635
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1155
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1005
         TabIndex        =   14
         Top             =   5550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterFTItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2655
         TabIndex        =   15
         Top             =   5550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterFTItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4305
         TabIndex        =   16
         Top             =   5550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   2310
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmAddEditMasterFTItem"
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

Private m_MasterFromToDetail As CMasterFromToDetail

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Public StepFlag  As Boolean

Public DocumentType As MASTER_COMMISSION_AREA
Private Sub cboGroupCom_Click()
   m_HasModify = True
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
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblFrom, MapText("จาก"))
   Call InitNormalLabel(lblTo, MapText("ถึง"))
   Call InitNormalLabel(lblGroupCom, MapText("กลุ่มคอมมิตชั่น"))
   
   If DocumentType = COMMISSION_TABLE Then
      Call InitNormalLabel(lblValue1, MapText("% คอมมิตชั่น"))
      Call InitNormalLabel(lblValue2, MapText("ยอดคอมมิตชั่น"))
   Else
      Call InitNormalLabel(lblValue1, MapText("%คงเหลือ"))
      Call InitNormalLabel(lblValue2, MapText("ยอดส่วนหัก"))
   End If
   
   Call InitNormalLabel(lblValue3, MapText("ยอดสะสมของช่วง"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   
   Call txtFrom.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTo.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtValue1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtValue2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtValue3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitOptionEx(SSOption1, "ช่วงจำนวน")
   Call InitOptionEx(SSOption2, "ช่วงยอดขาย")
   Call InitOptionEx(SSOption3, "%ยอดเป้าการขาย")
   Call InitOptionEx(SSOption7, "ช่วง%รับคืน/จำนวน")
   Call InitOptionEx(SSOption8, "ช่วง%รับคืน/ยอดขาย")
   
   If DocumentType = RETURN_TABLE Then
      SSOption1.Enabled = False
      SSOption2.Enabled = False
      SSOption3.Enabled = False
      SSOption5.Enabled = False
      SSOption7.Enabled = False
   ElseIf DocumentType = COMMISSION_TABLE Then
      SSOption7.Enabled = False
      SSOption8.Enabled = False
   End If
   
   Call InitOptionEx(SSOption4, "ไม่คูณ")
   Call InitOptionEx(SSOption5, "ยอดจำนวน")
   Call InitOptionEx(SSOption6, "ยอดขาย")
   Call InitCombo(cboGroupCom)
   
   SSOption1.Value = True
   SSOption4.Value = True
   
   If StepFlag Then
      lblValue3.Enabled = False
      txtValue3.Enabled = False
   End If
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CMasterFromToDetail
         
         Set BD = TempCollection.Item(ID)
                  
         SSOption1.Value = StringToCheckSSoption(BD.GetFieldValue("AMOUNT_FLAG"))
         SSOption2.Value = StringToCheckSSoption(BD.GetFieldValue("VALUE_FLAG"))
         SSOption3.Value = StringToCheckSSoption(BD.GetFieldValue("TAGET_VALUE_FLAG"))
         SSOption7.Value = StringToCheckSSoption(BD.GetFieldValue("AMOUNT_P_FLAG"))
         SSOption8.Value = StringToCheckSSoption(BD.GetFieldValue("VALUE_P_FLAG"))
         
         SSOption4.Value = StringToCheckSSoption(BD.GetFieldValue("NO_X_FLAG"))
         SSOption5.Value = StringToCheckSSoption(BD.GetFieldValue("AMOUNT_X_FLAG"))
         SSOption6.Value = StringToCheckSSoption(BD.GetFieldValue("VALUE_X_FLAG"))
         
         txtFrom.Text = BD.GetFieldValue("MASTER_FROMTO_DETAIL_FROM")
         txtTo.Text = BD.GetFieldValue("MASTER_FROMTO_DETAIL_TO")
         txtValue1.Text = BD.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1")
         txtValue2.Text = BD.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE2")
         txtValue3.Text = BD.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE3")
         
         cboGroupCom.ListIndex = IDToListIndex(cboGroupCom, BD.GetFieldValue("GROUP_COM_ID"))
         
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
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      txtFrom.Text = ""
      txtTo.Text = ""
      txtValue1.Text = ""
      txtValue2.Text = ""
      txtValue3.Text = ""
   End If
   Call QueryData(True)
   
   Call txtFrom.SetFocus
   
   Call ParentForm.RefreshGrid
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblFrom, txtFrom, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblTo, txtTo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblValue3, txtValue3, Not (txtValue3.Enabled)) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim BD As CMasterFromToDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CMasterFromToDetail
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   Call BD.SetFieldValue("AMOUNT_FLAG", CheckSSoptionToString(SSOption1.Value))
   Call BD.SetFieldValue("VALUE_FLAG", CheckSSoptionToString(SSOption2.Value))
   Call BD.SetFieldValue("TAGET_VALUE_FLAG", CheckSSoptionToString(SSOption3.Value))
   Call BD.SetFieldValue("AMOUNT_P_FLAG", CheckSSoptionToString(SSOption7.Value))
   Call BD.SetFieldValue("VALUE_P_FLAG", CheckSSoptionToString(SSOption8.Value))
   
   Call BD.SetFieldValue("NO_X_FLAG", CheckSSoptionToString(SSOption4.Value))
   Call BD.SetFieldValue("AMOUNT_X_FLAG", CheckSSoptionToString(SSOption5.Value))
   Call BD.SetFieldValue("VALUE_X_FLAG", CheckSSoptionToString(SSOption6.Value))
   
   Call BD.SetFieldValue("MASTER_FROMTO_DETAIL_FROM", Val(txtFrom.Text))
   Call BD.SetFieldValue("MASTER_FROMTO_DETAIL_TO", Val(txtTo.Text))
   Call BD.SetFieldValue("MASTER_FROMTO_DETAIL_VALUE1", Val(txtValue1.Text))  '%
   Call BD.SetFieldValue("MASTER_FROMTO_DETAIL_VALUE2", Val(txtValue2.Text))  'ยอดเงิน
   Call BD.SetFieldValue("MASTER_FROMTO_DETAIL_VALUE3", Val(txtValue3.Text))  'ยอดสะสม
   
   Call BD.SetFieldValue("GROUP_COM_ID", cboGroupCom.ItemData(Minus2Zero(cboGroupCom.ListIndex)))
   Call BD.SetFieldValue("GROUP_COM_DESC", cboGroupCom.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(cboGroupCom, , , , MASTER_GROUP_COM)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      'm_HasModify = False
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_MasterFromToDetail = New CMasterFromToDetail
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_MasterFromToDetail = Nothing

End Sub

Private Sub SSOption1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSOption1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption3_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption4_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption5_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption6_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub SSOption6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption7_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub SSOption8_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub SSOption8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtFrom_Change()
   m_HasModify = True
End Sub

Private Sub txtTo_Change()
   m_HasModify = True
End Sub

Private Sub txtValue1_Change()
   m_HasModify = True
End Sub

Private Sub txtValue2_Change()
   m_HasModify = True
End Sub

Private Sub txtValue3_Change()
   m_HasModify = True
End Sub
