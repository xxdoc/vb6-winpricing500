VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExtraDiscount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditExtraDiscount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame Frame1 
      Height          =   3285
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5794
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit 
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
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2640
         Width           =   1245
      End
      Begin Xivess.uctlTextLookup uctlDetail 
         Height          =   435
         Left            =   2400
         TabIndex        =   1
         Top             =   1200
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtNo 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtAmount 
         Height          =   435
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtExtraDiscountNo 
         Height          =   435
         Left            =   2400
         TabIndex        =   2
         Top             =   1680
         Width           =   2115
         _ExtentX        =   2302
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtExtraDiscountDesc 
         Height          =   435
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   767
      End
      Begin VB.Label lblExtraDiscountDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label lblExtraDiscountNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2205
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   2640
         Width           =   2325
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2205
      End
      Begin VB.Label lblNo 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         TabIndex        =   12
         Top             =   690
         Width           =   2175
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   825
      Left            =   0
      TabIndex        =   9
      Top             =   3840
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1455
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3510
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditExtraDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public SupplierID As Long
Public DocumentType As INVENTORY_DOCTYPE
Public ParentForm As Object

Private m_ExtraDiscount As Collection


Private Sub cboUnit_Click()
    m_HasModify = True
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
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

   Call InitNormalLabel(lblNo, "ลำดับ")
   Call InitNormalLabel(lblDetail, "รายการ")
   Call InitNormalLabel(lblExtraDiscountNo, "เลขที่ส่วนลด")
   Call InitNormalLabel(lblExtraDiscountDesc, "รายละเอียด")
   Call InitNormalLabel(lblAmount, "จำนวน")
   Call InitNormalLabel(lblUnit, "หน่วย")

   Call txtNo.SetTextLenType(TEXT_INTEGER, glbSetting.ID_TYPE)
   Call txtAmount.SetTextLenType(TEXT_STRING, glbSetting.ID_TYPE)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
   Call InitCombo(cboUnit)


   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)

      If ShowMode = SHOW_EDIT Then
         Dim ExtraDiscount As CExtraDiscount

         Set ExtraDiscount = TempCollection.Item(ID)
          txtNo.Text = ExtraDiscount.GetFieldValue("ITEM")
          uctlDetail.MyCombo.ListIndex = IDToListIndex(uctlDetail.MyCombo, ExtraDiscount.GetFieldValue("DISCOUNT_TYPE_ID"))
          txtAmount.Text = ExtraDiscount.GetFieldValue("EXTRA_DISCOUNT_VALUE")
          txtExtraDiscountNo.Text = ExtraDiscount.GetFieldValue("EXTRA_DISCOUNT_NO")
          txtExtraDiscountDesc.Text = ExtraDiscount.GetFieldValue("EXTRA_DISCOUNT_DESC")
          cboUnit.ListIndex = IDToListIndex(cboUnit, ExtraDiscount.GetFieldValue("UNIT_TYPE"))
      End If
   End If

   Call EnableForm(Me, True)
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
   If Not VerifyTextControl(lblNo, txtNo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblDetail, uctlDetail.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If


   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If


   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   Dim ExtraDiscount As CExtraDiscount
   If ShowMode = SHOW_ADD Then
      Set ExtraDiscount = New CExtraDiscount
      ExtraDiscount.Flag = "A"
      Call TempCollection.add(ExtraDiscount)
   Else
      Set ExtraDiscount = TempCollection.Item(ID)
      If ExtraDiscount.Flag <> "A" Then
         ExtraDiscount.Flag = "E"
      End If
   End If

   Call ExtraDiscount.SetFieldValue("EXTRA_DISCOUNT_VALUE", Val(txtAmount.Text))
   Call ExtraDiscount.SetFieldValue("UNIT_TYPE", cboUnit.ListIndex)
   Call ExtraDiscount.SetFieldValue("ITEM", Val(txtNo.Text))
   Call ExtraDiscount.SetFieldValue("DISCOUNT_TYPE_ID", uctlDetail.MyCombo.ItemData(Minus2Zero(uctlDetail.MyCombo.ListIndex)))
   Call ExtraDiscount.SetFieldValue("EXTRA_DISCOUNT_NO", txtExtraDiscountNo.Text)
   Call ExtraDiscount.SetFieldValue("EXTRA_DISCOUNT_DESC", txtExtraDiscountDesc.Text)
   Call ExtraDiscount.SetFieldValue("ITEM", Val(txtNo.Text))
   Call ExtraDiscount.SetFieldValue("KEY_CODE", uctlDetail.MyTextBox.Text)
   Call ExtraDiscount.SetFieldValue("KEY_NAME", uctlDetail.MyCombo.Text)
       
   Set ExtraDiscount = Nothing
   SaveData = True
End Function


Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call LoadMaster(uctlDetail.MyCombo, m_ExtraDiscount, , , MASTER_DISCOUNT)
      Set uctlDetail.MyCollection = m_ExtraDiscount
      
      Call InitUnitOrderBy(cboUnit)

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdNext_Click
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
   Set m_ExtraDiscount = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_ExtraDiscount = Nothing
End Sub


Private Sub txtAmount_Change()
 m_HasModify = True
End Sub

Private Sub txtExtraDiscountDesc_Change()
 m_HasModify = True
End Sub

Private Sub txtExtraDiscountNo_Change()
 m_HasModify = True
End Sub

Private Sub txtNo_Change()
 m_HasModify = True
End Sub

Private Sub uctlDetail_Change()
 m_HasModify = True
End Sub
