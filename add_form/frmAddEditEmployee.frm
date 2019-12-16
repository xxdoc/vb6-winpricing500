VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditEmployee 
   BackColor       =   &H80000000&
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13740
   Icon            =   "frmAddEditEmployee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13740
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   17013
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDealerType 
         Height          =   315
         Left            =   6570
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1980
         Width           =   2955
      End
      Begin Xivess.uctlTextBox txtJointCode 
         Height          =   375
         Left            =   1860
         TabIndex        =   4
         Top             =   2410
         Width           =   2995
         _ExtentX        =   5292
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1980
         Width           =   2955
      End
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
         Width           =   2955
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtLastName 
         Height          =   435
         Left            =   6540
         TabIndex        =   2
         Top             =   1440
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4395
         Left            =   120
         TabIndex        =   18
         Top             =   3405
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   7752
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
         Column(1)       =   "frmAddEditEmployee.frx":27A2
         Column(2)       =   "frmAddEditEmployee.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditEmployee.frx":290E
         FormatStyle(2)  =   "frmAddEditEmployee.frx":2A6A
         FormatStyle(3)  =   "frmAddEditEmployee.frx":2B1A
         FormatStyle(4)  =   "frmAddEditEmployee.frx":2BCE
         FormatStyle(5)  =   "frmAddEditEmployee.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditEmployee.frx":2D5E
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   13515
         _ExtentX        =   23839
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
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   21
         Top             =   7950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEmployee.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   20
         Top             =   7950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEmployee.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   19
         Top             =   7950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblDealerType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   17
         Top             =   1980
         Width           =   1575
      End
      Begin Threed.SSCheck chkNotShowReturn 
         Height          =   405
         Left            =   7560
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJointCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   2520
         Width           =   1575
      End
      Begin Threed.SSCheck chkMainSale 
         Height          =   405
         Left            =   5280
         TabIndex        =   5
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4890
         TabIndex        =   13
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   11760
         TabIndex        =   7
         Top             =   7920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10065
         TabIndex        =   6
         Top             =   7920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEmployee.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee
Private m_Employees As Collection
Private m_MasterRef As CMasterRef

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboDealerType_Click()
   m_HasModify = True
End Sub
Private Sub chkMainSale_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkNotShowReturn_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditEmployeeDealer.ParentForm = Me
      Set frmAddEditEmployeeDealer.TempCollection = m_Employee.CollEmpDealer
      frmAddEditEmployeeDealer.ShowMode = SHOW_ADD
      frmAddEditEmployeeDealer.HeaderText = MapText("เพิ่มประเภทตัวแทน")
      Load frmAddEditEmployeeDealer
      frmAddEditEmployeeDealer.Show 1

      OKClick = frmAddEditEmployeeDealer.OKClick

      Unload frmAddEditEmployeeDealer
      Set frmAddEditEmployeeDealer = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.CollEmpDealer)
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
         m_Employee.CollEmpDealer.Remove (ID2)
      Else
         m_Employee.CollEmpDealer.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.CollEmpDealer)
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
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID1 = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditEmployeeDealer.ID = ID1
      Set frmAddEditEmployeeDealer.ParentForm = Me
      Set frmAddEditEmployeeDealer.TempCollection = m_Employee.CollEmpDealer
      frmAddEditEmployeeDealer.HeaderText = MapText("แก้ไขประเภทตัวแทน")
      frmAddEditEmployeeDealer.ShowMode = SHOW_EDIT
      Load frmAddEditEmployeeDealer
      frmAddEditEmployeeDealer.Show 1

      OKClick = frmAddEditEmployeeDealer.OKClick

      Unload frmAddEditEmployeeDealer
      Set frmAddEditEmployeeDealer = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.CollEmpDealer)
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
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = ID
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Employee.PopulateFromRS(1, m_Rs)
      
      txtCode.Text = m_Employee.EMP_CODE
      txtName.Text = m_Employee.EMP_NAME
      txtLastName.Text = m_Employee.EMP_LNAME
      cboPosition.ListIndex = IDToListIndex(cboPosition, Val(m_Employee.CURRENT_POSITION))
      chkMainSale.Value = FlagToCheck(m_Employee.MAINSALE_FLAG)
      chkNotShowReturn.Value = FlagToCheck(m_Employee.NOT_SHOW_RETURN)
      txtJointCode.Text = m_Employee.JOINT_CODE
      cboDealerType.ListIndex = IDToListIndex(cboDealerType, Val(m_Employee.DEALER_TYPE))
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
   
   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If

   If Not CheckUniqueNs(EMPCODE_UNIQUE, txtCode.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.ShowMode = ShowMode
   m_Employee.EMP_ID = ID
   m_Employee.EMP_CODE = txtCode.Text
   m_Employee.EMP_NAME = txtName.Text
   m_Employee.EMP_LNAME = txtLastName.Text
   m_Employee.CURRENT_POSITION = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
   m_Employee.MAINSALE_FLAG = Check2Flag(chkMainSale.Value)
   m_Employee.NOT_SHOW_RETURN = Check2Flag(chkNotShowReturn.Value)
   m_Employee.JOINT_CODE = txtJointCode.Text
   m_Employee.DEALER_TYPE = cboDealerType.ItemData(Minus2Zero(cboDealerType.ListIndex))
   
   m_Employee.EmpName.ShowMode = ShowMode
   m_Employee.EName.ShowMode = ShowMode
   Call m_Employee.EName.SetFieldValue("LONG_NAME", txtName.Text)
   Call m_Employee.EName.SetFieldValue("LAST_NAME", txtLastName.Text)
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      'Call EnableForm(Me, False)
      
      Call LoadMaster(cboPosition, , , , MASTER_POSITION)
      
      Call LoadDealerType(cboDealerType)
      
      If ShowMode = SHOW_EDIT Then
         m_Employee.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Employee.QueryFlag = 0
         Call QueryData(False)
      End If
      
      'Call EnableForm(Me, True)
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
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblCode, MapText("รหัสพนักงาน"))
   Call InitNormalLabel(lblName, MapText("ชื่อ"))
   Call InitNormalLabel(lblLastName, MapText("นามสกุล"))
   Call InitNormalLabel(lblPosition, MapText("ตำแหน่ง"))
   Call InitCheckBox(chkMainSale, "ใช้เป็นเซลล์หลัก")
   Call InitCheckBox(chkNotShowReturn, "ออกรายงานไม่รวมเอกสาร CN")
   Call InitNormalLabel(lblJointCode, MapText("รหัสอีกบริษัท"))
   Call InitNormalLabel(lblDealerType, MapText("ประเภทตัวแทน"))
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtLastName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtJointCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitCombo(cboPosition)
   Call InitCombo(cboDealerType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ประเภทตัวแทน")
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
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   Set m_MasterRef = New CMasterRef
   Set m_Employees = New Collection
   
   m_HasActivate = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_MasterRef = Nothing
   Set m_Employees = Nothing
End Sub
Private Sub txtJointCode_Change()
   m_HasModify = True
End Sub
Private Sub txtLastName_Change()
   m_HasModify = True
End Sub
Private Sub txtCode_Change()
   m_HasModify = True
End Sub
Private Sub txtName_Change()
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
   Col.Width = 2000
   Col.Caption = MapText("เดือน/ปี")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 5000
   Col.Caption = MapText("ประเภทตัวแทน")
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
      If m_Employee.CollEmpDealer Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim EmpDl As CEmployeeDealer
      If m_Employee.CollEmpDealer.Count <= 0 Then
         Exit Sub
      End If
      Set EmpDl = GetItem(m_Employee.CollEmpDealer, RowIndex, RealIndex)
      If EmpDl Is Nothing Then
         Exit Sub
      End If
      Values(1) = EmpDl.EMPLOYEE_DEALER_ID
      Values(2) = RealIndex
      Values(3) = Right(EmpDl.YYYYMM, 2) & "/" & (Val(Left(EmpDl.YYYYMM, 4)) + 543)
      Values(4) = DealerTypeToString(EmpDl.DEALER_TYPE)
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

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Employee.CollEmpDealer)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub
Public Sub RefreshGrid()
   m_HasModify = True
   GridEX1.ItemCount = CountItem(m_Employee.CollEmpDealer)
   GridEX1.Rebind
End Sub
