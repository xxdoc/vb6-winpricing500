VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEmployee 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmEmployee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1980
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   2955
      End
      Begin Xivess.uctlTextBox txtEmpName 
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   1530
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   8
         Top             =   2640
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8916
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
         Column(1)       =   "frmEmployee.frx":27A2
         Column(2)       =   "frmEmployee.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmEmployee.frx":290E
         FormatStyle(2)  =   "frmEmployee.frx":2A6A
         FormatStyle(3)  =   "frmEmployee.frx":2B1A
         FormatStyle(4)  =   "frmEmployee.frx":2BCE
         FormatStyle(5)  =   "frmEmployee.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmEmployee.frx":2D5E
      End
      Begin Xivess.uctlTextBox txtEmpCode 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   1080
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtLastName 
         Height          =   435
         Left            =   6150
         TabIndex        =   3
         Top             =   1530
         Width           =   2955
         _ExtentX        =   5265
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdOption 
         Height          =   525
         Left            =   6840
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEmployee.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4620
         TabIndex        =   22
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label lblEmpCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4620
         TabIndex        =   20
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   19
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label lblEmpName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1170
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEmployee.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   1740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEmployee.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEmployee.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   14
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmEmployee.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Employee As CEmployee
Private m_TempEmployee As CEmployee
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_MasterRef As CMasterRef

Private CurrentIndex As Long
Public OKClick As Boolean
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditEmployee.HeaderText = MapText("������ѡ�ҹ")
   frmAddEditEmployee.ShowMode = SHOW_ADD
   Load frmAddEditEmployee
   frmAddEditEmployee.Show 1
   
   OKClick = frmAddEditEmployee.OKClick
   
   Unload frmAddEditEmployee
   Set frmAddEditEmployee = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtEmpName.Text = ""
   txtEmpCode.Text = ""
   txtLastName.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboPosition.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   CurrentIndex = GridEX1.Row
    
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_Employee.EMP_ID = ID
   If Not glbDaily.DeleteEmployee(m_Employee, IsOK, True, glbErrorLog) Then
      m_Employee.EMP_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   CurrentIndex = GridEX1.Row - 1
   Call QueryData(True)
   Call EnableForm(Me, True)
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

   ID = Val(GridEX1.Value(1))
   CurrentIndex = GridEX1.Row
   
   frmAddEditEmployee.ID = ID
   frmAddEditEmployee.HeaderText = MapText("��䢾�ѡ�ҹ")
   frmAddEditEmployee.ShowMode = SHOW_EDIT
   Load frmAddEditEmployee
   frmAddEditEmployee.Show 1
   
   OKClick = frmAddEditEmployee.OKClick
   
   Unload frmAddEditEmployee
   Set frmAddEditEmployee = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdOption_Click()
   frmUpdateEmployeeBranch.HeaderText = MapText("�Ѿഴ/�͹ �ҢҨҡ��ѡ�ҹ��� ��ա��ѡ�ҹ���˹��")
   Load frmUpdateEmployeeBranch
   frmUpdateEmployeeBranch.Show 1
   
   OKClick = frmUpdateEmployeeBranch.OKClick
   
   Unload frmUpdateEmployeeBranch
   Set frmUpdateEmployeeBranch = Nothing
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
            
      Call LoadMaster(cboPosition, , , , MASTER_POSITION)
      
      Call InitEmployeeOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = -1
      m_Employee.EMP_CODE = PatchWildCard(txtEmpCode.Text)
      m_Employee.EMP_NAME = PatchWildCard(txtEmpName.Text)
      m_Employee.EMP_LNAME = txtLastName.Text
      m_Employee.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Employee.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      m_Employee.CURRENT_POSITION = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
    
   Call EnableForm(Me, True)
   
   If CurrentIndex > 0 Then
      Call GridEX1.SetFocus
      Call GridEX1.MoveToRowIndex(CurrentIndex)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 1620
   Col.Caption = MapText("���ʾ�ѡ�ҹ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3780
   Col.Caption = MapText("����")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3555
   Col.Caption = MapText("���ʡ��")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2520
   Col.Caption = MapText("���˹�")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("��ѡ�ҹ��ѡ")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1620
   Col.Caption = MapText("���������ѹ")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2000
   Col.Caption = MapText("���������᷹")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("�����ž�ѡ�ҹ")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblEmpName, MapText("����"))
   Call InitNormalLabel(lblLastName, MapText("���ʡ��"))
   Call InitNormalLabel(lblPosition, MapText("���˹�"))
   Call InitNormalLabel(lblEmpCode, MapText("���ʾ�ѡ�ҹ"))
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))
   
   Call txtEmpName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboPosition)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOption.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
   Call InitMainButton(cmdOption, MapText("OPTION"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_Employee = New CEmployee
   Set m_TempEmployee = New CEmployee
   Set m_Rs = New ADODB.Recordset

   Set m_MasterRef = New CMasterRef
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_MasterRef = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   '''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempEmployee.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempEmployee.EMP_ID
   Values(2) = m_TempEmployee.EMP_CODE
   Values(3) = m_TempEmployee.EMP_NAME
   Values(4) = m_TempEmployee.EMP_LNAME
   Values(5) = m_TempEmployee.POSITION_NAME
   Values(6) = m_TempEmployee.MAINSALE_FLAG
   Values(7) = m_TempEmployee.JOINT_CODE
   Values(8) = DealerTypeToString(m_TempEmployee.DEALER_TYPE)
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdOption.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdOption.Left = cmdOK.Left - cmdOption.Width - 50
End Sub

