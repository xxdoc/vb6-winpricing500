VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditEnterprise 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditEnterprise.frx":0000
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
      TabIndex        =   12
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   2100
         TabIndex        =   1
         Top             =   1440
         Width           =   7335
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtShortName 
         Height          =   435
         Left            =   2100
         TabIndex        =   0
         Top             =   990
         Width           =   3585
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBranchCode 
         Height          =   435
         Left            =   2100
         TabIndex        =   2
         Top             =   1890
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtBranchName 
         Height          =   435
         Left            =   2100
         TabIndex        =   3
         Top             =   2340
         Width           =   7335
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtTaxID 
         Height          =   435
         Left            =   2100
         TabIndex        =   4
         Top             =   2790
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   3510
         Width           =   11685
         _ExtentX        =   20611
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   3795
         Left            =   120
         TabIndex        =   6
         Top             =   4050
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   6694
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
         Column(1)       =   "frmAddEditEnterprise.frx":27A2
         Column(2)       =   "frmAddEditEnterprise.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditEnterprise.frx":290E
         FormatStyle(2)  =   "frmAddEditEnterprise.frx":2A6A
         FormatStyle(3)  =   "frmAddEditEnterprise.frx":2B1A
         FormatStyle(4)  =   "frmAddEditEnterprise.frx":2BCE
         FormatStyle(5)  =   "frmAddEditEnterprise.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditEnterprise.frx":2D5E
      End
      Begin VB.Label lblTaxID 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   2910
         Width           =   1995
      End
      Begin VB.Label lblBranchName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label lblBranchCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   2010
         Width           =   1995
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1500
         Width           =   1995
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   1080
         Width           =   1995
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEnterprise.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   14
         Top             =   1650
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   9
         Top             =   7920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEnterprise.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   7
         Top             =   7920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEnterprise.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   8
         Top             =   7920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10215
         TabIndex        =   11
         Top             =   7920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8565
         TabIndex        =   10
         Top             =   7920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEnterprise.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditEnterprise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Enterprise As CEnterprise

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_Enterprise.SetFieldValue("ENTERPRISE_ID", ID)
      m_Enterprise.QueryFlag = 1
      If Not glbDaily.QueryEnterprise(m_Enterprise, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      ShowMode = SHOW_EDIT
      Call m_Enterprise.PopulateFromRS(1, m_Rs)
      Call m_Enterprise.SetFieldValue("ENTERPRISE_ID", m_Enterprise.GetFieldValue("ENTERPRISE_ID"))
                       
      Dim Name As cName
      Dim EnpName As CEnterpriseName

      If (Not m_Enterprise.EnpNames Is Nothing) And (m_Enterprise.EnpNames.Count > 0) Then
         Set EnpName = m_Enterprise.EnpNames(1)
         Set Name = EnpName.Names(1)
         txtTaxID.Text = m_Enterprise.GetFieldValue("TAX_ID")
         txtBranchCode.Text = m_Enterprise.GetFieldValue("BRANCH_CODE")
         txtName.Text = m_Enterprise.GetFieldValue("ENTERPRISE_NAME")
         txtShortName.Text = m_Enterprise.GetFieldValue("SHORT_NAME")
         txtBranchName.Text = m_Enterprise.GetFieldValue("BRANCH_NAME")
      Else
         txtName.Text = ""
         txtShortName.Text = ""
      End If
      
      Call TabStrip1_Click
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Enterprise.ShowMode = ShowMode
   Call m_Enterprise.SetFieldValue("SETUP_DATE", -1)
   Call m_Enterprise.SetFieldValue("TAX_ID", txtTaxID.Text)
   Call m_Enterprise.SetFieldValue("EMAIL", "")
   Call m_Enterprise.SetFieldValue("POLICY", "")
   Call m_Enterprise.SetFieldValue("WEBSITE", "")
   Call m_Enterprise.SetFieldValue("BUSINESS_TYPE", -1)
   Call m_Enterprise.SetFieldValue("ENTERPRISE_TYPE", -1)
   Call m_Enterprise.SetFieldValue("BRANCH_CODE", txtBranchCode.Text)
   Call m_Enterprise.SetFieldValue("BRANCH_NAME", txtBranchName.Text)
   
   Dim EnpName As CEnterpriseName
   If m_Enterprise.EnpNames.Count <= 0 Then
      Set EnpName = New CEnterpriseName
      EnpName.Flag = "A"
      Call m_Enterprise.EnpNames.add(EnpName)
   Else
      Set EnpName = m_Enterprise.EnpNames.Item(1)
      EnpName.Flag = "E"
   End If
   
   Dim Name As cName
   If EnpName.Names.Count <= 0 Then
      Set EnpName.Names = New Collection
      Set Name = New cName
      Call Name.SetFieldValue("LONG_NAME", txtName.Text)
      Call Name.SetFieldValue("SHORT_NAME", txtShortName.Text)
      Name.Flag = "A"
      Call EnpName.Names.add(Name)
      Set Name = Nothing
   Else
      Set Name = EnpName.Names.Item(1)
      Call Name.SetFieldValue("LONG_NAME", txtName.Text)
      Call Name.SetFieldValue("SHORT_NAME", txtShortName.Text)
      Name.Flag = "E"
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEnterprise(m_Enterprise, IsOK, True, glbErrorLog) Then
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

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditEnpAddress.TempCollection = m_Enterprise.EnpAddresses
      frmAddEditEnpAddress.ShowMode = SHOW_ADD
      frmAddEditEnpAddress.HeaderText = MapText("�����������")
      Load frmAddEditEnpAddress
      frmAddEditEnpAddress.Show 1
      
      OKClick = frmAddEditEnpAddress.OKClick
      
      Unload frmAddEditEnpAddress
      Set frmAddEditEnpAddress = Nothing
   
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Enterprise.EnpAddresses)
         GridEX1.Rebind
      End If
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Set frmAddEditEnpPerson.TempCollection = m_Enterprise.EnpPersons
'      frmAddEditEnpPerson.ShowMode = SHOW_ADD
'      frmAddEditEnpPerson.HeaderText = MapText("����������ü��Ѵ���")
'      Load frmAddEditEnpPerson
'      frmAddEditEnpPerson.Show 1
'
'      OKClick = frmAddEditEnpPerson.OKClick
'
'      Unload frmAddEditEnpPerson
'      Set frmAddEditEnpPerson = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Enterprise.EnpPersons)
'         GridEX1.Rebind
'      End If
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
         m_Enterprise.EnpAddresses.Remove (ID2)
      Else
         m_Enterprise.EnpAddresses.Item(ID2).Flag = "D"
      End If
      
      GridEX1.ItemCount = CountItem(m_Enterprise.EnpAddresses)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_Enterprise.EnpPersons.Remove (ID2)
      Else
         m_Enterprise.EnpPersons.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Enterprise.EnpPersons)
      GridEX1.Rebind
      m_HasModify = True
   End If
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

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditEnpAddress.ID = ID
      Set frmAddEditEnpAddress.TempCollection = m_Enterprise.EnpAddresses
      frmAddEditEnpAddress.HeaderText = MapText("��䢷������")
      frmAddEditEnpAddress.ShowMode = SHOW_EDIT
      Load frmAddEditEnpAddress
      frmAddEditEnpAddress.Show 1
         
      OKClick = frmAddEditEnpAddress.OKClick
      
      Unload frmAddEditEnpAddress
      Set frmAddEditEnpAddress = Nothing
      
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Enterprise.EnpAddresses)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      frmAddEditEnpPerson.ID = ID
'      Set frmAddEditEnpPerson.TempCollection = m_Enterprise.EnpPersons
'      frmAddEditEnpPerson.HeaderText = MapText("��䢡�����ü��Ѵ���")
'      frmAddEditEnpPerson.ShowMode = SHOW_EDIT
'      Load frmAddEditEnpPerson
'      frmAddEditEnpPerson.Show 1
'
'      OKClick = frmAddEditEnpPerson.OKClick
'
'      Unload frmAddEditEnpPerson
'      Set frmAddEditEnpPerson = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_Enterprise.EnpPersons)
'         GridEX1.Rebind
'      End If
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)

     Call QueryData(True)
      Call TabStrip1_Click
      
      m_HasModify = False
      Call EnableForm(Me, True)
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
   
   Set m_Enterprise = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid2()
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
   Col.Visible = False
   Col.Caption = MapText("ID")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("Real ID")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2835
   Col.Caption = MapText("����")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2745
   Col.Caption = MapText("���ʡ��")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2535 + 3450
   Col.Caption = MapText("�������")

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
   Col.Visible = False
   Col.Caption = MapText("ID")
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("Real ID")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 11550
   Col.Caption = MapText("�������")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblName, MapText("���ͺ���ѷ"))
   Call InitNormalLabel(lblShortName, MapText("���ʺ���ѷ"))
   Call InitNormalLabel(lblBranchCode, MapText("�����Ң�"))
   Call InitNormalLabel(lblBranchName, MapText("�����Ң�"))
   Call InitNormalLabel(lblTaxID, MapText("�����Ţ�����������"))
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBranchCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBranchName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTaxID.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
'   txtBranchName.Enabled = False
'   txtBranchCode.Enabled = False
   
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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("�������")
'   TabStrip1.Tabs.Add().Caption = MapText("������ü��Ѵ���")
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
   Set m_Enterprise = New CEnterprise
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Enterprise.EnpAddresses Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CEnterpriseAddress
      Dim Addr As CAddress
      If m_Enterprise.EnpAddresses.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Enterprise.EnpAddresses, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses(1)
      
      Values(1) = Addr.GetFieldValue("ADDRESS_ID")
      Values(2) = RealIndex
      Values(3) = Addr.PackAddress
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_Enterprise.EnpPersons Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim Ep As CEnterprisePerson
'      Dim N As cName
'      If m_Enterprise.EnpPersons.count <= 0 Then
'         Exit Sub
'      End If
'      Set Ep = GetItem(m_Enterprise.EnpPersons, RowIndex, RealIndex)
'      If Ep Is Nothing Then
'         Exit Sub
'      End If
'      Set N = Ep.Name
'
'      Values(1) = N.NAME_ID
'      Values(2) = RealIndex
'      Values(3) = N.LONG_NAME
'      Values(4) = N.LAST_NAME
'      Values(5) = N.EMAIL
'      Values(6) = Ep.CONTACT_POSITION
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Enterprise.EnpAddresses)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_Enterprise.EnpPersons)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtBranchCode_Change()
   m_HasModify = True
End Sub

Private Sub txtBranchName_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_Change()
   m_HasModify = True
End Sub

Private Sub txtSlogan_Change()
   m_HasModify = True
End Sub

Private Sub txtTaxID_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
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

