VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmApArMas 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmApArMas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCustomerGrade 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1860
         Width           =   2955
      End
      Begin VB.ComboBox cboCustomerType 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1860
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   2955
      End
      Begin Xivess.uctlTextBox txtCustomerName 
         Height          =   435
         Left            =   1560
         TabIndex        =   1
         Top             =   1410
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4725
         Left            =   180
         TabIndex        =   8
         Top             =   3000
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8334
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
         Column(1)       =   "frmApArMas.frx":27A2
         Column(2)       =   "frmApArMas.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmApArMas.frx":290E
         FormatStyle(2)  =   "frmApArMas.frx":2A6A
         FormatStyle(3)  =   "frmApArMas.frx":2B1A
         FormatStyle(4)  =   "frmApArMas.frx":2BCE
         FormatStyle(5)  =   "frmApArMas.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmApArMas.frx":2D5E
      End
      Begin Xivess.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdCopyBranch 
         Height          =   525
         Left            =   5040
         TabIndex        =   23
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmApArMas.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkBranch 
         Height          =   495
         Left            =   9240
         TabIndex        =   22
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblCustomerGrade 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblCustomerType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   19
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   18
         Top             =   2340
         Width           =   1365
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2340
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmApArMas.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   1650
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
         MouseIcon       =   "frmApArMas.frx":356A
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
         MouseIcon       =   "frmApArMas.frx":3884
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmApArMas.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmApArMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Customer As CAPARMas
Private m_TempCustomer As CAPARMas
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_BranchColl As Collection
Public OKClick As Boolean

Public HeaderText As String
Public ApArInd As Long
Private ApArText As String

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "ADD", "����") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   frmAddEditApArMas.ApArInd = ApArInd
   frmAddEditApArMas.HeaderText = MapText("����" & ApArText)
   frmAddEditApArMas.ShowMode = SHOW_ADD
   Load frmAddEditApArMas
   frmAddEditApArMas.Show 1
   
   OKClick = frmAddEditApArMas.OKClick
   
   Unload frmAddEditApArMas
   Set frmAddEditApArMas = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtCustomerName.Text = ""
   txtCustomerCode.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboCustomerGrade.ListIndex = -1
   cboCustomerType.ListIndex = -1
End Sub

Private Sub cmdCopyBranch_Click()
       If ApArInd = 1 Then '�١���
         If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "COPY", "�Ѵ�͡������") Then
              Call EnableForm(Me, True)
              Exit Sub
         End If
      End If
      frmCopyApArBranch.HeaderText = "Copy �������Ң��١���"
      
      Load frmCopyApArBranch
      frmCopyApArBranch.Show 1
      
      Unload frmCopyApArBranch
      Set frmCopyApArBranch = Nothing
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
    If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "DELETE", "ź") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   m_Customer.APAR_MAS_ID = ID
   If Not glbDaily.DeleteCustomer(m_Customer, IsOK, True, glbErrorLog) Then
      m_Customer.APAR_MAS_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call EnableForm(Me, True)
   
   Dim CUS As CAPARMas
   Set CUS = New CAPARMas
   CUS.APAR_IND = 1
   Call LoadApArMas(CUS, Nothing, m_CustomerColl)
   Set CUS = Nothing
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
    If ApArInd = 1 Then '�١���
      If Not VerifyAccessRight("MAIN_CUSTOMER" & "_" & "EDIT", "���") Then
            Call EnableForm(Me, True)
            Exit Sub
      End If
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditApArMas.ApArInd = ApArInd
   frmAddEditApArMas.ID = ID
   frmAddEditApArMas.HeaderText = MapText("���" & ApArText)
   frmAddEditApArMas.ShowMode = SHOW_EDIT
   Load frmAddEditApArMas
   frmAddEditApArMas.Show 1
   
   OKClick = frmAddEditApArMas.OKClick
   
   Unload frmAddEditApArMas
   Set frmAddEditApArMas = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      If ApArInd = 1 Then
         Call LoadMaster(cboCustomerType, , , , MASTER_CUSTYPE)
               
         Call LoadMaster(cboCustomerGrade, , , , MASTER_CUSGRADE)
      
         Call InitCustomerOrderBy(cboOrderBy)
      ElseIf ApArInd = 2 Then
         Call LoadMaster(cboCustomerType, , , , MASTER_SUPTYPE)
         
         Call LoadMaster(cboCustomerGrade, , , , MASTER_SUPGRADE)
      
         Call InitSupplierOrderBy(cboOrderBy)
      End If
      
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
      
      m_Customer.APAR_MAS_ID = -1
      m_Customer.APAR_IND = ApArInd
      m_Customer.APAR_CODE = PatchWildCard(txtCustomerCode.Text)
      m_Customer.APAR_NAME = PatchWildCard(txtCustomerName.Text)
      m_Customer.APAR_GRADE = cboCustomerGrade.ItemData(Minus2Zero(cboCustomerGrade.ListIndex))
      m_Customer.APAR_TYPE = cboCustomerType.ItemData(Minus2Zero(cboCustomerType.ListIndex))
      m_Customer.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Customer.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If Not glbDaily.QueryCustomer(m_Customer, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If chkBranch.Value = ssCBChecked Then
      Call GenerateBranchID
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
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
   Set fmsTemp = GridEX1.FormatStyles.add("Y")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 1620
   Col.Caption = MapText("����" & ApArText)
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 5000
   Col.Caption = MapText("����" & ApArText)
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2500
   Col.Caption = MapText("�������")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 4000
   Col.Caption = MapText("�����͡���")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 3000
   Col.Caption = MapText("������" & ApArText)
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 5500
   Col.Caption = MapText("�Ң�")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Caption = MapText("FLAG EDIT")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1200
   Col.Caption = MapText("������١���")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   If ApArInd = 1 Then
      ApArText = MapText("�١���")
      Call txtCustomerCode.SetKeySearch("CUSTOMER_CODE")
   ElseIf ApArInd = 2 Then
      ApArText = MapText("�����")
   End If
   
   Call InitGrid
   
   Call InitNormalLabel(lblCustomerName, MapText("����" & ApArText))
   Call InitNormalLabel(lblCustomerGrade, MapText("�дѺ" & ApArText))
   Call InitNormalLabel(lblCustomerType, MapText("������" & ApArText))
   Call InitNormalLabel(lblCustomerCode, MapText("����" & ApArText))
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))
   
   Call txtCustomerName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboCustomerGrade)
   Call InitCombo(cboCustomerType)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitCheckBox(chkBranch, "�ʴ��Ң�")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCopyBranch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
   Call InitMainButton(cmdCopyBranch, MapText("Copy �Ң�"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_Customer = New CAPARMas
   Set m_TempCustomer = New CAPARMas
   Set m_Rs = New ADODB.Recordset

   Set m_BranchColl = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_BranchColl = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(8)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim Mr As CMasterRef
   
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
   Call m_TempCustomer.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempCustomer.APAR_MAS_ID
   Values(2) = m_TempCustomer.APAR_CODE
   Values(3) = m_TempCustomer.APAR_NAME
   Values(4) = m_TempCustomer.APAR_SHORT_NAME
   Values(5) = m_TempCustomer.BILL_NAME
   Values(6) = m_TempCustomer.APAR_TYPE_NAME

   If chkBranch.Value = ssCBChecked Then
      For Each Mr In m_BranchColl
         If Mr.PARENT_EX_ID2 = m_TempCustomer.APAR_MAS_ID Then
            Values(7) = Mr.KEY_CODE & " " & Mr.KEY_NAME
            Exit For
         End If
      Next Mr
   End If
   Values(8) = m_TempCustomer.CANCEL_OUT_DOCUMENT
   
   Values(9) = m_TempCustomer.APAR_MAS_GROUP_CODE
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
   cmdCopyBranch.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub GenerateBranchID()
Dim m_Rs1 As ADODB.Recordset
      
   Set m_BranchColl = Nothing
   Set m_BranchColl = New Collection
   
   Call LoadMaster(Nothing, m_BranchColl, , , MASTER_APARMAS_BRANCH)
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

