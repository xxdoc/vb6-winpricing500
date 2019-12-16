VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrganize 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmOrganize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1508
      _Version        =   131073
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7905
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   13944
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   360
         Top             =   870
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOrganize.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOrganize.frx":11A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   6375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   11245
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3330
         TabIndex        =   9
         Top             =   7290
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmOrganize.frx":1A7E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   60
         TabIndex        =   8
         Top             =   7290
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmOrganize.frx":1D98
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1680
         TabIndex        =   7
         Top             =   7290
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   10260
         TabIndex        =   6
         Top             =   7290
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSOption radPosition 
         Height          =   405
         Left            =   3270
         TabIndex        =   5
         Top             =   270
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption radDepartment 
         Height          =   405
         Left            =   210
         TabIndex        =   4
         Top             =   270
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
   End
End
Attribute VB_Name = "frmOrganize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmCustomerInfo"
Private Const ROOT_TREE = "R"
Private HasActivate As Boolean
Public HeaderText As String

Private m_Rs As ADODB.Recordset

Private m_Positions As Collection
Private m_Organizes As Collection
Private m_Sections As Collection

Private Function GetHRChart(TempID As Long, TempCol As Collection) As CHRChart
Dim L As CHRChart

   For Each L In TempCol
      If L.HR_CHART_ID = TempID Then
         Set GetHRChart = L
         Exit Function
      End If
   Next L
End Function

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim TableName As String
Dim L As CHRChart

   If Not VerifyAccessRight("EMPLOYEE_ORG_DELETE", "ź������Ἱ���е��˹觾�ѡ�ҹ") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   If trvMain.SelectedItem Is Nothing Then
'      glbErrorLog.LocalErrorMsg = GetTextMessage("TEXT-KEY569")
'      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If trvMain.Nodes.count <= 0 Then
      Exit Sub
   End If
   
   ID = Val(trvMain.SelectedItem.Tag)
   If radDepartment.Value Then
      Set L = GetHRChart(ID, m_Organizes)
   ElseIf radPosition.Value Then
      Set L = GetHRChart(ID, m_Positions)
   End If

   glbErrorLog.LocalErrorMsg = "��ͧ���ź������ " & L.ORGANIZE_NAME & " ��������� ?"
   If glbErrorLog.AskMessage = vbNo Then
      Exit Sub
   End If

   If Not glbDaily.DeleteHRChart(ID, IsOK, glbErrorLog, L.PARENT_ID) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call glbDatabaseMngr.UnLockTable(TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call LoadHRChart(Nothing, m_Organizes, DEPARTMENT_ARA)
   Call LoadHRChart(Nothing, m_Positions, POSITION_ARA)
   If radDepartment.Value Then
      Call InitMainTreeview("", m_Organizes)
   ElseIf radPosition.Value Then
      Call InitMainTreeview("", m_Positions)
   End If
   
   Call glbDatabaseMngr.UnLockTable(TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim itemcount As Long

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadHRChart(Nothing, m_Organizes, DEPARTMENT_ARA)
      Call LoadHRChart(Nothing, m_Positions, POSITION_ARA)
      Call InitMainTreeview("", m_Organizes)
   End If
   
   Call EnableForm(Me, True)
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.Name
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
      'Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
      'Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
      'Call cmdPrint_Click
   End If
End Sub

Private Sub cmdAdd_Click()
Dim itemcount As Long

   If Not VerifyAccessRight("EMPLOYEE_ORG_ADD", "����������Ἱ���е��˹觾�ѡ�ҹ") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   Call EnableForm(Me, False)

   If radDepartment.Value Then
      frmAddEditOrganize.HeaderText = "����������˹��§ҹ"
      frmAddEditOrganize.OrganizeArea = DEPARTMENT_ARA
   ElseIf radPosition.Value Then
      frmAddEditOrganize.HeaderText = "���������ŵ��˹�"
      frmAddEditOrganize.OrganizeArea = POSITION_ARA
   End If
   frmAddEditOrganize.ShowMode = SHOW_ADD
   Load frmAddEditOrganize
   frmAddEditOrganize.Show 1

   If frmAddEditOrganize.OKClick Then
      Call EnableForm(Me, False)
      Call LoadHRChart(Nothing, m_Organizes, DEPARTMENT_ARA)
      Call LoadHRChart(Nothing, m_Positions, POSITION_ARA)
      
      If radDepartment.Value Then
         Call InitMainTreeview("", m_Organizes)
      ElseIf radPosition.Value Then
         Call InitMainTreeview("", m_Positions)
      End If
      
      Call EnableForm(Me, True)
   End If

   Unload frmAddEditOrganize
   Set frmAddEditOrganize = Nothing
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
Dim ID As Long
Dim TableName As String

      If Not VerifyAccessRight("EMPLOYEE_ORG_QUERY", "���Ң�����Ἱ���е��˹觾�ѡ�ҹ") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
   If trvMain.SelectedItem Is Nothing Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If trvMain.Nodes.count <= 0 Then
      Exit Sub
   End If
   
   ID = Val(trvMain.SelectedItem.Tag)
   If Not glbDatabaseMngr.LockTable(TableName, ID, IsCanLock, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   Else
      If Not IsCanLock Then
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
            
   Call EnableForm(Me, False)
   If radDepartment.Value Then
      frmAddEditOrganize.HeaderText = "��䢢�����˹��§ҹ"
      frmAddEditOrganize.OrganizeArea = DEPARTMENT_ARA
   ElseIf radPosition.Value Then
      frmAddEditOrganize.HeaderText = "��䢢����ŵ��˹�"
      frmAddEditOrganize.OrganizeArea = POSITION_ARA
   End If
   frmAddEditOrganize.ShowMode = SHOW_EDIT
   frmAddEditOrganize.OrganizeID = ID
   Load frmAddEditOrganize
   frmAddEditOrganize.Show 1

   If frmAddEditOrganize.OKClick Then
      Call EnableForm(Me, False)
      Call LoadHRChart(Nothing, m_Organizes, DEPARTMENT_ARA)
      Call LoadHRChart(Nothing, m_Positions, POSITION_ARA)
      
      If radDepartment.Value Then
         Call InitMainTreeview("", m_Organizes)
      ElseIf radPosition.Value Then
         Call InitMainTreeview("", m_Positions)
      End If
      Call EnableForm(Me, True)
   End If

   Unload frmAddEditOrganize
   Set frmAddEditOrganize = Nothing

   Call glbDatabaseMngr.UnLockTable(TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Function GetIconNo(O As CHRChart) As Long
   If O.CHILD_COUNT = 0 Then
      GetIconNo = 2
   Else
      GetIconNo = 1
   End If
End Function

Private Sub GenerateTree(Organizes As Collection, N As Node, NodeID As String, PID As Long, Level As Long)
Dim O As CHRChart
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In Organizes
      If O.PARENT_ID = PID Then
         If Level = 0 Then
            Set Node = trvMain.Nodes.Add(, tvwFirst, NodeID & O.HR_CHART_ID, O.ORGANIZE_NO & " : " & O.ORGANIZE_NAME, GetIconNo(O))
            Node.Tag = O.HR_CHART_ID
            Call GenerateTree(Organizes, Node, NodeID & O.HR_CHART_ID, O.HR_CHART_ID, Level + 1)
         Else
            NewNodeID = NodeID & "-" & O.HR_CHART_ID
            Set Node = trvMain.Nodes.Add(N, tvwChild, NewNodeID, O.ORGANIZE_NO & " : " & O.ORGANIZE_NAME, GetIconNo(O))
            Node.Tag = O.HR_CHART_ID
            Call GenerateTree(Organizes, Node, NewNodeID, O.HR_CHART_ID, Level + 1)
         End If
         Node.Expanded = True
      End If
   Next O
End Sub

Private Sub InitMainTreeview(Caption As String, Organizes As Collection)
   If Organizes Is Nothing Then
      Exit Sub
   End If
   
   ClearTreeView (trvMain.hwnd)
   Call GenerateTree(Organizes, Nothing, "ROOT", -1, 0)
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Me.KeyPreview = True
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_FORM_COLOR
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitOptionEx(radDepartment, "˹��§ҹ")
   radDepartment.Value = True
   Call InitOptionEx(radPosition, "���˹�")
   
   Call EnableForm(Me, True)
   
   HasActivate = False
   Set m_Rs = New ADODB.Recordset
   Set m_Positions = New Collection
   Set m_Organizes = New Collection
   Set m_Sections = New Collection
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   HasActivate = False
   Me.Caption = HeaderText
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"
   Call InitFormLayout
   
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Positions = Nothing
   Set m_Organizes = Nothing
   Set m_Sections = Nothing
End Sub

Private Sub radDepartment_Click(Value As Integer)
   Call InitMainTreeview("", m_Organizes)
End Sub

Private Sub radPosition_Click(Value As Integer)
   Call InitMainTreeview("", m_Positions)
End Sub

Private Sub trvMain_DblClick()
'   Call cmdEdit_Click
End Sub
