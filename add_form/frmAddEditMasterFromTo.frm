VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditMasterFromTo 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditMasterFromTo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   10
         Top             =   2325
         Width           =   11595
         _ExtentX        =   20452
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
      Begin Xivess.uctlTextBox txtMasterFromToNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   2385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2685
         Left            =   150
         TabIndex        =   11
         Top             =   2880
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4736
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
         Column(1)       =   "frmAddEditMasterFromTo.frx":27A2
         Column(2)       =   "frmAddEditMasterFromTo.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditMasterFromTo.frx":290E
         FormatStyle(2)  =   "frmAddEditMasterFromTo.frx":2A6A
         FormatStyle(3)  =   "frmAddEditMasterFromTo.frx":2B1A
         FormatStyle(4)  =   "frmAddEditMasterFromTo.frx":2BCE
         FormatStyle(5)  =   "frmAddEditMasterFromTo.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditMasterFromTo.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtValue1 
         Height          =   435
         Left            =   6960
         TabIndex        =   7
         Top             =   1800
         Width           =   765
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtMasterFromToDesc 
         Height          =   435
         Left            =   5760
         TabIndex        =   1
         Top             =   840
         Width           =   6020
         _ExtentX        =   10610
         _ExtentY        =   767
      End
      Begin Xivess.uctlDate uctlToDate 
         Height          =   405
         Left            =   7920
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtIncludeSub 
         Height          =   435
         Left            =   3720
         TabIndex        =   5
         Top             =   1800
         Width           =   765
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Threed.SSOption SSOption3 
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   1860
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption2 
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   1860
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   1860
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCheck ChkIncludeSub 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1860
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   22
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblMasterFromToDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   20
         Top             =   1380
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterFromTo.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   16
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterFromTo.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterFromTo.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblMasterFromToNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditMasterFromTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_MasterFromTo As CMasterFromTo

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As MASTER_COMMISSION_AREA
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_ID", ID)
      If Not glbDaily.QueryMasterFromTo(m_MasterFromTo, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_MasterFromTo.PopulateFromRS(1, m_Rs)
      
      txtMasterFromToNo.Text = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_NO")
      txtMasterFromToDesc.Text = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_DESC")
      uctlFromDate.ShowDate = m_MasterFromTo.GetFieldValue("VALID_FROM")
      uctlToDate.ShowDate = m_MasterFromTo.GetFieldValue("VALID_TO")
      
      ChkIncludeSub.Value = FlagToCheck(m_MasterFromTo.GetFieldValue("INCLUDE_SUB_FLAG"))
      txtIncludeSub.Text = m_MasterFromTo.GetFieldValue("INCLUDE_SUB_PERCENT")
      SSOption1.Value = StringToCheckSSoption(m_MasterFromTo.GetFieldValue("MULTIPLE_FLAG"))
      txtValue1.Text = m_MasterFromTo.GetFieldValue("MULTIPLE_PERCENT")
      SSOption2.Value = StringToCheckSSoption(m_MasterFromTo.GetFieldValue("STEP_FLAG"))
      SSOption3.Value = StringToCheckSSoption(m_MasterFromTo.GetFieldValue("TIER_FLAG"))
   
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblMasterFromToNo, txtMasterFromToNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If

   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If Not CheckUniqueNs(MASTER_FT_UNIQUE, txtMasterFromToNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtMasterFromToNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Call txtMasterFromToNo.SetFocus
      Exit Function
   End If
   
   m_MasterFromTo.ShowMode = ShowMode
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_ID", ID)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_NO", txtMasterFromToNo.Text)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_DESC", txtMasterFromToDesc.Text)
   Call m_MasterFromTo.SetFieldValue("VALID_FROM", uctlFromDate.ShowDate)
   Call m_MasterFromTo.SetFieldValue("VALID_TO", uctlToDate.ShowDate)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_TYPE", DocumentType)
   
   Call m_MasterFromTo.SetFieldValue("INCLUDE_SUB_FLAG", Check2Flag(ChkIncludeSub.Value))
   Call m_MasterFromTo.SetFieldValue("INCLUDE_SUB_PERCENT", Val(txtIncludeSub.Text))
   Call m_MasterFromTo.SetFieldValue("MULTIPLE_FLAG", CheckSSoptionToString(SSOption1.Value))
   Call m_MasterFromTo.SetFieldValue("MULTIPLE_PERCENT", Val(txtValue1.Text))
   Call m_MasterFromTo.SetFieldValue("STEP_FLAG", CheckSSoptionToString(SSOption2.Value))
   Call m_MasterFromTo.SetFieldValue("TIER_FLAG", CheckSSoptionToString(SSOption3.Value))
      
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterFromTo(m_MasterFromTo, IsOK, True, glbErrorLog) Then
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

Private Sub ChkIncludeSub_Click(Value As Integer)
   If ChkIncludeSub.Value = ssCBChecked Then
      txtIncludeSub.Enabled = True
   Else
      txtIncludeSub.Enabled = False
   End If
End Sub

Private Sub ChkIncludeSub_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditMasterFTItem.ParentForm = Me
      Set frmAddEditMasterFTItem.TempCollection = m_MasterFromTo.Details
      frmAddEditMasterFTItem.StepFlag = SSOption2.Value
      frmAddEditMasterFTItem.ShowMode = SHOW_ADD
      frmAddEditMasterFTItem.DocumentType = DocumentType
      frmAddEditMasterFTItem.HeaderText = MapText("����" & Comissiontype2Text(DocumentType))
      Load frmAddEditMasterFTItem
      frmAddEditMasterFTItem.Show 1

      OKClick = frmAddEditMasterFTItem.OKClick

      Unload frmAddEditMasterFTItem
      Set frmAddEditMasterFTItem = Nothing
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      Call RefreshGrid
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

'   If Not cmdDelete.Enabled Then
'      Exit Sub
'   End If
   
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
         m_MasterFromTo.Details.Remove (ID2)
      Else
         m_MasterFromTo.Details.Item(ID2).Flag = "D"
      End If
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   Call RefreshGrid
   
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditMasterFTItem.ParentForm = Me
      frmAddEditMasterFTItem.ID = ID
      frmAddEditMasterFTItem.StepFlag = SSOption2.Value
      Set frmAddEditMasterFTItem.TempCollection = m_MasterFromTo.Details
      frmAddEditMasterFTItem.DocumentType = DocumentType
      frmAddEditMasterFTItem.HeaderText = MapText("���" & Comissiontype2Text(DocumentType))
      frmAddEditMasterFTItem.ShowMode = SHOW_EDIT
      Load frmAddEditMasterFTItem
      frmAddEditMasterFTItem.Show 1
      
      OKClick = frmAddEditMasterFTItem.OKClick

      Unload frmAddEditMasterFTItem
      Set frmAddEditMasterFTItem = Nothing
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   If OKClick Then
      Call RefreshGrid
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As CPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("�ѹ�֡", "-", "�ѹ�֡����͡�ҡ˹�Ҩ�")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID")
      m_MasterFromTo.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_MasterFromTo.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_MasterFromTo.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_MasterFromTo = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim I As Byte
   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   I = 6
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = (ScaleWidth - 600) / 7
   Col.Caption = MapText("���������Ե���")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = (ScaleWidth - 600) / 7
   Col.Caption = MapText("������")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = (ScaleWidth - 600) / 7
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ҡ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = (ScaleWidth - 600) / 7
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�֧")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = (ScaleWidth - 600) / 7
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("%��Ҥ��")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = (ScaleWidth - 600) / 7
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ʹ��Ҥ��")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = (ScaleWidth - 600) / 7
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("����")
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMasterFromToNo, MapText("�����Ţ"))
   Call InitNormalLabel(lblMasterFromToDesc, MapText("��������´"))
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToDate, MapText("�֧�ѹ���"))
   
   Call txtMasterFromToNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtIncludeSub.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeSub.Enabled = False
   Call txtValue1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtValue1.Enabled = False
   
   If DocumentType = RETURN_TABLE Then
      SSOption3.Enabled = False
   End If
   SSOption1.Value = True
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   
   Call InitCheckBox(ChkIncludeSub, "����ӴѺ���� %")
   Call InitOptionEx(SSOption1, "��Ҥ���� %")
   Call InitOptionEx(SSOption2, "�ʵ�")
   Call InitOptionEx(SSOption3, "�����")
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("��������´")
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
   Set m_MasterFromTo = New CMasterFromTo
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
      If m_MasterFromTo.Details Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CMasterFromToDetail
      If m_MasterFromTo.Details.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_MasterFromTo.Details, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.GetFieldValue("MASTER_FROMTO_DETAIL_ID")
      Values(2) = RealIndex
      Values(3) = CR.GetFieldValue("GROUP_COM_DESC")
      If CR.GetFieldValue("AMOUNT_FLAG") = "Y" Then
         Values(4) = "�ӹǹ"
      ElseIf CR.GetFieldValue("VALUE_FLAG") = "Y" Then
         Values(4) = "�ʹ���"
      ElseIf CR.GetFieldValue("TAGET_VALUE_FLAG") = "Y" Then
         Values(4) = "%�ʹ��ҡ�â��"
      ElseIf CR.GetFieldValue("AMOUNT_P_FLAG") = "Y" Then
         Values(4) = "%��ǧ%�Ѻ�׹/�ӹǹ"
      ElseIf CR.GetFieldValue("VALUE_P_FLAG") = "Y" Then
         Values(4) = "%��ǧ%�Ѻ�׹/�ʹ���"
      End If
      Values(5) = FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_FROM"))
      Values(6) = FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_TO"))
      
      If CR.GetFieldValue("NO_X_FLAG") = "Y" Then
         Values(7) = ""
      ElseIf CR.GetFieldValue("AMOUNT_X_FLAG") = "Y" Then
         Values(7) = "�ӹǹ �ٳ " & FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"))
      ElseIf CR.GetFieldValue("VALUE_X_FLAG") = "Y" Then
         Values(7) = "�ʹ��� �ٳ " & FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"))
      End If
      Values(8) = FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE2"))
      Values(9) = FormatNumber(CR.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE3"))
      
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

Private Sub SSOption1_Click(Value As Integer)
   Call SetEnableOption
End Sub
Private Sub SSOption2_Click(Value As Integer)
   Call SetEnableOption
End Sub

Private Sub SSOption3_Click(Value As Integer)
   Call SetEnableOption
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
  If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_MasterFromTo.Details)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtMasterFromToDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtMasterFromToNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMasterFromToNo_LostFocus()
   If Not CheckUniqueNs(MASTER_FT_UNIQUE, txtMasterFromToNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtMasterFromToNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Call txtMasterFromToNo.SetFocus
      Exit Sub
   End If
End Sub

Private Sub txtValue1_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_MasterFromTo.Details)
   GridEX1.Rebind
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
Private Sub SetEnableOption()
   If SSOption1.Value Then
      txtValue1.Enabled = True
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
   Else
      txtValue1.Enabled = False
      txtValue1.Text = ""
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
   End If
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
