VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMasterRefItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMasterRefItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   4471
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboMasterCusType 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4395
      End
      Begin VB.ComboBox cboMasterCusGroup 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4395
      End
      Begin VB.Label lblMasterCusGroup 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2325
         TabIndex        =   2
         Top             =   1350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterRefItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3975
         TabIndex        =   3
         Top             =   1350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterRefItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5625
         TabIndex        =   4
         Top             =   1350
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblMasterCusType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   780
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditMasterRefItem"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_Products As Collection
Private Sc As CStockCode
Private Sub cboMasterCusGroup_Click()
   m_HasModify = True
End Sub
Private Sub cboMasterCusGroup_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboMasterCusType_Click()
   m_HasModify = True
End Sub
Private Sub cboMasterCusType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
   
   Call InitNormalLabel(lblMasterCusGroup, MapText("������١���"))
   Call InitNormalLabel(lblMasterCusType, MapText("�������١���"))
   
   Call InitCombo(cboMasterCusGroup)
   Call InitCombo(cboMasterCusType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdNext, MapText("�Ѵ�"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CMasterRefDetail
         
         Set BD = TempCollection.Item(ID)
         
         cboMasterCusGroup.ListIndex = IDToListIndex(cboMasterCusGroup, BD.MASTER_CUSGROUP_ID)
         cboMasterCusType.ListIndex = IDToListIndex(cboMasterCusType, BD.MASTER_CUSTYPE_ID)
         
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
         glbErrorLog.LocalErrorMsg = "�֧�ä�����ش��������"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
       cboMasterCusGroup.ListIndex = -1
      cboMasterCusType.ListIndex = -1
   End If
   Call QueryData(True)
   
   Call cboMasterCusGroup.SetFocus
   
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
Dim I As Long

   If Not VerifyCombo(lblMasterCusGroup, cboMasterCusGroup, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblMasterCusType, cboMasterCusType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim CheckBd As CMasterRefDetail
   For Each CheckBd In TempCollection
      I = I + 1
      If CheckBd.MASTER_CUSGROUP_ID = cboMasterCusGroup.ItemData(Minus2Zero(cboMasterCusGroup.ListIndex)) And CheckBd.MASTER_CUSTYPE_ID = cboMasterCusType.ItemData(Minus2Zero(cboMasterCusType.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("�բ����š���� : ") & cboMasterCusGroup.Text & " ��л����� " & cboMasterCusType.Text & " " & MapText("������к�����")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   Next CheckBd
   
   Dim BD As CMasterRefDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CMasterRefDetail
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.MASTER_CUSGROUP_ID = cboMasterCusGroup.ItemData(Minus2Zero(cboMasterCusGroup.ListIndex))
   BD.MASTER_CUSTYPE_ID = cboMasterCusType.ItemData(Minus2Zero(cboMasterCusType.ListIndex))
   BD.MASTER_CUSGROUP_NAME = cboMasterCusGroup.Text
   BD.MASTER_CUSTYPE_NAME = cboMasterCusType.Text
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboMasterCusGroup, Nothing, , , MASTER_CUSGROUP)
      Call LoadMaster(cboMasterCusType, Nothing, , , MASTER_CUSTYPE)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
