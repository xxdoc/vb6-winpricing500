VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmLoginfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2685
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   4736
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjWINPricing500.uctlTextBox txtUserName 
         Height          =   435
         Left            =   1830
         TabIndex        =   0
         Top             =   930
         Width           =   4275
         _ExtentX        =   7117
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjWINPricing500.uctlTextBox txtOldPassword 
         Height          =   435
         Left            =   1830
         TabIndex        =   1
         Top             =   1380
         Width           =   3525
         _ExtentX        =   7541
         _ExtentY        =   767
      End
      Begin VB.Label lblOldPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   7
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   990
         Width           =   1665
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3495
         TabIndex        =   3
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1845
         TabIndex        =   2
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmLoginfrm.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Public OKClick As Boolean
'
'Private Sub cmdPasswd_Click()
'
'End Sub
'
'Private Sub cmdOK_Click()
'Dim IsCanLogin As Boolean
'
'   Call EnableForm(Me, False)
'   If Not glbAdmin.DBLogin(txtUserName.Text, txtOldPassword.Text, IsCanLogin, glbUser, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'
'      Call EnableForm(Me, True)
'      txtUserName.SetFocus
'      Exit Sub
'   End If
'
'   If Not IsCanLogin Then
'      glbErrorLog.ShowUserError
'
'      Call EnableForm(Me, True)
'      txtUserName.SetFocus
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, True)
'
'   OKClick = True
'   Unload Me
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If Shift = 1 And KeyCode = DUMMY_KEY Then
'      glbErrorLog.LocalErrorMsg = Me.Name
'      glbErrorLog.ShowUserError
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 116 Then
''      Call cmdSearch_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 115 Then
''      Call cmdClear_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 118 Then
''      Call cmdAdd_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
''      Call cmdDelete_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 114 Then
''      Call cmdEdit_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 121 Then
''      Call cmdPrint_Click
'      KeyCode = 0
'   End If
'End Sub
'
'Private Sub InitFormLayout()
'   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
'   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
'
'   Me.Caption = MapText("��ͤ�Թ�������к�")
'   pnlHeader.Caption = Me.Caption
'
'   Call InitNormalLabel(lblUsername, "���ͼ���� ")
'   Call InitNormalLabel(lblOldPassword, "���ʼ�ҹ")
'
'   Call txtUserName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
'   Call txtUserName.SetTextType(1)
'   Call txtOldPassword.SetTextLenType(TEXT_STRING, glbSetting.PASSWORD_TYPE)
'   txtOldPassword.PasswordChar = "*"
'
'   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
'
'   pnlHeader.Font.Name = GLB_FONT
'   pnlHeader.Font.Bold = True
'   pnlHeader.Font.Size = 19
'
'   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
'
'   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
'   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
'End Sub
'
'Private Sub cmdExit_Click()
'   OKClick = False
'   Unload Me
'End Sub
'
'Private Sub Form_Load()
'   Call InitFormLayout
'End Sub
