VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMasterMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMasterMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   2
         Top             =   7800
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10095
            TabIndex        =   7
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1770
            TabIndex        =   6
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   150
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3420
            TabIndex        =   4
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2ABC
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
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
                  Picture         =   "frmMasterMain.frx":2DD6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":36B2
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":39CE
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6945
         Left            =   0
         TabIndex        =   3
         Top             =   870
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   12250
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   6915
         Left            =   4440
         TabIndex        =   8
         Top             =   900
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   12197
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
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
         Column(1)       =   "frmMasterMain.frx":3CE8
         Column(2)       =   "frmMasterMain.frx":3DB0
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMasterMain.frx":3E54
         FormatStyle(2)  =   "frmMasterMain.frx":3FB0
         FormatStyle(3)  =   "frmMasterMain.frx":4060
         FormatStyle(4)  =   "frmMasterMain.frx":4114
         FormatStyle(5)  =   "frmMasterMain.frx":41EC
         ImageCount      =   0
         PrinterProperties=   "frmMasterMain.frx":42A4
      End
   End
End
Attribute VB_Name = "frmMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String
Private m_MasterRef As CMasterRef
Private m_MasterRef1 As CMasterRef
Private m_TempArea As MASTER_TYPE

Public HeaderText As String
Public MasterMode As Long
Private m_FieldLists As Collection
Private CurrentIndex As Long
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim MI As CMenuItem
   
   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set MI = m_FieldLists(trvMaster.SelectedItem.Key)
   
   frmAddEditMaster1.KEY_CODE = MI.KEYWORD
   frmAddEditMaster1.KEY_NAME = MI.MENU_TEXT
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.MasterMode = MasterMode
   frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
   frmAddEditMaster1.ShowMode = SHOW_ADD
   frmAddEditMaster1.HeaderText = MapText("��������ѡ") & " " & trvMaster.SelectedItem.Text
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
   
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub


Private Sub InitTreeView()
Dim Node As Node
Dim Key As String

   trvMaster.Font.Name = GLB_FONT_EX
   trvMaster.Font.Size = 14
   
   If MasterMode = 1 Then '��ǹ��ҧ
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Key = MASTER_PREFIX & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�ӹ�˹�Ҫ���"), 1, 2)
      Call AddMenuItem(MapText("���ʤӹ�˹�Ҫ���"), MapText("�ӹ�˹�Ҫ���"), Key)
      Node.Expanded = False
      
      Key = MASTER_COUNTRY & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�����"), 1, 2)
      Call AddMenuItem(MapText("���ʻ����"), MapText("�����"), Key)
      Node.Expanded = False
      
      Key = MASTER_CUSGROUP & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("������١���"), 1, 2)
      Call AddMenuItem(MapText("���ʡ�����١���"), MapText("������١���"), Key)
      Node.Expanded = False
      
      Key = MASTER_CUSTYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������١���"), 1, 2)
      Call AddMenuItem(MapText("���ʻ������١���"), MapText("�������١���"), Key)
      Node.Expanded = False
   
      Key = MASTER_CUSGRADE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�дѺ�١���"), 1, 2)
      Call AddMenuItem(MapText("�����дѺ�١���"), MapText("�дѺ�١���"), Key)
      Node.Expanded = False
         
      Key = MASTER_SUPTYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�����������"), 1, 2)
      Call AddMenuItem(MapText("���ʻ����������"), MapText("�����������"), Key)
      Node.Expanded = False
      
      Key = MASTER_SUPGRADE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�дѺ�����"), 1, 2)
      Call AddMenuItem(MapText("�����дѺ�����"), MapText("�дѺ�١���"), Key)
      Node.Expanded = False
      
      Key = MASTER_POSITION & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���˹�"), 1, 2)
      Call AddMenuItem(MapText("���ʵ��˹�"), MapText("���˹�"), Key)
      Node.Expanded = False
      
'      Key = MASTER_LOCATION_SALE & "-X"
'      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("ࢵ��â��"), 1, 2)
'      Call AddMenuItem(MapText("����ࢵ��â��"), MapText("ࢵ��â��"), Key)
'      Node.Expanded = False
      
      Key = MASTER_CUSTOMER_BLOCK & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���ͤ�١���"), 1, 2)
      Call AddMenuItem(MapText("���ʺ��ͤ�١���"), MapText("���ͤ�١���"), Key)
      Node.Expanded = False
      
      Key = MASTER_APARMAS_BRANCH & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�Ң��١���"), 1, 2)
      Call AddMenuItem(MapText("���������Ң�"), MapText("�Ң�"), Key)
      Node.Expanded = False

      Key = MASTER_DRIVER & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���Ѻö"), 1, 2)
      Call AddMenuItem(MapText("���ʤ��Ѻö"), MapText("���Ѻö"), Key)
      Node.Expanded = False
      
      Key = MASTER_TRANSPORTOR & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("����"), 1, 2)
      Call AddMenuItem(MapText("���ʢ���"), MapText("����"), Key)
      Node.Expanded = False
      
      Key = MASTER_CAR_LICENSE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("����¹ö"), 1, 2)
      Call AddMenuItem(MapText("���ʷ���¹ö"), MapText("����¹ö"), Key)
      Node.Expanded = False
      
      Key = MASTER_TRANSPORT_CYCLE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�ͺ����"), 1, 2)
      Call AddMenuItem(MapText("�����ͺ����"), MapText("�ͺ����"), Key)
      Node.Expanded = False
      
      Key = MASTER_GROUP_COM & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���������Ե���"), 1, 2)
      Call AddMenuItem(MapText("���ʡ��������Ե���"), MapText("���������Ե���"), Key)
      Node.Expanded = False
      
      Key = MASTER_PAYMENT_BY & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�����Թ��"), 1, 2)
      Call AddMenuItem(MapText("���ʪ����Թ��"), MapText("�����Թ��"), Key)
      Node.Expanded = False
   ElseIf MasterMode = 2 Then '�ѭ���¡������
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Key = MASTER_JOURNAL & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��ش����ѹ"), 1, 2)
      Call AddMenuItem(MapText("������ش����ѹ"), MapText("��ش����ѹ"), Key)
      Node.Expanded = False
   
      Key = MASTER_DEPARTMENT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("Ἱ�"), 1, 2)
      Call AddMenuItem(MapText("����Ἱ�"), MapText("Ἱ�"), Key)
      Node.Expanded = False
   ElseIf MasterMode = 3 Then '��ѧ
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Key = MASTER_DEPARTMENT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("Ἱ�"), 1, 2)
      Call AddMenuItem(MapText("����Ἱ�"), MapText("Ἱ�"), Key)
      Node.Expanded = False
      
      Key = MASTER_UNIT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("˹����Ѵ"), 1, 2)
      Call AddMenuItem(MapText("����˹����Ѵ"), MapText("˹����Ѵ"), Key)
      Node.Expanded = False
   
      Key = MASTER_STOCKGROUP & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��������ʤ�ѧ"), 1, 2)
      Call AddMenuItem(MapText("���ʡ�������ʤ�ѧ"), MapText("��������ʤ�ѧ"), Key)
      Node.Expanded = False
   
      Key = MASTER_STOCKTYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���������ʤ�ѧ"), 1, 2)
      Call AddMenuItem(MapText("���ʻ��������ʤ�ѧ"), MapText("���������ʤ�ѧ"), Key)
      Node.Expanded = False
      
      Key = MASTER_STOCKTYPE_SUB & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������������¤�ѧ"), 1, 2)
      Call AddMenuItem(MapText("���ʻ������������¤�ѧ"), MapText("�������������¤�ѧ"), Key)
      Node.Expanded = False
      
      Key = MASTER_LOCATION & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("ʶҹ���Ѵ��"), 1, 2)
      Call AddMenuItem(MapText("����ʶҹ���Ѵ��"), MapText("ʶҹ���Ѵ��"), Key)
      Node.Expanded = False
      
      Key = MASTER_INVENTORY_SUB_TYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������͡�������"), 1, 2)
      Call AddMenuItem(MapText("���ʻ������͡�������"), MapText("�������͡�������"), Key)
      Node.Expanded = False
      
      Key = MASTER_INVENTORY_SALE_GROUP & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�����ʶҹ���Ѵ��"), 1, 2)
      Call AddMenuItem(MapText("���ʤ�ѧ��ѡ�ҹ���"), MapText("�����ʶҹ���Ѵ��"), Key)
      Node.Expanded = False
      
      
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Key = MASTER_DOCTYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������͡���"), 1, 2)
      Call AddMenuItem(MapText("�����͡���"), MapText("�������͡���"), Key)
      Node.Expanded = False
   
      Key = MASTER_CHEQUE_TYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��������"), 1, 2)
      Call AddMenuItem(MapText("���ʻ�����"), MapText("��������"), Key)
      Node.Expanded = False
   
      Key = MASTER_BANK & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��Ҥ��"), 1, 2)
      Call AddMenuItem(MapText("���ʸ�Ҥ��"), MapText("��Ҥ��"), Key)
      Node.Expanded = False
      
      Key = MASTER_BBRANCH & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�ҢҸ�Ҥ��"), 1, 2)
      Call AddMenuItem(MapText("�����Ң�"), MapText("�ҢҸ�Ҥ��"), Key)
      Node.Expanded = False
      
      Key = MASTER_BACCOUNT_TYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������ѭ��"), 1, 2)
      Call AddMenuItem(MapText("���ʻ������ѭ��"), MapText("�������ѭ��"), Key)
      Node.Expanded = False
      
      Key = MASTER_BANK_ACCOUNT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�ѭ�ո�Ҥ��"), 1, 2)
      Call AddMenuItem(MapText("���ʺѭ�ո�Ҥ��"), MapText("�ѭ�ո�Ҥ��"), Key)
      Node.Expanded = False
      
      Key = MASTER_CNDN_REASON & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("���˵ء������Ŵ˹��"), 1, 2)
      Call AddMenuItem(MapText("���ʡ������Ŵ˹��"), MapText("���˵ء������Ŵ˹��"), Key)
      Node.Expanded = False
      
      Key = MASTER_INVOICE_SUB & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("����Թ�������"), 1, 2)
      Call AddMenuItem(MapText("��������Թ�������"), MapText("����Թ�������"), Key)
      Node.Expanded = False
      
      Key = MASTER_INVOICE_RETURN & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("����Թ��ҹӡ�Ѻ"), 1, 2)
      Call AddMenuItem(MapText("��������Թ��ҹӡ�Ѻ"), MapText("����Թ��ҹӡ�Ѻ"), Key)
      Node.Expanded = False
      
      Key = MASTER_SUBTRACT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��¡���ѡ"), 1, 2)
      Call AddMenuItem(MapText("������¡���ѡ"), MapText("��¡���ѡ"), Key)
      Node.Expanded = False
   
      Key = MASTER_ADDITION & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��¡������"), 1, 2)
      Call AddMenuItem(MapText("������¡������"), MapText("��¡������"), Key)
      Node.Expanded = False
      
      Key = MASTER_DISCOUNT & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��¡����ǹŴ"), 1, 2)
      Call AddMenuItem(MapText("������¡����ǹŴ"), MapText("��¡����ǹŴ"), Key)
      Node.Expanded = False
      
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Key = MASTER_PRODUCTION_LOST & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("��¡���ٹ�����¨ҡ��ü�Ե"), 1, 2)
      Call AddMenuItem(MapText("������¡��"), MapText("��¡��"), Key)
      
      Key = MASTER_PRODUCTION_LOCATION & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("ʶҹ����Ե"), 1, 2)
      Call AddMenuItem(MapText("����ʶҹ����Ե"), MapText("ʶҹ����Ե"), Key)
      
      Key = MASTER_PRODUCTION_TYPE & "-X"
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, Key, MapText("�������Թ��Ҽ�Ե"), 1, 2)
      Call AddMenuItem(MapText("���ʻ������Թ��Ҽ�Ե"), MapText("�������Թ��Ҽ�Ե"), Key)
      
      Node.Expanded = False
   ElseIf MasterMode = 7 Then
   ElseIf MasterMode = 8 Then
   End If
End Sub

Private Sub AddMenuItem(KeyCode As String, KeyName As String, Key As String)
Dim MI As CMenuItem

   Set MI = New CMenuItem
   MI.KEYWORD = KeyCode
   MI.MENU_TEXT = KeyName
   
   Call m_FieldLists.add(MI, Key)
   
   Set MI = Nothing
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
Dim Status As Boolean
Dim IsOK As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
      
   m_MasterRef.KEY_ID = TempID
   Status = glbDaily.DeleteMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   If Status Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   Else
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Exit Sub
   
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim TempID As Long
Dim MI As CMenuItem
'Dim TempRow As Long
   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   'TempRow = GridEX1.RowIndex(GridEX1.Row)
   CurrentIndex = TempID
   
   Set MI = m_FieldLists(trvMaster.SelectedItem.Key)
      
   frmAddEditMaster1.KEY_CODE = MI.KEYWORD
   frmAddEditMaster1.KEY_NAME = MI.MENU_TEXT
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.MasterMode = MasterMode
   frmAddEditMaster1.ID = TempID
   frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
   frmAddEditMaster1.ShowMode = SHOW_EDIT
   frmAddEditMaster1.HeaderText = MapText("��������ѡ") & " " & trvMaster.SelectedItem.Text
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
   If OKClick Then
'      Call trvMaster_NodeClick(trvMaster.SelectedItem)
'      GridEX1.Row = TempRow
'      GridEX1.Refresh
''      GridEX1.MoveToRowIndex (TempRow)
''      GridEX1.SetFocus
''      GridEX1.MoveLast
'      GridEX1.View = jgexTable
   End If
End Sub
Private Sub Form_Activate()
Dim itemcount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      Call QueryData(True)
      m_HasActivate = True
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
'      Call cmdOK_Click
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
   
   Set m_MasterRef = Nothing
   Set m_MasterRef1 = Nothing
   Set m_FieldLists = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid(Key As String)
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle
Dim MI As CMenuItem

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   If (Key <> "") And (Key <> "Root") Then
      Set MI = m_FieldLists(Key)
      
      Set Col = GridEX1.Columns.add '1
      Col.Width = 0
      Col.Caption = "ID"
   
      Set Col = GridEX1.Columns.add '2
      Col.Width = 2235
      Col.Caption = MI.KEYWORD
         
      Set Col = GridEX1.Columns.add '3
      Col.Width = 5100
      Col.Caption = MI.MENU_TEXT
      
      If Key = "21-X" Then
         Set Col = GridEX1.Columns.add '
         Col.Width = 2000
         Col.Caption = "�����١���"
         
         Set Col = GridEX1.Columns.add '
         Col.Width = 2000
         Col.Caption = "���ʾ�ѡ�ҹ"
         
         Set Col = GridEX1.Columns.add '
         Col.Width = 2000
         Col.Caption = "���ʵ��᷹"
      ElseIf Key = "36-X" Then
         Set Col = GridEX1.Columns.add '
         Col.Width = 1800
         Col.Caption = "�Դ�Թ���·ҧ"
         
         Set Col = GridEX1.Columns.add '
         Col.Width = 1800
         Col.Caption = "�ʴ��������"
         
         Set Col = GridEX1.Columns.add '
         Col.Width = 2000
         Col.Caption = "�������"
      ElseIf Key = "14-X" Then
         Set Col = GridEX1.Columns.add '
         Col.Width = 1800
         Col.Caption = "�����ʶҹ���Ѵ��"
      End If
   End If
   
   GridEX1.itemcount = 0
   GridEX1.Rebind
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR

   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdExit, MapText("�͡ (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
   Call InitGrid("")
   
'   lsvMaster.Font.NAME = GLB_FONT
'   lsvMaster.Font.Size = 14
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   m_HasActivate = False
   m_TableName = "SYSTEM_PARAM"
   Set m_Rs = New ADODB.Recordset
   
   Set m_MasterRef = New CMasterRef
   Set m_FieldLists = New Collection
   
   Call InitFormLayout
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   End If
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
   Call m_MasterRef.PopulateFromRS(1, m_Rs)
      
   Values(1) = m_MasterRef.KEY_ID
   Values(2) = m_MasterRef.KEY_CODE
   Values(3) = m_MasterRef.KEY_NAME
   
   If m_TempArea = MASTER_TRANSPORTOR Then
      Values(4) = m_MasterRef.CASH_DELIVERY_FLAG
      Values(5) = m_MasterRef.INDEX_LINK
      Values(6) = m_MasterRef.SHORT_CODE
  ElseIf m_TempArea = MASTER_APARMAS_BRANCH Then
      Values(4) = m_MasterRef.APAR_CODE
      Values(5) = m_MasterRef.EMP_CODE
      Values(6) = m_MasterRef.DEALER_CODE
   ElseIf m_TempArea = MASTER_LOCATION Then
      Values(4) = m_MasterRef.KEY_NAME2
   End If
   
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Private Sub LoadListView(Rs As ADODB.Recordset, FieldName As String, IDName As String)
'Dim Lst As ListItem
'
'   While Not Rs.EOF
'      Set Lst = lsvMaster.ListItems.Add(, , NVLS(Rs(FieldName), ""), 1, 1)
'      Lst.Tag = NVLI(Rs(IDName), 0)
'      Rs.MoveNext
'   Wend
'End Sub

Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim itemcount As Long
Dim QueryFlag As Boolean
   
'   If LastKey = Node.Key Then
'      Exit Sub
'   End If

   Status = True
   QueryFlag = False
   
   m_TempArea = Val(Node.Key)
   Call InitGrid(Node.Key)
   
   If m_TempArea > 0 Then
      Dim Mr As CMasterRef
      Set Mr = New CMasterRef
      Mr.KEY_ID = -1
      Mr.MASTER_AREA = m_TempArea
      Call Mr.QueryData(1, m_Rs, itemcount, True)
      GridEX1.itemcount = itemcount
      GridEX1.Rebind
      Set Mr = Nothing
      
      If CurrentIndex > 0 Then
         Call GridEX1.SetFocus
         Call GridEX1.Find(1, jgexEqual, CurrentIndex)
      End If
   End If
End Sub
Private Sub Form_Resize()
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   trvMaster.Width = (1 / 3) * ScaleWidth
   trvMaster.Height = ScaleHeight - pnlHeader.Height - pnlFooter.Height
   GridEX1.Left = trvMaster.Width
   GridEX1.Width = ScaleWidth - trvMaster.Width
   GridEX1.Height = trvMaster.Height
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.Height
   
   cmdExit.Left = ScaleWidth - cmdExit.Width - 20
   
End Sub

