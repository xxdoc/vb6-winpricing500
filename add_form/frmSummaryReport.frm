VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSummaryReport 
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   Icon            =   "frmSummaryReport.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   13755
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10275
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   18124
      _Version        =   131073
      Begin Threed.SSPanel pnlFooter 
         Height          =   825
         Left            =   0
         TabIndex        =   7
         Top             =   9480
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   1455
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   10140
            TabIndex        =   14
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReport.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   11790
            TabIndex        =   15
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdConfig 
            Height          =   525
            Left            =   8520
            TabIndex        =   13
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   8775
         Left            =   7080
         TabIndex        =   8
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   15478
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   4800
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   10
            Top             =   4200
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.ComboBox cboGeneric 
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
            Index           =   0
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   630
            Visible         =   0   'False
            Width           =   3855
         End
         Begin Xivess.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   2400
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin Xivess.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   2400
            TabIndex        =   1
            Top             =   240
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin Threed.SSCheck chkCommit 
            Height          =   300
            Index           =   0
            Left            =   2400
            TabIndex        =   4
            Top             =   1560
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   529
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   2205
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   735
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   1296
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   4080
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":2ABC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":3398
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":36B4
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   3480
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
                  Picture         =   "frmSummaryReport.frx":3F8E
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin Xivess.uctlTextBox txtSpace 
            Height          =   435
            Left            =   1230
            TabIndex        =   12
            Top             =   165
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
         End
         Begin VB.Label lblSpace 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   1155
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   8835
         Left            =   0
         TabIndex        =   0
         Top             =   750
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   15584
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
   End
End
Attribute VB_Name = "frmSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String

Public HeaderText As String
Public MasterMode As Long

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_TaxDocs  As Collection
Private m_Checks As Collection

Private m_FromDate As Date
Private m_ToDate As Date
Private m_FromRcp As Date
Private m_ToRcp As Date

Private m_MonthID As Long
Private m_YearNo As String
Private TEMP_ROOT_TREE As String

Private m_CyclePerMonth As Long
Private C As CReportControl

Private SupLookup As Collection
Private EmpLookup As Collection
Private PartLookup As Collection

Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT_EX
   trvMaster.Font.Size = 14
  
  If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-1", MapText("��§ҹ��������ѡ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-2", MapText("��§ҹ�������١�������Ң�"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-3", MapText("��§ҹ�����ž�ѡ�ҹ�������Ң�"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PT-1", MapText("��§ҹ����������ͧ�����"), 1, 2)
      Node.Expanded = True
      
      
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-1", MapText("��§ҹ�������١���"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-2", MapText("��§ҹ�����ž�ѡ�ҹ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-3", MapText("��§ҹ�����ūѾ���������"), 1, 2)
      Node.Expanded = True

   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-1", MapText("��§ҹ��ü�Ե��Ш��ѹ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-1", tvwChild, ROOT_TREE & " PD-1-1", MapText("��§ҹ㺼�Ե��Ш��ѹ(PD001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-2", MapText("��§ҹ��ü�Ե�����ü�Ե��ԧ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-2", tvwChild, ROOT_TREE & " PD-1-4", MapText("��§ҹ%�ʹ��Ե���LOT��Ե  �������ѵ�شԺ �ѹ���LOT ��������Ե (PD002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-2", tvwChild, ROOT_TREE & " PD-1-6", MapText("��§ҹ%�ʹ��Ե����ӹǹ��Ե �������ѵ�شԺ �ѹ���LOT ��������Ե (PD003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-3", MapText("��§ҹ��ü�Ե����ѹ����Ե"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-2", MapText("��§ҹ%�ʹ��Ե �������ѹ����Ե (PD004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-3", MapText("��§ҹ%�ʹ��Ե �������ѹ����Ե ��������Ե (PD005)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-5", MapText("��§ҹ%�ʹ��Ե �������ѵ�شԺ �ѹ����Ե ��������Ե (PD006)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-4", MapText("��§ҹ��ü�Ե����ѹ��� LOT�����"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-4", tvwChild, ROOT_TREE & " PD-1-8", MapText("��§ҹ%�ʹ��Ե����ӹǹ��Ե  �������ѵ�شԺ �ѹ���LOT (PD007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-5", MapText("��§ҹ��ü�Ե��������Ţ��Ե"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-5", tvwChild, ROOT_TREE & " PD-1-7", MapText("��§ҹ%�ʹ��Ե����ӹǹ��Ե  �����������Ţ��Ե ��������Ե (PD008)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-6", MapText("��§ҹ��ü�Ե���ʶҹ����Ե��ѡ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-6", tvwChild, ROOT_TREE & " PD-1-9", MapText("��§ҹ%�ʹ��Ե����ӹǹ��Ե  �������ѵ�شԺ ʶҹ����Ե��ѡ (PD009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-7", MapText("��§ҹ��ҡ�ü�Ե"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-7", tvwChild, ROOT_TREE & " PD-7-1", MapText("��§ҹ�ʹ��Ե���º��º��ҡ�ü�Ե��� �ѹ����Ե (PD010)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-1", MapText("�к������š���Թ"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-1", MapText("��§ҹ����¹�礨��� CHQ001"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-2", MapText("��§ҹ����¹���Ѻ CHQ002"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-3", MapText("��§ҹ����¹�礨��µ�����˹�� CHQ003"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-4", MapText("��§ҹ����¹���Ѻ����١˹�� CHQ004"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PON-1", MapText("����Ѻ �س��س��"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PON-1", tvwChild, ROOT_TREE & " PON-1-1", MapText("��§ҹ���º��º�ʹ��� �¡���������Թ��� �������ѷ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PON-1", tvwChild, ROOT_TREE & " PON-1-2", MapText("��§ҹ���º��º�ʹ��� �¡���������Թ��� �������ѷ (੾�������͹)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " SL-1", MapText("����Ѻ��ѡ�ҹ���"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-1", MapText("��§ҹ�ʹ������º��º"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �١��� �Թ��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-1", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �Թ��� �١��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-2", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �١��� �Թ��� ����繧Ǵ ������������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-3", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �Թ��� �١��� ����繧Ǵ ������������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-4", MapText("��§ҹ���º��º�ʹ��� �¡�����ѡ�ҹ��� �١��� �Թ��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-5", MapText("��§ҹ���º��º�ʹ��� �¡�����ѡ�ҹ��� �Թ��� �١��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-6", MapText("��§ҹ�ʹ��� �¡���������١��� ��ѡ�ҹ��� �١��� �Թ��� ������������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-7", MapText("��§ҹ�ʹ��� �¡���������١��� ��ѡ�ҹ��� ������Թ��� �Թ��� ������������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-8", MapText("��§ҹ�ʹ��� �¡����١��� ������Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-8-1", MapText("��§ҹ�ʹ��� �¡����١��� �Թ���"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-9", MapText("��§ҹ��èѴ�ӴѺ�ʹ��� ����Թ���"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-10", MapText("��§ҹ��èѴ�ӴѺ�ʹ��� ����١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-11", MapText("��§ҹ���º��º��ҡ�â�� �ʹ��� �¡�����ѡ�ҹ��� ����ѹ ����"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " ST-1", MapText("�к��ҹ��� TOP"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-1", MapText("��§ҹ�ʹ��� TOP �ء�Թ��� ���¡����Ң�"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-2", MapText("��§ҹ�ʹ��� TOP �¡����Թ�������Ң�"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-3", MapText("��§ҹ�ʹ��� TOP �¡�����ѡ�ҹ�������Ң�"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " S-1", MapText("�к��ҹ���"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-1", MapText("��§ҹ PO ����ѹ����͡���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-2", MapText("��§ҹ PO ����ѹ����觢ͧ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3", MapText("��§ҹ����ԡ�Թ����¡��� �ѹ��� ���Ѻ ����¹ ���� (���� PO)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3-1", MapText("��§ҹ����ԡ�Թ�����ػ��� �ѹ��� ���Ѻ ����¹ ���� (���� PO)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-4", MapText("��§ҹ����ԡ�Թ����¡��� �ѹ��� ���Ѻ ����¹ ���� (���� �觢ͧ��� ���ʴ)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3-5", MapText("��§ҹ��������´ �ѹ��� ���Ѻ ����¹ ���� (�����觢ͧ)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-5", MapText("��§ҹ˹��¢����Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-6", MapText("��§ҹ PO ��ҧ��"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1", MapText("����Թ����繪ش"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-8", MapText("����Թ����繪ش (1/2 Letter)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-2", MapText("����Թ����繪ش (�������������)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-7", MapText("��������� ���Ѵ���Թ����繪ش"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-3", MapText("��Ѻ�׹�Թ����繪ش (�������������)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-4", MapText("�Ŵ˹���繪ش (�������������)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-5", MapText("���ػ�ҧ����繪ش (�������������)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-1", MapText("㺽ҡ���/����Թ��Ҫ��Ǥ����繪ش"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-6", MapText("��§ҹ���������"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-2", MapText("��§ҹ�ʹ��� ��� �����Թ��� �ѹ�����"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-4", MapText("��§ҹ�͡����礵���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-9", MapText("��§ҹ����ѵԡ�â�� �¡����Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-10", MapText("��§ҹ����ѵԡ�â�� �¡����١��� �Թ��� �͡���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-10-1", MapText("��§ҹ����ѵԡ�â�� �¡����١��� �͡��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-12", MapText("��§ҹ�ʹ��»�ШӧǴ (�ʹ��¡�͹�Դ VAT)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-13", MapText("��§ҹ��ػ�ʹ��������¡����١˹��"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-14", MapText("��§ҹ��ػ�ʹ��� ᨡᨧ�繧Ǵ �¡�����Ǵ�Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-22", MapText("��§ҹ����Թʴ ���§����ѹ���"), 1, 2)
      Node.Expanded = False
     
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-21", MapText("��§ҹ�������"), 1, 2)
      Node.Expanded = False
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-21", MapText("��§ҹ������� "), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-34", MapText("��§ҹ�ҡ��� ���§����ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-16", MapText("��§ҹ��èѴ�ӴѺ�ʹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-17", MapText("��§ҹ��ػ�ʹ��� ᨡᨧ�繧Ǵ �¡����١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-17-1", MapText("��§ҹ��ػ�ʹ��� ᨡᨧ�繧Ǵ �¡��� ࢵ��â�� �١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-18", MapText("��§ҹ��ػ�ʹ��� �¡�����Ǵ�Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-18-1", MapText("��§ҹ��ػ�ʹ��� �¡�����Ǵ�Թ��� �������Թ�������"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-19", MapText("��§ҹ��ػ�ʹ��� �¡����١���"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-20", MapText("��§ҹʶҹС�����Թ��� �¡����١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-23", MapText("��§ҹ���º��º�ʹ��� �¡���������١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-24", MapText("��§ҹ���º��º�ʹ��� �¡����������١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-25", MapText("��§ҹ���º��º�ʹ��� �¡����Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-27", MapText("��§ҹ���º��º�ʹ��� �¡���������١��� �������͡��� �������١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-28", MapText("��§ҹ�ʹ��� �¡���������١��� �������͡��� �������١��� ������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-30", MapText("��§ҹ�ʹ��� �¡���������١��� �������͡��� ������ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-32", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �١��� �Թ��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-32-1", MapText("��§ҹ�ʹ��� �¡�����ѡ�ҹ��� �Թ��� ����繧Ǵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " T-1", MapText("�к�����"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " T-1", tvwChild, ROOT_TREE & " T-1-1", MapText("��§ҹ���բ��"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " D-1", MapText("�к��١˹��"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-1", MapText("��§ҹʶҹ��١˹�� (AR001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-2", MapText("��§ҹ�١˹�餧��ҧ (AR002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-2-2", MapText("��§ҹ�١˹�餧��ҧ ���§�����ѡ�ҹ���(AR002.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-5", MapText("��§ҹ�������͹����١˹�� (AR003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-4", MapText("��§ҹ㺡ӡѺ���֧��˹����Թ �¡����١��� (AR004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-6", MapText("��§ҹ㺡ӡѺ���֧��˹����Թ �¡����١��� �ѹ��� (AR005)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11", MapText("��§ҹ�������������١˹�� �¡����١���"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-1", MapText("��§ҹ�������������١˹�� �¡�����ѡ�ҹ��� �١���"), 1, 2)
'      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-2", MapText("��§ҹ�������������١˹�� �¡����١��� (AR006)"), 1, 2)
      Node.Expanded = False
    
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-3", MapText("��§ҹ�������������١˹�� �¡�����ѡ�ҹ��� �١���  (AR007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-4", MapText("��§ҹ�������������١˹�� �¡�����ѡ�ҹ��� �١��� (�кت�ǧ�ѹ���) (AR008)"), 1, 2)
      Node.Expanded = False
       
            
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-6", MapText("��§ҹ�Ŵ˹��/�Ѻ�׹�Թ��� �¡����١���  (AR009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-8", MapText("��§ҹ�Ŵ˹��/�Ѻ�׹�Թ��� ���§���ࢵ��â��,��ѡ�ҹ���  (AR010)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31", MapText("��§ҹ�Ŵ˹��/�Ѻ�׹�Թ��� ���§����ѹ��� (AR011)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31-2", MapText("��§ҹ�Ŵ˹��/�Ѻ�׹�Թ��� ���§����Թ��� (AR011.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31-1", MapText("��§ҹ�Ŵ˹�� ���§����ѹ��� (AR012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-33", MapText("��§ҹ�������������١˹�� �¡�����ѡ�ҹ��� �١��� �������١��� (AR013)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " R-1", MapText("�к�����Ѻ����˹��"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-2", MapText("��§ҹ�Ѻ�����¡����١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-4", MapText("��§ҹ�Ѻ�������§����ѹ��������"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-6", MapText("��§ҹ�Ѻ�����¡�����ѡ�ҹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " B-1", MapText("�к���ë���"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-1", MapText("��§ҹ����ѵԡ�ë��� �¡����Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-2", MapText("��§ҹ����ѵԡ�ë��� �¡����Ѿ���������"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-3", MapText("��§ҹ��Ѻ�Թ��� ���§����ѹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-4", MapText("��§ҹ��ػ�ʹ���� ᨡᨧ�繧Ǵ �¡����Ѿ��������� �Թ���"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-10", MapText("��§ҹ�����Թ�������ѵ�شԺ (ST001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-1-1", MapText("��§ҹ�ӹǹ����ѧ����ʹ��Ǩ�Ѻ (ST002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2", MapText("��§ҹ�������͹��Ǥ�ѧ Ẻ 1 (ST003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-1", MapText("��§ҹ�������͹��Ǥ�ѧ Ẻ 2 (ST004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-2", MapText("��§ҹ�������͹����Թ��� �¡�����ѧ�Թ��� ᨡᨧ�ѹ��� (ST004-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3", MapText("��§ҹ��ػ�ʹ����͹����Թ��� �¡�����ѧ�Թ��� (ST005)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-1", MapText("��§ҹ��ػ�ʹ�Թ��� �¡�����ѧ�Թ��� (ST005-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-2", MapText("��§ҹ��ػ�ʹ�Թ��� ���§�����ѧ�Թ��� (ST005-2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-3", MapText("��§ҹ��ػ�ʹ�Թ��� ���§����Թ��� (ST005-3)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-4", MapText("��§ҹ��ػ�ʹ�Թ��Ҥ������ ���§�����ѧ (ST005-4)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-5", MapText("��§ҹ��ػ�ʹ����͹����Թ��� �¡����������ѧ�Թ��� (ST005-5)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4", MapText("��§ҹ�͡��á���͹ �¡�����ѧ�Թ��� (ST006)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4-1", MapText("��§ҹ�͡��á���͹ ����Ѻ�׹  (ST006-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-5", MapText("��§ҹ�͡��á���ԡ (ST007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6", MapText("��§ҹ�͡��á���ԡ ᨡᨧ����������͡������� (ST008)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-2", MapText("��§ҹ�͡��á���ԡ ᨧᨧ����������ԡ (ST008.1)"), 1, 2)
      Node.Expanded = False
      
      
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-1", MapText("��§ҹ�͡��á���ԡ ᨧᨧ��������Ţ�͡��� (ST009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-3", MapText("��§ҹ�ʹ�ԡ�ѵ�شԺ  �¡�����͹ (ST009.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-7", MapText("��§ҹ�͡��á���Ѻ��� �¡�����ѧ�Թ��� (ST010)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-8", MapText("��§ҹ��ü�Ե �ҡ BARCODE (ST011)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-8-1", MapText("��§ҹ % ��ü�Ե �ҡ BARCODE �¡����ѵ�شԺ (ST012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-9", MapText("����������(ST101)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 7 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-1", MapText("��§ҹ��͹��ѵ��ԡ���¤�� Commission"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-2", MapText("��§ҹ���º��º�ʹ��� �¡�����ѡ�ҹ��� ���Ἱ���Ծ�ѡ�ҹ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-3", MapText("��§ҹ��è��� Rebate ������᷹��˹���"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 8 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-1", MapText("��§ҹ��ҡ�â�� �¡��� ࢵ ��ѡ�ҹ��� �١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-2", MapText("��§ҹ��ҡ�â�� �¡��� ࢵ ��ѡ�ҹ��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-4", MapText("��§ҹ��ҡ�â�� �¡��� ࢵ �١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-5", MapText("��§ҹ��ҡ�â�� �¡��� ࢵ �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-6", MapText("��§ҹ��ҡ�â�� �¡��� ��ѡ�ҹ��� �١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-7", MapText("��§ҹ��ҡ�â�� �¡��� ��ѡ�ҹ��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-8", MapText("��§ҹ��ҡ�â�� �¡��� �١��� �Թ���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-3", MapText("��§ҹ��ҡ�â�� �¡��� �������Թ��� �Թ���"), 1, 2)
      Node.Expanded = False
   End If
End Sub
Private Sub FillReportInput(R As CReportInterface)
On Error Resume Next
   
   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
         
         If C.Param2 = "MONTH_ID" Then
            m_MonthID = cboGeneric(C.ControlIndex).ListIndex
         End If
         
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
         
         If Len(txtGeneric(C.ControlIndex).Text) = 0 Then
            If C.Param2 = "YEAR_NO" Then
               txtGeneric(C.ControlIndex).Text = Year(Now)
            End If
         End If
         
         If C.Param2 = "YEAR_NO" Then
            m_YearNo = txtGeneric(C.ControlIndex).Text
         End If
         
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_BILL_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_BILL_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "FROM_RPC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "TO_RPC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               End If
            End If
            If C.Param2 = "FROM_BILL_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_BILL_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "FROM_RCP_DATE" Then
               m_FromRcp = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_RCP_DATE" Then
               m_ToRcp = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
      
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param2)
         End If
      End If
      
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
On Error Resume Next
   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   Next C
   VerifyReportInput = True
End Function
Private Sub cboGeneric_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkCommit_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim ReportMode As Long
   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMaster.SelectedItem.Key
   
   Set Rc = New CReportConfig
   Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
   Call Rc.QueryData(1, m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.GetFieldValue("REPORT_CONFIG_ID")
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   If ReportKey = "Root S-1-1" Or ReportKey = "Root S-1-1-7" Or ReportKey = "Root S-1-1-8" Then
      ReportMode = 2
   Else
      ReportMode = 1
   End If
   frmReportConfig.ReportMode = ReportMode
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMaster.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim ClassName As String
   
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   Key = trvMaster.SelectedItem.Key
   Name = trvMaster.SelectedItem.Text
    
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Key = "Root A-1-1" Then

      Set Report = New CReportCheque001
      ClassName = "CReportCheque001"
   ElseIf Key = "Root A-1-2" Then
      
      Set Report = New CReportCheque002
      ClassName = "CReportCheque002"
   ElseIf Key = "Root A-1-3" Then
   
      Set Report = New CReportCheque003
      ClassName = "CReportCheque003"
   ElseIf Key = "Root A-1-4" Then
      Set Report = New CReportCheque004
      ClassName = "CReportCheque004"
   
   ElseIf Key = "Root SL-1-1" Then
      Set Report = New CReportBilling038
      ClassName = "CReportBilling038"
      
   ElseIf Key = "Root S-1-1" Then
      Set Report = New CReportFormDO002
      ClassName = "CReportFormDO002"
   ElseIf Key = "Root S-1-1-8" Then
      Set Report = New CReportFormDO002_2
      ClassName = "CReportFormDO002_2"
      
   ElseIf Key = "Root S-1-1-7" Then
      Set Report = New CReportPrintLabel006
      ClassName = "CReportPrintLabel006"
   
   ElseIf Key = "Root S-1-1-2" Then
      Set Report = New CReportNormalRcp001_2
      ClassName = "CReportNormalRcp001_2"

  ElseIf Key = "Root S-1-1-3" Then 'pui ����    ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
      Set Report = New CReportNormalRcp001_3
      ClassName = "CReportNormalRcp001_3"
      
   ElseIf Key = "Root S-1-1-4" Then 'pui ����    ' pui ���� ����Ѻ �Ŵ˹���繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->�Ŵ˹���繪ش(���������)    CReportNormalRcp001_4
      Set Report = New CReportNormalRcp001_4
      ClassName = "CReportNormalRcp001_4"
  ElseIf Key = "Root S-1-1-5" Then 'pui ����    ' pui ���� ����Ѻ ���ػ�ҧ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->���ػ�ҧ����繪ش(���������)    CReportNormalRcp001_5
      Set Report = New CReportNormalRcp001_5
      ClassName = "CReportNormalRcp001_5"

   ElseIf Key = "Root S-1-1-1" Then
      Set Report = New CReportNormalPO001
      ClassName = "CReportNormalPO001"
      
   ElseIf Key = "Root S-1-1-6" Then
      Set Report = New CReportPrintLabel005
      ClassName = "CReportPrintLabel005"
' S -1 - 1 - 6
   ElseIf Key = "Root S-1-2" Then
      Set Report = New CReportBilling002
      ClassName = "CReportBilling002"
   ElseIf Key = "Root S-2-1" Then
      Set Report = New CReportBilling003
      ClassName = "CReportBilling003"
   ElseIf Key = "Root S-2-2" Then
      Set Report = New CReportBilling004
      ClassName = "CReportBilling004"
   ElseIf Key = "Root S-2-3" Then
      Set Report = New CReportBilling005
      ClassName = "CReportBilling005"
   ElseIf Key = "Root S-2-4" Then
      Set Report = New CReportBilling006
      ClassName = "CReportBilling006"
   ElseIf Key = "Root S-2-6" Then
      Set Report = New CReportBilling007_1
      ClassName = "CReportBilling007_1"
   ElseIf Key = "Root S-2-8" Then
      Set Report = New CReportBilling007_3
      ClassName = "CReportBilling007_3"
   ElseIf Key = "Root S-2-9" Then
      Set Report = New CReportBilling010
      ClassName = "CReportBilling010"
   ElseIf Key = "Root S-2-10" Then
      Set Report = New CReportBilling011
      ClassName = "CReportBilling011"
   ElseIf Key = "Root S-2-10-1" Then
      Set Report = New CReportBilling011_1
      ClassName = "CReportBilling011_1"
   ElseIf Key = "Root S-2-11" Then
      Set Report = New CReportBilling012
      ClassName = "CReportBilling012"
   ElseIf Key = "Root S-2-11-1" Then
      Set Report = New CReportBilling012_2
      ClassName = "CReportBilling012_2"
  ElseIf Key = "Root S-2-11-2" Then
      Set Report = New CReportBillingD002_1
      ClassName = "CReportBillingD002_1"
  ElseIf Key = "Root S-2-11-3" Then
      Set Report = New CReportBillingD002_2
      ClassName = "CReportBillingD002_2"
  ElseIf Key = "Root S-2-11-4" Then
      Set Report = New CReportBillingD002_3
      ClassName = "CReportBillingD002_3"
   ElseIf Key = "Root S-2-33" Then
      Set Report = New CReportBilling012_1
      ClassName = "CReportBilling012_1"
   ElseIf Key = "Root S-2-12" Then
      Set Report = New CReportBilling013
      ClassName = "CReportBilling013"
   ElseIf Key = "Root S-2-13" Then
      Set Report = New CReportBilling014
      ClassName = "CReportBilling014"
   ElseIf Key = "Root S-2-14" Then
      Set Report = New CReportBilling015
      ClassName = "CReportBilling015"
   ElseIf Key = "Root S-2-16" Then
      Set Report = New CReportBilling017
      ClassName = "CReportBilling017"
   ElseIf Key = "Root S-2-17" Then
      Set Report = New CReportBilling018
      ClassName = "CReportBilling018"
   ElseIf Key = "Root S-2-17-1" Then
      Set Report = New CReportBilling040
      ClassName = "CReportBilling040"
   ElseIf Key = "Root S-2-18" Then
      Set Report = New CReportBilling019
      ClassName = "CReportBilling019"
   ElseIf Key = "Root S-2-18-1" Then
      Set Report = New CReportBilling019_1
      ClassName = "CReportBilling019_1"
   ElseIf Key = "Root S-2-19" Then
      Set Report = New CReportBilling023
      ClassName = "CReportBilling023"
   ElseIf Key = "Root S-2-20" Then
      Set Report = New CReportBilling024
      ClassName = "CReportBilling024"
   ElseIf Key = "Root S-2-21" Then
      Set Report = New CReportBilling021
      ClassName = "CReportBilling021"
      Call Report.AddParam(1, "AREA") '��觢��
   ElseIf Key = "Root S-2-22" Then
      Set Report = New CReportBilling022
      ClassName = "CReportBilling022"
   ElseIf Key = "Root S-2-23" Then
      Set Report = New CReportBilling030
      ClassName = "CReportBilling030"
  ElseIf Key = "Root S-2-34" Then
      Set Report = New CReportBilling042
      ClassName = "CReportBilling042"
      Call Report.AddParam(1, "AREA") '��觢��
   ElseIf Key = "Root PON-1-1" Then
      Set Report = New CReportBilling030_1
      ClassName = "CReportBilling030_1"
   ElseIf Key = "Root PON-1-2" Then
      Set Report = New CReportBilling030_2
      ClassName = "CReportBilling030_2"
   ElseIf Key = "Root S-2-24" Then
      Set Report = New CReportBilling031
      ClassName = "CReportBilling031"
   ElseIf Key = "Root S-2-25" Then
      Set Report = New CReportBilling032
      ClassName = "CReportBilling032"
   ElseIf Key = "Root S-2-27" Then
      Set Report = New CReportBilling034
      ClassName = "CReportBilling034"
   ElseIf Key = "Root S-2-28" Then
      Set Report = New CReportBilling035
      ClassName = "CReportBilling035"
   ElseIf Key = "Root S-2-32" Then
      Set Report = New CReportBilling036_9
      ClassName = "CReportBilling036_9"
   ElseIf Key = "Root S-2-32-1" Then
      Set Report = New CReportBilling036_13
      ClassName = "CReportBilling036_13"
   ElseIf Key = "Root SL-1-2" Then
      Set Report = New CReportBilling036
      ClassName = "CReportBilling036"
   ElseIf Key = "Root SL-1-2-1" Then
      Set Report = New CReportBilling036_1
      ClassName = "CReportBilling036_1"
   ElseIf Key = "Root SL-1-2-2" Then
      Set Report = New CReportBilling036_2
      ClassName = "CReportBilling036_2"
   ElseIf Key = "Root SL-1-2-3" Then
      Set Report = New CReportBilling036_3
      ClassName = "CReportBilling036_3"
   ElseIf Key = "Root SL-1-2-4" Then
      Set Report = New CReportBilling036_4
      ClassName = "CReportBilling036_4"
   ElseIf Key = "Root SL-1-2-5" Then
      Set Report = New CReportBilling036_5
      ClassName = "CReportBilling036_5"
   ElseIf Key = "Root SL-1-2-6" Then
      Set Report = New CReportBilling036_6
      ClassName = "CReportBilling036_6"
   ElseIf Key = "Root SL-1-2-7" Then
      Set Report = New CReportBilling036_7
      ClassName = "CReportBilling036_7"
   ElseIf Key = "Root SL-1-2-8" Then
      Set Report = New CReportBilling036_8
      ClassName = "CReportBilling036_8"
   ElseIf Key = "Root SL-1-2-8-1" Then
      Set Report = New CReportBilling036_8_1
      ClassName = "CReportBilling036_8_1"
   ElseIf Key = "Root SL-1-2-9" Then
      Set Report = New CReportBilling036_10
      ClassName = "CReportBilling036_10"
   ElseIf Key = "Root SL-1-2-10" Then
      Set Report = New CReportBilling036_11
      ClassName = "CReportBilling036_11"
   ElseIf Key = "Root SL-1-2-11" Then
      Set Report = New CReportBilling036_12
      ClassName = "CReportBilling036_12"
   ElseIf Key = "Root S-2-30" Then
      Set Report = New CReportBilling035_1
      ClassName = "CReportBilling035_1"
   ElseIf Key = "Root S-2-31" Then
      Set Report = New CReportBilling007_4
      ClassName = "CReportBilling007_4"
   ElseIf Key = "Root S-2-31-1" Then
      Set Report = New CReportBilling007_5
      ClassName = "CReportBilling007_5"
   ElseIf Key = "Root S-2-31-2" Then
      Set Report = New CReportBilling007_2
      ClassName = "CReportBilling007_2"
   ElseIf Key = "Root P-1-1" Then
      Set Report = New CReportBilling008
      ClassName = "CReportBilling008"
   ElseIf Key = "Root P-1-2" Then
      Set Report = New CReportBilling009
      ClassName = "CReportBilling009"
   ElseIf Key = "Root P-1-3" Then
      Set Report = New CReportBilling037
      ClassName = "CReportBilling037"
      Call Report.AddParam(PO_DOCTYPE, "DOCUMENT_TYPE")
   ElseIf Key = "Root P-1-3-1" Then
      Set Report = New CReportBilling037_1
      ClassName = "CReportBilling037_1"
      Call Report.AddParam(PO_DOCTYPE, "DOCUMENT_TYPE")
   ElseIf Key = "Root P-1-3-5" Then
      Set Report = New CReportBilling037_3
      ClassName = "CReportBilling037_3"
   ElseIf Key = "Root P-1-4" Then
      Set Report = New CReportBilling037
      ClassName = "CReportBilling037"
      Call Report.AddParam(INVOICE_DOCTYPE, "DOCUMENT_TYPE")
   ElseIf Key = "Root P-1-5" Then
      Set Report = New CReportBilling039
      ClassName = "CReportBilling039"
   ElseIf Key = "Root P-1-6" Then
      Set Report = New CReportBillingPo001
      ClassName = "CReportBillingPo001"
   ElseIf Key = "Root T-1-1" Then
      Set Report = New CReportBillingT01
      ClassName = "CReportBillingT01"
   ElseIf Key = "Root D-1-1" Then
      Set Report = New CReportBillingD001
      ClassName = "CReportBillingD001"
   ElseIf Key = "Root D-1-2" Then
      Set Report = New CReportBillingD002
      ClassName = "CReportBillingD002"
   ElseIf Key = "Root D-1-2-2" Then
      Set Report = New CReportBillingD002_4
      ClassName = "CReportBillingD002_4"
   ElseIf Key = "Root D-1-4" Then
      Set Report = New CReportBillingD004
      ClassName = "CReportBillingD004"
   ElseIf Key = "Root D-1-5" Then
      Set Report = New CReportBillingD005
      ClassName = "CReportBillingD005"
   ElseIf Key = "Root D-1-6" Then
      Set Report = New CReportBillingD006
      ClassName = "CReportBillingD006"
   ElseIf Key = "Root R-1-2" Then
      Set Report = New CReportBilling025
      ClassName = "CReportBilling025"
   ElseIf Key = "Root R-1-4" Then
      Set Report = New CReportBilling027
      ClassName = "CReportBilling027"
   ElseIf Key = "Root R-1-6" Then
      Set Report = New CReportBilling029
      ClassName = "CReportBilling029"
   ElseIf Key = "Root 6-1-1" Then
      Set Report = New CReportInventoryDoc1_1
      ClassName = "CReportInventoryDoc1_1"
   ElseIf Key = "Root 6-2" Then
      Set Report = New CReportInventoryDoc2
      ClassName = "CReportInventoryDoc2"
   ElseIf Key = "Root 6-2-1" Then
      Set Report = New CReportInventoryDoc2_1
      ClassName = "CReportInventoryDoc2_1"
   ElseIf Key = "Root 6-2-2" Then
      Set Report = New CReportInventoryDoc2_2
      ClassName = "CReportInventoryDoc2_2"
   ElseIf Key = "Root 6-3" Then
      Set Report = New CReportInventoryDoc3
      ClassName = "CReportInventoryDoc3"
   ElseIf Key = "Root 6-3-1" Then
      Set Report = New CReportInventoryDoc3_1
      ClassName = "CReportInventoryDoc3_1"
   ElseIf Key = "Root 6-3-2" Then
      Set Report = New CReportInventoryDoc3_2_1
      ClassName = "CReportInventoryDoc3_2_1"
   ElseIf Key = "Root 6-3-3" Then
      Set Report = New CReportInventoryDoc3_3_1
      ClassName = "CReportInventoryDoc3_3_1"
   ElseIf Key = "Root 6-3-4" Then
      Set Report = New CReportInventoryDoc3_4_1
      ClassName = "CReportInventoryDoc3_4_1"
   ElseIf Key = "Root 6-3-5" Then
      Set Report = New CReportInventoryDoc3_5_1
      ClassName = "CReportInventoryDoc3_5_1"
   ElseIf Key = "Root 6-4" Then
      Set Report = New CReportInventoryDoc4
      ClassName = "CReportInventoryDoc4"
   ElseIf Key = "Root 6-4-1" Then
      Set Report = New CReportInventoryDoc4_1
      ClassName = "CReportInventoryDoc4_1"
   ElseIf Key = "Root 6-5" Then
      Set Report = New CReportInventoryDoc5
      ClassName = "CReportInventoryDoc5"
   ElseIf Key = "Root 6-6" Then
      Set Report = New CReportInventoryDoc6
      ClassName = "CReportInventoryDoc6"
   ElseIf Key = "Root 6-6-1" Then
      Set Report = New CReportInventoryDoc6_1
      ClassName = "CReportInventoryDoc6_1"
   ElseIf Key = "Root 6-6-2" Then
      Set Report = New CReportInventoryDoc6_2
      ClassName = "CReportInventoryDoc6_2"
   ElseIf Key = "Root 6-6-3" Then
      Set Report = New CReportInventoryDoc6_3
      ClassName = "CReportInventoryDoc6_3"
   ElseIf Key = "Root 6-7" Then
      Set Report = New CReportInventoryDoc7
      ClassName = "CReportInventoryDoc7"
   ElseIf Key = "Root 6-8" Then
      Set Report = New CReportInventoryDoc11
      ClassName = "CReportInventoryDoc11"
   ElseIf Key = "Root 6-8-1" Then
      Set Report = New CReportInventoryDoc11_1
      ClassName = "CReportInventoryDoc11_1"
   ElseIf Key = "Root 6-9" Then
      Set Report = New CReportInventoryDoc9
      ClassName = "CReportInventoryDoc9"
   ElseIf Key = "Root 6-10" Then
      Set Report = New CReportInventoryDoc10
      ClassName = "CReportInventoryDoc10"
   ElseIf Key = "Root 7-1" Then
      Set Report = New CReportCommission001
      ClassName = "CReportCommission001"
   ElseIf Key = "Root 7-2" Then
      Set Report = New CReportCommission002
      ClassName = "CReportCommission002"
   ElseIf Key = "Root 7-3" Then
      Set Report = New CReportCommission003
      ClassName = "CReportCommission003"
   ElseIf Key = "Root 8-1" Then
      Set Report = New CReportTaget001
      ClassName = "CReportTaget001"
   ElseIf Key = "Root 8-2" Then
      Set Report = New CReportTaget002
      ClassName = "CReportTaget002"
   ElseIf Key = "Root 8-3" Then
      Set Report = New CReportTaget003
      ClassName = "CReportTaget003"
   ElseIf Key = "Root 8-4" Then
      Set Report = New CReportTaget004
      ClassName = "CReportTaget004"
   ElseIf Key = "Root 8-5" Then
      Set Report = New CReportTaget005
      ClassName = "CReportTaget005"
   ElseIf Key = "Root 8-6" Then
      Set Report = New CReportTaget006
      ClassName = "CReportTaget006"
   ElseIf Key = "Root 8-7" Then
      Set Report = New CReportTaget007
      ClassName = "CReportTaget007"
   ElseIf Key = "Root 8-8" Then
      Set Report = New CReportTaget008
      ClassName = "CReportTaget008"
   ElseIf Key = "Root PD-1-1" Then
      Set Report = New CReportProduct001
      ClassName = "CReportProduct001"
   ElseIf Key = "Root PD-1-2" Then
      Set Report = New CReportProduct002
      ClassName = "CReportProduct002"
   ElseIf Key = "Root PD-1-3" Then
      Set Report = New CReportProduct003
      ClassName = "CReportProduct003"
   ElseIf Key = "Root PD-1-4" Then
      Set Report = New CReportProduct004
      ClassName = "CReportProduct004"
   ElseIf Key = "Root PD-1-5" Then
      Set Report = New CReportProduct003_1
      ClassName = "CReportProduct003_1"
   ElseIf Key = "Root PD-1-6" Then
      Set Report = New CReportProduct004_1
      ClassName = "CReportProduct004_1"
   ElseIf Key = "Root PD-1-7" Then
      Set Report = New CReportProduct005
      ClassName = "CReportProduct005"
   ElseIf Key = "Root PD-1-8" Then
      Set Report = New CReportProduct006
      ClassName = "CReportProduct006"
   ElseIf Key = "Root PD-1-9" Then
      Set Report = New CReportProduct007
      ClassName = "CReportProduct007"
   ElseIf Key = "Root PD-7-1" Then
      Set Report = New CReportProduct008
      ClassName = "CReportProduct008"
   ElseIf Key = "Root B-1-1" Then
      Set Report = New CReportBillingSub001
      ClassName = "CReportBillingSub001"
   ElseIf Key = "Root B-1-2" Then
      Set Report = New CReportBillingSup002
      ClassName = "CReportBillingSup002"
   ElseIf Key = "Root B-1-3" Then
      Set Report = New CReportBilling021
      ClassName = "CReportBilling021"
      
      Call Report.AddParam(2, "AREA") '��觫���
   ElseIf Key = "Root B-1-4" Then
      Set Report = New CReportBilling041
      ClassName = "CReportBilling041"
   ElseIf Key = "Root MS-1" Then
      Set Report = New CReportMaster001
      ClassName = "CReportMaster001"
   ElseIf Key = "Root MS-2" Then
      Set Report = New CReportMain002
      ClassName = "CReportMain002"
    ElseIf Key = "Root MS-3" Then
      Set Report = New CReportMain003
      ClassName = "CReportMain003"
   ElseIf Key = "Root PT-1" Then
      Set Report = New CReportPrinter001
      ClassName = "CReportPrinter001"
   ElseIf Key = "Root MN-1" Then
      Set Report = New CReportMain001
      ClassName = "CReportMain001"
   ElseIf Key = "Root MN-2" Then
      Set Report = New CReportMN02
      ClassName = "CReportMN02"
   ElseIf Key = "Root MN-3" Then
      Set Report = New CReportMN03
      ClassName = "CReportMN03"
   End If
   
   SelectFlag = True
   
   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If
      
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")
      Set frmReport.ReportObject = Report
      frmReport.ClassName = ClassName
      frmReport.Space = Val(txtSpace.Text)
      frmReport.HeaderText = MapText("�������§ҹ")
      Load frmReport
      frmReport.Show 1
      
      Unload frmReport
      Set frmReport = Nothing
   End If
   
   txtSpace.Text = ""
End Sub

Private Sub Form_Activate()
Dim itemcount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      m_HasActivate = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub Form_Resize()
   pnlHeader.Width = ScaleWidth
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   If ScaleWidth <= 0 Then
      trvMaster.Width = 0
   Else
      trvMaster.Width = ScaleWidth - SSFrame2.Width
   End If
   SSFrame2.Left = trvMaster.Width
   If ScaleHeight <= 0 Then
      trvMaster.Height = 0
   Else
      trvMaster.Height = ScaleHeight - pnlHeader.Height - pnlFooter.Height
   End If
   SSFrame2.Height = trvMaster.Height
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.Height
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20
   cmdConfig.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20 - cmdConfig.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_Checks = Nothing
   
   Set SupLookup = Nothing
   Set EmpLookup = Nothing
   Set PartLookup = Nothing
End Sub
Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   SSFrame2.BackColor = GLB_FORM_COLOR
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitNormalLabel(lblSpace, MapText("������ҧ"))
   Call InitMainButton(cmdOK, MapText("����� (F10)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("����� (F10)"))
   Call InitMainButton(cmdConfig, MapText("��Ѻ���"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   Call InitTreeView
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Call InitFormLayout
   
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   

   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_Checks = New Collection
   
   Set SupLookup = New Collection
   Set EmpLookup = New Collection
   Set PartLookup = New Collection
   
End Sub

Private Sub UnloadAllControl()
Dim I As Long
Dim j As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_Checks.Count
   While I > 0
      Call Unload(m_Checks(I))
      Call m_Checks.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub

Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long


   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            If C.OldLine Then
               m_Combos(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Combos(C.ControlIndex).Top = PrevTop - m_Combos(C.ControlIndex - 1).Height
            Else
               m_Combos(C.ControlIndex).Left = PrevLeft
               m_Combos(C.ControlIndex).Top = PrevTop
            End If
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            If C.OldLine Then
               PrevLeft = m_Combos(C.ControlIndex).Left - CurWidth - 20
            Else
               PrevLeft = m_Combos(C.ControlIndex).Left
            End If
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            If C.OldLine Then
               m_Texts(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Texts(C.ControlIndex).Top = PrevTop - txtGeneric(0).Height
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
               m_Texts(C.ControlIndex).Width = C.Width
            Else
               m_Texts(C.ControlIndex).Left = PrevLeft
               m_Texts(C.ControlIndex).Top = PrevTop
               m_Texts(C.ControlIndex).Width = C.Width
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
                              
               CurTop = PrevTop
               CurLeft = PrevLeft
               CurWidth = PrevWidth
               
               PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
               PrevLeft = m_Texts(C.ControlIndex).Left
               PrevWidth = C.Width
            End If
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            m_Checks(C.ControlIndex).Left = PrevLeft
            m_Checks(C.ControlIndex).Top = PrevTop + 100
            m_Checks(C.ControlIndex).Width = C.Width
            m_Checks(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Checks(C.ControlIndex).Top + m_Checks(C.ControlIndex).Height
            PrevLeft = m_Checks(C.ControlIndex).Left
            PrevWidth = C.Width
         End If
      
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
End Sub

Private Sub LoadComboData()
Dim Mr As CMasterRef
Dim Comb As ComboBox
   
   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
      
         Set Mr = New CMasterRef
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
            If C.ComboLoadID = 1 Then
               Call InitUserGroupOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
            If C.ComboLoadID = 1 Then
               'Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitUserOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
            If C.ComboLoadID = 1 Then
               'Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               'Call InitLoginOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Then
            If C.ComboLoadID = 1 Then
               'Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
            If C.ComboLoadID = 1 Then
               'Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_SUPTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
            If C.ComboLoadID = 1 Then
               'Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_EMPPOSITION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
            If C.ComboLoadID = 0 Then
               'Call InitTaxType(m_Combos(C.ControlIndex))
            ElseIf (C.ComboLoadID = 1) Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_SUPTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportChequeBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportChequeBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_SUPTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportChequeBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
                  
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportChequeBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-7" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_DRIVER)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_TRANSPORTOR)
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportS_1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         
            If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-6" Then
      
           If C.ComboLoadID = 1 Then
               Call InitReportS_1_1_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
            If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-3" Then     ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
              If C.ComboLoadID = 1 Then
               Call InitReportS_1_3Orderby(m_Combos(C.ControlIndex))
             ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
          End If

          If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-4" Then     ' pui ���� ����Ѻ �Ŵ˹���繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->�Ŵ˹���繪ش(���������)    CReportNormalRcp001_4
            If C.ComboLoadID = 1 Then
               Call InitReportS_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-5" Then     ' pui ���� ����Ѻ ���ػ�ҧ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->���ػ�ҧ����繪ش(���������)    CReportNormalRcp001_5
            If C.ComboLoadID = 1 Then
               Call InitReportS_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
     End If



         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , 2, MASTER_APARMAS_BRANCH)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportS_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , 2, MASTER_APARMAS_BRANCH)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportS_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , 2, MASTER_APARMAS_BRANCH)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportS_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-18" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-18-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , 1, MASTER_STOCKGROUP)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-23" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-27" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-28" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-30" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKTYPE)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 3 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " PON-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " PON-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-11" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-4" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-5" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-6" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-7" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-8" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-8-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-33" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-33" Or trvMaster.SelectedItem.Key = ROOT_TREE & " D-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " D-1-2-2" _
           Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-11-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-11-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-11-4" Then
            If C.ComboLoadID = 3 Then
               Call InitShortCode(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = 0
            ElseIf C.ComboLoadID = 4 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSTYPE)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-16" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-9" Or trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-2-10" Then
            If C.ComboLoadID = 1 Then
               Call InitReportS_2_16Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-17" Or trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-17-1" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-13" Or trvMaster.SelectedItem.Key = ROOT_TREE & " B-1-4" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-14" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
         
      If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-21" Then
         If C.ComboLoadID = 1 Then
             Call InitReportS_2_21Orderby(m_Combos(C.ControlIndex))
         End If
      End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " B-1-3" Then
         If C.ComboLoadID = 1 Then
             Call InitReportS_2_21Orderby(m_Combos(C.ControlIndex))
         End If
      End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportNullOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-4" Or trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-3-1" Then
            If C.ComboLoadID = 1 Or C.ComboLoadID = 2 Or C.ComboLoadID = 3 Or C.ComboLoadID = 4 Or C.ComboLoadID = 5 _
            Or C.ComboLoadID = 6 Or C.ComboLoadID = 7 Or C.ComboLoadID = 8 Or C.ComboLoadID = 9 Or C.ComboLoadID = 10 _
            Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_UNIT)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-3-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_TRANSPORTOR)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_TRANSPORTOR)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_TRANSPORTOR)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_TRANSPORTOR)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " P-1-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), Nothing, , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportNullOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
'         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1-1" Then
'            If C.ComboLoadID = 1 Then
'               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
'            End If
'         End If
'-------�¡  6-2 ���ҧ����
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-7" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-4" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-5" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-8" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-8-1" Then
            If C.ComboLoadID = 1 Then
'               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6-3" Then
         If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-4" Then
            If C.ComboLoadID = 1 Or C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_INVENTORY_SUB_TYPE, , , , EXPORT_DOCTYPE)
                If MASTER_INVENTORY_SUB_TYPE = 34 Then ' ������͡��á���ԡ ST007  ��������� InventoryDoc_Sub_Type =null (��ҷ���͡�� �� -1) ��� ����ԡ���������ԡ�ҡDatabase 㹷������ԡ�ҡ barcode
                       'Comb.AddItem ("��͡����ԡ����")
                       m_Combos(C.ControlIndex).AddItem ("��͡����ԡ����")
                        m_Combos(C.ControlIndex).ItemData(m_Combos(C.ControlIndex).ListCount - 1) = 999999999
                 End If
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-10" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKGROUP)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 7-1" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-2" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-3" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-4" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-5" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-6" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-7" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-8" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               If m_MonthID > 0 Then
                  cboGeneric(C.ControlIndex).ListIndex = m_MonthID
               Else
                  cboGeneric(C.ControlIndex).ListIndex = Month(Now)
               End If
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-1" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-3" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-5" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_PRODUCTION_LOCATION)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-6" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-8" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-7" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-9" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " PD-1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_PRODUCTION_LOCATION)
            ElseIf C.ComboLoadID = 2 Or C.ComboLoadID = 3 Or C.ComboLoadID = 4 Or C.ComboLoadID = 5 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_PRODUCTION_TYPE)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " PD-7-1" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_PRODUCTION_LOCATION)
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MS-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMasterTypeName(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitMasterOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MN-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " MS-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSGROUP)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MN-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_POSITION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MN-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_SUPGRADE)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_SUPTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MS-3" Then
            If C.ComboLoadID = 1 Then
              ' Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_EMPPOSITION)
              Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " SL-1-1" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               If m_MonthID > 0 Then
                  cboGeneric(C.ControlIndex).ListIndex = m_MonthID
               Else
                  cboGeneric(C.ControlIndex).ListIndex = Month(Now)
               End If
            End If
         End If
         
         Set Mr = Nothing
      End If
   Next C
   Call EnableForm(Me, True)

End Sub
Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "", Optional KeySearch As String, Optional OldLine As Boolean = False, Optional ToolTipText As String)
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim ChIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChIdx = m_Checks.Count + 1
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
      lblGeneric(LblIdx).ToolTipText = ToolTipText
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
      C.OldLine = OldLine
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
      C.OldLine = OldLine
      txtGeneric(TxtIdx).SetKeySearch (KeySearch)
      
      If Param1 = "YEAR_NO" Then
         If Len(m_YearNo) > 0 Then
            txtGeneric(TxtIdx).Text = m_YearNo
         Else
            txtGeneric(TxtIdx).Text = Year(Now) + 543
         End If
      End If
      
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
      
      
      If DateIdx = 1 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      ElseIf DateIdx = 2 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         End If
      ElseIf DateIdx = 3 Then
         If m_FromRcp > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromRcp
         Else
            Call GetFirstLastDate(Now, m_FromRcp, m_ToRcp)
            uctlGenericDate(DateIdx).ShowDate = m_FromRcp
         End If
      ElseIf DateIdx = 4 Then
         If m_ToRcp > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToRcp)
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         End If
      ElseIf DateIdx = 5 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      End If
      
   ElseIf ControlType = "LU" Then
'         Load uctlTextLookup(LkupIdx)
'         Call m_TextLookups.Add(uctlTextLookup(LkupIdx))
'         C.ControlIndex = LkupIdx
   ElseIf ControlType = "CH" Then
      Load chkCommit(ChIdx)
      Call m_Checks.add(chkCommit(ChIdx))
      Call InitCheckBox(chkCommit(ChIdx), TextMsg)
      C.ControlIndex = ChIdx
   End If
   
   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.add(C)
   Set C = Nothing
End Sub

Private Sub InitReport1_1()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���͡����"))

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_2()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͼ����"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���͡����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_3()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͼ����"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim itemcount As Long
Dim QueryFlag As Boolean
   
'   If LastKey = Node.Key Then
'      Exit Sub
'   End If
   
   LastKey = Node.Key
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   
   If Node.Children > 0 Then
      cmdOK.Enabled = False
      Exit Sub
   End If
   
   If MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MAIN_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("COMMISSION_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("TAGET_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   End If
   
   cmdOK.Enabled = True
   
   If Node.Key = ROOT_TREE & " 1-1" Then
      Call InitReport1_1
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitReport1_2
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitReport1_3
   ElseIf Node.Key = ROOT_TREE & " MS-1" Then
      Call InitReportMS_1
  ElseIf Node.Key = ROOT_TREE & " MS-2" Then
      Call InitReportMS_2
  ElseIf Node.Key = ROOT_TREE & " MS-3" Then
      Call InitReportMS_3
  ElseIf Node.Key = ROOT_TREE & " PT-1" Then
      Call InitReportPT
   ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      Call InitReport3_1
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      Call InitReport3_2
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      Call InitReport3_3
  ElseIf Node.Key = ROOT_TREE & " 3-4" Then
      Call InitReport3_4
  ElseIf Node.Key = ROOT_TREE & " A-1-1" Then
      Call InitReportA_1_1
   ElseIf Node.Key = ROOT_TREE & " A-1-2" Then
      Call InitReportA_1_2
   ElseIf Node.Key = ROOT_TREE & " A-1-3" Then
      Call InitReportA_1_3
   ElseIf Node.Key = ROOT_TREE & " A-1-4" Then
      Call InitReportA_1_4
   ElseIf Node.Key = ROOT_TREE & " SL-1-1" Then
      Call InitReportSL_1_1
   ElseIf Node.Key = ROOT_TREE & " S-1-1" Or Node.Key = ROOT_TREE & " S-1-1-1" Or Node.Key = ROOT_TREE & " S-1-1-2" Or Node.Key = ROOT_TREE & " S-1-1-8" Then
      Call InitReportS_1_1
   ElseIf Node.Key = ROOT_TREE & " S-1-1-7" Then
      Call InitReportS_1_1_7
    ElseIf Node.Key = ROOT_TREE & " S-1-1-6" Then
      Call InitReportS_1_1_6
    ElseIf Node.Key = ROOT_TREE & " S-1-1-3" Then     ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
      Call InitReportS_1_3
     ElseIf Node.Key = ROOT_TREE & " S-1-1-4" Then     ' pui ���� ����Ѻ �Ŵ˹���繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->�Ŵ˹���繪ش(���������)    CReportNormalRcp001_4
      Call InitReportS_1_4
    ElseIf Node.Key = ROOT_TREE & " S-1-1-5" Then     ' pui ���� ����Ѻ ���ػ�ҧ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->���ػ�ҧ����繪ش(���������)    CReportNormalRcp001_5
      Call InitReportS_1_5
   ElseIf Node.Key = ROOT_TREE & " S-1-2" Then
      Call InitReportS_1_2
   ElseIf Node.Key = ROOT_TREE & " S-2-1" Then
      Call InitReportS_2_1
   ElseIf Node.Key = ROOT_TREE & " S-2-2" Then
      Call InitReportS_2_1
   ElseIf Node.Key = ROOT_TREE & " S-2-3" Then
      Call InitReportS_2_3
   ElseIf Node.Key = ROOT_TREE & " S-2-4" Then
      Call InitReportS_2_4_1
   ElseIf Node.Key = ROOT_TREE & " S-2-6" Then
      Call InitReportS_2_12
   ElseIf Node.Key = ROOT_TREE & " S-2-8" Then
      Call InitReportS_2_12
   ElseIf Node.Key = ROOT_TREE & " S-2-9" Then
      Call InitReportS_2_9_1
   ElseIf Node.Key = ROOT_TREE & " S-2-10" Then
      Call InitReportS_2_9
   ElseIf Node.Key = ROOT_TREE & " S-2-10-1" Then
      Call InitReportS_2_9
'   ElseIf Node.Key = ROOT_TREE & " S-2-11" Then
'      Call InitReportS_2_11
'   ElseIf Node.Key = ROOT_TREE & " S-2-11-1" Then
'      Call InitReportS_2_11_2
   ElseIf Node.Key = ROOT_TREE & " S-2-11-2" Then
      Call InitReportS_2_33
  ElseIf Node.Key = ROOT_TREE & " S-2-11-3" Then
      Call InitReportS_2_32
  ElseIf Node.Key = ROOT_TREE & " S-2-11-4" Then
      Call InitReportS_2_33
   ElseIf Node.Key = ROOT_TREE & " S-2-33" Then
      Call InitReportS_2_11_1
   ElseIf Node.Key = ROOT_TREE & " S-2-12" Then
      Call InitReportS_2_12
   ElseIf Node.Key = ROOT_TREE & " S-2-13" Then
      Call InitReportS_2_13
   ElseIf Node.Key = ROOT_TREE & " S-2-14" Then
      Call InitReportS_2_14
   ElseIf Node.Key = ROOT_TREE & " S-2-16" Then
      Call InitReportS_2_16
   ElseIf Node.Key = ROOT_TREE & " S-2-17" Then
      Call InitReportS_2_17
   ElseIf Node.Key = ROOT_TREE & " S-2-17-1" Then
      Call InitReportS_2_17_1
   ElseIf Node.Key = ROOT_TREE & " S-2-18" Then
      Call InitReportS_2_18
   ElseIf Node.Key = ROOT_TREE & " S-2-18-1" Then
      Call InitReportS_2_18
   ElseIf Node.Key = ROOT_TREE & " S-2-19" Then
      Call InitReportS_2_4_1
   ElseIf Node.Key = ROOT_TREE & " S-2-20" Then
      Call InitReportS_2_4_1
   ElseIf Node.Key = ROOT_TREE & " S-2-21" Then '����¹ ����Ѻ  CRBilling021  �¨����� combobox  �����ͧ Sorting by 1.����ѹ��� ����͡���  2. ����͡���
        TEMP_ROOT_TREE = " S-2-21"
         Call InitReportS_2_15
   ElseIf Node.Key = ROOT_TREE & " S-2-22" Then
      TEMP_ROOT_TREE = " S-2-22"
      Call InitReportS_2_15
      TEMP_ROOT_TREE = ""
   ElseIf Node.Key = ROOT_TREE & " S-2-23" Then
      TEMP_ROOT_TREE = " S-2-23"
      Call InitReportS_2_23
      TEMP_ROOT_TREE = ""
   ElseIf Node.Key = ROOT_TREE & " S-2-24" Then
      Call InitReportS_2_24
   ElseIf Node.Key = ROOT_TREE & " S-2-25" Then
      Call InitReportS_2_25
   ElseIf Node.Key = ROOT_TREE & " S-2-27" Then
      Call InitReportS_2_23
   ElseIf Node.Key = ROOT_TREE & " S-2-28" Then
      Call InitReportS_2_28
   ElseIf Node.Key = ROOT_TREE & " S-2-32" Then
      Call InitReportS_2_28_1
   ElseIf Node.Key = ROOT_TREE & " S-2-32-1" Then
      TEMP_ROOT_TREE = " S-2-32-1"
      Call InitReportS_2_28_1
      TEMP_ROOT_TREE = ""
   ElseIf Node.Key = ROOT_TREE & " S-2-34" Then
      Call InitReportS_2_15
   ElseIf Node.Key = ROOT_TREE & " PON-1-1" Then
      Call InitReportS_2_23_1
   ElseIf Node.Key = ROOT_TREE & " PON-1-2" Then
      Call InitReportS_2_23_1
   ElseIf Node.Key = ROOT_TREE & " SL-1-2" Then
      Call InitReportSL_1_2
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-1" Then
      Call InitReportS_2_29
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-2" Then
      Call InitReportS_2_29_1
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-3" Then
      Call InitReportS_2_29_1
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-4" Then
      Call InitReportS_2_29_2
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-5" Then
      Call InitReportS_2_29_2
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-6" Then
      Call InitReportS_2_29_3
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-7" Then
      Call InitReportS_2_29_4
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-8" Then
      Call InitReportS_2_29_5
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-8-1" Then
      TEMP_ROOT_TREE = " SL-1-2-8-1"
      Call InitReportS_2_29_5
      TEMP_ROOT_TREE = ""
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-9" Then
      Call InitReportS_2_29_6
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-10" Then
      Call InitReportS_2_29_6
   ElseIf Node.Key = ROOT_TREE & " SL-1-2-11" Then
      Call InitReportS_2_29_7
   ElseIf Node.Key = ROOT_TREE & " S-2-30" Then
      Call InitReportS_2_28
   ElseIf Node.Key = ROOT_TREE & " S-2-31" Then
      TEMP_ROOT_TREE = " S-2-31"
      Call InitReportS_2_31
      TEMP_ROOT_TREE = ""
  ElseIf Node.Key = ROOT_TREE & " S-2-31-2" Then
      TEMP_ROOT_TREE = " S-2-31-2"
      Call InitReportS_2_31
      TEMP_ROOT_TREE = ""
   ElseIf Node.Key = ROOT_TREE & " S-2-31-1" Then
      Call InitReportS_2_31
   ElseIf Node.Key = ROOT_TREE & " P-1-1" Then
      Call InitReportS_2_4
   ElseIf Node.Key = ROOT_TREE & " P-1-2" Then
      Call InitReportS_2_4
   ElseIf Node.Key = ROOT_TREE & " P-1-3" Then
      Call InitReportP_1_3
   ElseIf Node.Key = ROOT_TREE & " P-1-3-1" Then
      Call InitReportP_1_3
   ElseIf Node.Key = ROOT_TREE & " P-1-3-5" Then
      Call InitReportP_1_3_5
   ElseIf Node.Key = ROOT_TREE & " P-1-4" Then
      Call InitReportP_1_3
   ElseIf Node.Key = ROOT_TREE & " P-1-5" Then
      Call InitReportP_1_5
   ElseIf Node.Key = ROOT_TREE & " P-1-6" Then
      Call InitReportP_1_6
   ElseIf Node.Key = ROOT_TREE & " T-1-1" Then
      Call InitReportT_1_1
   ElseIf Node.Key = ROOT_TREE & " D-1-1" Then
      Call InitReportD_1_1
   ElseIf Node.Key = ROOT_TREE & " D-1-2" Then
      Call InitReportD_1_2
   ElseIf Node.Key = ROOT_TREE & " D-1-2-2" Then
      Call InitReportD_1_2
   ElseIf Node.Key = ROOT_TREE & " D-1-4" Then
      Call InitReportD_1_4
   ElseIf Node.Key = ROOT_TREE & " D-1-5" Then
      Call InitReportD_1_1
   ElseIf Node.Key = ROOT_TREE & " D-1-6" Then
      Call InitReportD_1_4
   ElseIf Node.Key = ROOT_TREE & " R-1-2" Then
      Call InitReportR_1_1
   ElseIf Node.Key = ROOT_TREE & " R-1-4" Then
      Call InitReportR_1_4
   ElseIf Node.Key = ROOT_TREE & " R-1-6" Then
      Call InitReportR_1_4
   ElseIf Node.Key = ROOT_TREE & " 6-1-1" Then
      Call InitReport6_1
   ElseIf Node.Key = ROOT_TREE & " 6-2" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-2-1" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-2-2" Then
      Call InitReport6_2_2
   ElseIf Node.Key = ROOT_TREE & " 6-3" Or Node.Key = ROOT_TREE & " 6-3-1" Then
      Call InitReport6_3
   ElseIf Node.Key = ROOT_TREE & " 6-3-2" Then
      Call InitReport6_3_1
   ElseIf Node.Key = ROOT_TREE & " 6-3-3" Then
      Call InitReport6_3_2
   ElseIf Node.Key = ROOT_TREE & " 6-3-4" Then
      Call InitReport6_3_3
   ElseIf Node.Key = ROOT_TREE & " 6-3-5" Then
      Call InitReport6_3
   ElseIf Node.Key = ROOT_TREE & " 6-4" Then
      Call InitReport6_4
   ElseIf Node.Key = ROOT_TREE & " 6-4-1" Then
      Call InitReport6_4_1
   ElseIf Node.Key = ROOT_TREE & " 6-5" Then
      Call InitReport6_5
      '--------------------�¡ Node.Key = ROOT_TREE & " 6-8"  ���ҧ   Call InitReport6_8 ����-------
   ElseIf Node.Key = ROOT_TREE & " 6-6" Then
      Call InitReport6_6
  ElseIf Node.Key = ROOT_TREE & " 6-8" Then
      Call InitReport6_8
   ElseIf Node.Key = ROOT_TREE & " 6-8-1" Then
      Call InitReport6_8_1
   ElseIf Node.Key = ROOT_TREE & " 6-6-1" Then 'Or Node.Key = ROOT_TREE & " 6-6-2"
      Call InitReport6_6_1
  ElseIf Node.Key = ROOT_TREE & " 6-6-2" Then
      Call InitReport6_2_2_2
   ElseIf Node.Key = ROOT_TREE & " 6-6-3" Then 'For CReportInventory6_3  ��§ҹ�ʹ���ԡ�ѵ�شԺ �¡�����͹(ST009.1) ____User :P'��� QMC ___ by pui
      Call InitReport6_6_3
   ElseIf Node.Key = ROOT_TREE & " 6-7" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-10" Then
      Call InitReport6_10
   ElseIf Node.Key = ROOT_TREE & " 6-9" Then
      Call InitReport6_10
   ElseIf Node.Key = ROOT_TREE & " 7-1" Then
      Call InitReport7_1
   ElseIf Node.Key = ROOT_TREE & " 7-2" Then
      Call InitReport7_2
   ElseIf Node.Key = ROOT_TREE & " 7-3" Then
      Call InitReport7_3
   ElseIf Node.Key = ROOT_TREE & " 8-1" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-2" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-3" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-4" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-5" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-6" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-7" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-8" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " PD-1-1" Then
      Call InitReportPD_1_1
   ElseIf Node.Key = ROOT_TREE & " PD-1-2" Then
      Call InitReportPD_1_2
   ElseIf Node.Key = ROOT_TREE & " PD-1-3" Then
      Call InitReportPD_1_2
   ElseIf Node.Key = ROOT_TREE & " PD-1-4" Then
      Call InitReportPD_1_4
   ElseIf Node.Key = ROOT_TREE & " PD-1-5" Then
      Call InitReportPD_1_5
   ElseIf Node.Key = ROOT_TREE & " PD-1-6" Then
      Call InitReportPD_1_4
   ElseIf Node.Key = ROOT_TREE & " PD-1-7" Then
      Call InitReportPD_1_7
   ElseIf Node.Key = ROOT_TREE & " PD-1-8" Then
      Call InitReportPD_1_8
   ElseIf Node.Key = ROOT_TREE & " PD-1-9" Then
      Call InitReportPD_1_9
   ElseIf Node.Key = ROOT_TREE & " PD-7-1" Then
      Call InitReportPD_7_1
   ElseIf Node.Key = ROOT_TREE & " B-1-1" Then
      Call InitReportB_1_1
   ElseIf Node.Key = ROOT_TREE & " B-1-2" Then
      Call InitReportB_1_1
   ElseIf Node.Key = ROOT_TREE & " B-1-3" Then
      Call InitReportB_1_3
   ElseIf Node.Key = ROOT_TREE & " B-1-4" Then
      Call InitReportB_1_4
   ElseIf Node.Key = ROOT_TREE & " MN-1" Then
      Call InitReportMN_1
   ElseIf Node.Key = ROOT_TREE & " MN-2" Then
      Call InitReportMN_2
   ElseIf Node.Key = ROOT_TREE & " MN-3" Then
      Call InitReportMN_3
  End If
End Sub

Private Sub InitReportA_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���������˹��"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportA_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ�����"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١˹��"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportA_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ�����"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���������˹��"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportA_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ�����"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١˹��"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReport3_1()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_2()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ���������"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͫѾ���������"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Ѿ �"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_3()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ�ž�ѡ�ҹ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_4()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
   '1 =============================
'   Call LoadControl("LU", uctlTextLookup(0).Width, True, "", 0, "")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͺ���ѷ (˹��§ҹ)"))
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 0, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("  Ẻ������������"))
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
 '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   '5 =============================
    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '64 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub InitReportS_2_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STOCK_GROUP", "STOCK_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������Թ���"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
      
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��ӹǹ", , "SHOW_AMOUNT")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReport6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("� �ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '6 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2_2_2()  'Pui ����
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_6_3()  '���� For  CReportInventoryDoc6_2_3 ��§ҹ�ʹ�ԡ�ѵ�شԺ �¡�����͹(ST009.1)__user: P'��� QMC__By Pui
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
   End Sub
Private Sub InitReport6_6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_8() 'Pui ����
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��鹷ع", , "SHOW_COSTS")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��鹷ع", , "SHOW_COSTS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�˹���", , "SHOW_UNIT_NAME")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѧ"))
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѧ"))
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))
   
   '2 =============================

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѧ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
'   '8 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
'
'   '9 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "FROM_LOCATION_ID", "FROM_LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FROM_LOCATION_NO", "FROM_LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��ѧ"))
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "TO_LOCATION_ID", "TO_LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TO_LOCATION_NO", "TO_LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
  Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", "", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID2", "", "LOCATION_NO", True)
   
   '7 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", "", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("令�ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID2", "", "LOCATION_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ҡ���", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "INVENTORY_SUB_TYPE", "INVENTORY_SUB_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������͡�������"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_6()  'Pui ����
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ȹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportS_2_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportP_1_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_PO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ������觫���"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�����觫���", , "SHOW_PO_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�ú�Ţ��", , "SHOW_INV_DETAIL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ҡ���", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_9_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������١���", , "SHOW_CUSTOMER_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ҡ���", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��ʹ���Թ", , "SHOW_RCP")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   
   If TEMP_ROOT_TREE = " S-2-22" Then
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   End If
   
    
    
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   If TEMP_ROOT_TREE = " S-2-21" Then ' ���� combobox Sorting By ����Ѻ ��¡�â������ CReportBilling021  �������§��� 1.����ѹ��� ����Ţ��� 2.����ѹ���
'    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_TYPE", "ORDER_TYPE_NAME")
'    Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_23()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ������º��º"))
      
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������º��º"))
   Else
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������º��º"))
   End If
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_FROM_STOCK_NO", , "NOT_STOCK_NO")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������ѵ�شԺ"))
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_TO_STOCK_NO", , "NOT_STOCK_NO", True)
   End If
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���1"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "STOCK_TYPE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´�¡������", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡���Ѻ�׹", , "INCLUDE_RETURN")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ�ط��", , "SUMMARY_NET")
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ҡ���", , "CONSIGNMENT")
      
      Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ�ҡ���", , "SUMMARY_CONSIGNMENT")
      Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ�ҡ����ط��", , "SUMMARY_CONSIGNMENT_NET")
      
      Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
   End If
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_23_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ������º��º"))
      
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������º��º"))
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ҡ���", , "CONSIGNMENT")
   '��§ҹ���2
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ѹ��ش Repost2"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ѹ��ش Repost3"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE2", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE2", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���2"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO2", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ2"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO2", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE2", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���2"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE2", , "SALE_CODE", True)
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_24()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������º��º"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_25()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������º��º"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReport7_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReport7_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))

   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ѹ��ش"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN", , "NOT_SHOW_RETURN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��������ҧ�дѺ Sale", , "SHOW_COLOR")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReport7_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))

   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʵ��᷹"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN", , "NOT_SHOW_RETURN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��������ҧ�дѺ Sale", , "SHOW_COLOR")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportT_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportD_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportD_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´�������ҧ��0", , "SHOW_DETAIL_ZERO")
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
      '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١˹��"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾������͹�����ͺ�ѹ", , "SHOW_ONLY_MOVE")
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_MOVE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������͹���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_MOVE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������͹���"))
      
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportD_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���ú��˹�"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����ú��˹�"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 1, "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
         
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 1, "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
      
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_16()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ѹ��ش"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)

   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReport8_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportR_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���ǹ������ǹ�ѡ(Ẻ�¡)", , "SHOW_DETAIL_ADDSUB")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´����", , "SHOW_PAYMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��Ţ���ѭ�ա���͹", , "SHOW_ACCOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�ê���", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���ػ", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportR_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���ǹ������ǹ�ѡ(Ẻ�¡)", , "SHOW_DETAIL_ADDSUB")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´����", , "SHOW_PAYMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��Ţ���ѭ�ա���͹", , "SHOW_ACCOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�ê���", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportPD_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportPD_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ�Ե�ѳ��"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� %", , "SHOW_PERCENT")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportPD_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��§ҹ����ѹ", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾����ػ", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportPD_1_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ�Ե�ѳ��"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��§ҹ����ѹ", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾����ػ", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportPD_1_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ�Ե�ѳ��"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportPD_1_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ�Ե�ѳ��"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��§ҹ����ѹ", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾����ػ", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportPD_1_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "SUB_FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ǧ�ҡ(�ѹ)"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "ADD_TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ǧ�֧(�ѹ)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ�Ե�ѳ��"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ʴ���������´4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��§ҹ����ѹ", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾����ػ", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportPD_7_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ҡ�ü�Ե��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ҡ�ü�Ե��"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ����Ե"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_28()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���1"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "STOCK_TYPE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�¡�������Թ���3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� GP", , "SHOW_GP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_28_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_FROM_STOCK_NO", , "NOT_STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_TO_STOCK_NO", , "NOT_STOCK_NO", True)
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �鹷ع", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� GP", , "SHOW_GP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   If TEMP_ROOT_TREE <> " S-2-32-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   If TEMP_ROOT_TREE <> " S-2-32-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "����¡�Ң��١���", , "SHOW_GROUP_CUSTOMER")
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportSL_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������˵�", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 3, "BIRTH_MONTH_ID", "BIRTH_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))
   
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_29()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������˵�", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
      
      
      
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�ѹ��ش"))
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������˵�", , "NOTE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�੾�Ъ��� Sale", , "SHOW_SALE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
   
   '���º��º
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE_COMPARE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ������º��º"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE_COMPARE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ������º��º(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������˵�", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ��Թ���", , "NOT_SHOW_CARGO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ��١��� ����Թ���", , "NOT_SHOW_CUSTOMERS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�SALE, �١��� ����Թ���", , "NOT_SHOW_SALE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ��Թ���", , "NOT_SHOW_CARGO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�������Թ��� ����Թ���", , "NOT_SHOW_CARGO_GROUP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�SALE, ������Թ��� ����Թ���", , "NOT_SHOW_SALE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������ CN *੾�� SALE ����˹� ", , "NOT_SHOW_RETURN")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_29_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ� �ʹ���", , "SHOW_PRICE")

   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ��١���", , "NOT_SHOW_CUSTOMERS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������١������᷹�������", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
      
  If TEMP_ROOT_TREE = " SL-1-2-8-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��ȹ���", , "SHOW_DECIMAL")
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������١���", , "SHOW_CUS_CODE")
   End If
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportB_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ���������"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
    
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportB_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ���������"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
      
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
    
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_11_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ���(7,7,7,7,2)"))
      
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_32()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
    Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReportS_2_33()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ����"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ú��˹�(60<,30,30)"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Թ��˹�(7,8,15,>30)"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
    Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_EXCEL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub


Private Sub InitReportS_2_11_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")

   Call LoadControl("CH", cboGeneric(0).Width, True, "�������ѧ���", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
         
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ��ӹǹ", , "NOT_SHOW_AMOUNT")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_31()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   If TEMP_ROOT_TREE = " S-2-31" Or TEMP_ROOT_TREE = " S-2-31-2" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   End If
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReport6_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_GROUP", "STOCK_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������Թ���/�ѵ�شԺ"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE", "STOCK_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���/�ѵ�شԺ"))
   
   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "¡��ԡ", , "EXCEPTION_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportP_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ����觢ͧ"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_CHANGE1", "UNIT_CHANGE_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���1"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "UNIT_CHANGE2", "UNIT_CHANGE_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "UNIT_CHANGE3", "UNIT_CHANGE_NAME3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���3"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "UNIT_CHANGE4", "UNIT_CHANGE_NAME4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���4"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "UNIT_CHANGE5", "UNIT_CHANGE_NAME5")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���5"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "UNIT_CHANGE6", "UNIT_CHANGE_NAME6")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���6"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "UNIT_CHANGE7", "UNIT_CHANGE_NAME7")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���7"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "UNIT_CHANGE8", "UNIT_CHANGE_NAME8")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���8"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 9, "UNIT_CHANGE9", "UNIT_CHANGE_NAME9")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���9"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 10, "UNIT_CHANGE10", "UNIT_CHANGE_NAME10")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹���10"))
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportP_1_3_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))   '---------------------------------------------------
         
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ����", , "NOT_SHOW_BILL_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportPT()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_AREA", "MASTER_AREA_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ǣ�͢�������ѡ"))
'
'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMS_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_AREA", "MASTER_AREA_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ǣ�͢�������ѡ"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMN_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������١˹��"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١˹��"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ����", , "NOT_SHOW_HEAD_FLAG")
   
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportMN_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
    '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "EMP_NAME", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
    '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "EMP_LASTNAME", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ��"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportMN_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "APAR_CODE", , "APAR_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
    '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "APAR_NAME", , "APAR_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͼ����"))
   
    '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_SUPGRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "MASTER_SUPTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��Ţ�����������", , "SHOW_TAX")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ��������", , "SHOW_ADDRESS")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportMS_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������١˹��"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١˹��"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMS_3() '��ѡ�ҹ��� sale
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub cboGeneric_Click(Index As Integer)
Dim Node As Node
Dim TempID As Long

   Set Node = trvMaster.SelectedItem
   
   If (Node.Key = ROOT_TREE & " MN-1") Then
      If Index = 1 Then
         TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
         If TempID > 0 Then
            Call LoadMaster(cboGeneric(Index + 1), , , , MASTER_CUSTYPE, , TempID)
         End If
      End If
   End If
End Sub
Private Sub InitReportS_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_SEARCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���Ѻ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����")) ''''''''''''''
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_1_1_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_SEARCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_FROM")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_TO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���Ѻ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����")) ''''''''''''''
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_1_1_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
'   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
  '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))


   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_1_3() ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���Ѻ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
  Call LoadComboData
   
End Sub
Private Sub InitReportS_1_4() ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���Ѻ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
  Call LoadComboData
   
End Sub

Private Sub InitReportS_1_5() ' pui ���� ����Ѻ 㺤׹�Թ����繪ش(���������) -----------path------------  �к��ѭ����С���Թ --- >��§ҹ�к��ѭ����С���Թ----->㺤׹�Թ����繪ش(���������)    CReportNormalRcp001_3
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�͡���"), , , , , , "������ҧ ���ʴ HS,��觢ͧ IV55,㺡ӡѺ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���Ѻ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
  Call LoadComboData
   
End Sub
Private Sub InitReportSL_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ҡ�â����͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ҡ�â�»�"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ������º��º"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������º��º"))
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ӹǹ", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʹ���", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ӹǹ�׹", , "SHOW_RETURN_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʹ�׹", , "SHOW_RETURN_PRICE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportP_1_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TRANSPORTOR1", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����1"))
      
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR1_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��¨��¢���1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR2", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����2"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR2_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��¨��¢���2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "TRANSPORTOR3", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����3"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR3_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��¨��¢���3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�Դ�ʹ�ҡ� PO", , "PO_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�Դ�ʹ�ҡ��觢ͧ��Т��ʴ", , "INVOICE_FLAG")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_17_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ���"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportB_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ���������"))
         
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����¡�âͧ��", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub

Private Sub InitReport6_8_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ѧ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO_PERCENT", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ%"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO_PERCENT", , "STOCK_NO", True)
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
  Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub

