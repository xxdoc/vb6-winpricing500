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
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-1", MapText("รายงานข้อมูลหลัก"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-2", MapText("รายงานข้อมูลลูกค้าและสาขา"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-3", MapText("รายงานข้อมูลพนักงานขายและสาขา"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PT-1", MapText("รายงานข้อมูลเครื่องพิมพ์"), 1, 2)
      Node.Expanded = True
      
      
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-1", MapText("รายงานข้อมูลลูกค้า"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-2", MapText("รายงานข้อมูลพนักงาน"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-3", MapText("รายงานข้อมูลซัพพลายเออร์"), 1, 2)
      Node.Expanded = True

   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-1", MapText("รายงานการผลิตประจำวัน"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-1", tvwChild, ROOT_TREE & " PD-1-1", MapText("รายงานใบผลิตประจำวัน(PD001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-2", MapText("รายงานการผลิตตามการผลิตจริง"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-2", tvwChild, ROOT_TREE & " PD-1-4", MapText("รายงาน%ยอดผลิตตามLOTผลิต  รวมตามวัตถุดิบ วันที่LOT ประเภทผลิต (PD002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-2", tvwChild, ROOT_TREE & " PD-1-6", MapText("รายงาน%ยอดผลิตตามจำนวนผลิต รวมตามวัตถุดิบ วันที่LOT ประเภทผลิต (PD003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-3", MapText("รายงานการผลิตตามวันที่ผลิต"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-2", MapText("รายงาน%ยอดผลิต รวมตามวันที่ผลิต (PD004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-3", MapText("รายงาน%ยอดผลิต รวมตามวันที่ผลิต ประเภทผลิต (PD005)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-3", tvwChild, ROOT_TREE & " PD-1-5", MapText("รายงาน%ยอดผลิต รวมตามวัตถุดิบ วันที่ผลิต ประเภทผลิต (PD006)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-4", MapText("รายงานการผลิตตามวันที่ LOTนำเข้า"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-4", tvwChild, ROOT_TREE & " PD-1-8", MapText("รายงาน%ยอดผลิตตามจำนวนผลิต  รวมตามวัตถุดิบ วันที่LOT (PD007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-5", MapText("รายงานการผลิตตามหมายเลขผลิต"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-5", tvwChild, ROOT_TREE & " PD-1-7", MapText("รายงาน%ยอดผลิตตามจำนวนผลิต  รวมตามหมายเลขผลิต ประเภทผลิต (PD008)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-6", MapText("รายงานการผลิตตามสถานที่ผลิตหลัก"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-6", tvwChild, ROOT_TREE & " PD-1-9", MapText("รายงาน%ยอดผลิตตามจำนวนผลิต  รวมตามวัตถุดิบ สถานที่ผลิตหลัก (PD009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-7", MapText("รายงานเป้าการผลิต"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PD-7", tvwChild, ROOT_TREE & " PD-7-1", MapText("รายงานยอดผลิตเปรียบเทียบเป้าการผลิตตาม วันที่ผลิต (PD010)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-1", MapText("ระบบข้อมูลการเงิน"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-1", MapText("รายงานทะเบียนเช็คจ่าย CHQ001"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-2", MapText("รายงานทะเบียนเช็ครับ CHQ002"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-3", MapText("รายงานทะเบียนเช็คจ่ายตามเจ้าหนี้ CHQ003"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " A-1-4", MapText("รายงานทะเบียนเช็ครับตามลูกหนี้ CHQ004"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PON-1", MapText("สำหรับ คุณอรุณพร"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PON-1", tvwChild, ROOT_TREE & " PON-1-1", MapText("รายงานเปรียบเทียบยอดขาย แยกตามกลุ่มสินค้า รวมบริษัท"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " PON-1", tvwChild, ROOT_TREE & " PON-1-2", MapText("รายงานเปรียบเทียบยอดขาย แยกตามกลุ่มสินค้า รวมบริษัท (เฉพาะสิ้นเดือน)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " SL-1", MapText("สำหรับพนักงานขาย"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-1", MapText("รายงานยอดขายเปรียบเทียบ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2", MapText("รายงานยอดขาย แยกตามพนักงานขาย ลูกค้า สินค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-1", MapText("รายงานยอดขาย แยกตามพนักงานขาย สินค้า ลูกค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-2", MapText("รายงานยอดขาย แยกตามพนักงานขาย ลูกค้า สินค้า รวมเป็นงวด รวมตามกลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-3", MapText("รายงานยอดขาย แยกตามพนักงานขาย สินค้า ลูกค้า รวมเป็นงวด รวมตามกลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-4", MapText("รายงานเปรียบเทียบยอดขาย แยกตามพนักงานขาย ลูกค้า สินค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-5", MapText("รายงานเปรียบเทียบยอดขาย แยกตามพนักงานขาย สินค้า ลูกค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-6", MapText("รายงานยอดขาย แยกตามกลุ่มลูกค้า พนักงานขาย ลูกค้า สินค้า รวมตามกลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-7", MapText("รายงานยอดขาย แยกตามกลุ่มลูกค้า พนักงานขาย กลุ่มสินค้า สินค้า รวมตามกลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-8", MapText("รายงานยอดขาย แยกตามลูกค้า กลุ่มสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-8-1", MapText("รายงานยอดขาย แยกตามลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-9", MapText("รายงานการจัดลำดับยอดขาย ตามสินค้า"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-10", MapText("รายงานการจัดลำดับยอดขาย ตามลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " SL-1", tvwChild, ROOT_TREE & " SL-1-2-11", MapText("รายงานเปรียบเทียบเป้าการขาย ยอดขาย แยกตามพนักงานขาย รายวัน สะสม"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " ST-1", MapText("ระบบงานขาย TOP"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-1", MapText("รายงานยอดขาย TOP ทุกสินค้า โดยแยกตามสาขา"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-2", MapText("รายงานยอดขาย TOP แยกตามสินค้าและสาขา"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " ST-1", tvwChild, ROOT_TREE & " S-2-3", MapText("รายงานยอดขาย TOP แยกตามพนักงานขายและสาขา"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " S-1", MapText("ระบบงานขาย"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-1", MapText("รายงาน PO ตามวันที่เอกสาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-2", MapText("รายงาน PO ตามวันที่ส่งของ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3", MapText("รายงานการเบิกสินค้าแยกตาม วันที่ คนขับ ทะเบียน ขนส่ง (ตามใบ PO)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3-1", MapText("รายงานการเบิกสินค้าสรุปตาม วันที่ คนขับ ทะเบียน ขนส่ง (ตามใบ PO)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-4", MapText("รายงานการเบิกสินค้าแยกตาม วันที่ คนขับ ทะเบียน ขนส่ง (ตามใบ ส่งของและ ขายสด)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-3-5", MapText("รายงานรายละเอียด วันที่ คนขับ ทะเบียน ขนส่ง (ตามใบส่งของ)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-5", MapText("รายงานหน่วยขนส่งสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " P-1-6", MapText("รายงาน PO ค้างส่ง"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1", MapText("ใบส่งสินค้าเป็นชุด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-8", MapText("ใบส่งสินค้าเป็นชุด (1/2 Letter)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-2", MapText("ใบส่งสินค้าเป็นชุด (พิมพ์ฟอร์มเปล่า)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-7", MapText("พิมพ์ลาเบล ที่จัดส่งสินค้าเป็นชุด"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-3", MapText("ใบรับคืนสินค้าเป็นชุด (พิมพ์ฟอร์มเปล่า)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-4", MapText("ใบลดหนี้เป็นชุด (พิมพ์ฟอร์มเปล่า)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-5", MapText("ใบสรุปวางบิลเป็นชุด (พิมพ์ฟอร์มเปล่า)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-1", MapText("ใบฝากขาย/ใบส่งสินค้าชั่วคราวเป็นชุด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-6", MapText("รายงานพิมพ์ลาเบล"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-2", MapText("รายงานยอดขาย ตาม รหัสสินค้า วันที่ขาย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-4", MapText("รายงานเอกสารแพ็คตั๋ว"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-9", MapText("รายงานประวัติการขาย แยกตามสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-10", MapText("รายงานประวัติการขาย แยกตามลูกค้า สินค้า เอกสาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-10-1", MapText("รายงานประวัติการขาย แยกตามลูกค้า เอกสาร สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-12", MapText("รายงานยอดขายประจำงวด (ยอดขายก่อนคิด VAT)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-13", MapText("รายงานสรุปยอดขายเชื่อแยกตามลูกหนี้"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-14", MapText("รายงานสรุปยอดขาย แจกแจงเป็นงวด แยกตามหมวดสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-22", MapText("รายงานขายเงินสด เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
     
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-21", MapText("รายงานขายเชื่อ"), 1, 2)
      Node.Expanded = False
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-21", MapText("รายงานขายเชื่อ "), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-34", MapText("รายงานฝากขาย เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-16", MapText("รายงานการจัดลำดับยอดขาย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-17", MapText("รายงานสรุปยอดขาย แจกแจงเป็นงวด แยกตามลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-17-1", MapText("รายงานสรุปยอดขาย แจกแจงเป็นงวด แยกตาม เขตการขาย ลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-18", MapText("รายงานสรุปยอดขาย แยกตามหมวดสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-18-1", MapText("รายงานสรุปยอดขาย แยกตามหมวดสินค้า ประเภทสินค้าย่อย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-19", MapText("รายงานสรุปยอดขาย แยกตามลูกค้า"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-20", MapText("รายงานสถานะการส่งสินค้า แยกตามลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-23", MapText("รายงานเปรียบเทียบยอดขาย แยกตามกลุ่มลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-24", MapText("รายงานเปรียบเทียบยอดขาย แยกตามประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-25", MapText("รายงานเปรียบเทียบยอดขาย แยกตามสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-27", MapText("รายงานเปรียบเทียบยอดขาย แยกตามกลุ่มลูกค้า ประเภทเอกสาร ประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-28", MapText("รายงานยอดขาย แยกตามกลุ่มลูกค้า ประเภทเอกสาร ประเภทลูกค้า กลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-30", MapText("รายงานยอดขาย แยกตามกลุ่มลูกค้า ประเภทเอกสาร กลุ่มวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-32", MapText("รายงานยอดขาย แยกตามพนักงานขาย ลูกค้า สินค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-32-1", MapText("รายงานยอดขาย แยกตามพนักงานขาย สินค้า รวมเป็นงวด"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " T-1", MapText("ระบบภาษี"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " T-1", tvwChild, ROOT_TREE & " T-1-1", MapText("รายงานภาษีขาย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " D-1", MapText("ระบบลูกหนี้"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-1", MapText("รายงานสถานะลูกหนี้ (AR001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-2", MapText("รายงานลูกหนี้คงค้าง (AR002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-2-2", MapText("รายงานลูกหนี้คงค้าง เรียงตามพนักงานขาย(AR002.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-5", MapText("รายงานการเคลื่อนไหวลูกหนี้ (AR003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-4", MapText("รายงานใบกำกับที่ถึงกำหนดเก็บเงิน แยกตามลูกค้า (AR004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " D-1-6", MapText("รายงานใบกำกับที่ถึงกำหนดเก็บเงิน แยกตามลูกค้า วันที่ (AR005)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามลูกค้า"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-1", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามพนักงานขาย ลูกค้า"), 1, 2)
'      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-2", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามลูกค้า (AR006)"), 1, 2)
      Node.Expanded = False
    
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-3", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามพนักงานขาย ลูกค้า  (AR007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-11-4", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามพนักงานขาย ลูกค้า (ระบุช่วงวันที่) (AR008)"), 1, 2)
      Node.Expanded = False
       
            
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-6", MapText("รายงานใบลดหนี้/รับคืนสินค้า แยกตามลูกค้า  (AR009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-8", MapText("รายงานใบลดหนี้/รับคืนสินค้า เรียงตามเขตการขาย,พนักงานขาย  (AR010)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31", MapText("รายงานใบลดหนี้/รับคืนสินค้า เรียงตามวันที่ (AR011)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31-2", MapText("รายงานใบลดหนี้/รับคืนสินค้า เรียงตามสินค้า (AR011.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-31-1", MapText("รายงานใบลดหนี้ เรียงตามวันที่ (AR012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " D-1", tvwChild, ROOT_TREE & " S-2-33", MapText("รายงานวิเคราะห์อายุลูกหนี้ แยกตามพนักงานขาย ลูกค้า ประเภทลูกค้า (AR013)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " R-1", MapText("ระบบการรับชำระหนี้"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-2", MapText("รายงานรับชำระแยกตามลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-4", MapText("รายงานรับชำระเรียงตามวันที่ใบเสร็จ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " R-1", tvwChild, ROOT_TREE & " R-1-6", MapText("รายงานรับชำระแยกตามพนักงานขาย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " B-1", MapText("ระบบการซื้อ"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-1", MapText("รายงานประวัติการซื้อ แยกตามสินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-2", MapText("รายงานประวัติการซื้อ แยกตามซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-3", MapText("รายงานใบรับสินค้า เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " B-1", tvwChild, ROOT_TREE & " B-1-4", MapText("รายงานสรุปยอดซื้อ แจกแจงเป็นงวด แยกตามซัพพลายเออร์ สินค้า"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-10", MapText("รายงานรหัสสินค้าและวัตถุดิบ (ST001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-1-1", MapText("รายงานจำนวนคงคลังและยอดตรวจนับ (ST002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2", MapText("รายงานการเคลื่อนไหวคลัง แบบ 1 (ST003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-1", MapText("รายงานการเคลื่อนไหวคลัง แบบ 2 (ST004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-2", MapText("รายงานการเคลื่อนไหวสินค้า แยกตามคลังสินค้า แจกแจงวันที่ (ST004-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3", MapText("รายงานสรุปยอดเคลื่อนไหวสินค้า แยกตามคลังสินค้า (ST005)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-1", MapText("รายงานสรุปยอดสินค้า แยกตามคลังสินค้า (ST005-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-2", MapText("รายงานสรุปยอดสินค้า เรียงตามคลังสินค้า (ST005-2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-3", MapText("รายงานสรุปยอดสินค้า เรียงตามสินค้า (ST005-3)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-4", MapText("รายงานสรุปยอดสินค้าคงเหลือ เรียงตามคลัง (ST005-4)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3-5", MapText("รายงานสรุปยอดเคลื่อนไหวสินค้า แยกตามกลุ่มคลังสินค้า (ST005-5)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4", MapText("รายงานเอกสารการโอน แยกตามคลังสินค้า (ST006)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4-1", MapText("รายงานเอกสารการโอน และรับคืน  (ST006-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-5", MapText("รายงานเอกสารการเบิก (ST007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6", MapText("รายงานเอกสารการเบิก แจกแจงตามประเภทเอกสารย่อย (ST008)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-2", MapText("รายงานเอกสารการเบิก แจงแจงตามประเภทเบิก (ST008.1)"), 1, 2)
      Node.Expanded = False
      
      
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-1", MapText("รายงานเอกสารการเบิก แจงแจงตามหมายเลขเอกสาร (ST009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-6-3", MapText("รายงานยอดเบิกวัตถุดิบ  แยกตามเดือน (ST009.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-7", MapText("รายงานเอกสารการรับเข้า แยกตามคลังสินค้า (ST010)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-8", MapText("รายงานการผลิต จาก BARCODE (ST011)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-8-1", MapText("รายงาน % การผลิต จาก BARCODE แยกตามวัตถุดิบ (ST012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-9", MapText("พิมพ์บาร์โค๊ด(ST101)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 7 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-1", MapText("รายงานขออนุมัติเบิกจ่ายค่า Commission"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-2", MapText("รายงานเปรียบเทียบยอดขาย แยกตามพนักงานขาย ตามแผนภูมิพนักงานขาย"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 7-3", MapText("รายงานการจ่าย Rebate ตามตัวแทนจำหน่าย"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 8 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-1", MapText("รายงานเป้าการขาย แยกตาม เขต พนักงานขาย ลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-2", MapText("รายงานเป้าการขาย แยกตาม เขต พนักงานขาย สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-4", MapText("รายงานเป้าการขาย แยกตาม เขต ลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-5", MapText("รายงานเป้าการขาย แยกตาม เขต สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-6", MapText("รายงานเป้าการขาย แยกตาม พนักงานขาย ลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-7", MapText("รายงานเป้าการขาย แยกตาม พนักงานขาย สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-8", MapText("รายงานเป้าการขาย แยกตาม ลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-3", MapText("รายงานเป้าการขาย แยกตาม ประเภทสินค้า สินค้า"), 1, 2)
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

  ElseIf Key = "Root S-1-1-3" Then 'pui เพิ่ม    ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
      Set Report = New CReportNormalRcp001_3
      ClassName = "CReportNormalRcp001_3"
      
   ElseIf Key = "Root S-1-1-4" Then 'pui เพิ่ม    ' pui เพิ่ม สำหรับ ใบลดหนี้เป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบลดหนี้เป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_4
      Set Report = New CReportNormalRcp001_4
      ClassName = "CReportNormalRcp001_4"
  ElseIf Key = "Root S-1-1-5" Then 'pui เพิ่ม    ' pui เพิ่ม สำหรับ ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_5
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
      Call Report.AddParam(1, "AREA") 'ฝั่งขาย
   ElseIf Key = "Root S-2-22" Then
      Set Report = New CReportBilling022
      ClassName = "CReportBilling022"
   ElseIf Key = "Root S-2-23" Then
      Set Report = New CReportBilling030
      ClassName = "CReportBilling030"
  ElseIf Key = "Root S-2-34" Then
      Set Report = New CReportBilling042
      ClassName = "CReportBilling042"
      Call Report.AddParam(1, "AREA") 'ฝั่งขาย
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
      
      Call Report.AddParam(2, "AREA") 'ฝั่งซื้อ
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
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
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
   
   Call InitNormalLabel(lblSpace, MapText("ระยะห่าง"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   
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
         
            If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-3" Then     ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
              If C.ComboLoadID = 1 Then
               Call InitReportS_1_3Orderby(m_Combos(C.ControlIndex))
             ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
          End If

          If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-4" Then     ' pui เพิ่ม สำหรับ ใบลดหนี้เป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบลดหนี้เป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_4
            If C.ComboLoadID = 1 Then
               Call InitReportS_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-1-1-5" Then     ' pui เพิ่ม สำหรับ ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_5
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
'-------แยก  6-2 สร้างใหม่
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
                If MASTER_INVENTORY_SUB_TYPE = 34 Then ' เพิ่มในเอกสารการเบิก ST007  เพื่อให้ค่า InventoryDoc_Sub_Type =null (ค่าที่ออกมา เป็น -1) คือ การเบิกที่ไม่ได้เบิกจากDatabase ในที่นี้จะเบิกจาก barcode
                       'Comb.AddItem ("ใบเอกสารเบิกอื่นๆ")
                       m_Combos(C.ControlIndex).AddItem ("ใบเอกสารเบิกอื่นๆ")
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
    ElseIf Node.Key = ROOT_TREE & " S-1-1-3" Then     ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
      Call InitReportS_1_3
     ElseIf Node.Key = ROOT_TREE & " S-1-1-4" Then     ' pui เพิ่ม สำหรับ ใบลดหนี้เป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบลดหนี้เป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_4
      Call InitReportS_1_4
    ElseIf Node.Key = ROOT_TREE & " S-1-1-5" Then     ' pui เพิ่ม สำหรับ ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบสรุปวางบิลเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_5
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
   ElseIf Node.Key = ROOT_TREE & " S-2-21" Then 'เปลี่ยน สำหรับ  CRBilling021  โดยจะเพิ่ม combobox  ในเรื่อง Sorting by 1.ตามวันที่ ตามเอกสาร  2. ตามเอกสาร
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
      '--------------------แยก Node.Key = ROOT_TREE & " 6-8"  สร้าง   Call InitReport6_8 ใหม่-------
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
   ElseIf Node.Key = ROOT_TREE & " 6-6-3" Then 'For CReportInventory6_3  รายงานยอดใช้เบิกวัตถุดิบ แยกตามเดือน(ST009.1) ____User :P'แอล QMC ___ by pui
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เช็ค"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เช็ค"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เช็ค"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้ค้า"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเจ้าหนี้"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เช็ค"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เช็ค"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เช็ค"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เช็ค"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เช็ค"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เช็ค"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้ค้า"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเจ้าหนี้"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เช็ค"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เช็ค"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ดิวเช็ค"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_CHEQUE_Q")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ดิวเช็ค"))

   '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เช็ค"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อซัพพลายเออร์"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทซัพ ฯ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อพนักงาน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("นามสกุลพนักงาน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ตำแหน่ง"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
   '1 =============================
'   Call LoadControl("LU", uctlTextLookup(0).Width, True, "", 0, "")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อบริษัท (หน่วยงาน)"))
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 0, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("  แบบที่ใช้ยื่นภาษี"))
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
 '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   '5 =============================
    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '64 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STOCK_GROUP", "STOCK_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสินค้า"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))

   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงจำนวน", , "SHOW_AMOUNT")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ณ วันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '6 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2_2_2()  'Pui เพิ่ม
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_6_3()  'เพิ่ม For  CReportInventoryDoc6_2_3 รายงานยอดเบิกวัตถุดิบ แยกตามเดือน(ST009.1)__user: P'แอล QMC__By Pui
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
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_8() 'Pui เพิ่ม
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COSTS")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มคลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COSTS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหน่วย", , "SHOW_UNIT_NAME")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มคลัง"))
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มคลัง"))
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))
   
   '2 =============================

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", , "LOCATION_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_GROUP_NO", "LOCATION_GROUP_NAME", "LOCATION_GROUP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มคลัง"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
'   '8 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '9 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "FROM_LOCATION_ID", "FROM_LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FROM_LOCATION_NO", "FROM_LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากคลัง"))
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "TO_LOCATION_ID", "TO_LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TO_LOCATION_NO", "TO_LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไปยังคลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
  Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID", "", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากคลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_ID2", "", "LOCATION_NO", True)
   
   '7 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID", "", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไปคลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_ID2", "", "LOCATION_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ฝากขาย", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "INVENTORY_SUB_TYPE", "INVENTORY_SUB_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเอกสารย่อย"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_6()  'Pui เพิ่ม
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_PO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่ใบสั่งซื้อ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายการใบสั่งซื้อ", , "SHOW_PO_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายการบิลขาย", , "SHOW_INV_DETAIL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ฝากขาย", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงชื่อลูกค้า", , "SHOW_CUSTOMER_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ฝากขาย", , "CONSIGNMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงยอดเก็บเงิน", , "SHOW_RCP")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   
   If TEMP_ROOT_TREE = " S-2-22" Then
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   End If
   
    
    
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
   
   If TEMP_ROOT_TREE = " S-2-21" Then ' เพิ่ม combobox Sorting By สำหรับ รายการขายเชื่อ CReportBilling021  ที่จะเรียงตาม 1.ตามวันที่ ตามเลขที่ 2.ตามวันที่
'    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_TYPE", "ORDER_TYPE_NAME")
'    Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เปรียบเทียบ"))
      
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เปรียบเทียบ"))
   Else
      Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่เปรียบเทียบ"))
   End If
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_FROM_STOCK_NO", , "NOT_STOCK_NO")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมรหัสวัตถุดิบ"))
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_TO_STOCK_NO", , "NOT_STOCK_NO", True)
   End If
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า1"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "STOCK_TYPE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียดแยกประเภท", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายการรับคืน", , "INCLUDE_RETURN")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
   If TEMP_ROOT_TREE = " S-2-23" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "สรุปสุทธิ", , "SUMMARY_NET")
      Call LoadControl("CH", cboGeneric(0).Width, True, "ฝากขาย", , "CONSIGNMENT")
      
      Call LoadControl("CH", cboGeneric(0).Width, True, "สรุปฝากขาย", , "SUMMARY_CONSIGNMENT")
      Call LoadControl("CH", cboGeneric(0).Width, True, "สรุปฝากขายสุทธิ", , "SUMMARY_CONSIGNMENT_NET")
      
      Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เปรียบเทียบ"))
      
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เปรียบเทียบ"))
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "ฝากขาย", , "CONSIGNMENT")
   'รายงานที่2
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนของเป้าการขาย"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีของเป้าการขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนวันหยุด Repost2"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนวันหยุด Repost3"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE2", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE2", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า2"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO2", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ2"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO2", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE2", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย2"))
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่เปรียบเทียบ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่เปรียบเทียบ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))

   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนวันหยุด"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN", , "NOT_SHOW_RETURN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงสีระหว่างระดับ Sale", , "SHOW_COLOR")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))

   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสตัวแทน"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN", , "NOT_SHOW_RETURN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงสีระหว่างระดับ Sale", , "SHOW_COLOR")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียดรวมคงค้างเป็น0", , "SHOW_DETAIL_ZERO")
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
      '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะเคลื่อนไหวในรอบวัน", , "SHOW_ONLY_MOVE")
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_MOVE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เคลื่อนไหว"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_MOVE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เคลื่อนไหว"))
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ครบกำหนด"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ที่ครบกำหนด"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
         
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width, False, "", , "SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE", "ORDER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนวันหยุด"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)

   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนของเป้าการขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีของเป้าการขาย"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ขาย"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ขาย"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงส่วนเพิ่มส่วนหัก(แบบแยก)", , "SHOW_DETAIL_ADDSUB")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียดชำระ", , "SHOW_PAYMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเลขที่บัญชีการโอน", , "SHOW_ACCOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายการชำระ", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงสรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงส่วนเพิ่มส่วนหัก(แบบแยก)", , "SHOW_DETAIL_ADDSUB")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียดชำระ", , "SHOW_PAYMENT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเลขที่บัญชีการโอน", , "SHOW_ACCOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายการชำระ", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผลิตภัณฑ์"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง %", , "SHOW_PERCENT")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รายงานรายวัน", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะสรุป", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผลิตภัณฑ์"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รายงานรายวัน", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะสรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผลิตภัณฑ์"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
      Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผลิตภัณฑ์"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รายงานรายวัน", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะสรุป", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "SUB_FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เพิ่มช่วงจาก(วัน)"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "ADD_TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เพิ่มช่วงถึง(วัน)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "BATCH_NO_SET")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("BATCH NO"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO1", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผลิตภัณฑ์"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO1", , "STOCK_NO", True)
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SHOW_DETAIL1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHOW_DETAIL2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "SHOW_DETAIL3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด3"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "SHOW_DETAIL4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แสดงรายละเอียด4"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รายงานรายวัน", , "DAIRY_REPORT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะสรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการผลิตเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการผลิตปี"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
      
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width / 2, False, "", 2, "PRODUCTION_LOCATION", "PRODUCTION_LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่ผลิต"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า1"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "STOCK_TYPE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แยกประเภทสินค้า3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง GP", , "SHOW_GP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_FROM_STOCK_NO", , "NOT_STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมรหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "NOT_TO_STOCK_NO", , "NOT_STOCK_NO", True)
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง GP", , "SHOW_GP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   If TEMP_ROOT_TREE <> " S-2-32-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
   If TEMP_ROOT_TREE <> " S-2-32-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แยกสาขาลูกค้า", , "SHOW_GROUP_CUSTOMER")
   End If
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหมายเหตุ", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 3, "BIRTH_MONTH_ID", "BIRTH_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))
   
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหมายเหตุ", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
      
      
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนของเป้าการขาย"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีของเป้าการขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "HOLIDAY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนวันหยุด"))
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหมายเหตุ", , "NOTE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเฉพาะชื่อ Sale", , "SHOW_SALE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
   
   'เปรียบเทียบ
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE_COMPARE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เปรียบเทียบ"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE_COMPARE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่เปรียบเทียบ(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหมายเหตุ", , "NOTE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงสินค้า", , "NOT_SHOW_CARGO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงลูกค้า และสินค้า", , "NOT_SHOW_CUSTOMERS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงSALE, ลูกค้า และสินค้า", , "NOT_SHOW_SALE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนของเป้าการขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีของเป้าการขาย"))
   
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงสินค้า", , "NOT_SHOW_CARGO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงกลุ่มสินค้า และสินค้า", , "NOT_SHOW_CARGO_GROUP")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงSALE, กลุ่มสินค้า และสินค้า", , "NOT_SHOW_SALE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวม CN *เฉพาะ SALE ที่กำหนด ", , "NOT_SHOW_RETURN")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือนของเป้าการขาย"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีของเป้าการขาย"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดง ยอดขาย", , "SHOW_PRICE")

   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงลูกค้า", , "NOT_SHOW_CUSTOMERS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ใช้ชื่อลูกค้าย่อแทนชื่อเต็ม", , "SHORT_NAME")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
      
  If TEMP_ROOT_TREE = " SL-1-2-8-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงทศนิยม", , "SHOW_DECIMAL")
      Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรหัสลูกค้า", , "SHOW_CUS_CODE")
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
    
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_LOCATION_NO", , "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_LOCATION_NO", , "LOCATION_NO", True)
      
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
    
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงวันที่(7,7,7,7,2)"))
      
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
    Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_RCP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่รับชำระ"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ครบกำหนด(60<,30,30)"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PERIOD_DATE2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เกินกำหนด(7,8,15,>30)"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
    Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SHORT_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_EXCEL")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")

   Call LoadControl("CH", cboGeneric(0).Width, True, "พิมพ์ไปยังไฟล์", , "PRINT_TO_FILE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
         
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงจำนวน", , "NOT_SHOW_AMOUNT")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   If TEMP_ROOT_TREE = " S-2-31" Or TEMP_ROOT_TREE = " S-2-31-2" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_GROUP", "STOCK_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสินค้า/วัตถุดิบ"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE", "STOCK_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสินค้า/วัตถุดิบ"))
   
   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยกเลิก", , "EXCEPTION_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่ส่งของ"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_CHANGE1", "UNIT_CHANGE_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย1"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "UNIT_CHANGE2", "UNIT_CHANGE_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "UNIT_CHANGE3", "UNIT_CHANGE_NAME3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย3"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "UNIT_CHANGE4", "UNIT_CHANGE_NAME4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย4"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "UNIT_CHANGE5", "UNIT_CHANGE_NAME5")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย5"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "UNIT_CHANGE6", "UNIT_CHANGE_NAME6")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย6"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "UNIT_CHANGE7", "UNIT_CHANGE_NAME7")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย7"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "UNIT_CHANGE8", "UNIT_CHANGE_NAME8")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย8"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 9, "UNIT_CHANGE9", "UNIT_CHANGE_NAME9")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย9"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 10, "UNIT_CHANGE10", "UNIT_CHANGE_NAME10")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย10"))
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง"))   '---------------------------------------------------
         
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงบิล", , "NOT_SHOW_BILL_FLAG")
   
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
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หัวข้อข้อมูลหลัก"))
'
'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หัวข้อข้อมูลหลัก"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มลูกหนี้"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงหัว", , "NOT_SHOW_HEAD_FLAG")
   
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
    '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "EMP_NAME", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อ"))
   
    '5 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "EMP_LASTNAME", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("นามสกุล"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ตำแหน่ง"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้ค้า"))
   
    '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "APAR_NAME", , "APAR_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ค้า"))
   
    '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_SUPGRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ระดับผู้ค้า"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "MASTER_SUPTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้ค้า"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงเลขผู้เสียภาษี", , "SHOW_TAX")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงที่อยู่", , "SHOW_ADDRESS")
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มลูกหนี้"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMS_3() 'พนักงานขาย sale
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_SEARCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คนขับ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง")) ''''''''''''''
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_SEARCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_FROM")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากหมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   
      '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_TO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงหมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คนขับ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง")) ''''''''''''''
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
'   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
  '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))


   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_1_3() ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คนขับ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
  Call LoadComboData
   
End Sub
Private Sub InitReportS_1_4() ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คนขับ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
  Call LoadComboData
   
End Sub

Private Sub InitReportS_1_5() ' pui เพิ่ม สำหรับ ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า) -----------path------------  ระบบบัญชีและการเงิน --- >รายงานระบบบัญชีและการเงิน----->ใบคืนสินค้าเป็นชุด(ฟอร์มเปล่า)    CReportNormalRcp001_3
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_DOCUMENT_NO", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO", , "DOCUMENT_NO", True)
   
'   '4 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_DOCUMENT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
'
 




'   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DRIVER_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คนขับ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง"))
'
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

  ' 2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการขายเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการขายปี"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เปรียบเทียบ"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เปรียบเทียบ"))
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "จำนวนคืน", , "SHOW_RETURN_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยอดคืน", , "SHOW_RETURN_PRICE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TRANSPORTOR1", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง1"))
      
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR1_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR2", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง2"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR2_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "TRANSPORTOR3", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง3"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR3_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "คิดยอดจากใบ PO", , "PO_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "คิดยอดจากใบส่งของและขายสด", , "INVOICE_FLAG")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
      
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SHOW_SUMMARY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "SUPPLIER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "SUPPLIER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
         
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
'   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LOCATION_NO", "LOCATION_NAME", "LOCATION_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO_PERCENT", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วง%"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO_PERCENT", , "STOCK_NO", True)
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
  Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "SHOW_OUTLAY")
   
   Call ShowControl
   Call LoadComboData
End Sub

