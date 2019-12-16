VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPlanning 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "frmImportPlanning.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8685
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   15319
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2310
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   2760
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   10
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   8
         Top             =   3090
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlDate uctlPlanningDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1005
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlPlanningTo 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1875
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlPlanningFrom 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   4035
         Left            =   480
         TabIndex        =   19
         Top             =   4440
         Width           =   12585
      End
      Begin VB.Label lblPlanningDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningDate"
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   1095
         Width           =   1305
      End
      Begin VB.Label lblPlanningFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningFrom"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1575
         Width           =   1665
      End
      Begin VB.Label lblPlanningTo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlanningTo"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1980
         Width           =   1665
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   12480
         TabIndex        =   3
         Top             =   2325
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlanning.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   4
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlanning.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   14
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2820
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2340
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10935
         TabIndex        =   6
         Top             =   3630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9285
         TabIndex        =   5
         Top             =   3630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportPlanning.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public PlanningArea As Long

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private PartUctlColls As Collection
Private PartColls As Collection
Private PartLabColls  As Collection
Private PartLabUpdateColls  As Collection

Private m_PartItems As Collection
Private Sub cmdFileName_Click()
 On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If PlanningArea = 1 Or PlanningArea = 3 Then
      If Not VerifyDate(lblPlanningDate, uctlPlanningDate, False) Then
         Exit Sub
      End If
      
      If Not CheckUniqueNs(PLANNING_UNIQUE, Trim(DateToStringInt(uctlPlanningDate.ShowDate)), ID, Trim(Str(PlanningArea))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlPlanningDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   If PlanningArea = 2 Then
      If Not VerifyDate(lblPlanningFrom, uctlPlanningFrom, False) Then
         Exit Sub
      End If
      
      If Not VerifyDate(lblPlanningTo, uctlPlanningTo, False) Then
         Exit Sub
      End If
      
      If Not CheckUniqueNs(PLANNING_UNIQUE, Trim(DateToStringInt(uctlPlanningFrom.ShowDate)), ID, Trim(Str(PlanningArea))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlPlanningFrom.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
      
   
      
      
   Call LoadPartItem(Nothing, PartUctlColls, , , , 1)
   Call LoadPartItem(Nothing, PartColls, , , , 2)
   Call LoadPartItem(Nothing, PartLabColls, , , , 4)
   
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
   Call ImportPlanning
      
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
End Sub

Private Sub ImportPlanning()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim Pn As CPlanning
Dim Pni As CPlanningItem
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
   
   
   HasBegin = False

   ID = 1
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   ReDim FieldNames(MaxCol)
   ReDim FieldTypes(MaxCol)
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = (MaxRow * 3) + 1
   
   Set Pn = New CPlanning
   Pn.AddEditMode = SHOW_ADD
   If PlanningArea = 1 Then
      Pn.PLANNING_DATE = uctlPlanningDate.ShowDate
      Pn.PLANNING_FROM = uctlPlanningDate.ShowDate
      Pn.PLANNING_TO = uctlPlanningDate.ShowDate
   Else
      Pn.PLANNING_FROM = uctlPlanningFrom.ShowDate
      Pn.PLANNING_TO = uctlPlanningTo.ShowDate
      Pn.PLANNING_DATE = uctlPlanningFrom.ShowDate
   End If
   Pn.PLANNING_AREA = PlanningArea
   Pn.PLANNING_DESC = "IMPORTED" & Now
   
   'รอบแรก วัตถุดิบ หน่วยตัน ต้องคูณ 1000
   For row = 4 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(Val(m_ExcelSheet.Cells(row, 3).Value)) > 0 Then
         Set Pni = New CPlanningItem
         Pni.Flag = "A"
         
         Pni.PLAN_AMOUNT = Val(m_ExcelSheet.Cells(row, 3).Value) * 1000
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 1).Value), Trim(m_ExcelSheet.Cells(row, 2).Value)) Then
            Call EnableForm(Me, True)
            'cmdStart.Enabled = True
            cmdExit.Enabled = True
            cmdOK.Enabled = True
            Exit Sub
         End If
         
         Pni.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         If PlanningArea = 1 Or PlanningArea = 2 Then
            Call Pn.CollPartUse.add(Pni)
         Else
            Call Pn.CollPartSup.add(Pni)
         End If
         Set Pni = Nothing
      End If
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   'รอบสอง Premix หน่วย กก
   For row = 4 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(Val(m_ExcelSheet.Cells(row, 7).Value)) > 0 Then
         Set Pni = New CPlanningItem
         Pni.Flag = "A"
         
         Pni.PLAN_AMOUNT = Val(m_ExcelSheet.Cells(row, 7).Value)
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 5).Value), Trim(m_ExcelSheet.Cells(row, 6).Value)) Then
            Call EnableForm(Me, True)
            'cmdStart.Enabled = True
            cmdExit.Enabled = True
            cmdOK.Enabled = True
            Exit Sub
         End If
         
         Pni.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         If PlanningArea = 1 Or PlanningArea = 2 Then
            Call Pn.CollPartUse.add(Pni)
         End If
         Set Pni = Nothing
      End If
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   'รอบสาม อาหาร หน่วย ตันคูณ 1000
   For row = 4 To MaxRow
      DoEvents
      Me.Refresh
      
      If Val(Val(m_ExcelSheet.Cells(row, 9).Value)) > 0 Then
         Set Pni = New CPlanningItem
         Pni.Flag = "A"
         
         Pni.PLAN_AMOUNT = Val(m_ExcelSheet.Cells(row, 9).Value) * 1000
         If Not SearchLabCode(SearchItemNo, Trim(m_ExcelSheet.Cells(row, 8).Value), Trim(m_ExcelSheet.Cells(row, 8).Value)) Then
            Call EnableForm(Me, True)
            'cmdStart.Enabled = True
            cmdExit.Enabled = True
            cmdOK.Enabled = True
            Exit Sub
         End If
               
         Pni.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         If PlanningArea = 1 Or PlanningArea = 2 Then
            Call Pn.CollProductGet.add(Pni)
         End If
         Set Pni = Nothing
      End If
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   Call glbPlanning.AddEditPlanning(Pn, IsOK, False, glbErrorLog)
   
   
   For Each SearchItemNo In PartLabUpdateColls
      Call SearchItemNo.UpdateLabPartNo
   Next SearchItemNo
   
   Set Pn = Nothing
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   
   'cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   Call EnableForm(Me, True)
   
   'cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If PlanningArea = 1 Or PlanningArea = 3 Then
         uctlPlanningFrom.Enable = False
         uctlPlanningFrom.TabStop = False
         uctlPlanningTo.Enable = False
         uctlPlanningTo.TabStop = False
         uctlPlanningDate.SetFocus
      Else
         uctlPlanningDate.Enable = False
         uctlPlanningDate.TabStop = False
         uctlPlanningFrom.SetFocus
      End If
      
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
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
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
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "อิมพอร์ต" & HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblPlanningDate, MapText("วันที่ประมาณ"))
   Call InitNormalLabel(lblPlanningFrom, MapText("จากวันที่ประมาณ"))
   Call InitNormalLabel(lblPlanningTo, MapText("ถึงวันที่ประมาณ"))
   
   Call InitNormalLabel(lblNote, "- เริ่ม Import ที่ Row ที่ 4" & vbCrLf & "Col1 = รหัสวัตถุดิบ,Col2 รายละเอียดวัตถุดิบ,Col3 จำนวนเป็นตัน" & vbCrLf & "Col5 = รหัสวัตถุดิบพรีมิกซ์,Col6 รายละเอียดวัตถุดิบพรีมิกซ์,Col7 จำนวนเป็น กิโลกรัม" & vbCrLf & "Col = รหัสสินค้า,Col8 รายละเอียดสินค้า,Col9 จำนวนเป็นตัน" & vbCrLf & vbCrLf & "กรณีที่เป็นข้อมูลประมาณการรับวัตถุดิบจากซัพพลายเออร์นั้นให้มีเฉพาะ" & vbCrLf & "Col1 = รหัสวัตถุดิบ,Col2 รายละเอียดวัตถุดิบ,Col3 จำนวนเป็นตัน")
   
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set PartUctlColls = New Collection
   Set PartColls = New Collection
   Set PartLabColls = New Collection
   Set PartLabUpdateColls = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PartColls = Nothing
   Set PartLabColls = Nothing
   Set PartLabUpdateColls = Nothing
   Set PartUctlColls = Nothing
End Sub
Private Function SearchLabCode(SearchItemNo As CPartItem, PartNo As String, PartName As String) As Boolean
   SearchLabCode = True
   Set SearchItemNo = GetObject("CPartItem", PartColls, PartNo, False)
   If SearchItemNo Is Nothing Then
      Set SearchItemNo = GetObject("CPartItem", PartLabColls, PartNo, False)
      If SearchItemNo Is Nothing Then
         Set SearchItemNo = GetObject("CPartItem", PartLabUpdateColls, Trim(PartNo), False)
         If SearchItemNo Is Nothing Then
            'LoadForm
            Set SearchItemNo = New CPartItem
            Set frmMapPlcProductItem.PartItem = SearchItemNo
            Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
            If Trim(PartNo) = Trim(PartName) Then
               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo)
            Else
               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo & "-" & PartName)
            End If
            frmMapPlcProductItem.ShowMode = SHOW_ADD
            Load frmMapPlcProductItem
            frmMapPlcProductItem.Show 1
            
            OKClick = frmMapPlcProductItem.OKClick
               
            Unload frmMapPlcProductItem
            Set frmMapPlcProductItem = Nothing
      
            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & PartNo & "-" & PartName
               glbErrorLog.ShowUserError
               
               SearchLabCode = False
               Exit Function
            End If
            SearchItemNo.NUMBER_LAB_ID = Trim(PartNo)
            Call PartLabUpdateColls.add(SearchItemNo, Trim(PartNo))
         End If
      End If
   End If
End Function
