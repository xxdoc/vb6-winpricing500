VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImportCustomer 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "frmImportCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8685
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   15319
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtPercent 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
      End
      Begin Xivess.uctlTextBox txtFileName 
         Height          =   405
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   873
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   12480
         TabIndex        =   0
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportCustomer.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1920
         TabIndex        =   1
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportCustomer.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   10
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10680
         TabIndex        =   3
         Top             =   2640
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9000
         TabIndex        =   2
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportCustomer.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportCustomer"
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
Private m_Data As Collection
Private m_Customer As CAPARMas
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
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
   Call ImportData
   Call SaveData
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim temp_Cus As CAPARMas
Dim temp_CusAddr As CApArAddress
Dim j As Integer
Dim dataEdit As String
dataEdit = ""
j = 0
If Not m_HasModify Then
   SaveData = True
   Exit Function
End If
For Each temp_Cus In m_Data
DoEvents
Set m_Customer = New CAPARMas
j = j + 1
   If Not CheckUniqueNs(APARCODE_UNIQUE, temp_Cus.APAR_CODE, ID) Then
       ShowMode = SHOW_EDIT
       Debug.Print temp_Cus.APAR_CODE
   Else
      ShowMode = SHOW_ADD
      m_Customer.ShowMode = ShowMode
      m_Customer.APAR_CODE = temp_Cus.APAR_CODE
      m_Customer.APAR_GRADE = -1
      m_Customer.APAR_TYPE = 827 'บุคคลทั่วไป
      m_Customer.CREDIT = 0
      m_Customer.TAX_ID = ""
      m_Customer.BIRTH_DATE = temp_Cus.BIRTH_DATE
      m_Customer.EMAIL = ""
      m_Customer.WEBSITE = ""
      m_Customer.PASSWD = ""
      m_Customer.BUSINESS_TYPE = -1
      m_Customer.BUSINESS_DESC = ""
      m_Customer.NORMAL_DISCOUNT = 0
      m_Customer.APAR_IND = 19999 'ลูกค้า
      m_Customer.PACKAGE_ID = -1
      m_Customer.LABEL_FLAG = "N"
      m_Customer.ADD_BRANCH_NAME = "N"
      m_Customer.FLAG_EDIT = "N"
      m_Customer.APAR_MAS_GROUP_CODE = ""
      m_Customer.APAR_MAS_GROUP_NAME = ""
      m_Customer.CANCEL_OUT_DOCUMENT = "N"
      m_Customer.CONSIGNMENT_FLAG = "N"
      m_Customer.BASKET_FIX_AMOUNT = 0
      
      Dim CstName As CApArName
      If m_Customer.CstNames.Count <= 0 Then
         Set CstName = New CApArName
         CstName.Flag = "A"
         Call m_Customer.CstNames.add(CstName)
      Else
         Set CstName = m_Customer.CstNames.Item(1)
         CstName.Flag = "E"
      End If
   
      Dim Name As cName
      Set Name = CstName.Name
      If m_Customer.CstNames.Count <= 0 Then
         Set Name = CstName.Name
         Call Name.SetFieldValue("LONG_NAME", temp_Cus.APAR_NAME)
         Call Name.SetFieldValue("SHORT_NAME", temp_Cus.APAR_NAME)
         Call Name.SetFieldValue("BILL_NAME", temp_Cus.APAR_NAME)
         Name.Flag = "A"
      Else
         Set Name = CstName.Name
         Call Name.SetFieldValue("LONG_NAME", temp_Cus.APAR_NAME)
         Call Name.SetFieldValue("SHORT_NAME", temp_Cus.APAR_NAME)
         Call Name.SetFieldValue("BILL_NAME", temp_Cus.APAR_NAME)
         Name.Flag = "E"
      End If
      
      Dim CstAddress As CApArAddress
      If m_Customer.CstAddresses.Count <= 0 Then
         Set CstAddress = New CApArAddress
         CstAddress.Flag = "A"
         Call m_Customer.CstAddresses.add(CstAddress)
      Else
         Set CstAddress = m_Customer.CstAddresses.Item(1)
         CstAddress.Flag = "E"
      End If

      Dim Address As CAddress
      If m_Customer.CstAddresses.Count <= 0 Then
         Set Address = CstAddress.Addresses
         Call Address.SetFieldValue("PHONE1", temp_Cus.PHONE)
         Call Address.SetFieldValue("BANGKOK_FLAG", "N")
         Address.Flag = "A"
      Else
         Set Address = CstAddress.Addresses
         Call Address.SetFieldValue("PHONE1", temp_Cus.PHONE)
         Call Address.SetFieldValue("BANGKOK_FLAG", "N")
         Address.Flag = "E"
      End If
     End If
    
     If Not glbDaily.AddEditCustomer(m_Customer, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   txtPercent.Text = j
   Set m_Customer = Nothing
Next


   Call EnableForm(Me, False)
  
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub ImportData()
'On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim APM As CAPARMas
Dim ADDS As CAddress
Dim glbCustomer As CAPARMas

Dim IsOK As Boolean



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

   prgProgress.Min = 1
   prgProgress.Max = (MaxRow * 3) + 1

   Set APM = New CAPARMas
   APM.ShowMode = SHOW_ADD
   For Row = 4 To MaxRow
      DoEvents
       Set glbCustomer = New CAPARMas

       glbCustomer.APAR_CODE = m_ExcelSheet.Cells(Row, 2).Value
       glbCustomer.APAR_NAME = m_ExcelSheet.Cells(Row, 3).Value
       glbCustomer.PHONE = m_ExcelSheet.Cells(Row, 4).Value
       glbCustomer.BIRTH_DATE = InternalDateToDateExGrid(CStr(m_ExcelSheet.Cells(Row, 5).Value)) 'm_ExcelSheet.Cells(Row, 5).Value
       
       Call m_Data.add(glbCustomer)
      Set glbCustomer = Nothing
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next Row

'   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
    
   prgProgress.Value = prgProgress.Max

   Call EnableForm(Me, True)
'   glbDatabaseMngr.DBConnection.CommitTrans
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

   glbErrorLog.LocalErrorMsg = err.Description
   glbErrorLog.ShowUserError
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
'      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
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
   Set m_Data = New Collection
   Set m_Customer = New CAPARMas
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PartColls = Nothing
   Set PartLabColls = Nothing
   Set PartLabUpdateColls = Nothing
   Set m_Data = Nothing
   Set m_Customer = Nothing
   Set PartUctlColls = Nothing
End Sub
'Private Function SearchLabCode(SearchItemNo As CPartItem, PartNo As String, PartName As String) As Boolean
'   SearchLabCode = True
'   Set SearchItemNo = GetObject("CPartItem", PartColls, PartNo, False)
'   If SearchItemNo Is Nothing Then
'      Set SearchItemNo = GetObject("CPartItem", PartLabColls, PartNo, False)
'      If SearchItemNo Is Nothing Then
'         Set SearchItemNo = GetObject("CPartItem", PartLabUpdateColls, Trim(PartNo), False)
'         If SearchItemNo Is Nothing Then
'            'LoadForm
'            Set SearchItemNo = New CPartItem
'            Set frmMapPlcProductItem.PartItem = SearchItemNo
'            Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
'            If Trim(PartNo) = Trim(PartName) Then
'               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo)
'            Else
'               frmMapPlcProductItem.HeaderText = MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo & "-" & PartName)
'            End If
'            frmMapPlcProductItem.ShowMode = SHOW_ADD
'            Load frmMapPlcProductItem
'            frmMapPlcProductItem.Show 1
'
'            OKClick = frmMapPlcProductItem.OKClick
'
'            Unload frmMapPlcProductItem
'            Set frmMapPlcProductItem = Nothing
'
'            'AddDataTo PartPlcUpdateColls
'            If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
'               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & PartNo & "-" & PartName
'               glbErrorLog.ShowUserError
'
'               SearchLabCode = False
'               Exit Function
'            End If
'            SearchItemNo.NUMBER_LAB_ID = Trim(PartNo)
'            Call PartLabUpdateColls.add(SearchItemNo, Trim(PartNo))
'         End If
'      End If
'   End If
'End Function
Private Sub Label1_Click()

End Sub
