VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditJob1 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditJob1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTime uctlTime 
         Height          =   405
         Left            =   8760
         TabIndex        =   3
         Top             =   720
         Width           =   1215
         _extentx        =   2143
         _extenty        =   714
      End
      Begin VB.ComboBox cboLocation 
         Height          =   315
         Left            =   10080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   1755
      End
      Begin VB.ComboBox cboProductionLocation 
         Height          =   315
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1395
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   5535
         Left            =   10080
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3060
         Width           =   495
      End
      Begin Threed.SSFrame FraBorder 
         Height          =   6855
         Left            =   0
         TabIndex        =   25
         Top             =   3045
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12091
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSFrame fraInner 
            Height          =   4995
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   8811
            _Version        =   131073
            PictureBackgroundStyle=   2
            Begin Xivess.uctlTextBox txtPartNo 
               Height          =   435
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Width           =   1275
               _extentx        =   3096
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtPartDesc 
               Height          =   435
               Index           =   0
               Left            =   1380
               TabIndex        =   12
               Top             =   480
               Width           =   1725
               _extentx        =   3043
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtItemAmount 
               Height          =   435
               Index           =   0
               Left            =   3120
               TabIndex        =   13
               Top             =   480
               Width           =   1275
               _extentx        =   3096
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtLocationName 
               Height          =   435
               Index           =   0
               Left            =   8385
               TabIndex        =   17
               Top             =   480
               Width           =   1635
               _extentx        =   2884
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtWeightAmount 
               Height          =   435
               Index           =   0
               Left            =   4440
               TabIndex        =   14
               Top             =   480
               Width           =   1275
               _extentx        =   13361
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtItemAmountSTD 
               Height          =   435
               Index           =   0
               Left            =   5760
               TabIndex        =   15
               Top             =   480
               Width           =   1275
               _extentx        =   13361
               _extenty        =   767
            End
            Begin Xivess.uctlTextBox txtWeightAmountSTD 
               Height          =   435
               Index           =   0
               Left            =   7080
               TabIndex        =   16
               Top             =   480
               Width           =   1275
               _extentx        =   13361
               _extenty        =   767
            End
            Begin VB.Label lblLocation 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   8400
               TabIndex        =   29
               Top             =   120
               Width           =   1605
            End
            Begin VB.Label lblWeightSTD 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   7080
               TabIndex        =   30
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lblAmountSTD 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   5760
               TabIndex        =   31
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lblWeight 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   4440
               TabIndex        =   32
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lblAmount 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   3120
               TabIndex        =   35
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lblName 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   1440
               TabIndex        =   34
               Top             =   120
               Width           =   2085
            End
            Begin VB.Label lblNo 
               Alignment       =   2  'Center
               Caption         =   "Label1"
               Height          =   435
               Left            =   120
               TabIndex        =   33
               Top             =   120
               Width           =   1245
            End
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   720
         Width           =   1785
         _extentx        =   13361
         _extenty        =   767
      End
      Begin Xivess.uctlDate uctlJobDate 
         Height          =   405
         Left            =   4920
         TabIndex        =   2
         Top             =   720
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin Xivess.uctlTextBox txtTxAmount 
         Height          =   435
         Left            =   3645
         TabIndex        =   10
         Top             =   1680
         Width           =   1740
         _extentx        =   13361
         _extenty        =   767
      End
      Begin Xivess.uctlTextBox txtLotItemAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   1680
         Width           =   1755
         _extentx        =   13361
         _extenty        =   767
      End
      Begin Xivess.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   5295
         _extentx        =   9340
         _extenty        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   885
         Left            =   0
         TabIndex        =   38
         Top             =   2160
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1561
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
         Column(1)       =   "frmAddEditJob1.frx":27A2
         Column(2)       =   "frmAddEditJob1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJob1.frx":290E
         FormatStyle(2)  =   "frmAddEditJob1.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJob1.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJob1.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJob1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJob1.frx":2D5E
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   10605
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblLocationIn 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9360
         TabIndex        =   37
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label lblProductionLocation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   36
         Top             =   1320
         Width           =   1185
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   375
         Left            =   10320
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   27
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblInventoryDocNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   26
         Top             =   1800
         Width           =   1245
      End
      Begin Threed.SSCommand cmdBrowse 
         Height          =   405
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   780
         Width           =   915
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1320
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   23
         Top             =   840
         Width           =   1245
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   7920
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10600
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7320
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJob1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Job   As CJob

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private TempCollection As Collection

'Private TempLotItemID  As Long
'Private TempDocumentDate As String
'Private TempDocumentNo As String
'Private TempPartItemID As Long
'Private TempLocationID As Long
'Private TempPartNo As String
'Private TempPartDesc  As String
'Private TempLotItemAmount  As Double
'Private TempTxAmount  As Double
'Private TempUnitID As Long

Private m_Gui(100) As Long
Private m_PartItems As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Ji As CJobItem
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_Job.JOB_ID = ID
      m_Job.QueryFlag = 1
      If Not glbDaily.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Job.PopulateFromRS(1, m_Rs)

      txtJobNo.Text = m_Job.JOB_NO
      uctlJobDate.ShowDate = m_Job.JOB_DATETIME
      uctlTime.HR = Hour(m_Job.JOB_DATETIME)
      uctlTime.MI = Minute(m_Job.JOB_DATETIME)
      uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, m_Job.PRODUCT_ID)
      
      chkCommit.Value = FlagToCheck(m_Job.COMMIT_FLAG)
      cboProductionLocation.ListIndex = IDToListIndex(cboProductionLocation, m_Job.PRD_LOCATION_ID)
      cboLocation.ListIndex = IDToListIndex(cboLocation, m_Job.PRODUCT_LOCATION_ID)
      
      txtLotItemAmount.Text = m_Job.LOT_ITEM_AMOUNT
      txtTxAmount.Text = m_Job.PRODUCT_AMOUNT
      
'      TempDocumentDate = Ji.DOCUMENT_DATE                '  วันที่ใน INVENTORY_DOC
'      TempDocumentNo = Ji.DOCUMENT_NO
'      TempPartItemID = Ji.PART_ITEM_ID
'      TempLotItemAmount = m_Job.LOT_ITEM_AMOUNT
'      TempTxAmount = Ji.TX_AMOUNT
      'cboLocation.ListIndex = IDToListIndex(cboLocation, Ji.LOCATION_ID)
      
      
   End If
   
   If m_Job.JobOutItems.Count > 0 Then
      Call LoadControl(m_Job.JobOutItems)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call InitGrid
      
   If m_Job.JobInItems.Count > 1 Then
      Call Form_Resize
   End If
   
   GridEX1.ItemCount = CountItem(m_Job.JobInItems)
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
   
   If Not VerifyTextControl(lblJobNo, txtJobNo) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblJobDate, uctlJobDate) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblProductionLocation, cboProductionLocation, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(JOB_NO_UNIQUE, txtJobNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not VerifyTime(lblJobDate, uctlTime) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Job.AddEditMode = ShowMode
   m_Job.JOB_NO = txtJobNo.Text
   m_Job.JOB_DATE = uctlJobDate.ShowDate
   m_Job.JOB_DATETIME = uctlJobDate.ShowDate
   m_Job.JOB_DATETIME = DateAdd("h", uctlTime.HR, m_Job.JOB_DATETIME)
   m_Job.JOB_DATETIME = DateAdd("n", uctlTime.MI, m_Job.JOB_DATETIME)
   m_Job.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   
   m_Job.PRODUCT_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   
   m_Job.PRD_LOCATION_ID = cboProductionLocation.ItemData(Minus2Zero(cboProductionLocation.ListIndex))
   m_Job.PRODUCT_LOCATION_ID = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
   
   m_Job.LOT_ITEM_AMOUNT = Val(txtLotItemAmount.Text)
   m_Job.PRODUCT_AMOUNT = Val(txtTxAmount.Text)
   
   If Not SaveOutItem(m_Job.JobOutItems) Then
      'แสดงว่า ไมผ่านมาตรฐาน
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   
   If ShowMode = SHOW_ADD Then
      Call MergeInput
   
      Call PopulateGuiID(m_Job)
   End If
   
   Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1000, TempCollection, m_PartItems)             'ใบสั่งผลิต
   
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   m_Job.INVENTORY_DOC_ID = Ivd.GetFieldValue("INVENTORY_DOC_ID")
   
   If ShowMode = SHOW_ADD Then
      Dim Ji As CJobItem
      Dim Lt As CLotItem
      For Each Ji In m_Job.JobInItems
         For Each Lt In Ivd.ImportExportItems
            If Lt.PART_ITEM_ID > 0 Then
               If Ji.PART_ITEM_ID = Lt.PART_ITEM_ID Then
                  Ji.EXPORT_LOT_ITEM_ID = Lt.LOT_ITEM_ID
                  Exit For
               End If
            End If
         Next Lt
      Next Ji
   End If
   If Not glbDaily.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   Call glbDaily.CommitTransaction
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cboLocation_Click()
   m_HasModify = True
End Sub
Private Sub cboLocation_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboProductionLocation_Click()
   m_HasModify = True
End Sub
Private Sub cboProductionLocation_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblPartItem, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblProductionLocation, cboProductionLocation, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblLocation, cboLocation, False) Then
      Exit Sub
   End If
   
   frmAddLotItem.TempPartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   frmAddLotItem.DocumentDate = uctlJobDate.ShowDate
   frmAddLotItem.TempLocationID = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
   Set frmAddLotItem.TempCollection = TempCollection
   frmAddLotItem.CanEditPartFlag = True
   frmAddLotItem.HeaderText = Caption
   Load frmAddLotItem
   frmAddLotItem.Show 1
      
   Unload frmAddLotItem
   Set frmAddLotItem = Nothing
   
   Call GenerateJobItems
   If m_Job.JobOutItems.Count > 0 Then
      cmdBrowse.Enabled = False
      cmdAdd.Enabled = True
   End If
   
   Call Form_Resize
   
   GridEX1.ItemCount = CountItem(m_Job.JobInItems)
   GridEX1.Rebind
   
End Sub
Private Sub cmdBrowse_Click()
            
   If Not VerifyCombo(lblPartItem, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblProductionLocation, cboProductionLocation, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblLocation, cboLocation, False) Then
      Exit Sub
   End If
   
   frmAddLotItem.TempPartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   frmAddLotItem.DocumentDate = uctlJobDate.ShowDate
   frmAddLotItem.TempLocationID = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
   Set frmAddLotItem.TempCollection = TempCollection
   frmAddLotItem.HeaderText = Caption
   frmAddLotItem.CanEditPartFlag = True
   Load frmAddLotItem
   frmAddLotItem.Show 1
   
   Unload frmAddLotItem
   Set frmAddLotItem = Nothing
   
   Call GenerateJobItems
   If m_Job.JobOutItems.Count > 0 Then
      cmdBrowse.Enabled = False
      cmdAdd.Enabled = True
   End If
   
   GridEX1.ItemCount = CountItem(m_Job.JobInItems)
   GridEX1.Rebind
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
   
      Call EnableForm(Me, False)
      
      Call LoadStockCode(uctlPartLookup.MyCombo, m_PartItems)
      Set uctlPartLookup.MyCollection = m_PartItems
      
      Call LoadMaster(cboProductionLocation, , , , MASTER_PRODUCTION_LOCATION)
      
      Call LoadMaster(cboLocation, , , , MASTER_LOCATION)
      
      If ShowMode = SHOW_EDIT Then
         cmdBrowse.Enabled = False
         cmdAdd.Enabled = False
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         Call QueryData(False)
         uctlJobDate.ShowDate = Now
         uctlTime.HR = Hour(Now)
         uctlTime.MI = Minute(Now)
      End If

      m_HasModify = False
      Call EnableForm(Me, True)
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
   ElseIf Shift = 0 And KeyCode = 33 Then
      Call VsDown
   ElseIf Shift = 0 And KeyCode = 34 Then
      Call Vsup
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Job = Nothing
   Set m_PartItems = Nothing
   Set TempCollection = Nothing
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   fraInner.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   FraBorder.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblJobNo, MapText("JOB NO"))
   Call InitNormalLabel(lblJobDate, MapText("วันเวลา"))
   Call InitNormalLabel(lblInventoryDocNo, MapText("รายละเอียด"))
   Call InitNormalLabel(lblPartItem, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblProductionLocation, MapText("สถานที่ผลิต"))
   Call InitNormalLabel(lblLocationIn, MapText("คลัง"))
   
   Call InitNormalLabel(lblNo, MapText("รหัส"))
   Call InitNormalLabel(lblName, MapText("สินค้า"))
   Call InitNormalLabel(lblAmount, MapText("จน.จริง"))
   Call InitNormalLabel(lblWeight, MapText("นน.จริง"))
   Call InitNormalLabel(lblAmountSTD, MapText("จน.STD"))
   Call InitNormalLabel(lblWeightSTD, MapText("นน.STD"))
   Call InitNormalLabel(lblLocation, MapText("คลัง"))
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtLotItemAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTxAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   txtLotItemAmount.Enabled = False
   txtTxAmount.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboProductionLocation)
   Call InitCombo(cboLocation)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdBrowse.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก"))
   Call InitMainButton(cmdOK, MapText("ตกลง"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdBrowse, MapText("B"))

   Call InitCheckBox(chkCommit, "ห้ามแก้ไข")
   
   Call InitGrid
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
   Set m_Job = New CJob
   Set m_PartItems = New Collection
   Set TempCollection = New Collection
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   
   If m_Job.JobInItems.Count > 1 Then
      GridEX1.Height = 885 + (360 * (m_Job.JobInItems.Count - 1))
      If GridEX1.Height >= 5000 Then
         GridEX1.Height = 5000
      End If
   End If
   FraBorder.Top = GridEX1.Top + GridEX1.Height
   FraBorder.Height = ScaleHeight - FraBorder.Top
   FraBorder.Width = ScaleWidth - 1800
   VScroll1.Top = FraBorder.Top
   VScroll1.Height = FraBorder.Height - 20
   fraInner.Top = 0
   If m_Job.JobInItems.Count <= 0 Then
      fraInner.Height = FraBorder.Height
   End If
   fraInner.Width = ScaleWidth - 1800
   cmdAdd.Top = ScaleHeight - 1780
   cmdOK.Top = ScaleHeight - 1180
   cmdExit.Top = ScaleHeight - 580
   VScroll1.Left = ScaleWidth - cmdOK.Width - VScroll1.Width - 40
   cmdAdd.Left = VScroll1.Left + VScroll1.Width - 10
   cmdExit.Left = VScroll1.Left + VScroll1.Width - 10
   cmdOK.Left = VScroll1.Left + VScroll1.Width - 10
      
   GridEX1.Width = ScaleWidth - 50
   
   txtPartDesc(0).Width = fraInner.Width - txtPartNo(0).Width - txtItemAmount(0).Width - txtWeightAmount(0).Width - txtItemAmountSTD(0).Width - txtWeightAmountSTD(0).Width - txtLocationName(0).Width - 300
   txtItemAmount(0).Left = txtPartDesc(0).Left + txtPartDesc(0).Width + 20
   txtWeightAmount(0).Left = txtItemAmount(0).Left + txtItemAmount(0).Width + 20
   txtItemAmountSTD(0).Left = txtWeightAmount(0).Left + txtWeightAmount(0).Width + 20
   txtWeightAmountSTD(0).Left = txtItemAmountSTD(0).Left + txtItemAmountSTD(0).Width + 20
   txtLocationName(0).Left = txtWeightAmountSTD(0).Left + txtWeightAmountSTD(0).Width + 20
   
   lblNo.Left = txtPartNo(0).Left
   lblName.Left = txtPartDesc(0).Left
   lblAmount.Left = txtItemAmount(0).Left
   lblWeight.Left = txtWeightAmount(0).Left
   lblAmountSTD.Left = txtItemAmountSTD(0).Left
   lblWeightSTD.Left = txtWeightAmountSTD(0).Left
   lblLocation.Left = txtLocationName(0).Left
   
   
End Sub
Private Sub txtItemAmount_Change(Index As Integer)
On Error GoTo err
Dim I As Long
Dim Sum1 As Double
Dim Sum2  As Double
Dim MainCmp As Double
Dim ShowSum  As Double
Dim Ji As CJobItem
   
   m_HasModify = True
   If txtItemAmount.Count <> txtLocationName.Count Then
      Exit Sub
   End If
   
   Sum1 = 0
   Sum2 = 0
   MainCmp = Val(txtTxAmount.Text)
   For I = 1 To txtItemAmount.Count - 1
      Set Ji = m_Job.JobOutItems(I)
      If Ji.SUM_FLAG = "Y" Then
         txtItemAmount(I).Text = Sum1
         Sum1 = 0
         MainCmp = Sum2
         Sum2 = 0
      Else
         Set Ji = m_Job.JobOutItems(I + 1)
         If Ji.SUM_FLAG = "Y" Then
            txtItemAmount(I).Text = Val(MainCmp - Sum1)
         End If
         Sum1 = Sum1 + Val(txtItemAmount(I).Text)
         If txtLocationName.Item(I).Text = "------>" Then
            Sum2 = Sum2 + Val(txtItemAmount(I).Text)
         End If
      End If
   Next I
   
   Exit Sub
err:
   glbErrorLog.LocalErrorMsg = "กรุณาใส่รวมให้ครบทุก PROCESS"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtItemAmountSTD_Change(Index As Integer)
Dim I As Long
Dim Sum1 As Double
Dim Sum2  As Double
Dim MainCmp As Double
Dim ShowSum  As Double
   
   m_HasModify = True
   If txtItemAmount.Count <> txtLocationName.Count Then
      Exit Sub
   End If
   
   Sum1 = 0
   Sum2 = 0
   MainCmp = Val(txtTxAmount.Text)
   For I = 1 To txtItemAmountSTD.Count - 1
      If txtPartDesc.Item(I).Text = "รวม" Then
         txtItemAmountSTD(I).Text = Sum1
         Sum1 = 0
         MainCmp = Sum2
         
      Else
         If txtPartDesc.Item(I + 1).Text = "รวม" Then
            txtItemAmountSTD(I).Text = MainCmp - Sum1
         End If
         Sum1 = Sum1 + Val(txtItemAmountSTD(I).Text)
         If txtLocationName.Item(I).Text = "------>" Then
            Sum2 = Sum2 + Val(txtItemAmountSTD(I).Text)
         End If
      End If
   Next I
   
End Sub

Private Sub txtJobNo_Change()
   m_HasModify = True
End Sub
Private Sub txtLotItemAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPartDesc_Change(Index As Integer)
   m_HasModify = True
End Sub
Private Sub txtPartNo_Change(Index As Integer)
   m_HasModify = True
End Sub

Private Sub txtTxAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightAmount_Change(Index As Integer)
Dim I As Long
Dim Sum1 As Double
Dim Sum2  As Double
Dim MainCmp As Double
Dim ShowSum  As Double

   m_HasModify = True
   If txtWeightAmount.Count <> txtLocationName.Count Then
      Exit Sub
   End If
   
   Sum1 = 0
   Sum2 = 0
   'MainCmp = Val(txtTxAmount.Text)
   For I = 1 To txtWeightAmount.Count - 1
      If txtPartDesc.Item(I).Text = "รวม" Then
         txtWeightAmount(I).Text = Sum1
         Sum1 = 0
         'MainCmp = Sum2
         
      Else
'         If txtPartDesc.Item(I + 1).Text = "รวม" Then
'            txtWeightAmount(I).Text = MainCmp - Sum1
'         End If
         Sum1 = Sum1 + Val(txtWeightAmount(I).Text)
         If txtLocationName.Item(I).Text = "------>" Then
            Sum2 = Sum2 + Val(txtWeightAmount(I).Text)
         End If
      End If
   Next I
End Sub

Private Sub txtWeightAmountSTD_Change(Index As Integer)
Dim I As Long
Dim Sum1 As Double
Dim Sum2  As Double
Dim MainCmp As Double
Dim ShowSum  As Double

   m_HasModify = True
   If txtWeightAmountSTD.Count <> txtLocationName.Count Then
      Exit Sub
   End If
   
   Sum1 = 0
   Sum2 = 0
   'MainCmp = Val(txtTxAmount.Text)
   For I = 1 To txtWeightAmountSTD.Count - 1
      If txtPartDesc.Item(I).Text = "รวม" Then
         txtWeightAmountSTD(I).Text = Sum1
         Sum1 = 0
         'MainCmp = Sum2
         
      Else
'         If txtPartDesc.Item(I + 1).Text = "รวม" Then
'            txtWeightAmountSTD(I).Text = MainCmp - Sum1
'         End If
         Sum1 = Sum1 + Val(txtWeightAmountSTD(I).Text)
         If txtLocationName.Item(I).Text = "------>" Then
            Sum2 = Sum2 + Val(txtWeightAmountSTD(I).Text)
         End If
      End If
   Next I
End Sub

Private Sub uctlJobDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTime_HasChange()
   m_HasModify = True
End Sub

Private Sub VScroll1_Change()
   If fraInner.Height > FraBorder.Height Then
      VScroll1.Min = 0
      VScroll1.Max = 100
      
      ''debug.print (VScroll1.Value)

      fraInner.Top = (FraBorder.Height - fraInner.Height) * VScroll1.Value / 100
   End If
End Sub
Private Sub GenerateJobItems()
Dim TempRs  As ADODB.Recordset
Dim ItemCount  As Long
Dim Fm   As CFormula
Dim TempFi As CFormulaItem
Dim TempJi As CJobItem
Dim IsOK As Boolean
Dim PrevKey1  As String
Dim Lk As CLotItemLink
Dim PartItemID As Long
   
   Set TempRs = New ADODB.Recordset
   
   '<--------------------------------------------------------------------------------------------------------------------------------------------
   Set m_Job.JobInItems = Nothing
   Set m_Job.JobInItems = New Collection
   For Each Lk In TempCollection
      Set TempJi = New CJobItem
      TempJi.Flag = "A"
      TempJi.TX_TYPE = "I"
      TempJi.PART_ITEM_ID = Lk.PART_ITEM_ID
      
      PartItemID = Lk.PART_ITEM_ID
      
      TempJi.PART_NO = Lk.PART_NO
      TempJi.PART_DESC = Lk.PART_DESC
      TempJi.TX_AMOUNT = Lk.IMPORT_AMOUNT
      TempJi.IMPORT_AMOUNT = Lk.LOT_ITEM_AMOUNT
      TempJi.LOCATION_ID = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
      
      TempJi.DOCUMENT_DATE = Lk.DOCUMENT_DATE
      TempJi.DOCUMENT_NO = Lk.DOCUMENT_NO
      
      TempJi.UNIT_ID = Lk.UNIT_ID
      TempJi.NEXT_FLAG = "N"
      
      Call m_Job.JobInItems.add(TempJi)
   Next Lk
   Call CalcualteLotItemAmount
   '-------------------------------------------------------------------------------------------------------------------------------------------->
   
   If m_Job.JobOutItems.Count = 0 And m_Job.JobInItems.Count > 0 And PartItemID > 0 Then
      Set Fm = New CFormula
      Fm.FORMULA_ID = -1
      Fm.PART_ITEM_ID = PartItemID
      Fm.PRD_LOCATION_ID = cboProductionLocation.ItemData(Minus2Zero(cboProductionLocation.ListIndex))
      Fm.QueryFlag = 1
      If Not glbDaily.QueryFormula(Fm, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      For Each TempFi In Fm.FormulaItems
         Set TempJi = New CJobItem
         TempJi.Flag = "A"
          
         TempJi.TX_TYPE = "E"
         TempJi.BATCH_NO = TempFi.BATCH_NO
         TempJi.ORDER_NO = TempFi.ORDER_NO
         TempJi.PART_ITEM_ID = TempFi.PART_ITEM_ID
         TempJi.PART_NO = TempFi.PART_NO
         TempJi.PART_DESC = TempFi.PART_DESC
         TempJi.UNIT_ID = TempFi.UNIT_ID
         
         TempJi.LOCATION_ID = TempFi.LOCATION_ID
         TempJi.LOCATION_NO = TempFi.LOCATION_NO
         TempJi.LOCATION_NAME = TempFi.LOCATION_NAME
         
         TempJi.LOST_ID = TempFi.LOST_ID
         TempJi.PROBLEM_DESC = TempFi.PROBLEM_DESC
         TempJi.PROBLEM_LIMIT_PERCENT = TempFi.PROBLEM_LIMIT_PERCENT
         TempJi.SUM_FLAG = TempFi.SUM_FLAG
         TempJi.NEXT_FLAG = TempFi.NEXT_FLAG
         TempJi.PRODUCTION_TYPE = TempFi.PRODUCTION_TYPE
         
         Call m_Job.JobOutItems.add(TempJi)
      Next TempFi
      
      Call LoadControl(m_Job.JobOutItems)
   End If
   
   If m_Job.JobInItems.Count > 1 Then
      Call Form_Resize
   End If
End Sub
Private Sub LoadControl(TempCol As Collection)
Dim Ji As CJobItem
Dim I As Long
Dim PrevKey1  As String
   
   I = 0
   txtPartNo(I).Visible = False
   txtPartDesc(I).Visible = False
   txtItemAmount(I).Visible = False
   txtWeightAmount(I).Visible = False
   txtItemAmountSTD(I).Visible = False
   txtWeightAmountSTD(I).Visible = False
   txtLocationName(I).Visible = False
   
   For Each Ji In TempCol
      
      
      I = I + 1
      
      m_Gui(I) = Ji.PART_ITEM_ID
      
      If Ji.PART_ITEM_ID > 0 Then
         Load txtPartNo(I)
         Call txtPartNo(I).SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
         txtPartNo(I).Visible = True
         txtPartNo(I).Enabled = False
         txtPartNo(I).Top = txtPartNo(0).Top + txtPartNo(0).Height * (I - 1)
         txtPartNo(I).Text = Ji.PART_NO
      End If
            
      Load txtPartDesc(I)
      Call txtPartDesc(I).SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
      txtPartDesc(I).Visible = True
      txtPartDesc(I).Enabled = False
      txtPartDesc(I).Top = txtPartDesc(0).Top + txtPartDesc(0).Height * (I - 1)
      If Ji.PART_ITEM_ID > 0 Then
         txtPartDesc(I).Text = Ji.PART_DESC
      Else
         txtPartDesc(I).Text = Ji.PROBLEM_DESC
      End If
      
      Load txtItemAmount(I)
      Call txtItemAmount(I).SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      txtItemAmount(I).Visible = True
      txtItemAmount(I).Enabled = True
      txtItemAmount(I).Top = txtItemAmount(0).Top + txtItemAmount(0).Height * (I - 1)
      txtItemAmount(I).Text = Ji.TX_AMOUNT
      txtItemAmount(I).TabIndex = txtItemAmount(0).TabIndex + (2 * I)
      
      Load txtWeightAmount(I)
      Call txtWeightAmount(I).SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      txtWeightAmount(I).Visible = True
      txtWeightAmount(I).Enabled = True
      txtWeightAmount(I).Top = txtWeightAmount(0).Top + txtWeightAmount(0).Height * (I - 1)
      txtWeightAmount(I).Text = Ji.TX_WEIGHT
      txtWeightAmount(I).TabIndex = txtWeightAmount(0).TabIndex + (2 * I)
      
      Load txtItemAmountSTD(I)
      Call txtItemAmountSTD(I).SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      txtItemAmountSTD(I).Visible = True
      txtItemAmountSTD(I).Enabled = True
      txtItemAmountSTD(I).Top = txtItemAmountSTD(0).Top + txtItemAmountSTD(0).Height * (I - 1)
      txtItemAmountSTD(I).Text = Ji.TX_AMOUNT_STD
      
      Load txtWeightAmountSTD(I)
      Call txtWeightAmountSTD(I).SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      txtWeightAmountSTD(I).Visible = True
      txtWeightAmountSTD(I).Enabled = True
      txtWeightAmountSTD(I).Top = txtWeightAmountSTD(0).Top + txtWeightAmountSTD(0).Height * (I - 1)
      txtWeightAmountSTD(I).Text = Ji.TX_WEIGHT_STD
      
      Load txtLocationName(I)
      Call txtLocationName(I).SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
      txtLocationName(I).Visible = True
      txtLocationName(I).Enabled = False
      txtLocationName(I).Top = txtLocationName(0).Top + txtLocationName(0).Height * (I - 1)
      If Ji.LOCATION_ID > 0 Then
         txtLocationName(I).Text = Ji.LOCATION_NAME
      ElseIf Ji.NEXT_FLAG = "Y" Then
         txtLocationName(I).Text = "------>"
      Else
         txtLocationName(I).Text = ""
      End If
      
      If (txtPartDesc(I).Top + txtPartDesc(I).Height) > fraInner.Height Then
         fraInner.Height = fraInner.Height + txtPartDesc(I).Height
      End If
   Next Ji
   
End Sub
Private Function SaveOutItem(TempCol As Collection) As Boolean
Dim Ji As CJobItem
Dim I As Long
   
   SaveOutItem = False
   I = 0
   For Each Ji In TempCol
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
      I = I + 1
      
      If Ji.PART_ITEM_ID <= 0 And Ji.PROBLEM_LIMIT_PERCENT > 0 Then
         If (Val(txtItemAmount(I).Text) > MyDiff(Val(txtTxAmount.Text) * Ji.PROBLEM_LIMIT_PERCENT, 100)) Then
            glbErrorLog.LocalErrorMsg = "รายการผลิตที่ผิดพลาด (" & Ji.PROBLEM_DESC & ")  เกินค่ามาตรฐาน " & MyDiff(Val(txtTxAmount.Text) * Ji.PROBLEM_LIMIT_PERCENT, 100)
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            
            If Not VerifyAccessRight("PRODUCT_VERIFY") Then
               frmVerifyAccRight.AccName = "PRODUCT_VERIFY"
               Load frmVerifyAccRight
               frmVerifyAccRight.Show 1
               
               If frmVerifyAccRight.GrantRight Then
                  Unload frmVerifyAccRight
                  Set frmVerifyAccRight = Nothing
               Else
                  Unload frmVerifyAccRight
                  Set frmVerifyAccRight = Nothing
                  Exit Function
               End If
            End If
         End If
      End If
      
      If Val(txtItemAmount(I).Text) < 0 Then
         glbErrorLog.LocalErrorMsg = "รายการ " & txtPartDesc(I).Text & ")" & " ติดลบ"
         glbErrorLog.ShowErrorLog (LOG_MSGBOX)
         Exit Function
      End If
      
      Ji.TX_AMOUNT = Val(txtItemAmount(I).Text)
      Ji.TX_WEIGHT = Val(txtWeightAmount(I).Text)
      Ji.TX_AMOUNT_STD = Val(txtItemAmountSTD(I).Text)
      Ji.TX_WEIGHT_STD = Val(txtWeightAmountSTD(I).Text)
      
   Next Ji
   SaveOutItem = True
End Function
Private Sub Vsup()
   If (VScroll1.Value + 25) <= VScroll1.Max Then
      VScroll1.Value = VScroll1.Value + 25
   End If
End Sub
Private Sub VsDown()
   If (VScroll1.Value - 25) >= VScroll1.Min Then
      VScroll1.Value = VScroll1.Value - 25
   End If
End Sub
Private Sub PopulateGuiID(BD As CJob)
Dim Di As CJobItem
      
   For Each Di In BD.JobInItems
      If Di.LOCATION_ID > 0 Then
         If Di.Flag = "A" Then
            Di.LINK_ID = GetNextGuiID(BD)
         End If
      End If
   Next Di
   
   For Each Di In BD.JobOutItems
      If Di.LOCATION_ID > 0 Then
         If Di.Flag = "A" Then
            Di.LINK_ID = GetNextGuiID(BD)
         End If
      End If
   Next Di
   
End Sub
Private Function GetNextGuiID(BD As CJob) As Long
Dim Di As CJobItem
Dim MaxId As Long

   MaxId = 0
   
   For Each Di In BD.JobInItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In BD.JobOutItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("รหัสวัตถุดิบ")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = ScaleWidth - 7200
   Col.Caption = MapText("วัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("หมายเลขเอกสาร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.Caption = MapText("วันที่")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("จำนวน")
   Col.TextAlignment = jgexAlignRight
   
   GridEX1.ItemCount = 0
   
   Set Col = Nothing
   Set fmsTemp = Nothing

End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim Ji As CJobItem
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Job.JobInItems Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If

     
   If m_Job.JobInItems.Count <= 0 Then
      Exit Sub
   End If
   Set Ji = GetItem(m_Job.JobInItems, RowIndex, RealIndex)
   If Ji Is Nothing Then
      Exit Sub
   End If

   Values(1) = Ji.JOB_ITEM_ID
   Values(2) = RealIndex
   Values(3) = Ji.PART_NO
   Values(4) = Ji.PART_DESC
   Values(5) = Ji.DOCUMENT_NO
   Values(6) = DateToStringExtEx2(Ji.DOCUMENT_DATE)
   Values(7) = FormatNumber(Ji.TX_AMOUNT)
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
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
Private Sub CalcualteLotItemAmount()
Dim Ji As CJobItem
Dim Sum1  As Double
Dim Sum2  As Double
   Sum1 = 0
   Sum2 = 0
   For Each Ji In m_Job.JobInItems
      Sum1 = Sum1 + Ji.TX_AMOUNT
      Sum2 = Sum2 + Ji.IMPORT_AMOUNT
   Next Ji
   txtLotItemAmount.Text = Sum2
   txtTxAmount.Text = Sum1
End Sub
Private Sub MergeInput()
Dim Ji  As CJobItem
Dim TempJi As CJobItem
Dim TempCollectionM As Collection
   
   If m_Job.JobInItems.Count > 1 Then
      Set TempCollectionM = New Collection
      For Each Ji In m_Job.JobInItems
         Set TempJi = GetObject("CJobItem", TempCollectionM, Trim(Str(Ji.PART_ITEM_ID)), False)
         If TempJi Is Nothing Then
            Set TempJi = New CJobItem
            TempJi.Flag = "A"
            TempJi.TX_TYPE = "I"
            TempJi.PART_ITEM_ID = Ji.PART_ITEM_ID
            
            TempJi.TX_AMOUNT = Ji.TX_AMOUNT
            TempJi.LOCATION_ID = Ji.LOCATION_ID
            
            TempJi.DOCUMENT_DATE = Ji.DOCUMENT_DATE
            TempJi.DOCUMENT_NO = Ji.DOCUMENT_NO
         
            TempJi.UNIT_ID = Ji.UNIT_ID
            TempJi.NEXT_FLAG = Ji.NEXT_FLAG
            Call TempCollectionM.add(TempJi, Trim(Str(Ji.PART_ITEM_ID)))
            Set Ji = Nothing
         Else
            TempJi.TX_AMOUNT = TempJi.TX_AMOUNT + Ji.TX_AMOUNT
         End If
      Next Ji
         
      Set m_Job.JobInItems = Nothing
      
      Set m_Job.JobInItems = TempCollectionM
      
      Set TempCollectionM = Nothing
   End If
End Sub
