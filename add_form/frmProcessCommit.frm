VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessCommit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmProcessCommit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6218
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   2400
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Xivess.uctlDate uctlToDate 
         Height          =   405
         Left            =   7170
         TabIndex        =   1
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1980
         TabIndex        =   2
         Top             =   1440
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtFromStockNo 
         Height          =   465
         Left            =   1980
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Xivess.uctlTextBox txtToStockNo 
         Height          =   465
         Left            =   5460
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin VB.Label lblFromStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lblToStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5880
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   7
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessCommit.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   14
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   12
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   8
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcessCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ProcessMode As Long

Private m_Products    As Collection
Private DocLinkLot As Collection
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim PartItemID As Long
   Call glbDaily.StartTransaction
   
   Me.Enabled = False
   
   If ProcessMode = 1 Then     '��Ѻ �鹷ع����� ��е鹷ع���
      Status = UpdateCapitalMovement(uctlFromDate.ShowDate, uctlToDate.ShowDate, Me.prgProgress, Me.txtPercent)
   ElseIf ProcessMode = 2 Then   ' ź������ ���͵���ʹ
      Status = ClearDataBillingDocStockCash(uctlFromDate.ShowDate, uctlToDate.ShowDate, Me.prgProgress, Me.txtPercent)
   End If
   
   Me.Enabled = True
   
   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "����ѿഴ��������ó�"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "����ѿഴ ERROR"
      glbErrorLog.ShowUserError
   End If
   
   Call cmdOK_Click
   Exit Sub
   
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Dim FirstDate As Date
      Dim LastDate As Date
      uctlFromDate.ShowDate = DateAdd("D", -15, Now)
      uctlToDate.ShowDate = Now
      m_HasModify = False
      
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products, , "N")
      Set uctlProductLookup.MyCollection = m_Products
      
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
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitNormalLabel(lblFromDate, "�ҡ�ѹ���", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "�����׺˹��")
   Call InitNormalLabel(lblPercent, "����ૹ��")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "�֧�ѹ���", RGB(255, 0, 0))
   Call InitNormalLabel(lblProduct, "�����Թ���")
   Call InitNormalLabel(lblFromStockNo, "�ҡ�����Թ���")
   Call InitNormalLabel(lblToStockNo, "�֧�����Թ���")
   
   uctlProductLookup.MyTextBox.SetKeySearch ("STOCK_NO")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFromStockNo.SetKeySearch("STOCK_NO")
   Call txtToStockNo.SetKeySearch("STOCK_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
  ' cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
  ' Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdStart, MapText("�����"))
   
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
      
   Set m_Products = New Collection
   Set DocLinkLot = New Collection
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Products = Nothing
   Set DocLinkLot = Nothing
End Sub
Public Function UpdateCapitalMovement(Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional Parent As Object = Nothing, Optional ParentEx As Object = Nothing) As Boolean
On Error GoTo ErrorHandler
Dim m_Rs As ADODB.Recordset
' Load �ʹ¡�� �ҡ���ҧ LOTITEM
Dim BalanceCollold As Collection
'Load �鹷ع �ҡ���ҧ CapitalMovement
Dim CapitalCollold As Collection
'�纡������͹��Ǣͧ�鹷ع
Dim CapitalCollmovement As Collection
'�� ID �ͧ  LINK ImportExport
Dim ExPortIDColl As Collection
   
If FromDate = -1 Then
      FromDate = Now
   End If
   
   If ToDate = -1 Then
      ToDate = Now
   End If
   
   Set m_Rs = New ADODB.Recordset
   
   Set BalanceCollold = New Collection
   Set CapitalCollold = New Collection
   Set CapitalCollmovement = New Collection
   Set ExPortIDColl = New Collection
   
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim PERCENT As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean
Dim Lt As CLotItem
Dim Cm As CCapitalMovement
   
   If Not (Parent Is Nothing) Then
      Parent.Max = 100
      Parent.Min = 0
   End If
   
   
   
   Set Lt = New CLotItem
   Lt.LOT_ITEM_ID = -1
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.CANCEL_FLAG = "N"
   Lt.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Lt.FROM_STOCK_NO = txtFromStockNo.Text
   Lt.TO_STOCK_NO = txtToStockNo.Text
   Call Lt.QueryData(2, m_Rs, iCount, False)
   
   '                                      1                             ���ӡ�� load LotItem ����������բ����
   Parent.Value = MyDiff(I, m_Rs.RecordCount) * 10
   ParentEx.Text = FormatNumber(Parent.Value)
   Parent.Refresh
   ParentEx.Refresh
   
   Call GetDocItemIDLinkLotItemID(DocLinkLot, -1, txtFromStockNo.Text, txtToStockNo.Text, Lt.PART_ITEM_ID, "N", FromDate, ToDate)
   
   Parent.Value = MyDiff(I, m_Rs.RecordCount) * 50
   ParentEx.Text = FormatNumber(Parent.Value)
   Parent.Refresh
   ParentEx.Refresh
   
   HasBegin = True
   
   Set Cm = New CCapitalMovement
   Cm.FROM_DATE = FromDate
   Cm.TO_DATE = ToDate
   Cm.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Cm.FROM_STOCK_NO = txtFromStockNo.Text
   Cm.TO_STOCK_NO = txtToStockNo.Text
   Call Cm.ClearData
   
   '                                      2                             ź������� Balance Accumn
   
'   Dim strDate As String
   'Call glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog)
   ''debug.print (strDate)
   Call LoadCapitalMovementLocation(CapitalCollold, DateAdd("D", -1, FromDate), DateAdd("D", -1, FromDate), Cm.PART_ITEM_ID, , txtFromStockNo.Text, txtToStockNo.Text) '�Ҥ�����¢ͧ�Թ��� � �ѹ �������ش ����Ф�ѧ
   
   Parent.Value = MyDiff(I, m_Rs.RecordCount) * 60
   ParentEx.Text = FormatNumber(Parent.Value)
   Parent.Refresh
   ParentEx.Refresh
   
   Call CopyOldToMovement(CapitalCollold, CapitalCollmovement)
   
   Call LoadLeftAmountLocation(BalanceCollold, , DateAdd("D", -1, FromDate), , txtFromStockNo.Text, txtToStockNo.Text, Lt.PART_ITEM_ID)                                  'Load �����Ţͧ����ҳ � �ѹ �������ش ����Ф�ѧ
   
   'Call glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog)
   ''debug.print (strDate)
   
   '                                      3
   
   '                                      4                                ǹ Loop ������� LotItem
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      If Not (Parent Is Nothing) Then
         Parent.Value = 70 + MyDiff(I, m_Rs.RecordCount) * 25
         ParentEx.Text = FormatNumber(Parent.Value)
         Parent.Refresh
         ParentEx.Refresh
         DoEvents
      End If
      
      Set Lt = New CLotItem
      
      Call Lt.PopulateFromRS(2, m_Rs)
      '                                5                                            Update ������� LotItem
      Call GenerateMovementItem(Lt, BalanceCollold, CapitalCollold, CapitalCollmovement, ExPortIDColl)
      
      Set Lt = Nothing
      
      m_Rs.MoveNext
   Wend
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   '                                   6                                            Save UpDate ŧ DataBase
   Call SaveCMToDbBill(CapitalCollmovement, ToDate, Parent, ParentEx)
   
   If Not (Parent Is Nothing) Then
      Parent.Value = 100
      ParentEx.Text = 100
   End If
   
   Set Lt = Nothing
   Set Cm = Nothing
   
   Set BalanceCollold = Nothing
   Set CapitalCollold = Nothing
   Set CapitalCollmovement = Nothing
   
   Set m_Rs = Nothing
   Set ExPortIDColl = Nothing
   
   UpdateCapitalMovement = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
   End If
   
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.RoutineName = "UpdateCapitalMovement"
   glbErrorLog.ModuleName = "FrmProcessCommit"
   glbErrorLog.LocalErrorMsg = "Eror"
   glbErrorLog.ShowErrorLog (LOG_MSGBOX)
   
   Set Lt = Nothing
   Set Cm = Nothing
   
   Set BalanceCollold = Nothing
   Set CapitalCollold = Nothing
   Set CapitalCollmovement = Nothing
   Set ExPortIDColl = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   UpdateCapitalMovement = False
End Function
Private Sub GenerateMovementItem(Lt As CLotItem, BalanceCollold As Collection, CapitalCollold As Collection, CapitalCollmovement As Collection, ExPortIDColl As Collection)
Dim BalanceLt As CLotItem
Dim PrevAmount As Double      '�ʹ�����͹�ǡź�ӹǹ����
Dim ExID As CExportID
Dim Sc As CStockCode
   
   Set BalanceLt = GetObject("CLotItem", BalanceCollold, Trim(Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID), False)
   '      ' Load �ʹ¡��
   If BalanceLt Is Nothing Then
      PrevAmount = 0
      Set BalanceLt = New CLotItem
      BalanceLt.PART_ITEM_ID = Lt.PART_ITEM_ID
      BalanceLt.LOCATION_ID = Lt.LOCATION_ID
      BalanceLt.SUM_AMOUNT = Lt.SUM_AMOUNT                                    '�ʹ � �Ѩ�غѹ �ͧ �����ѵ�شԺ
      Call BalanceCollold.add(BalanceLt, Trim(Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID))
   Else
      PrevAmount = BalanceLt.SUM_AMOUNT
      BalanceLt.SUM_AMOUNT = BalanceLt.SUM_AMOUNT + Lt.SUM_AMOUNT       '�ʹ � �Ѩ�غѹ �ͧ �����ѵ�شԺ
   End If
   
   Dim CmMovement As CCapitalMovement
   If BalanceLt.SUM_AMOUNT >= 0 Then
      Set CmMovement = GetObject("CCapitalMovement", CapitalCollold, Trim(Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID), False)
      If CmMovement Is Nothing Then
         Set CmMovement = New CCapitalMovement
         If Lt.TX_TYPE = "I" And Lt.DOCUMENT_TYPE = TRANSFER_DOCTYPE Then
            '�Ҥ���������� = ((�Ҥ��������� X �ӹǹ���) + (�Ҥҷ���ըҡ����͹�͡ X �ӹǹ����͹���)/(�ӹǹ��� + �ӹǹ����͹���))
            Set ExID = GetObject("CExportID", ExPortIDColl, Lt.INVENTORY_DOC_ID & "-" & Lt.LINK_ID)
            If ExID.AVG_PRICE > 0 Then
               CmMovement.CAPITAL_AMOUNT = MyDiffEx(((CmMovement.CAPITAL_AMOUNT * MinustoZero(PrevAmount)) + (ExID.AVG_PRICE * Lt.TX_AMOUNT)), MinustoZero(PrevAmount) + Lt.TX_AMOUNT)
               Lt.AVG_PRICE = ExID.AVG_PRICE
               Lt.TOTAL_INCLUDE_PRICE = ExID.AVG_PRICE * Lt.TX_AMOUNT
            ElseIf CmMovement.CAPITAL_AMOUNT > 0 Then
               Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
               Lt.TOTAL_INCLUDE_PRICE = CmMovement.CAPITAL_AMOUNT * Lt.TX_AMOUNT
            Else
               Set Sc = GetObject("CStockcode", m_Products, Trim(Str(Lt.PART_ITEM_ID)))
               CmMovement.CAPITAL_AMOUNT = Sc.COST_PER_AMOUNT                                    '�ʹ � �Ѩ�غѹ �ͧ �����ѵ�شԺ
            End If
            Set ExID = Nothing
         ElseIf Lt.AVG_PRICE > 0 And Lt.TX_TYPE = "I" Then
            CmMovement.CAPITAL_AMOUNT = Lt.AVG_PRICE                                    '�ʹ � �Ѩ�غѹ �ͧ �����ѵ�شԺ
         Else
            Set Sc = GetObject("CStockcode", m_Products, Trim(Str(Lt.PART_ITEM_ID)))
            CmMovement.CAPITAL_AMOUNT = Sc.COST_PER_AMOUNT                                    '�ʹ � �Ѩ�غѹ �ͧ �����ѵ�شԺ
         End If
         CmMovement.PART_ITEM_ID = Lt.PART_ITEM_ID
         CmMovement.LOCATION_ID = Lt.LOCATION_ID
         Call CapitalCollold.add(CmMovement, Trim(Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID))
      Else
         If Lt.DOCUMENT_TYPE = IMPORT_DOCTYPE Or Lt.DOCUMENT_TYPE = 11 Or Lt.DOCUMENT_TYPE = 22 Or Lt.DOCUMENT_TYPE = 30 Then
            ' #11 ��Ѻ�Թ��� #22 ����稫���ʴ #30 ��Ѻ�׹�Թ��Ң��
            '�Ҥ���������� = ((�Ҥ��������� X �ӹǹ���) + (�Ҥ����� X �������)/(�ӹǹ��� + �ӹǹ�������))
            '��ҵ鹷ع����� 0 ���ըӹǹ�Թ��Ҥ�����������Ҥ�����ͧ�Թ��ҷ���Ѻ��������Ҥҵ鹷ع�ͧ�Թ���������
            If CmMovement.CAPITAL_AMOUNT > 0 Then
               If Lt.AVG_PRICE > 0 Then         'Update �鹷ع੾�Сóշ���Ѻ��������Ѻ�׹�Թ��������յ鹷ع�Ѻ�׹
                  CmMovement.CAPITAL_AMOUNT = MyDiffEx((CmMovement.CAPITAL_AMOUNT * MinustoZero(PrevAmount)) + (Lt.AVG_PRICE * Lt.TX_AMOUNT), MinustoZero(PrevAmount) + Lt.TX_AMOUNT)
               End If
            Else
               CmMovement.CAPITAL_AMOUNT = Lt.AVG_PRICE
            End If
         ElseIf (Lt.DOCUMENT_TYPE = ADJUST_DOCTYPE And Lt.TX_TYPE = "I") Then
   '         If Lt.AVG_PRICE <= 0 Then
               '��һ�Ѻ�ʹ���������������Ҥ��Ѻ��Ҩж������Ҥ����Ҥ����
               Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
               Lt.TOTAL_INCLUDE_PRICE = Lt.AVG_PRICE * Lt.TX_AMOUNT
   '         Else
               '�Ҥ���������� = ((�Ҥ��������� X �ӹǹ���) + (�Ҥ����� X �������)/(�ӹǹ��� + �ӹǹ�������))
               '��ҵ鹷ع����� 0 ���ըӹǹ�Թ��Ҥ�����������Ҥ�����ͧ�Թ��ҷ���Ѻ��������Ҥҵ鹷ع�ͧ�Թ���������
   '            If CmMovement.CAPITAL_AMOUNT > 0 Then
   '               If Lt.AVG_PRICE > 0 Then         'Update �鹷ع੾�Сóշ���Ѻ��������Ѻ�׹�Թ��������յ鹷ع�Ѻ�׹
   '                  CmMovement.CAPITAL_AMOUNT = MyDiffEx((CmMovement.CAPITAL_AMOUNT * MinustoZero(PrevAmount)) + (Lt.AVG_PRICE * Lt.TX_AMOUNT), MinustoZero(PrevAmount) + Lt.TX_AMOUNT)
   '               End If
   '            Else
   '               CmMovement.CAPITAL_AMOUNT = Lt.AVG_PRICE
   '            End If
   '         End If
         ElseIf (Lt.DOCUMENT_TYPE = 1000 And Lt.TX_TYPE = "I") Then
            '�ѧ�����Դ�鹷ع��ü�Ե
            Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
            Lt.TOTAL_INCLUDE_PRICE = Lt.AVG_PRICE * Lt.TX_AMOUNT
         ElseIf Lt.DOCUMENT_TYPE = EXPORT_DOCTYPE Or (Lt.DOCUMENT_TYPE = ADJUST_DOCTYPE And Lt.TX_TYPE = "E") Or _
         Lt.DOCUMENT_TYPE = 10 Or Lt.DOCUMENT_TYPE = 21 Or Lt.DOCUMENT_TYPE = 31 Or (Lt.DOCUMENT_TYPE = 1000 And Lt.TX_TYPE = "E") Then
            '#10 ��觢ͧ #21 ����稢��ʴ #31㺤׹�Թ��ҫ��� #1000 㺼�Ե
            'Export �Ҥ�����������������¹�ŧ
         ElseIf Lt.DOCUMENT_TYPE = TRANSFER_DOCTYPE Then
            '�óշ���͹���������ҧ��ѧ �鹷ع�ж١����� �� ��� ��ҤԴ��ҵ鹷ع�ء��ѧ�ͧ�Թ����դ������Ǵѧ��� �����Թ�������ǡѹ������ѧ������ռŵ���Ҥ����������
            '��ǹ����͹Ẻ����¹�ѵ�شԺ����ѧ����ͧ�͹�ѵ�شԺ���
            If Lt.TX_TYPE = "E" Then
               Set ExID = New CExportID
               ExID.INVENTORY_DOC_ID = Lt.INVENTORY_DOC_ID
               ExID.LINK_ID = Lt.LINK_ID
               ExID.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
               
               Call ExPortIDColl.add(ExID, ExID.GetKey1)
               Set ExID = Nothing
            ElseIf Lt.TX_TYPE = "I" Then
               '�Ҥ���������� = ((�Ҥ��������� X �ӹǹ���) + (�Ҥҷ���ըҡ����͹�͡ X �ӹǹ����͹���)/(�ӹǹ��� + �ӹǹ����͹���))
               Set ExID = GetObject("CExportID", ExPortIDColl, Lt.INVENTORY_DOC_ID & "-" & Lt.LINK_ID)
               If ExID.AVG_PRICE > 0 Then
                  CmMovement.CAPITAL_AMOUNT = MyDiffEx(((CmMovement.CAPITAL_AMOUNT * MinustoZero(PrevAmount)) + (ExID.AVG_PRICE * Lt.TX_AMOUNT)), MinustoZero(PrevAmount) + Lt.TX_AMOUNT)
                  Lt.AVG_PRICE = ExID.AVG_PRICE
                  Lt.TOTAL_INCLUDE_PRICE = ExID.AVG_PRICE * Lt.TX_AMOUNT
               Else
                  Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
                  Lt.TOTAL_INCLUDE_PRICE = CmMovement.CAPITAL_AMOUNT * Lt.TX_AMOUNT
               End If
               Set ExID = Nothing
            End If
         End If
      End If
   Else
      Set CmMovement = GetObject("CCapitalMovement", CapitalCollold, Trim(Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID), False)
      If CmMovement Is Nothing Then
         Set CmMovement = New CCapitalMovement
      End If
   End If
   
   Lt.ShowMode = SHOW_EDIT
   If Lt.TX_TYPE = "I" Then
      Lt.NEW_AVG_PRICE = CmMovement.CAPITAL_AMOUNT
      Call Lt.UpdateAvg
      If Lt.DOCUMENT_TYPE = 30 Then  '�Ѻ�׹
         Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
         Lt.TOTAL_INCLUDE_PRICE = Lt.TX_AMOUNT * Lt.AVG_PRICE
         Call UpdateCapitalToDocItem(Lt)
      End If
   Else
      Lt.AVG_PRICE = CmMovement.CAPITAL_AMOUNT
      Lt.TOTAL_INCLUDE_PRICE = Lt.TX_AMOUNT * Lt.AVG_PRICE
      
      Lt.NEW_AVG_PRICE = Lt.AVG_PRICE
      Call Lt.UpdateAvg
      
      Call UpdateCapitalToDocItem(Lt)
   End If

Dim CmMovementDate As CCapitalMovement
   Set CmMovementDate = GetObject("CCapitalMovement", CapitalCollmovement, Trim(Lt.DOCUMENT_DATE & "-" & Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID), False)
   If CmMovementDate Is Nothing Then
      Set CmMovementDate = New CCapitalMovement
      CmMovementDate.DOCUMENT_DATE = Lt.DOCUMENT_DATE
      CmMovementDate.PART_ITEM_ID = Lt.PART_ITEM_ID
      CmMovementDate.LOCATION_ID = Lt.LOCATION_ID
      CmMovementDate.CAPITAL_AMOUNT = CmMovement.CAPITAL_AMOUNT
      Call CapitalCollmovement.add(CmMovementDate, Trim(Lt.DOCUMENT_DATE & "-" & Lt.LOCATION_ID & "-" & Lt.PART_ITEM_ID))
   Else
      CmMovementDate.CAPITAL_AMOUNT = CmMovement.CAPITAL_AMOUNT
   End If
      
      
   Set CmMovement = Nothing
   Set CmMovementDate = Nothing
   
   
   Set BalanceLt = Nothing
End Sub
Private Sub SaveCMToDbBill(CapitalCollmovement As Collection, Optional ToDate As Date = -1, Optional Parent As Object = Nothing, Optional ParentEx As Object = Nothing)
Dim Cm As CCapitalMovement
Dim ID As Long
Dim TempCM As CCapitalMovement
Dim FoundFlag As Boolean
Dim I  As Long
Dim k As Long
   If ToDate = -1 Then
      ToDate = Now
   End If
   
   I = 0
   For Each Cm In CapitalCollmovement
      I = I + 1
      If Not (Parent Is Nothing) Then
         Parent.Value = (MyDiff(I, CapitalCollmovement.Count) * 25) + 75
         ParentEx.Text = FormatNumber(Parent.Value)
         Parent.Refresh
         ParentEx.Refresh
         DoEvents
      End If
      
      FoundFlag = False
      Cm.AddEditMode = SHOW_ADD
      ''debug.print (CM.GetFieldValue("DOCUMENT_DATE") &  "-" & CM.GetFieldValue("PART_ITEM_ID"))
      If Cm.DOCUMENT_DATE >= uctlFromDate.ShowDate Then
         Call Cm.AddEditData
      End If
      k = 0
      While DateAdd("D", k + 1, Cm.DOCUMENT_DATE) <= ToDate And Not (FoundFlag)
         k = k + 1
         Set TempCM = GetObject("CCapitalMovement", CapitalCollmovement, DateAdd("D", k, Cm.DOCUMENT_DATE) & "-" & Cm.LOCATION_ID & "-" & Cm.PART_ITEM_ID, False)
         If TempCM Is Nothing Then
            Set TempCM = New CCapitalMovement
            TempCM.AddEditMode = SHOW_ADD
            TempCM.PART_ITEM_ID = Cm.PART_ITEM_ID
            TempCM.LOCATION_ID = Cm.LOCATION_ID
            TempCM.DOCUMENT_DATE = DateAdd("D", k, Cm.DOCUMENT_DATE)
            
            TempCM.CAPITAL_AMOUNT = Cm.CAPITAL_AMOUNT
                        
            Call TempCM.AddEditData
         Else
            FoundFlag = True
         End If
         Set TempCM = Nothing
      Wend
   Next Cm
   
   Set CapitalCollmovement = Nothing
   
End Sub
Private Sub UpdateCapitalToDocItem(Lt As CLotItem)
Dim TempLt As CLotItem
Dim Di As CDocItem
   Set TempLt = GetObject("CLotItem", DocLinkLot, Trim(Str(Lt.LOT_ITEM_ID)), False)
   If Not (TempLt Is Nothing) Then
      Set Di = New CDocItem
      Di.ShowMode = SHOW_EDIT
      Di.DOC_ITEM_ID = TempLt.DOC_ITEM_ID
      Di.CAPITAL_AMOUNT = Lt.AVG_PRICE
      Di.TOTAL_INCLUDE_PRICE = Lt.TOTAL_INCLUDE_PRICE
      Call Di.UpdateCapitalSell
   End If
End Sub
Private Function MinustoZero(Amount As Double) As Double
   If Amount > 0 Then
      MinustoZero = Amount
   Else
      MinustoZero = 0
   End If
End Function
Private Sub CopyOldToMovement(Old As Collection, Movement As Collection)
Dim Cm As CCapitalMovement
Dim CmMovementDate As CCapitalMovement
   For Each Cm In Old
      Set CmMovementDate = New CCapitalMovement
      CmMovementDate.DOCUMENT_DATE = Cm.DOCUMENT_DATE
      CmMovementDate.PART_ITEM_ID = Cm.PART_ITEM_ID
      CmMovementDate.LOCATION_ID = Cm.LOCATION_ID
      CmMovementDate.CAPITAL_AMOUNT = Cm.CAPITAL_AMOUNT
      Call Movement.add(CmMovementDate, Trim(Cm.DOCUMENT_DATE & "-" & Cm.LOCATION_ID & "-" & Cm.PART_ITEM_ID))
   Next Cm
End Sub

