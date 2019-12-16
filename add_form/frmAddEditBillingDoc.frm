VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditBillingDoc 
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   Icon            =   "frmAddEditBillingDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   13260
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   1560
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   33
      Top             =   -480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10455
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   18441
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin GridEX20.GridEX GridEX1 
         Height          =   3945
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4200
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   6959
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
         Column(1)       =   "frmAddEditBillingDoc.frx":27A2
         Column(2)       =   "frmAddEditBillingDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditBillingDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditBillingDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditBillingDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditBillingDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditBillingDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditBillingDoc.frx":2D5E
      End
      Begin VB.ComboBox cboDocumentReturn 
         Height          =   315
         Left            =   11040
         Style           =   2  'Dropdown List
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3645
         Width           =   13755
         _ExtentX        =   24262
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
      Begin VB.ComboBox cboBranchAddress 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   5385
      End
      Begin VB.ComboBox cboLocationSale 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1755
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   5385
      End
      Begin VB.ComboBox cboAparMasAddress 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   5385
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   11040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1515
      End
      Begin Xivess.uctlTextLookup uctlAparMasLookup 
         Height          =   435
         Left            =   1740
         TabIndex        =   5
         Top             =   1200
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Xivess.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   5040
         TabIndex        =   2
         Top             =   750
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2250
         TabIndex        =   0
         Top             =   750
         Width           =   2055
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
      Begin Threed.SSPanel pnlHeader 
         Height          =   615
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   1085
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtCredit 
         Height          =   435
         Left            =   8400
         TabIndex        =   9
         Top             =   1680
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin Xivess.uctlDate uctlDueDate 
         Height          =   405
         Left            =   8400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextBox txtReferText 
         Height          =   435
         Left            =   9240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   795
         _ExtentX        =   4154
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtNote 
         Height          =   435
         Left            =   11280
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3120
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtPoNo 
         Height          =   435
         Left            =   11040
         TabIndex        =   15
         Top             =   2640
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1185
         Left            =   0
         TabIndex        =   46
         Top             =   8160
         Visible         =   0   'False
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   2090
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Xivess.uctlTextBox txtIncludeDiscount 
            Height          =   435
            Left            =   5280
            TabIndex        =   47
            Top             =   120
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtTotalPrice 
            Height          =   435
            Left            =   1080
            TabIndex        =   48
            Top             =   120
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtItemDiscount 
            Height          =   435
            Left            =   3360
            TabIndex        =   49
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtExtraDiscountPercent 
            Height          =   435
            Left            =   8250
            TabIndex        =   50
            Top             =   120
            Width           =   555
            _ExtentX        =   1191
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtExtraDiscountAmount 
            Height          =   435
            Left            =   8760
            TabIndex        =   51
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtVatPercent 
            Height          =   435
            Left            =   8250
            TabIndex        =   52
            Top             =   600
            Width           =   555
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtVatAmount 
            Height          =   435
            Left            =   8760
            TabIndex        =   53
            Top             =   600
            Width           =   1035
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtTotal 
            Height          =   435
            Left            =   10680
            TabIndex        =   54
            Top             =   600
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtTotalAmount 
            Height          =   435
            Left            =   1080
            TabIndex        =   55
            Top             =   600
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtIncludeDiscountEx 
            Height          =   435
            Left            =   10680
            TabIndex        =   56
            Top             =   120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Threed.SSCommand cmdCalculate 
            Height          =   525
            Left            =   12960
            TabIndex        =   119
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditBillingDoc.frx":2F36
            ButtonStyle     =   3
         End
         Begin VB.Label lblItemDiscount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            TabIndex        =   63
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblTotalPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblIncludeDiscount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            TabIndex        =   61
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblVatPercent 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6960
            TabIndex        =   60
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label lblExtraDiscountPercent 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6960
            TabIndex        =   59
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label lblTOtal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9960
            TabIndex        =   58
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lblIncludeDiscountEx 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9840
            TabIndex        =   57
            Top             =   240
            Width           =   705
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1140
         Left            =   0
         TabIndex        =   64
         Top             =   8160
         Visible         =   0   'False
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   2011
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Xivess.uctlTextBox txtTotalDebt 
            Height          =   435
            Left            =   1080
            TabIndex        =   65
            Top             =   120
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtAdditionAmount 
            Height          =   435
            Left            =   6240
            TabIndex        =   66
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtPaidAmount 
            Height          =   435
            Left            =   3480
            TabIndex        =   67
            Top             =   120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtVatPercentEx 
            Height          =   435
            Left            =   1050
            TabIndex        =   71
            Top             =   600
            Width           =   555
            _ExtentX        =   1191
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtVatAmountEx 
            Height          =   435
            Left            =   1560
            TabIndex        =   72
            Top             =   600
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtTotalEx 
            Height          =   435
            Left            =   3480
            TabIndex        =   73
            Top             =   600
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtDebitAmount 
            Height          =   435
            Left            =   6240
            TabIndex        =   76
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtCreditAmount 
            Height          =   435
            Left            =   8280
            TabIndex        =   77
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtAfterDebitCredit 
            Height          =   435
            Left            =   10680
            TabIndex        =   80
            Top             =   600
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtAfterSubTract 
            Height          =   435
            Left            =   10680
            TabIndex        =   82
            Top             =   120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtSubTractAmount 
            Height          =   435
            Left            =   8280
            TabIndex        =   84
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
         End
         Begin VB.Label lblSubTractAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7440
            TabIndex        =   85
            Top             =   195
            Width           =   795
         End
         Begin VB.Label lblAfterSubTract 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9840
            TabIndex        =   83
            Top             =   195
            Width           =   795
         End
         Begin VB.Label lblAfterDebitCredit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9960
            TabIndex        =   81
            Top             =   675
            Width           =   675
         End
         Begin VB.Label lblDebitAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5280
            TabIndex        =   79
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lblCreditAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7440
            TabIndex        =   78
            Top             =   675
            Width           =   705
         End
         Begin VB.Label lblTOtalEx 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2640
            TabIndex        =   75
            Top             =   720
            Width           =   705
         End
         Begin VB.Label lblVatPercentEx 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   74
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblPaidAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2520
            TabIndex        =   70
            Top             =   195
            Width           =   915
         End
         Begin VB.Label lblAdditionAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5325
            TabIndex        =   69
            Top             =   195
            Width           =   900
         End
         Begin VB.Label lblTotalDebt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            TabIndex        =   68
            Top             =   195
            Width           =   825
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   3945
         Left            =   0
         TabIndex        =   34
         Top             =   4200
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   6959
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboPaymentBy 
            Height          =   315
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   2160
            Width           =   2385
         End
         Begin VB.ComboBox cboDiscountType 
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   1680
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.ComboBox cboTranSportCycle 
            Height          =   315
            Left            =   1750
            Style           =   2  'Dropdown List
            TabIndex        =   103
            Top             =   1680
            Width           =   2385
         End
         Begin VB.ComboBox cboCarLicense 
            Height          =   315
            Left            =   1750
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   720
            Width           =   2385
         End
         Begin VB.ComboBox cboTranSportor 
            Height          =   315
            Left            =   1750
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1200
            Width           =   2385
         End
         Begin VB.ComboBox cboDriver 
            Height          =   315
            Left            =   1750
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   240
            Width           =   2385
         End
         Begin Xivess.uctlTextBox txtDiscountAmount 
            Height          =   375
            Left            =   4440
            TabIndex        =   105
            Top             =   1080
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlDate uctlGetDate 
            Height          =   405
            Left            =   1755
            TabIndex        =   107
            Top             =   2640
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin VB.Label lblGetDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   118
            Top             =   2700
            Width           =   1515
         End
         Begin VB.Label lblPaymentBy 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   117
            Top             =   2160
            Width           =   1515
         End
         Begin VB.Label lblDiscountType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            TabIndex        =   116
            Top             =   2040
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblTranSportCycle 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label lblCarLicense 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   110
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label lblTranSportor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   109
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label lblDriver 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   1515
         End
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   585
         Left            =   0
         TabIndex        =   86
         Top             =   9240
         Visible         =   0   'False
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Xivess.uctlTextBox txtWHPercent 
            Height          =   435
            Left            =   1050
            TabIndex        =   87
            Top             =   70
            Width           =   555
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtWHAmount 
            Height          =   435
            Left            =   1560
            TabIndex        =   88
            Top             =   70
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtGetAmount 
            Height          =   435
            Left            =   3480
            TabIndex        =   89
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtDifRcp 
            Height          =   435
            Left            =   6240
            TabIndex        =   90
            Top             =   75
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtFromCashTran 
            Height          =   435
            Left            =   10680
            TabIndex        =   91
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Xivess.uctlTextBox txtFeeAmount 
            Height          =   435
            Left            =   8280
            TabIndex        =   96
            Top             =   75
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
         End
         Begin VB.Label lblFeeAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7440
            TabIndex        =   97
            ToolTipText     =   "ค่าธรรมเนียม"
            Top             =   195
            Width           =   765
         End
         Begin VB.Label lblWHPercent 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblGetAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2640
            TabIndex        =   94
            Top             =   195
            Width           =   765
         End
         Begin VB.Label lblDifRcp 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5400
            TabIndex        =   93
            Top             =   195
            Width           =   765
         End
         Begin VB.Label lblFromCashTran 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9840
            TabIndex        =   92
            Top             =   195
            Width           =   765
         End
      End
      Begin Xivess.uctlTextBox txtReferDesc 
         Height          =   435
         Left            =   8400
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2640
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlSale 
         Height          =   435
         Left            =   1740
         TabIndex        =   98
         Top             =   3120
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlDealer 
         Height          =   435
         Left            =   6360
         TabIndex        =   115
         Top             =   3120
         Width           =   3825
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblDealer 
         Height          =   315
         Left            =   5640
         TabIndex        =   114
         Top             =   3240
         Width           =   675
      End
      Begin Threed.SSCommand cmdSaveConsignment 
         Height          =   525
         Left            =   6720
         TabIndex        =   113
         Top             =   9840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkConsignmentFlag 
         Height          =   435
         Left            =   12480
         TabIndex        =   112
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   99
         Top             =   3180
         Width           =   1395
      End
      Begin Threed.SSCheck ChkCancelFlag 
         Height          =   435
         Left            =   10200
         TabIndex        =   45
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDocumentreturn 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   9840
         TabIndex        =   44
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10200
         TabIndex        =   41
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label lblPoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10080
         TabIndex        =   40
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label lblReferText 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         TabIndex        =   39
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label lblDueDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         TabIndex        =   38
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         TabIndex        =   37
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   9240
         TabIndex        =   36
         Top             =   1800
         Width           =   375
      End
      Begin Threed.SSCheck chkTicketFlag 
         Height          =   435
         Left            =   11280
         TabIndex        =   3
         Top             =   720
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblLocationSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7200
         TabIndex        =   35
         Top             =   1290
         Width           =   1095
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   8640
         TabIndex        =   22
         Top             =   9840
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1740
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   750
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   9000
         TabIndex        =   4
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label lblAparMasAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   10200
         TabIndex        =   30
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblAparMas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1260
         Width           =   1395
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   28
         Top             =   810
         Width           =   555
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10515
         TabIndex        =   23
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12165
         TabIndex        =   24
         Top             =   9825
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1650
         TabIndex        =   20
         Top             =   9825
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   30
         TabIndex        =   19
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3300
         TabIndex        =   21
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBillingDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Top             =   810
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmAddEditBillingDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_HasDiscountModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Customers As Collection
Private m_Employees As Collection
Private m_Resources As Collection
Private m_APARCancelOutDocument As Collection
Private m_LocationSales As Collection
Private m_Apm  As CAPARMas
Private m_Emp As CEmployee
Private m_Mr As CMasterRef
Private m_Adr As CAddress

Private AutoPrintMode As Boolean

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public Area As Long
Public DocumentType As SELL_BILLING_DOCTYPE
Public DocumentText As String
Public DocumentSubType As Long
Public ReceiptType As Long
Public ConsignmentNo As String
Public ConsignmentID As Long
Private CheckConsignmentNo As Long

Private Programowner As String
Private FileName As String
Private m_SumUnit As Double
Private FromCustomerLookup As Boolean

Private m_Cd As Collection
Private DocAdd As Long
Private LotItemLinkCollection As Collection
Dim flagVerify As Boolean
'Private IsCalculate As Boolean

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_BillingDoc.BILLING_DOC_ID = ID
      m_BillingDoc.DOC_ITEM_TYPE = -1
      m_BillingDoc.CANCEL_FLAG = ""
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      uctlAparMasLookup.MyCombo.ListIndex = IDToListIndex(uctlAparMasLookup.MyCombo, m_BillingDoc.APAR_MAS_ID)
      cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.DEPARTMENT_ID)
      cboAparMasAddress.ListIndex = IDToListIndex(cboAparMasAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)
      txtCredit.Text = m_BillingDoc.CREDIT
      DocumentSubType = m_BillingDoc.DOCUMENT_SUB_TYPE
      uctlDueDate.ShowDate = m_BillingDoc.Due_Date
      txtItemDiscount.Text = m_BillingDoc.DISCOUNT_AMOUNT
      txtExtraDiscountAmount.Text = m_BillingDoc.EXT_DISCOUNT_AMOUNT
      txtExtraDiscountPercent.Text = m_BillingDoc.EXT_DISCOUNT_PERCENT
      txtAdditionAmount.Text = m_BillingDoc.ADDITION_AMOUNT
      txtTotalAmount.Text = m_BillingDoc.TOTAL_AMOUNT
      txtTotalPrice.Text = m_BillingDoc.TOTAL_PRICE
      txtCredit.Text = m_BillingDoc.CREDIT
      uctlDueDate.ShowDate = m_BillingDoc.Due_Date
      If Len(Trim(m_BillingDoc.CONSIGNMENT_NO)) > 0 Then
         CheckConsignmentNo = 1
      Else
         CheckConsignmentNo = 0
      End If
      ConsignmentNo = m_BillingDoc.CONSIGNMENT_NO
      ConsignmentID = m_BillingDoc.CONSIGNMENT_ID
      
      
      If Val(txtTotalPrice.Text) > 0 Then
         txtVatPercent.Text = m_BillingDoc.VAT_PERCENT
         txtVatAmount.Text = m_BillingDoc.VAT_AMOUNT
      Else
         txtVatPercentEx.Text = m_BillingDoc.VAT_PERCENT
         txtVatAmountEx.Text = m_BillingDoc.VAT_AMOUNT
      End If
      
      txtWHPercent.Text = m_BillingDoc.WH_PERCENT
      txtWHAmount.Text = m_BillingDoc.WH_AMOUNT
      txtTotalDebt.Text = m_BillingDoc.PAY_AMOUNT
      txtPaidAmount.Text = m_BillingDoc.PAID_AMOUNT
      
      txtCreditAmount.Text = m_BillingDoc.CREDIT_AMOUNT
      txtDebitAmount.Text = m_BillingDoc.DEBIT_AMOUNT
      txtFeeAmount.Text = m_BillingDoc.FEE_AMOUNT
      
      txtNote.Text = m_BillingDoc.NOTE
      txtReferText.Text = m_BillingDoc.REFER_TEXT
      txtReferDesc.Text = m_BillingDoc.REFER_DESC
      cboLocationSale.ListIndex = IDToListIndex(cboLocationSale, m_BillingDoc.CUSTOMER_BRANCH)
      uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_BillingDoc.SALE_BY)
      txtPoNo.Text = m_BillingDoc.CUS_PO
      cboBranchAddress.ListIndex = IDToListIndex(cboBranchAddress, m_BillingDoc.BRANCH_ADDRESS)
      cboDocumentReturn.ListIndex = IDToListIndex(cboDocumentReturn, m_BillingDoc.DOCUMENT_RETURN)
      
      cboDriver.ListIndex = IDToListIndex(cboDriver, m_BillingDoc.DRIVER_ID)
      cboCarLicense.ListIndex = IDToListIndex(cboCarLicense, m_BillingDoc.CAR_LICENSE_ID)
      cboTranSportor.ListIndex = IDToListIndex(cboTranSportor, m_BillingDoc.TRANSPORTOR_ID)
      cboTranSportCycle.ListIndex = IDToListIndex(cboTranSportCycle, m_BillingDoc.TRANSPORT_CYCLE)
'      cboDiscountType.ListIndex = IDToListIndex(cboDiscountType, m_BillingDoc.DISCOUNT_TYPE_ID)
'      txtDiscountAmount.Text = m_BillingDoc.DISCOUNT_VALUE
      uctlDealer.MyCombo.ListIndex = IDToListIndex(uctlDealer.MyCombo, m_BillingDoc.DEALER_ID)
      cboPaymentBy.ListIndex = IDToListIndex(cboPaymentBy, m_BillingDoc.PAYMENT_BY)
      uctlGetDate.ShowDate = m_BillingDoc.GET_DATE
      
      chkCommit.Value = FlagToCheck(m_BillingDoc.OLD_COMMIT_FLAG)
      chkCommit.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      chkTicketFlag.Value = FlagToCheck(m_BillingDoc.TICKET_FLAG)
      ChkCancelFlag.Value = FlagToCheck(m_BillingDoc.CANCEL_FLAG)
      chkConsignmentFlag.Value = FlagToCheck(m_BillingDoc.CONSIGNMENT_FLAG)
'      cmdCalculate.Enabled = (m_BillingDoc.IS_CALCULATE = "N")
      Call EnableDisableButton(True)
      
      If DocumentType = RECEIPT2_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Then
         Call glbDaily.CreateBillingTransferItems(m_BillingDoc)
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CDocItem

   For Each Di In BD.DocItems
      If Di.Flag = "A" Then
         Call Di.SetFieldValue("LINK_ID", GetNextGuiID(BD))
      End If
   Next Di
End Sub
Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDocItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DocItems
      If Di.GetFieldValue("LINK_ID") > MaxId Then
         MaxId = Di.GetFieldValue("LINK_ID")
      End If
   Next Di

   GetNextGuiID = MaxId + 1
End Function
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim CRcp As CRcpCnDn_Item

   If Not (cmdOK.Enabled) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblAparMas, uctlAparMasLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDueDate, uctlDueDate, True) Then
      Exit Function
   End If
   
   If Not VerifyLockDate(uctlDocumentDate.ShowDate, m_BillingDoc.DOCUMENT_DATE) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      
      Exit Function
   End If
   
   If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = QUOATATION_DOCTYPE Or DocumentType = RETURN_DOCTYPE _
   Or DocumentType = S_PO_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE Or DocumentType = S_QUOATATION_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      If Not VerifyLockInvoiceDate(uctlDocumentDate.ShowDate, m_BillingDoc.DOCUMENT_DATE) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   Else
      If Not VerifyLockReceiptDate(uctlDocumentDate.ShowDate, m_BillingDoc.DOCUMENT_DATE) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   
   If Not CheckUniqueNs(DOCUMENT_NO_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   
     If (Len(txtPoNo.Text) > 0) Then
      If Not CheckUniqueNs(CUS_PO_UNIQUE, Trim(txtPoNo.Text), ID, Trim(Str(DocumentType))) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPoNo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   
   If DocumentType = RETURN_DOCTYPE Then
        If (Len(txtReferText.Text) > 0) Then
         If Not CheckUniqueNs(CUS_REFER_UNIQUE, Trim(txtReferText.Text), ID, Trim(Str(DocumentType))) Then
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtReferText.Text & " " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If DocumentType = RECEIPT2_DOCTYPE Or DocumentType = BILLS_DOCTYPE _
      Or DocumentType = RECEIPT2_DOCTYPE Or DocumentType = BILLS_DOCTYPE Then
      For Each CRcp In m_BillingDoc.RcpCnDnItems
         If CRcp.GetFieldValue("PAID_AMOUNT") <= 0 Then
            Call MsgBox("มีข้อมูลหมายเอกสาร " & CRcp.GetFieldValue("DOC_NO") & " ที่ไม่ได้ใส่จำนวนรับชำระ", vbOKOnly, PROJECT_NAME)
            Exit Function
         End If
      Next
   End If
   
   Dim DocItem As CDocItem
   Dim Pl As CPrintLabel
   
   For Each DocItem In m_BillingDoc.DocItems
      If DocItem.Flag = "I" Then
         DocItem.Flag = "E"
      End If
      Call DocItem.SetFieldValue("EXTRA_DISCOUNT_AMOUNT", FormatNumber((DocItem.GetFieldValue("TOTAL_PRICE") - DocItem.GetFieldValue("DISCOUNT_AMOUNT")) * Val(txtExtraDiscountPercent.Text) / 100, , False))
      DocItem.EXTRA_DISCOUNT_AMOUNT_DIF = ((DocItem.GetFieldValue("TOTAL_PRICE") - DocItem.GetFieldValue("DISCOUNT_AMOUNT")) * Val(txtExtraDiscountPercent.Text) / 100) - (FormatNumber((DocItem.GetFieldValue("TOTAL_PRICE") - DocItem.GetFieldValue("DISCOUNT_AMOUNT")) * Val(txtExtraDiscountPercent.Text) / 100, , False))
      'ส่วนต่างระหว่าง ไม่ปัดกับปัด
      
      For Each Pl In DocItem.PrintLabels
         If Pl.Flag = "I" Then
            Pl.Flag = "E"
         End If
         Call Pl.SetFieldValue("DISCOUNT_AMOUNT", FormatNumber(Pl.GetFieldValue("TOTAL_PRICE") * DocItem.GetFieldValue("DISCOUNT_PERCENT") / 100, , False))
         Call Pl.SetFieldValue("EXTRA_DISCOUNT_AMOUNT", FormatNumber((Pl.GetFieldValue("TOTAL_PRICE") - Pl.GetFieldValue("DISCOUNT_AMOUNT")) * Val(txtExtraDiscountPercent.Text) / 100, , False))
      Next Pl
   Next DocItem
   Set DocItem = Nothing
      
   If m_BillingDoc.DocItems.Count > 0 Then
      Call CheckDisCount
   End If
   
   If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      Dim Doc As CDocItem
      Dim RefOld As Long
      RefOld = 0
      For Each Doc In m_BillingDoc.DocItems
         If Doc.GetFieldValue("PO_ID") < 0 Or (RefOld <> Doc.GetFieldValue("PO_ID") And RefOld > 0) Then
            glbErrorLog.LocalErrorMsg = "สินค้าแต่ละรายการอ้างถึงใบกำกับต่างกัน"
            glbErrorLog.ShowUserError
            Exit Function
          End If
         If Doc.GetFieldValue("PO_ID") > 0 Then
            m_BillingDoc.SR_REF_DO_ID = Doc.GetFieldValue("PO_ID")
            m_BillingDoc.SR_REF_DO_NO = Doc.GetFieldValue("PO_NO")
            RefOld = Doc.GetFieldValue("PO_ID")
         Else
            m_BillingDoc.SR_REF_DO_ID = -1
            m_BillingDoc.SR_REF_DO_NO = ""
         End If
      Next Doc
      Set Doc = Nothing
   End If
   
   If (DocumentType = RECEIPT2_DOCTYPE) Or (DocumentType = RECEIPT1_DOCTYPE) _
   Or (DocumentType = S_RECEIPT2_DOCTYPE) Or (DocumentType = S_RECEIPT1_DOCTYPE) Then
      Call CreateBillDocCashTranItems
   End If
   
   m_BillingDoc.ShowMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_BillingDoc.APAR_MAS_ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
   m_BillingDoc.DEPARTMENT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   m_BillingDoc.DOCUMENT_TYPE = DocumentType
   If DocumentSubType > 1 Then
      m_BillingDoc.DOCUMENT_SUB_TYPE = DocumentSubType
   Else
      m_BillingDoc.DOCUMENT_SUB_TYPE = -1
   End If
   m_BillingDoc.SALE_BY = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   m_BillingDoc.BILLING_ADDRESS_ID = cboAparMasAddress.ItemData(Minus2Zero(cboAparMasAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.TICKET_FLAG = Check2Flag(chkTicketFlag.Value)
   m_BillingDoc.CANCEL_FLAG = Check2Flag(ChkCancelFlag.Value)
   m_BillingDoc.CONSIGNMENT_FLAG = Check2Flag(chkConsignmentFlag.Value)
   m_BillingDoc.DISCOUNT_AMOUNT = Val(txtItemDiscount.Text)
   m_BillingDoc.EXT_DISCOUNT_AMOUNT = Val(txtExtraDiscountAmount.Text)
   m_BillingDoc.EXT_DISCOUNT_PERCENT = Val(txtExtraDiscountPercent.Text)
   
   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_BillingDoc.TOTAL_PRICE = Val(txtTotalPrice.Text)
   m_BillingDoc.CREDIT = Val(txtCredit.Text)
   m_BillingDoc.Due_Date = uctlDueDate.ShowDate
   m_BillingDoc.CREDIT_AMOUNT = Val(txtCreditAmount.Text)
   m_BillingDoc.DEBIT_AMOUNT = Val(txtDebitAmount.Text)
   
   m_BillingDoc.RCP_CASH_TRAN = Val(txtFromCashTran.Text)
   m_BillingDoc.RCP_DIF = Val(txtDifRcp.Text)
   
   If Val(txtVatPercent.Text) > 0 Then
      m_BillingDoc.VAT_PERCENT = Val(txtVatPercent.Text)
   End If
   If Val(txtVatPercentEx.Text) > 0 Then
      m_BillingDoc.VAT_PERCENT = Val(txtVatPercentEx.Text)
   End If
   If Val(txtVatAmount.Text) > 0 Then
      m_BillingDoc.VAT_AMOUNT = Val(txtVatAmount.Text)
   End If
   If Val(txtVatAmountEx.Text) > 0 Then
      m_BillingDoc.VAT_AMOUNT = Val(txtVatAmountEx.Text)
   End If
   
   If Val(txtVatPercent.Text) = 0 And Val(txtVatPercentEx.Text) = 0 Then
      m_BillingDoc.VAT_PERCENT = 0
   End If
   If Val(txtVatAmount.Text) = 0 And Val(txtVatAmountEx.Text) = 0 Then
      m_BillingDoc.VAT_AMOUNT = 0
   End If
   
   m_BillingDoc.WH_PERCENT = Val(txtWHPercent.Text)
   m_BillingDoc.WH_AMOUNT = Val(txtWHAmount.Text)
   m_BillingDoc.PAID_AMOUNT = Val(txtPaidAmount.Text)
   m_BillingDoc.PAY_AMOUNT = Val(txtTotalDebt.Text)
   m_BillingDoc.SUBTRACT_AMOUNT = Val(txtSubTractAmount.Text)
   m_BillingDoc.ADDITION_AMOUNT = Val(txtAdditionAmount.Text)
   m_BillingDoc.NOTE = txtNote.Text
   m_BillingDoc.REFER_TEXT = txtReferText.Text
   m_BillingDoc.REFER_DESC = txtReferDesc.Text
   m_BillingDoc.CUSTOMER_BRANCH = cboLocationSale.ItemData(Minus2Zero(cboLocationSale.ListIndex))
   m_BillingDoc.CUS_PO = txtPoNo.Text
   m_BillingDoc.BRANCH_ADDRESS = cboBranchAddress.ItemData(Minus2Zero(cboBranchAddress.ListIndex))
   m_BillingDoc.DOCUMENT_RETURN = cboDocumentReturn.ItemData(Minus2Zero(cboDocumentReturn.ListIndex))
   m_BillingDoc.FEE_AMOUNT = Val(txtFeeAmount.Text)
   m_BillingDoc.DEALER_ID = uctlDealer.MyCombo.ItemData(Minus2Zero(uctlDealer.MyCombo.ListIndex))
   
   If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Then
      m_BillingDoc.DRIVER_ID = cboDriver.ItemData(Minus2Zero(cboDriver.ListIndex))
      m_BillingDoc.CAR_LICENSE_ID = cboCarLicense.ItemData(Minus2Zero(cboCarLicense.ListIndex))
      m_BillingDoc.TRANSPORTOR_ID = cboTranSportor.ItemData(Minus2Zero(cboTranSportor.ListIndex))
      m_BillingDoc.TRANSPORT_CYCLE = cboTranSportCycle.ItemData(Minus2Zero(cboTranSportCycle.ListIndex))
'      m_BillingDoc.DISCOUNT_TYPE_ID = cboDiscountType.ItemData(Minus2Zero(cboDiscountType.ListIndex))
'      m_BillingDoc.DISCOUNT_VALUE = Val(txtDiscountAmount.Text)
      m_BillingDoc.PAYMENT_BY = cboPaymentBy.ItemData(Minus2Zero(cboPaymentBy.ListIndex))
      m_BillingDoc.GET_DATE = uctlGetDate.ShowDate
   End If
   
   If DocumentType = PO_DOCTYPE Then
       m_BillingDoc.CONSIGNMENT_NO = ConsignmentNo
       m_BillingDoc.CONSIGNMENT_ID = ConsignmentID
   End If
   
'   If DocumentType = INVOICE_DOCTYPE Then
'      m_BillingDoc.IS_CALCULATE = IIf(IsCalculate, "Y", "N")
'   End If
         
   Call PopulateGuiID(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   Dim CheckBalance As Boolean
   CheckBalance = True
   If DocumentType = INVOICE_DOCTYPE And Area = 1 Then 'ใบส่งสินค้าขาย
      CheckBalance = glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 10)
   ElseIf DocumentType = S_INVOICE_DOCTYPE Then  'ใบรับสินค้าซื้อ
      Call glbDaily.RO2InventoryDoc(m_BillingDoc, Ivd, 11)
   ElseIf DocumentType = RECEIPT1_DOCTYPE And Area = 1 Then  'ใบเสร็จขายสด
      CheckBalance = glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 21)
   ElseIf DocumentType = S_RECEIPT1_DOCTYPE Then    'ใบเสร็จซื้อสด
      Call glbDaily.RO2InventoryDoc(m_BillingDoc, Ivd, 22)
   ElseIf DocumentType = RETURN_DOCTYPE And Area = 1 Then  'ใบรับคืนสินค้าขาย
      Call glbDaily.Rt2InventoryDoc(m_BillingDoc, Ivd, Area, 30)
   ElseIf DocumentType = S_RETURN_DOCTYPE And Area = 2 Then  'ใบส่งคืนสินค้าซื้อ
      CheckBalance = glbDaily.Rt2InventoryDoc(m_BillingDoc, Ivd, Area, 31)
   End If
   
   If Not CheckBalance Then
      SaveData = False
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   
   If DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE _
   Or DocumentType = RETURN_DOCTYPE _
   Or DocumentType = S_INVOICE_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE _
   Or DocumentType = S_RETURN_DOCTYPE Then
      
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      m_BillingDoc.INVENTORY_DOC_ID = Ivd.GetFieldValue("INVENTORY_DOC_ID") 'Ivd.GetFieldValue("INVENTORY_DOC_ID")
   End If
   
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   'If DocumentType = RECEIPT2_DOCTYPE Then
      'Call UpDateBDRcpCnDnItem(m_BillingDoc)  'ปิดไว้เนื่องจากว่ามนไม่จำเป็นที่จะต้องดูที่หน้านี้ดูจากรายงานน่าจะดีกว่า
   'End If
      
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub
Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboBranchAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboBranchAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboAparMasAddress_Click()
   m_HasModify = True
End Sub
Private Sub cboAparMasAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboCarLicense_Click()
   m_HasModify = True
End Sub
'Private Sub cboDiscountType_Click()
'   m_HasModify = True
'End Sub

'Private Sub cboDiscountType_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      SendKeys ("{TAB}")
'   End If
'End Sub

Private Sub cboDocumentReturn_Click()
   m_HasModify = True
End Sub

Private Sub cboDocumentReturn_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboDriver_Click()
   m_HasModify = True
End Sub

Private Sub cboDriver_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub
Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboLocationSale_Click()
Dim Branch As CMasterRef
Dim BranchID As Long

   If FromCustomerLookup Then
      Exit Sub
   End If

   BranchID = cboLocationSale.ItemData(Minus2Zero(cboLocationSale.ListIndex))
   
   If BranchID > 0 Then
      Set Branch = m_LocationSales(Trim(Str(BranchID)))
      
      If Branch.PARENT_EX_ID > 0 Then
         uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, Branch.PARENT_EX_ID)
      Else
         uctlSale.MyCombo.ListIndex = -1
      End If
      
      If Branch.PARENT_EX_ID3 > 0 Then
         cboBranchAddress.ListIndex = IDToListIndex(cboBranchAddress, Branch.PARENT_EX_ID3)
      Else
         cboBranchAddress.ListIndex = -1
      End If
      
      If Branch.DEALER_ID > 0 Then
         uctlDealer.MyCombo.ListIndex = IDToListIndex(uctlDealer.MyCombo, Branch.DEALER_ID)
      Else
         uctlDealer.MyCombo.ListIndex = -1
      End If
   End If
   
   Set Branch = Nothing
   m_HasModify = True
End Sub

Private Sub cboLocationSale_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkCancel_Click(Value As Integer)
   m_HasModify = True
End Sub


Private Sub cboPaymentBy_Click()
   m_HasModify = True
End Sub

Private Sub cboPaymentBy_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboTranSportCycle_Click()
   m_HasModify = True
End Sub

Private Sub cboTranSportCycle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboTransportor_Click()
   m_HasModify = True
End Sub

Private Sub cboTranSportor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkCancelFlag_Click(Value As Integer)

Dim strVerify As String
Dim TempUserName As String
If m_BillingDoc.CANCEL_FLAG = "Y" And m_HasModify = True Then
   Exit Sub
End If
If flagVerify = True Then
   flagVerify = False
   Exit Sub
End If
If Area = 1 Then
    strVerify = "LEDGER_SELL_" & DocumentText & "_CANCEL"
    If Not VerifyAccessRight(strVerify, "ยกเลิกเอกสาร") Then
            Call EnableForm(Me, True)
            flagVerify = True
   End If
ElseIf Area = 2 Then
    strVerify = "LEDGER_BUY_" & DocumentText & "_CANCEL"
    If Not VerifyAccessRight(strVerify, "ยกเลิกเอกสาร") Then
            Call EnableForm(Me, True)
            flagVerify = True
   End If
End If
If flagVerify = False Then 'ถ้ามีสิทธิ์ยกเลิกเอกสารได้
   m_HasModify = True
Else 'ถ้าไม่มีสิทธิ์ยกเลิกเอกสารให้แสดงช่อง สิทธิ์ยืนยัน
   frmVerifyAccRight.AccName = strVerify
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1

   If frmVerifyAccRight.GrantRight Then
      TempUserName = frmVerifyAccRight.UserName
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Call EnableForm(Me, True)
     ' Exit Sub
   End If

   If Not Len(TempUserName) > 0 Then
      If ChkCancelFlag.Value = ssCBChecked Then
         ChkCancelFlag.Value = ssCBUnchecked
      ElseIf ChkCancelFlag.Value = ssCBUnchecked Then
         ChkCancelFlag.Value = ssCBChecked
      End If
   Else
      m_HasModify = True
   End If
End If

   
End Sub
Private Sub ChkCancelFlag_KeyPress(KeyAscii As Integer)
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

Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
   GridEX1.Rebind
End Sub
Public Sub RefreshGridSub()
   Call GetTotalPriceReceipt

   GridEX1.itemcount = CountItem(m_BillingDoc.BillingSubTracts)
   GridEX1.Rebind
End Sub
Public Sub RefreshCashTran()
   Call GetTotalPriceReceipt

   GridEX1.itemcount = CountItem(m_BillingDoc.TransferItems)
   GridEX1.Rebind
End Sub

Private Sub chkConsignmentFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkTicketFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkTicketFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblAparMas, uctlAparMasLookup.MyCombo, False) Then
            Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = INVOICE_DOCTYPE & "-DTL" Or TabStrip1.SelectedItem.Tag = RECEIPT1_DOCTYPE & "-DTL" _
   Or TabStrip1.SelectedItem.Tag = S_INVOICE_DOCTYPE & "-DTL" Or TabStrip1.SelectedItem.Tag = S_RECEIPT1_DOCTYPE & "-DTL" Then
   
      Set oMenu = New CPopupMenu
      lMenuChosen = oMenu.Popup("เพิ่มรายการใหม่", "-", "เพิ่มจากใบ PO")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      If lMenuChosen = 1 Then
         frmAddEditDoItem.HeaderText = "เพิ่มรายการสินค้า"
         frmAddEditDoItem.ShowMode = SHOW_ADD
         frmAddEditDoItem.Area = Area
         frmAddEditDoItem.DocumentType = DocumentType
         frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
         Set frmAddEditDoItem.ParentForm = Me
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DocItems
         Set frmAddEditDoItem.LotItemLinkCollection = LotItemLinkCollection
         frmAddEditDoItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
         
         OKClick = frmAddEditDoItem.OKClick
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
         
         If OKClick Then
            m_HasModify = True
            GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
            GridEX1.Rebind
         
            Call GetTotalPrice
         End If
      ElseIf lMenuChosen = 3 Then
         frmAddPOItem.AparMasID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Set frmAddPOItem.TempCollection = m_BillingDoc.DocItems
         frmAddPOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddPOItem.Area = Area
         frmAddPOItem.ShowMode = SHOW_ADD
         frmAddPOItem.HeaderText = MapText("เพิ่มรายการจากใบ PO")
         Load frmAddPOItem
         frmAddPOItem.Show 1
   
         OKClick = frmAddPOItem.OKClick
         txtPoNo.Text = frmAddPOItem.AparMasPO
         txtVatPercent.Text = frmAddPOItem.VAT_PERCENT
         txtExtraDiscountPercent.Text = frmAddPOItem.EXT_DISCOUNT_PERCENT
         
         cboDriver.ListIndex = IDToListIndex(cboDriver, frmAddPOItem.DRIVER_ID)
         cboCarLicense.ListIndex = IDToListIndex(cboCarLicense, frmAddPOItem.CAR_LICENSE_ID)
         cboTranSportor.ListIndex = IDToListIndex(cboTranSportor, frmAddPOItem.TRANSPORTOR_ID)
         cboTranSportCycle.ListIndex = IDToListIndex(cboTranSportCycle, frmAddPOItem.TRANSPORT_CYCLE)
         
         cboPaymentBy.ListIndex = IDToListIndex(cboPaymentBy, frmAddPOItem.PAYMENT_BY)
         
         Unload frmAddPOItem
         Set frmAddPOItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
            GridEX1.Rebind
            
            m_HasModify = True
         End If
      
      End If
   ElseIf TabStrip1.SelectedItem.Tag = PO_DOCTYPE & "-DTL" Or TabStrip1.SelectedItem.Tag = S_PO_DOCTYPE & "-DTL" Then
      frmAddEditDoItem.HeaderText = "เพิ่มรายการสินค้า"
      frmAddEditDoItem.ShowMode = SHOW_ADD
      frmAddEditDoItem.Area = Area
      frmAddEditDoItem.DocumentType = DocumentType
      frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
      frmAddEditDoItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
      Set frmAddEditDoItem.ParentForm = Me
      Set frmAddEditDoItem.TempCollection = m_BillingDoc.DocItems
      
      Load frmAddEditDoItem
      frmAddEditDoItem.Show 1
      
      OKClick = frmAddEditDoItem.OKClick
      Unload frmAddEditDoItem
      Set frmAddEditDoItem = Nothing
      
      If OKClick Then
         m_HasModify = True
         GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
         GridEX1.Rebind
      
         Call GetTotalPrice
      End If
   ElseIf TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-RCP" Or TabStrip1.SelectedItem.Tag = BILLS_DOCTYPE & "-BILLS" _
   Or TabStrip1.SelectedItem.Tag = S_RECEIPT2_DOCTYPE & "-RCP" Or TabStrip1.SelectedItem.Tag = S_BILLS_DOCTYPE & "-BILLS" Then
      If TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-RCP" Or TabStrip1.SelectedItem.Tag = S_RECEIPT2_DOCTYPE & "-RCP" Then
         Set oMenu = New CPopupMenu
         lMenuChosen = oMenu.Popup("เพิ่มจากใบส่งสินค้า", "-", "เพิ่มจากใบวางบิล")
         If lMenuChosen = 0 Then
            Exit Sub
         End If
      End If
      If lMenuChosen = 1 Or TabStrip1.SelectedItem.Tag = BILLS_DOCTYPE & "-BILLS" Or TabStrip1.SelectedItem.Tag = S_BILLS_DOCTYPE & "-BILLS" Then
         frmAddReceiptItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Set frmAddReceiptItem.TempCollection = m_BillingDoc.RcpCnDnItems
         frmAddReceiptItem.DocumentType = DocumentType
         frmAddReceiptItem.ShowMode = SHOW_ADD
         frmAddReceiptItem.HeaderText = MapText("เพิ่มรายการ")
         
         Load frmAddReceiptItem
         frmAddReceiptItem.Show 1
         
         OKClick = frmAddReceiptItem.OKClick
         
         Unload frmAddReceiptItem
         Set frmAddReceiptItem = Nothing
      ElseIf lMenuChosen = 3 Then
         frmAddBillsItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Set frmAddBillsItem.TempCollection = m_BillingDoc.RcpCnDnItems
         frmAddBillsItem.DocumentType = DocumentType
         frmAddBillsItem.ShowMode = SHOW_ADD
         frmAddBillsItem.HeaderText = MapText("เพิ่มรายการ")
      
         Load frmAddBillsItem
         frmAddBillsItem.Show 1

         OKClick = frmAddBillsItem.OKClick

         Unload frmAddBillsItem
         Set frmAddBillsItem = Nothing
      
      End If
      
      If OKClick Then
         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = CN_DOCTYPE & "-CN" Or TabStrip1.SelectedItem.Tag = S_CN_DOCTYPE & "-CN" Then
      frmAddDebitCreditItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
      Set frmAddDebitCreditItem.TempCollection = m_BillingDoc.RcpCnDnItems
      frmAddDebitCreditItem.ShowMode = SHOW_ADD
      frmAddDebitCreditItem.HeaderText = MapText("เพิ่มรายการลดหนี้")
      
      Load frmAddDebitCreditItem
      frmAddDebitCreditItem.Show 1

      OKClick = frmAddDebitCreditItem.OKClick

      Unload frmAddDebitCreditItem
      Set frmAddDebitCreditItem = Nothing

      If OKClick Then
         Call GetTotalPriceDebitCredit

         GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DN_DOCTYPE & "-DN" Or TabStrip1.SelectedItem.Tag = S_DN_DOCTYPE & "-DN" Then
      frmAddDebitCreditItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
      Set frmAddDebitCreditItem.TempCollection = m_BillingDoc.RcpCnDnItems
      frmAddDebitCreditItem.ShowMode = SHOW_ADD
      frmAddDebitCreditItem.HeaderText = MapText("เพิ่มรายการเพิ่มหนี้")
      
      Load frmAddDebitCreditItem
      frmAddDebitCreditItem.Show 1

      OKClick = frmAddDebitCreditItem.OKClick

      Unload frmAddDebitCreditItem
      Set frmAddDebitCreditItem = Nothing

      If OKClick Then
         Call GetTotalPriceDebitCredit

         GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   
   ElseIf TabStrip1.SelectedItem.Tag = RETURN_DOCTYPE & "-DTL" Or TabStrip1.SelectedItem.Tag = S_RETURN_DOCTYPE & "-DTL" Then
      Set oMenu = New CPopupMenu
      lMenuChosen = oMenu.Popup("เพิ่มรายการคืนใหม่", "-", "เพิ่มรายการคืนจากใบ DO")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      If lMenuChosen = 1 Then
         frmAddEditDoItem.HeaderText = "เพิ่มรายการคืนสินค้า"
         frmAddEditDoItem.ShowMode = SHOW_ADD
         Set frmAddEditDoItem.ParentForm = Me
         frmAddEditDoItem.DocumentType = DocumentType
         frmAddEditDoItem.Area = Area
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DocItems
         frmAddEditDoItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
         
         OKClick = frmAddEditDoItem.OKClick
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
      Else
         frmAddReturnItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
         Set frmAddReturnItem.TempCollection = m_BillingDoc.DocItems
         frmAddReturnItem.DocumentType = DocumentType
         frmAddReturnItem.ShowMode = SHOW_ADD
         frmAddReturnItem.HeaderText = MapText("เพิ่มรายการจากใบ ส่งสินค้า")
         Load frmAddReturnItem
         frmAddReturnItem.Show 1
         
         OKClick = frmAddReturnItem.OKClick
         
         Unload frmAddReturnItem
         Set frmAddReturnItem = Nothing
      
      End If
      If OKClick Then
         Call GetTotalPrice

         GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-SUB" Or TabStrip1.SelectedItem.Tag = S_RECEIPT2_DOCTYPE & "-SUB" Then
      Set frmAddEditBillingSubTract.TempCollection = m_BillingDoc.BillingSubTracts
      frmAddEditBillingSubTract.ShowMode = SHOW_ADD
      Set frmAddEditBillingSubTract.ParentForm = Me
      frmAddEditBillingSubTract.HeaderText = MapText("เพิ่มรายการหัก")
     Load frmAddEditBillingSubTract
      frmAddEditBillingSubTract.Show 1
      
      OKClick = frmAddEditBillingSubTract.OKClick
      
      Unload frmAddEditBillingSubTract
      Set frmAddEditBillingSubTract = Nothing

      If OKClick Then
         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.BillingSubTracts)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-ADD" Or TabStrip1.SelectedItem.Tag = S_RECEIPT2_DOCTYPE & "-ADD" Then
      Set frmAddEditBillingAddition.TempCollection = m_BillingDoc.BillingAdditions
      frmAddEditBillingAddition.ShowMode = SHOW_ADD
      Set frmAddEditBillingAddition.ParentForm = Me
      frmAddEditBillingAddition.HeaderText = MapText("เพิ่มรายการส่วนเพิ่ม")
     Load frmAddEditBillingAddition
      frmAddEditBillingAddition.Show 1
      
      OKClick = frmAddEditBillingAddition.OKClick
      
      Unload frmAddEditBillingAddition
      Set frmAddEditBillingAddition = Nothing

      If OKClick Then
         Call GetTotalPriceReceipt
         
         GridEX1.itemcount = CountItem(m_BillingDoc.BillingAdditions)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.AparMasID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
      frmAddEditCashTran.HeaderText = "เพิ่มรายการการชำระเงิน"
      frmAddEditCashTran.DocumentType = DocumentType
      frmAddEditCashTran.ItemAmount = txtGetAmount.Text
      frmAddEditCashTran.ShowMode = SHOW_ADD
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.TransferItems
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
      
      If OKClick Then
         m_HasModify = True

         GridEX1.itemcount = CountItem(m_BillingDoc.TransferItems)
         Call GridEX1.Rebind
         
         Call GetTotalPriceReceipt
      End If 'RECEIPT1_DOCTYPE
   ElseIf TabStrip1.SelectedItem.Tag = INVOICE_DOCTYPE & "-DCO" Or TabStrip1.SelectedItem.Tag = RECEIPT1_DOCTYPE & "-DCO" Then
      Set frmAddEditExtraDiscount.TempCollection = m_BillingDoc.ExtraDiscounts
      frmAddEditExtraDiscount.ShowMode = SHOW_ADD
      Set frmAddEditExtraDiscount.ParentForm = Me
      frmAddEditExtraDiscount.HeaderText = MapText("เพิ่มรายการส่วนลด")
     Load frmAddEditExtraDiscount
      frmAddEditExtraDiscount.Show 1
      
      OKClick = frmAddEditExtraDiscount.OKClick
      
      Unload frmAddEditExtraDiscount
      Set frmAddEditExtraDiscount = Nothing

      If OKClick Then
'         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.ExtraDiscounts)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   End If
   
   Set oMenu = Nothing
End Sub



Private Sub cmdCalculate_Click()
Dim tDoItem As CDocItem
Dim tExtraDiscount As CExtraDiscount
Dim Amt As Double
Dim nTotalPrice As Double
Dim oTotalPrice As Double
Dim PercentDiscount As Double

oTotalPrice = Val(txtTotalPrice.Text) * 107 / 100
nTotalPrice = oTotalPrice
For Each tExtraDiscount In m_BillingDoc.ExtraDiscounts
  If tExtraDiscount.GetFieldValue("UNIT_TYPE") = 1 Then 'บาท
   nTotalPrice = nTotalPrice - tExtraDiscount.GetFieldValue("EXTRA_DISCOUNT_VALUE")
 ElseIf tExtraDiscount.GetFieldValue("UNIT_TYPE") = 2 Then 'เปอร์เซ็นต์
   Amt = nTotalPrice * (MyDiffEx(tExtraDiscount.GetFieldValue("EXTRA_DISCOUNT_VALUE"), 100))
   nTotalPrice = nTotalPrice - Format(Amt, "0.00")
 End If
Next tExtraDiscount

PercentDiscount = MyDiffEx(oTotalPrice - nTotalPrice, oTotalPrice) * 100

For Each tDoItem In m_BillingDoc.DocItems
   Amt = tDoItem.GetFieldValue("AVG_PRICE") * tDoItem.GetFieldValue("UNIT_MULTIPLE") '* 100) / 107
'   Call tDoItem.SetFieldValue("AVG_PRICE", Amt)
   
   Amt = MyDiffEx(tDoItem.GetFieldValue("ITEM_AMOUNT"), tDoItem.GetFieldValue("UNIT_MULTIPLE")) * Amt
'   Call tDoItem.SetFieldValue("TOTAL_PRICE", Amt)
   
   Amt = Amt * MyDiffEx(PercentDiscount, 100)
   'Call tDoItem.SetFieldValue("DISCOUNT_PERCENT", PercentDiscount)
   Call tDoItem.SetFieldValue("DISCOUNT_AMOUNT", Format(Amt, "0.00"))
   
Next tDoItem
   Call TabStrip1_Click
End Sub

Private Sub cmdSaveConsignment_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim TempStr As String
Dim TempID As Long
Dim m_InventoryDoc As CInventoryDoc
Dim IsOK As Boolean
   
   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("บันทึกเอกสาร PO และเอกสาร RQ และออกจากหน้าจอ", "-", "ลบเอกสาร RQ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
         If Not cmdSaveConsignment.Enabled Then
            Exit Sub
         End If
         
         If chkConsignmentFlag.Value = ssCBUnchecked Then
            Exit Sub
         End If
         
         If CheckConsignmentNo = 1 Then
            glbErrorLog.LocalErrorMsg = MapText("มีเลขที่เอกสารโอนสต็อก ฝากขาย " & m_BillingDoc.CONSIGNMENT_NO & " แล้ว")
            glbErrorLog.ShowUserError
            Exit Sub
         End If
         
         OKClick = False
         
         frmSaveConsignment.DocumentDate = uctlDocumentDate.ShowDate
         frmSaveConsignment.HeaderText = MapText("เพิ่มข้อมูลใบโอนสต็อก ฝากขาย")
         Set frmSaveConsignment.TempCollection = m_BillingDoc.DocItems
         Load frmSaveConsignment
         frmSaveConsignment.Show 1
         
         OKClick = frmSaveConsignment.OKClick
         
         Unload frmSaveConsignment
         Set frmSaveConsignment = Nothing
         
         If OKClick Then
               m_HasModify = True
               If Not SaveData Then
                  Exit Sub
               End If
               frmAddEditBillingDoc.OKClick = True
               Unload Me
         End If
   ElseIf lMenuChosen = 3 Then
         If CheckConsignmentNo = 0 Then
            glbErrorLog.LocalErrorMsg = MapText("ไม่มีเอกสารโอนสต็อก ฝากขาย ให้ลบ")
            glbErrorLog.ShowUserError
            Exit Sub
         End If
         
         If Not ConfirmDelete(m_BillingDoc.CONSIGNMENT_NO) Then
            Exit Sub
         End If
         
         ShowMode = SHOW_EDIT
         ConsignmentNo = ""
         TempStr = m_BillingDoc.CONSIGNMENT_NO
         ConsignmentID = -1
         TempID = m_BillingDoc.CONSIGNMENT_ID
         m_HasModify = True
         If Not SaveData Then
            Exit Sub
         End If
         m_HasModify = False
         
         Set m_InventoryDoc = New CInventoryDoc
         Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", TempID)
         
         If Not glbDaily.DeleteInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
            Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", -1)
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If

         CheckConsignmentNo = 0
         glbErrorLog.LocalErrorMsg = MapText("ลบเอกสารโอนสต็อก ฝากขาย " & TempStr & " แล้ว")
         glbErrorLog.ShowUserError
   End If
End Sub
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   ID = ConvertDocToConfigNo(1, DocumentType, DocumentSubType)
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
         m_BillingDoc.CONFIG_DOC_TYPE = ID
   
         Call txtDocumentNo.SetSelectText(Len(txtDocumentNo.Text) - Cd.GetFieldValue("DIGIT_AMOUNT"), Cd.GetFieldValue("DIGIT_AMOUNT"))
      Else
         txtDocumentNo.Text = ""
      End If
   Else
      txtDocumentNo.Text = ""
   End If
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)

   If TabStrip1.SelectedItem.Tag = DocumentType & "-DTL" Then
      If ID1 <= 0 Then
         m_BillingDoc.DocItems.Remove (ID2)
      Else
         m_BillingDoc.DocItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
      GridEX1.Rebind
      m_HasModify = True
   End If
   
      
   If TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-RCP" Or TabStrip1.SelectedItem.Tag = CN_DOCTYPE & "-CN" Or TabStrip1.SelectedItem.Tag = DN_DOCTYPE & "-DN" Or TabStrip1.SelectedItem.Tag = BILLS_DOCTYPE & "-BILLS" Then
      If ID1 <= 0 Then
         m_BillingDoc.RcpCnDnItems.Remove (ID2)
      Else
         m_BillingDoc.RcpCnDnItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-SUB" Then
      If ID1 <= 0 Then
         m_BillingDoc.BillingSubTracts.Remove (ID2)
      Else
         m_BillingDoc.BillingSubTracts.Item(ID2).Flag = "D"
      End If

      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.BillingSubTracts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Tag = RECEIPT2_DOCTYPE & "-ADD" Then
      If ID1 <= 0 Then
         m_BillingDoc.BillingAdditions.Remove (ID2)
      Else
         m_BillingDoc.BillingAdditions.Item(ID2).Flag = "D"
      End If
      
      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.BillingAdditions)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
      If ID1 <= 0 Then
         m_BillingDoc.TransferItems.Remove (ID2)
      Else
         m_BillingDoc.TransferItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.TransferItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DCO" Then
      If ID1 <= 0 Then
         m_BillingDoc.ExtraDiscounts.Remove (ID2)
      Else
         m_BillingDoc.ExtraDiscounts.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_BillingDoc.ExtraDiscounts)
      GridEX1.Rebind
      m_HasModify = True
   End If
   
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   
   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If Not VerifyCombo(lblAparMas, uctlAparMasLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-DTL" Then
      frmAddEditDoItem.HeaderText = "แก้ไขรายการสินค้า"
      frmAddEditDoItem.DocumentType = DocumentType
      frmAddEditDoItem.ShowMode = SHOW_EDIT
      frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
      frmAddEditDoItem.ID = ID
      frmAddEditDoItem.Area = Area
      Set frmAddEditDoItem.ParentForm = Me
      Set frmAddEditDoItem.TempCollection = m_BillingDoc.DocItems
      Set frmAddEditDoItem.LotItemLinkCollection = LotItemLinkCollection
      frmAddEditDoItem.CusID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
      
      Load frmAddEditDoItem
      frmAddEditDoItem.Show 1
      
      OKClick = frmAddEditDoItem.OKClick
      Unload frmAddEditDoItem
      Set frmAddEditDoItem = Nothing
      
      
      If OKClick Then
         m_HasModify = True
         GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
         GridEX1.Rebind
      
         Call GetTotalPrice
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-SUB" Then
      Set frmAddEditBillingSubTract.TempCollection = m_BillingDoc.BillingSubTracts
      frmAddEditBillingSubTract.ShowMode = SHOW_EDIT
      frmAddEditBillingSubTract.ID = ID
      Set frmAddEditBillingSubTract.ParentForm = Me
      frmAddEditBillingSubTract.HeaderText = MapText("แก้ไขรายการหัก")
     Load frmAddEditBillingSubTract
      frmAddEditBillingSubTract.Show 1
      
      OKClick = frmAddEditBillingSubTract.OKClick
      
      Unload frmAddEditBillingSubTract
      Set frmAddEditBillingSubTract = Nothing

      If OKClick Then
         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.BillingSubTracts)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-ADD" Then
      Set frmAddEditBillingAddition.TempCollection = m_BillingDoc.BillingAdditions
      frmAddEditBillingAddition.ShowMode = SHOW_EDIT
      frmAddEditBillingAddition.ID = ID
      Set frmAddEditBillingAddition.ParentForm = Me
      frmAddEditBillingAddition.HeaderText = MapText("แก้ไขรายการหัก")
     Load frmAddEditBillingAddition
      frmAddEditBillingAddition.Show 1
      
      OKClick = frmAddEditBillingAddition.OKClick
      
      Unload frmAddEditBillingAddition
      Set frmAddEditBillingAddition = Nothing

      If OKClick Then
         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.BillingAdditions)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.ID = ID
      frmAddEditCashTran.HeaderText = "แก้ไขรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_EDIT
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.TransferItems
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.itemcount = CountItem(m_BillingDoc.TransferItems)
         Call GridEX1.Rebind
         
         Call GetTotalPriceReceipt
      End If
   ElseIf TabStrip1.SelectedItem.Tag = INVOICE_DOCTYPE & "-DCO" Or TabStrip1.SelectedItem.Tag = RECEIPT1_DOCTYPE & "-DCO" Then
      Set frmAddEditExtraDiscount.TempCollection = m_BillingDoc.ExtraDiscounts
      frmAddEditExtraDiscount.ShowMode = SHOW_EDIT
      Set frmAddEditExtraDiscount.ParentForm = Me
      frmAddEditExtraDiscount.ID = ID
      frmAddEditExtraDiscount.HeaderText = MapText("แก้ไขรายการส่วนลด")
     Load frmAddEditExtraDiscount
      frmAddEditExtraDiscount.Show 1
      
      OKClick = frmAddEditExtraDiscount.OKClick
      
      Unload frmAddEditExtraDiscount
      Set frmAddEditExtraDiscount = Nothing

      If OKClick Then
'         Call GetTotalPriceReceipt

         GridEX1.itemcount = CountItem(m_BillingDoc.ExtraDiscounts)
         GridEX1.Rebind
         
         m_HasModify = True
      End If
   End If
End Sub
Private Sub cmdOK_Click()
Dim oMenu As CPopupMenu
Dim lMenuChosen  As Long
Dim iCount As Long
Dim m_Rs1 As ADODB.Recordset
Dim TempError As String


   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   
    If DocumentType = INVOICE_DOCTYPE Then
    'ตรวจเอกสาร : กรณีที่ทำการเปลี่ยนแปลงใบส่งของและใบกำกับภาษี หลังตัดยอดจากบิล(ชำระ) แล้ว ไม่ให้มีการเปลี่ยนแปลงในขณะที่ใบส่งของและใบกำกับภาษี ได้ถูกอ้างมาจาก ใบเสร็จชำระ
   If ID <> 0 Then
      Dim Ri As CRcpCnDn_Item
      Set Ri = New CRcpCnDn_Item
      Set m_Rs1 = New ADODB.Recordset
      Ri.RCPCNDN_ITEM_ID = -1
'      RI.BILLING_DOC_ID = ID
    Call Ri.SetFieldValue("DOC_ID", ID)
    
'      RI.DOC_ID = ID
      Call Ri.QueryData(3, m_Rs1, iCount)
      Set Ri = Nothing
      
      While Not m_Rs1.EOF
         Set Ri = New CRcpCnDn_Item
         Call Ri.PopulateFromRS(6, m_Rs1)
         
         If Len(Trim(TempError)) <= 0 Then
            TempError = Ri.DOCUMENT_NO
         Else
            TempError = TempError & ", " & Ri.DOCUMENT_NO
         End If
            
         Set Ri = Nothing
         m_Rs1.MoveNext
      Wend
   
      If Len(Trim(TempError)) > 0 Then
         glbErrorLog.LocalErrorMsg = MapText("มีการอ้างอิงกับเอกสาร " & TempError & " โปรดลบเอกสารอ้างอิงดังกล่าวก่อน เพิ่มเติม/แก้ไข เอกสาร " & txtDocumentNo.Text)
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
    
    End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_BillingDoc.BILLING_DOC_ID
      m_BillingDoc.QueryFlag = 1
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
Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
Dim ReportModeType As Long
   
   ReportMode = 1
   
   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      If Not SaveData Then
         Exit Sub
      End If
      ShowMode = SHOW_EDIT
      ID = m_BillingDoc.BILLING_DOC_ID
      m_BillingDoc.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   End If
   
   ReportFlag = False
   
   Set oMenu = New CPopupMenu
   If Not (AutoPrintMode) Then
      If DocumentType = INVOICE_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.DOPrintMenuItems)
      ElseIf DocumentType = PO_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.POPrintMenuItems)
      ElseIf DocumentType = RETURN_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RTPrintMenuItems)
      ElseIf DocumentType = S_RETURN_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RTPrintMenuItems2)
      ElseIf DocumentType = S_INVOICE_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RoPrintMenuItems)
      ElseIf DocumentType = BILLS_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.BSPrintMenuItems)
      ElseIf DocumentType = S_PO_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems2)
      ElseIf DocumentType = CN_DOCTYPE Then
           lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems3)
      ElseIf DocumentType = DN_DOCTYPE Then
          lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems4)
      Else
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems)
      End If
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalRcp001"

      Set Report = New CReportNormalRcp001
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   ElseIf lMenuChosen = 1 Or (AutoPrintMode And DocumentType = RETURN_DOCTYPE) Then
      ReportKey = "CReportNormalRcp001"
      
      AutoPrintMode = True
      
      Set Report = New CReportNormalRcp001
      ReportFlag = True
      Call Report.AddParam(2, "PREVIEW_TYPE")
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalRcp001"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)

      HeaderText = MapText("ปรับค่าหน้ากระดาษ")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportFormDO001"
      Set Report = New CReportFormDO001
      
      Call Report.AddParam(1, "PREVIEW_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 5 Or (AutoPrintMode And DocumentType = INVOICE_DOCTYPE) Then
      
      ReportKey = "CReportFormDO001"
      
      AutoPrintMode = True
      
      Set Report = New CReportFormDO001
      
      Call Report.AddParam(2, "PREVIEW_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 6 Then
      ReportKey = "CReportFormDO001"
      ReportMode = 2

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("ใบส่งสินค้า")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 8 Then
      ReportKey = "CReportPrintLabel001"

      Set Report = New CReportPrintLabel001
      ReportFlag = True
   ElseIf lMenuChosen = 9 Then
      ReportKey = "CReportPrintLabel001"
      
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("หน้า ลาเบล")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 11 Then
      ReportKey = "CReportPrintLabel002"

      Set Report = New CReportPrintLabel002
      ReportFlag = True
   ElseIf lMenuChosen = 12 Then
      ReportKey = "CReportPrintLabel002"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("ใบแยกถานที่จัดส่ง")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 13 Then
      ReportKey = "CReportNormalRcp001"

      Set Report = New CReportNormalRcp001
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   ElseIf lMenuChosen = 14 Then
      ReportKey = "CReportNormalRcp001"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("ใบรับของ")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 16 Then
      ReportKey = "CReportFormRv0001"

      Set Report = New CReportFormRv0001
      
      Call Report.AddParam(1, "PREVIEW_TYPE")
      ReportFlag = True
      
   ElseIf lMenuChosen = 17 Then
      ReportKey = "CReportFormRv0001"

      Set Report = New CReportFormRv0001
      
      Call Report.AddParam(2, "PREVIEW_TYPE")
      ReportFlag = True
      AutoPrintMode = True
   ElseIf lMenuChosen = 18 Then
      ReportMode = 2
      ReportKey = "CReportFormRv0001"
      
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("ใบสำคัญรับ")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 24 Then
      ReportKey = "CReportNormalRcp004"

      Set Report = New CReportNormalRcp004
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
   ElseIf lMenuChosen = 25 Then
      ReportKey = "CReportNormalRcp004"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)

      HeaderText = MapText("ปรับค่าหน้ากระดาษ")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 27 Then
      ReportKey = "CReportPrintLabel003"
      
      Set Report = New CReportPrintLabel003
      ReportFlag = True
      ReportModeType = 1
   ElseIf lMenuChosen = 28 Then
      ReportKey = "CReportPrintLabel003"
      
      Set Report = New CReportPrintLabel003
      ReportFlag = True
      ReportModeType = 2
   ElseIf lMenuChosen = 29 Then
      ReportKey = "CReportPrintLabel003"
      ReportMode = 2
      
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("หน้า ลาเบล(สติ๊กเกอร์)")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 31 Then              ' Add   by    Pat
      ReportKey = "CReportPrintLabel004"
      Set Report = New CReportPrintLabel004
      ReportFlag = True
   ElseIf lMenuChosen = 32 Then
      ReportKey = "CReportPrintLabel004"
      ReportMode = 2
      
      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("หน้า ลาเบล(สติ๊กเกอร์) อิออน")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
      
      
      
     ElseIf lMenuChosen = 33 Then                        '--------------------------pui เพิ่ม
      ReportKey = "CReportNormalRcp001_1"

      Set Report = New CReportNormalRcp001_1
      ReportFlag = True
      Call Report.AddParam(1, "PREVIEW_TYPE")
      

    ElseIf lMenuChosen = 34 Then
      ReportKey = "CReportNormalRcp001_1"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)

      HeaderText = MapText("ปรับค่าหน้ากระดาษ")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
   ElseIf lMenuChosen = 36 Then
      ReportKey = "CReportFormDO001_2"
      Set Report = New CReportFormDO001_2
      Call Report.AddParam(1, "PREVIEW_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 37 Or (AutoPrintMode And DocumentType = INVOICE_DOCTYPE) Then
      
      ReportKey = "CReportFormDO001_2"
      
      AutoPrintMode = True
      
      Set Report = New CReportFormDO001_2
      
      Call Report.AddParam(2, "PREVIEW_TYPE")
      ReportFlag = True
   ElseIf lMenuChosen = 38 Then
      ReportKey = "CReportFormDO001_2"
      ReportMode = 2

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)
      
      HeaderText = MapText("ใบส่งสินค้า")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   Else
      Exit Sub
   End If
   

   

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(m_BillingDoc.Due_Date, "TEMP_DUE_DATE")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(DocumentType, "DOCUMENT_TYPE")
      Call Report.AddParam(ReportModeType, "MODE_TYPE")
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
   End If

   Call EnableForm(Me, False)
   If ReportFlag Then
      frmReport.ClassName = ReportKey
      frmReport.AutoPrintMode = AutoPrintMode
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
      AutoPrintMode = False
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.GetFieldValue("REPORT_CONFIG_ID")
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1

      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
'      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(m_Adr, cboEnpAddress, , True)
      
      Call LoadConfigDoc(Nothing, m_Cd)

      If Area = 1 Then
         m_Apm.APAR_IND = 1
         Call LoadApArMas(m_Apm, uctlAparMasLookup.MyCombo, m_APARCancelOutDocument)
         Set uctlAparMasLookup.MyCollection = m_CustomerColl
      ElseIf Area = 2 Then
         m_Apm.APAR_IND = 2
         Call LoadApArMas(m_Apm, uctlAparMasLookup.MyCombo, m_Customers)
         Set uctlAparMasLookup.MyCollection = m_Customers
      End If
      
      Call LoadEmployee(m_Emp, uctlSale.MyCombo)
      Set uctlSale.MyCollection = m_EmployeeColl
      
      Call LoadEmployee(m_Emp, uctlDealer.MyCombo)
      Set uctlDealer.MyCollection = m_EmployeeColl
      
      Call LoadMaster(cboAccount, , , , MASTER_DEPARTMENT)
      
      Call LoadMaster(cboDocumentReturn, , , , MASTER_INVOICE_RETURN)
            
      If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Then
         Call LoadMaster(cboDriver, , , , MASTER_DRIVER)
         Call LoadMaster(cboCarLicense, , , , MASTER_CAR_LICENSE)
         Call LoadMaster(cboTranSportor, , , , MASTER_TRANSPORTOR)
         Call LoadMaster(cboTranSportCycle, , , , MASTER_TRANSPORT_CYCLE)
'         Call LoadMaster(cboDiscountType, , , , MASTER_DISCOUNT)
         Call LoadMaster(cboPaymentBy, , , , MASTER_PAYMENT_BY)
      End If
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         uctlDueDate.ShowDate = Now
         Call cmdAuto_Click
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
         If DocumentType = RETURN_DOCTYPE Then
            txtReferDesc.Text = "สินค้าหมดอายุ"
         End If
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
      m_HasDiscountModify = False
      cmdOK.Enabled = chkCommit.Enabled
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
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
      AutoPrintMode = False
      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 122 Then
      AutoPrintMode = True
      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 1 And KeyCode = 113 Then
      Call LoadCalculator
   End If
   
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
      
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
   Set m_APARCancelOutDocument = Nothing
   Set m_Apm = Nothing
   Set m_Emp = Nothing
   Set m_Mr = Nothing
   Set m_Adr = Nothing
   Set m_LocationSales = Nothing
   Set m_Cd = Nothing
   Set LotItemLinkCollection = Nothing
   
   Set uctlAparMasLookup.MyCollection = Nothing
   Set uctlSale.MyCollection = Nothing
   
   Unload Me
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   If TabStrip1.SelectedItem.Tag = DocumentType & "-DTL" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1600
      Col.Caption = MapText("รหัส")
      
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4575
      Col.Caption = MapText("รายละเอียด")
      
      Set Col = GridEX1.Columns.add '5
      Col.TextAlignment = jgexAlignRight
      Col.Width = 1200
      Col.Caption = MapText("จำนวน")
         
      Set Col = GridEX1.Columns.add '6
      Col.TextAlignment = jgexAlignRight
      Col.Width = 1500
      Col.Caption = MapText("ราคา/หน่วย")
      
      Set Col = GridEX1.Columns.add '7
      Col.TextAlignment = jgexAlignRight
      Col.Width = 1200
      Col.Caption = MapText("ส่วนลด")
      
      Set Col = GridEX1.Columns.add '8
      Col.TextAlignment = jgexAlignRight
      Col.Width = 1575
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 3000
      Col.Caption = MapText("เลขที่เอกสารอ้างอิง (PO/DO)")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-RCP" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 2220
      Col.Caption = MapText("เลขที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2730
      Col.Caption = MapText("วันที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ยอดหนี้")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ส่วนลดรับ")
   
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("รับชำระ")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("คงค้าง")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 3000
      Col.Caption = MapText("เลขที่เอกสารอ้างอิง (BL)")
      
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-JNL" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1965
      Col.Caption = MapText("รหัสบัญชี")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 5100
      Col.Caption = MapText("รายละเอียด")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เดบิต")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 2160
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("เครดิต")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 2415
      Col.Caption = MapText("ประเภทการชำระเงิน")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 3200
      Col.Caption = MapText("เลขที่เช็ค/บัญชี")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2000
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("ธนาคาร")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = ScaleWidth - 2415 - 3200 - 2000 - 1650
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("สาขาธนาคาร")
   
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-CN" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 2415
      Col.Caption = MapText("เลขที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2040
      Col.Caption = MapText("ยอดหนี้")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2160
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ยอดลดหนี้")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 4935
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สาเหตุการลดหนี้")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DN" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 2415
      Col.Caption = MapText("เลขที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2040
      Col.Caption = MapText("ยอดหนี้")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2160
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ยอดเพิ่มหนี้")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 4935
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("สาเหตุการเพิ่มหนี้")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-BILLS" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 2220
      Col.Caption = MapText("เลขที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2730
      Col.Caption = MapText("วันที่เอกสาร")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ยอดหนี้")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ส่วนลดรับ")
   
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("รับชำระ")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("คงค้าง")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-SUB" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 11500
      Col.Caption = MapText("รายการหัก")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-ADD" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 11500
      Col.Caption = MapText("รายการเพิ่ม")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")
  ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DCO" Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1000
      Col.Caption = MapText("ลำดับ")
      
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2500
      Col.Caption = MapText("รหัสประเภท")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4000
      Col.Caption = MapText("ประเภทส่วนลด")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 4000
      Col.Caption = MapText("เลขที่ส่วนลด")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 4000
      Col.Caption = MapText("รายละเอียด")
   
      Set Col = GridEX1.Columns.add '8
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 2000
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("หน่วย")
      
   End If
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2805
   Col.Caption = MapText("ชื่อส่วนลด")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 5055 + 1950
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("มูลค่าส่วนลด")
End Sub

Private Sub GetTotalPrice()
Dim II As CDocItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum5 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   Sum5 = 0

   For Each II In m_BillingDoc.DocItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.GetFieldValue("ITEM_AMOUNT")
         Sum2 = Sum2 + II.GetFieldValue("TOTAL_PRICE")
         Sum3 = Sum3 + II.GetFieldValue("DISCOUNT_AMOUNT")
      End If
   Next II

      txtTotalAmount.Text = Format(Sum1, "0.00")
      txtTotalPrice.Text = Format(Sum2, "0.00")
      txtItemDiscount.Text = Format(Sum3, "0.00")

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame5.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.Visible = True
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่"))
   Call InitNormalLabel(lblAccountNo, MapText("แผนก"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่"))
   If DocumentType = PO_DOCTYPE Or DocumentType = S_PO_DOCTYPE Then
      Call InitNormalLabel(lblDueDate, MapText("วันที่ส่งของ"))
   ElseIf DocumentType = BILLS_DOCTYPE Or DocumentType = S_BILLS_DOCTYPE Then
      Call InitNormalLabel(lblDueDate, MapText("วันนัดรับเงิน"))
   ElseIf DocumentType = RECEIPT2_DOCTYPE Or DocumentType = S_RECEIPT2_DOCTYPE Then
      Call InitNormalLabel(lblDueDate, MapText("วันที่รับเงิน"))
   Else
      Call InitNormalLabel(lblDueDate, MapText("ครบกำหนด"))
   End If
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(Label1, MapText("วัน"))
   Call InitNormalLabel(lblIncludeDiscount, MapText("รวม"))
   Call InitNormalLabel(lblIncludeDiscountEx, MapText("รวม"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblItemDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblExtraDiscountPercent, MapText("ส่วนลดรวม"))
   Call InitNormalLabel(lblVatPercent, MapText("VAT"))
   Call InitNormalLabel(lblTOtal, MapText("รวม"))
   Call InitNormalLabel(lblTotalDebt, MapText("ยอดหนี้"))
   Call InitNormalLabel(lblAdditionAmount, MapText("ส่วนเพิ่ม"))
   Call InitNormalLabel(lblPaidAmount, MapText("ยอดชำระ"))
   Call InitNormalLabel(lblWHPercent, MapText("หัก ณ ที่จ่าย"))
   Call InitNormalLabel(lblGetAmount, MapText("เหลือรับ"))
   
   Call InitNormalLabel(lblLocationSale, MapText("สาขา"))
   Call InitNormalLabel(lblRefertext, MapText("อ้างอิง"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   Call InitNormalLabel(lblPoNo, MapText("PO"))
   Call InitNormalLabel(lblDealer, MapText("ตัวแทน"))
   
   Call InitNormalLabel(lblAddress, MapText("สถานที่ส่งของ"))
   Call InitNormalLabel(lblDocumentreturn, MapText("ใบนำกลับ"))
   lblDocumentreturn.ToolTipText = "ประเภทเอกสารที่นำกลับมายังบริษทหลังจากที่ลูกค้าได้รับสินค้าและรับต้นฉบับไว้"
   
   Call InitNormalLabel(lblVatPercentEx, MapText("VAT"))
   Call InitNormalLabel(lblTOtalEx, MapText("รวม"))
   Call InitNormalLabel(lblDebitAmount, MapText("เพิ่มหนี้"))
   Call InitNormalLabel(lblCreditAmount, MapText("ลดหนี้"))
   Call InitNormalLabel(lblAfterDebitCredit, MapText("รวม"))
   Call InitNormalLabel(lblSubTractAmount, MapText("ส่วนหัก"))
   Call InitNormalLabel(lblAfterSubTract, MapText("คงเหลือ"))
   
   Call InitNormalLabel(lblFeeAmount, MapText("FEE."))
   Call InitNormalLabel(lblFromCashTran, MapText("รับจริง"))
   Call InitNormalLabel(lblDifRcp, MapText("ส่วนต่าง"))
   
   If Area = 1 Then
      Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   Else
      Call InitNormalLabel(lblSale, MapText("ผู้รับผิดชอบ"))
   End If

   uctlAparMasLookup.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   
   If DocumentType = INVOICE_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Then
      lblDocumentreturn.Visible = True
      cboDocumentReturn.Visible = True
      chkTicketFlag.Visible = True
   Else
      lblDocumentreturn.Visible = False
      cboDocumentReturn.Visible = False
      chkTicketFlag.Visible = False
   End If
   
   If Area = 1 Then
      Call InitNormalLabel(lblAparMasAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblAparMas, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่บริษัท"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblAparMasAddress, MapText("ที่อยู่ผู้ค้า"))
      Call InitNormalLabel(lblAparMas, MapText("รหัสผู้ค้า"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่บริษัท"))
   End If
   
   
   If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Then
         Call InitNormalLabel(lblDriver, MapText("คนขับ"))
         Call InitNormalLabel(lblCarLicense, MapText("ทะเบียน"))
         Call InitNormalLabel(lblTranSportor, MapText("ขนส่ง"))
         Call InitNormalLabel(lblTranSportCycle, MapText("รอบขนส่ง"))
'         Call InitNormalLabel(lblDiscountType, MapText("ประเภทส่วนลด"))
         Call InitNormalLabel(lblPaymentBy, MapText("วิธีการชำระเงิน"))
         Call InitNormalLabel(lblGetDate, MapText("วันที่ได้รับของ"))
         
         Call InitCombo(cboDriver)
         Call InitCombo(cboCarLicense)
         Call InitCombo(cboTranSportor)
         Call InitCombo(cboTranSportCycle)
'         Call InitCombo(cboDiscountType)
         Call InitCombo(cboPaymentBy)
   End If
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtIncludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeDiscount.Enabled = False
   Call txtIncludeDiscountEx.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeDiscountEx.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalPrice.Enabled = False
   Call txtItemDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtItemDiscount.Enabled = False
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   
   Call txtExtraDiscountPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtExtraDiscountAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtExtraDiscountAmount.Enabled = False
   Call txtVatPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtVatAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtVatAmount.Enabled = False
   Call txtTotalDebt.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalDebt.Enabled = False
   Call txtAdditionAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAdditionAmount.Enabled = False
   Call txtPaidAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPaidAmount.Enabled = False
   Call txtWHPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWHAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtWHAmount.Enabled = False
   Call txtGetAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtGetAmount.Enabled = False
   
   Call txtTotalEx.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalEx.Enabled = False
   Call txtVatPercentEx.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtVatPercentEx.Enabled = False
   Call txtVatAmountEx.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtVatAmountEx.Enabled = False
   Call txtDebitAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDebitAmount.Enabled = False
   Call txtCreditAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtCreditAmount.Enabled = False
   Call txtAfterDebitCredit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAfterDebitCredit.Enabled = False
   Call txtSubTractAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtSubTractAmount.Enabled = False
   Call txtAfterSubTract.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAfterSubTract.Enabled = False
   
   Call txtFromCashTran.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtFromCashTran.Enabled = False
   Call txtDifRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDifRcp.Enabled = False
   Call txtFeeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtFeeAmount.Enabled = False
   
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitCheckBox(chkTicketFlag, "ตั๋ว")
   Call InitCheckBox(ChkCancelFlag, "CANCEL")
   
   Call InitCheckBox(chkConsignmentFlag, "ฝากขาย")
   If DocumentType <> PO_DOCTYPE Then
      chkConsignmentFlag.Enabled = False
   End If
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboAparMasAddress)
   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboLocationSale)
    Call InitCombo(cboBranchAddress)
   Call InitCombo(cboDocumentReturn)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdCalculate.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSaveConsignment.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdCalculate, MapText("คำนวณ"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์ (F10/F11)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdSaveConsignment, MapText("บันทึก RQ"))
   
   If DocumentType <> PO_DOCTYPE Then
      cmdSaveConsignment.Visible = False
   End If
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   If DocumentType = PO_DOCTYPE Or DocumentType = S_PO_DOCTYPE Then
      SSFrame3.Visible = True
      txtCredit.Enabled = False
      uctlDueDate.Enable = True
      uctlDueDate.TabStop = True
      txtExtraDiscountPercent.Enabled = True
      txtVatPercent.Enabled = True
      txtWHPercent.Enabled = False
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการ")
      T.Tag = DocumentType & "-DTL"
         
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายละเอียด")
      T.Tag = DocumentType & "-INFO"


   
   ElseIf DocumentType = INVOICE_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Then
      SSFrame3.Visible = True
      txtWHPercent.Enabled = False
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการ")
      T.Tag = DocumentType & "-DTL"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายละเอียด")
      T.Tag = DocumentType & "-INFO"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนลดอื่นๆ")
      T.Tag = DocumentType & "-DCO"
'
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
      
   ElseIf DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE Then
      SSFrame3.Visible = True
      SSFrame5.Visible = True
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการ")
      T.Tag = DocumentType & "-DTL"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายละเอียด")
      T.Tag = DocumentType & "-INFO"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("การชำระเงิน")
      T.Tag = DocumentType & "-PMT"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนลดอื่นๆ")
      T.Tag = DocumentType & "-DCO"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
   ElseIf DocumentType = RECEIPT2_DOCTYPE Or DocumentType = S_RECEIPT2_DOCTYPE Then
      SSFrame4.Visible = True
      SSFrame5.Visible = True
      txtCredit.Enabled = False
      uctlDueDate.Enable = True
      txtExtraDiscountPercent.Enabled = False
      txtVatPercent.Enabled = True
      txtVatPercentEx.Enabled = True
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการใบเสร็จ")
      T.Tag = DocumentType & "-RCP"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนหัก")
      T.Tag = DocumentType & "-SUB"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนเพิ่ม")
      T.Tag = DocumentType & "-ADD"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("การชำระเงิน")
      T.Tag = DocumentType & "-PMT"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
   ElseIf DocumentType = CN_DOCTYPE Or DocumentType = S_CN_DOCTYPE Then
      SSFrame4.Visible = True
      txtCredit.Enabled = False
      uctlDueDate.Enable = False
      txtExtraDiscountPercent.Enabled = False
      txtVatPercent.Enabled = False
      txtWHPercent.Enabled = False
      txtVatPercentEx.Enabled = True
   
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการลดหนี้")
      T.Tag = DocumentType & "-CN"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
'
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
   ElseIf DocumentType = DN_DOCTYPE Or DocumentType = S_DN_DOCTYPE Then
      SSFrame4.Visible = True
      txtVatPercentEx.Enabled = True
      txtCredit.Enabled = False
      uctlDueDate.Enable = False
      txtWHPercent.Enabled = False
      txtExtraDiscountPercent.Enabled = False
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการเพิ่มหนี้")
      T.Tag = DocumentType & "-DN"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
'
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
   ElseIf DocumentType = BILLS_DOCTYPE Or DocumentType = S_BILLS_DOCTYPE Then
      SSFrame4.Visible = True
      txtVatPercent.Enabled = False
      txtExtraDiscountPercent.Enabled = False
      txtWHPercent.Enabled = False
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการวางบิล")
      T.Tag = DocumentType & "-BILLS"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
      
   ElseIf DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      txtReferText.TabStop = True
      txtReferDesc.TabStop = True
      SSFrame3.Visible = True
      txtWHPercent.Enabled = False
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการ")
      T.Tag = DocumentType & "-DTL"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
   ElseIf DocumentType = RECEIPT3_DOCTYPE Then
      SSFrame4.Visible = True
      SSFrame5.Visible = True
      
      txtCredit.Enabled = False
      uctlDueDate.Enable = False
      txtExtraDiscountPercent.Enabled = False
      txtVatPercent.Enabled = True
      txtVatPercentEx.Enabled = True
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการใบเสร็จ")
      T.Tag = DocumentType & "-RCP"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนหัก")
      T.Tag = DocumentType & "-SUB"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("ส่วนเพิ่ม")
      T.Tag = DocumentType & "-ADD"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("รายละเอียด")
'      T.Tag = DocumentType & "-INFO"
      
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("การชำระเงิน")
      T.Tag = DocumentType & "-PMT"
      
'      Set T = TabStrip1.Tabs.add()
'      T.Caption = MapText("สมุดรายวัน")
'      T.Tag = DocumentType & "-JNL"
   End If
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
   Set m_BillingDoc = New CBillingDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_Resources = New Collection
   Set m_APARCancelOutDocument = New Collection
   Set m_Apm = New CAPARMas
   Set m_Emp = New CEmployee
   Set m_Mr = New CMasterRef
   Set m_Adr = New CAddress
   Set m_LocationSales = New Collection
   Set m_Cd = New Collection
   Set LotItemLinkCollection = New Collection
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
   
  If TabStrip1.SelectedItem.Tag = DocumentType & "-DTL" Then
         If m_BillingDoc.DocItems Is Nothing Then
            Exit Sub
         End If
         
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim CR As CDocItem
         If m_BillingDoc.DocItems.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_BillingDoc.DocItems, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
         
         Values(1) = CR.GetFieldValue("DOC_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = CR.ShowDescCode
         Values(4) = CR.ShowDescText
         Values(5) = FormatNumber(MyDiff(CR.GetFieldValue("ITEM_AMOUNT"), CR.GetFieldValue("UNIT_MULTIPLE")))
         Values(6) = FormatNumber(CR.GetFieldValue("AVG_PRICE") * CR.GetFieldValue("UNIT_MULTIPLE"))
         Values(7) = FormatNumber(CR.GetFieldValue("DISCOUNT_AMOUNT"))
         Values(8) = FormatNumber(CR.GetFieldValue("TOTAL_PRICE") - CR.GetFieldValue("DISCOUNT_AMOUNT"))
         Values(9) = CR.GetFieldValue("PO_NO")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-JNL" Then
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-RCP" Or TabStrip1.SelectedItem.Tag = DocumentType & "-BILLS" Then
         If m_BillingDoc.RcpCnDnItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim Rc As CRcpCnDn_Item
         If m_BillingDoc.RcpCnDnItems.Count <= 0 Then
            Exit Sub
         End If
         Set Rc = GetItem(m_BillingDoc.RcpCnDnItems, RowIndex, RealIndex)
         If Rc Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Rc.GetFieldValue("RCPCNDN_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Rc.GetFieldValue("DOC_NO")
         Values(4) = DateToStringExtEx2(Rc.GetFieldValue("DOC_DATE"))
         If Rc.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Or Rc.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
            Values(5) = FormatNumber(-Rc.GetFieldValue("ITEM_AMOUNT"))
         Else
            Values(5) = FormatNumber(Rc.GetFieldValue("ITEM_AMOUNT"))
         End If
         Values(6) = FormatNumber(Rc.GetFieldValue("PAID_DISCOUNT"))
         If TabStrip1.SelectedItem.Tag = DocumentType & "-RCP" Then
            If Rc.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Or Rc.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
               Values(7) = FormatNumber(-Rc.GetFieldValue("PAID_AMOUNT"))
            Else
               Values(7) = FormatNumber(Rc.GetFieldValue("PAID_AMOUNT"))
            End If
            Values(8) = FormatNumber(Rc.GetFieldValue("ITEM_AMOUNT") - Rc.GetFieldValue("PAID_DISCOUNT") - Rc.GetFieldValue("PAID_AMOUNT"))
         Else
            Values(7) = FormatNumber(Rc.GetFieldValue("ITEM_AMOUNT") - Rc.GetFieldValue("PAID_DISCOUNT") - Rc.GetFieldValue("PAID_AMOUNT"))
            If Rc.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Or Rc.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
               Values(8) = FormatNumber(-Rc.GetFieldValue("PAID_AMOUNT"))
            Else
               Values(8) = FormatNumber(Rc.GetFieldValue("PAID_AMOUNT"))
            End If
         End If
         
         If TabStrip1.SelectedItem.Tag = DocumentType & "-RCP" Then
            Values(9) = Rc.GetFieldValue("BILLS_NO")
         End If
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-CN" Then
         If m_BillingDoc.RcpCnDnItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim CRT As CRcpCnDn_Item
         If m_BillingDoc.RcpCnDnItems.Count <= 0 Then
            Exit Sub
         End If
         Set CRT = GetItem(m_BillingDoc.RcpCnDnItems, RowIndex, RealIndex)
         If CRT Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = CRT.GetFieldValue("RCPCNDN_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = CRT.GetFieldValue("DOC_NO")
         Values(4) = FormatNumber(CRT.GetFieldValue("ITEM_AMOUNT"))
         Values(5) = FormatNumber(CRT.GetFieldValue("CNDN_AMOUNT"))
         Values(6) = CRT.GetFieldValue("CNDN_REASON_NAME")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DN" Then
         If m_BillingDoc.RcpCnDnItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim DNT As CRcpCnDn_Item
         If m_BillingDoc.RcpCnDnItems.Count <= 0 Then
            Exit Sub
         End If
         Set DNT = GetItem(m_BillingDoc.RcpCnDnItems, RowIndex, RealIndex)
         If DNT Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = DNT.GetFieldValue("RCPCNDN_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DNT.GetFieldValue("DOC_NO")
         Values(4) = FormatNumber(DNT.GetFieldValue("ITEM_AMOUNT"))
         Values(5) = FormatNumber(DNT.GetFieldValue("CNDN_AMOUNT"))
         Values(6) = DNT.GetFieldValue("CNDN_REASON_NAME")
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-SUB" Then
         If m_BillingDoc.RcpCnDnItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim BSub As CBillingSubTract
         If m_BillingDoc.BillingSubTracts.Count <= 0 Then
            Exit Sub
         End If
         Set BSub = GetItem(m_BillingDoc.BillingSubTracts, RowIndex, RealIndex)
         If BSub Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = BSub.GetFieldValue("BILLING_SUBTRACT_ID")
         Values(2) = RealIndex
         Values(3) = BSub.GetFieldValue("SUBTRACT_NAME")
         Values(4) = FormatNumber(BSub.GetFieldValue("ITEM_AMOUNT"))
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-ADD" Then
         If m_BillingDoc.RcpCnDnItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim BAdd As CBillingAddition
         If m_BillingDoc.BillingAdditions.Count <= 0 Then
            Exit Sub
         End If
         Set BAdd = GetItem(m_BillingDoc.BillingAdditions, RowIndex, RealIndex)
         If BAdd Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = BAdd.GetFieldValue("BILLING_ADDITION_ID")
         Values(2) = RealIndex
         Values(3) = BAdd.GetFieldValue("ADDITION_NAME")
         Values(4) = FormatNumber(BAdd.GetFieldValue("ITEM_AMOUNT"))
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
         If m_BillingDoc.TransferItems Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim TR As CCashTransferItem
         If m_BillingDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set TR = GetItem(m_BillingDoc.TransferItems, RowIndex, RealIndex)
         If TR Is Nothing Then
            Exit Sub
         End If
         
         Values(1) = TR.ImportItemEx.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = TR.ImportItemEx.GetFieldValue("PAYMENT_TYPE_NAME")
         If TR.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
            Values(7) = FormatNumber(TR.ImportItemEx.GetFieldValue("AMOUNT"))
         ElseIf TR.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = BANKTRF_PMT Then
            Values(4) = TR.ImportItemEx.GetFieldValue("ACCOUNT_NAME")
            Values(5) = TR.ImportItemEx.GetFieldValue("BANK_NAME")
            Values(6) = TR.ImportItemEx.GetFieldValue("BRANCH_NAME")
            Values(7) = FormatNumber(TR.ImportItemEx.GetFieldValue("AMOUNT"))
         ElseIf TR.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Then
            Values(4) = TR.ImportItemEx.Cheque.GetFieldValue("CHEQUE_NO")
            Values(5) = TR.ImportItemEx.Cheque.GetFieldValue("BANK_NAME")
            Values(6) = TR.ImportItemEx.Cheque.GetFieldValue("BRANCH_NAME")
            Values(7) = FormatNumber(TR.ImportItemEx.GetFieldValue("AMOUNT"))
         ElseIf TR.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
            Values(4) = TR.ImportItemEx.Cheque.GetFieldValue("CHEQUE_NO") & " ( " & TR.ImportItem.GetFieldValue("ACCOUNT_NAME") & " )"
            Values(5) = TR.ImportItem.GetFieldValue("BANK_NAME")
            Values(6) = TR.ImportItem.GetFieldValue("BRANCH_NAME")
            Values(7) = FormatNumber(TR.ImportItemEx.GetFieldValue("AMOUNT"))
         End If
  ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DCO" Then
         If m_BillingDoc.ExtraDiscounts Is Nothing Then
            Exit Sub
         End If
   
         If RowIndex <= 0 Then
            Exit Sub
         End If
   
         Dim Edc As CExtraDiscount
         If m_BillingDoc.ExtraDiscounts.Count <= 0 Then
            Exit Sub
         End If
         Set Edc = GetItem(m_BillingDoc.ExtraDiscounts, RowIndex, RealIndex)
         If Edc Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Edc.GetFieldValue("EXTRA_DISCOUNT_ID")
         Values(2) = RealIndex
         Values(3) = Edc.GetFieldValue("ITEM")
         Values(4) = Edc.GetFieldValue("KEY_CODE")
         Values(5) = Edc.GetFieldValue("KEY_NAME")
         Values(6) = Edc.GetFieldValue("EXTRA_DISCOUNT_NO")
         Values(7) = Edc.GetFieldValue("EXTRA_DISCOUNT_DESC")
         Values(8) = FormatNumber(Edc.GetFieldValue("EXTRA_DISCOUNT_VALUE"))
         Values(9) = UnitDiscount(Edc.GetFieldValue("UNIT_TYPE"))
   
   End If
      
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
         cmdEdit.Enabled = True '(m_BillingDoc.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      Else
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
      End If
   Else
      cmdAdd.Enabled = En
      cmdDelete.Enabled = En
      cmdEdit.Enabled = En
   End If
   
End Sub

Private Sub TabStrip1_Click()
   GridEX1.Visible = False
   GridEX1.Top = 4200
   GridEX1.Left = 50
      
   SSFrame2.Visible = False
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-INFO" Then
      Call EnableDisableButton(False)
      SSFrame2.Visible = True
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DTL" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      Call GetTotalPrice
      If DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE Then
         Call GetTotalPriceReceipt
      End If
      GridEX1.itemcount = CountItem(m_BillingDoc.DocItems)
      GridEX1.Rebind
 ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DCO" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

'      Call GetTotalPrice
'      If DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE Then
'         Call GetTotalPriceReceipt
'      End If
      GridEX1.itemcount = CountItem(m_BillingDoc.ExtraDiscounts)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-JNL" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      Call GetTotalPrice
      
'      GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-RCP" Or TabStrip1.SelectedItem.Tag = DocumentType & "-BILLS" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-PMT" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.TransferItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-CN" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-DN" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True

      GridEX1.itemcount = CountItem(m_BillingDoc.RcpCnDnItems)
     GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-SUB" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True
      
      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.BillingSubTracts)
     GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Tag = DocumentType & "-ADD" Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True
      
      Call GetTotalPriceReceipt
      GridEX1.itemcount = CountItem(m_BillingDoc.BillingAdditions)
     GridEX1.Rebind
   
   End If
End Sub

Private Sub txtAfterSubTract_Change()
   m_HasModify = True
End Sub

Private Sub txtAdditionAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtCredit_Change()
Dim NewDate As Date

   m_HasModify = True
   
   If DocumentType <> PO_DOCTYPE Then
      NewDate = DateAdd("D", Val(txtCredit.Text), uctlDocumentDate.ShowDate)
      uctlDueDate.ShowDate = NewDate
   End If
End Sub

'Private Sub txtDiscountAmount_Change()
' m_HasModify = True
'End Sub

Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(DOCUMENT_NO_UNIQUE, txtDocumentNo.Text, ID) Then
      If DocAdd = 0 Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
      End If
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      txtDocumentNo.SetFocus
      If Len(txtDocumentNo.Text) > 0 Then
         SendKeys ("{TAB}")
      End If
      Exit Sub
   End If
End Sub
Private Sub txtExtraDiscountPercent_Change()
   m_HasModify = True
   m_HasDiscountModify = True
   Call CalculateAmount
End Sub
Private Sub txtFeeAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtItemDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPaidAmount_Change()
   Call CalculateAmountRecript
   m_HasModify = True
End Sub

Private Sub txtPoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPoNo_LostFocus()
'   If Not CheckUniqueNs(CUS_PO_UNIQUE, txtPoNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPoNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      txtPoNo.SetFocus
'      Exit Sub
'    End If
End Sub

Private Sub txtReferDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtReferText_Change()
   m_HasModify = True
End Sub

Private Sub txtSubTractAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub

Private Sub CalculateAmount()
Dim TempAmt As Double
Dim OldExtraDiscountAmount  As Double
   
   TempAmt = FormatNumber(Val(txtTotalPrice.Text) - Val(txtItemDiscount.Text), , False)
   
   txtIncludeDiscount.Text = FormatNumber(TempAmt, , False)
   
   OldExtraDiscountAmount = Val(txtExtraDiscountAmount.Text)
   
   txtExtraDiscountAmount.Text = FormatNumber(TempAmt * Val(txtExtraDiscountPercent.Text) / 100, , False)
   
   If Trim(Str(OldExtraDiscountAmount)) <> Trim(Str(txtExtraDiscountAmount.Text)) Then
      m_HasDiscountModify = True
   End If
      
   txtIncludeDiscountEx.Text = FormatNumber(TempAmt - Val(txtExtraDiscountAmount.Text), , False)
   
   txtVatAmount.Text = FormatNumber((TempAmt - Val(txtExtraDiscountAmount.Text)) * Val(txtVatPercent.Text) / 100, , False)
   
   txtTotal.Text = FormatNumber(TempAmt - Val(txtExtraDiscountAmount.Text) + Val(txtVatAmount.Text), , False)
   
End Sub
Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub
Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalDebt_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub
Private Sub txtVatPercent_Change()
    m_HasModify = True
   If DocumentType = RECEIPT1_DOCTYPE Then
      Call CalculateAmountRecriptCash
   Else
      Call CalculateAmount
   End If
End Sub

Private Sub txtVatPercentEx_Change()
   Call CalculateAmountRecript
   m_HasModify = True
End Sub

Private Sub txtWHPercent_Change()
   If DocumentType = RECEIPT1_DOCTYPE Then
      Call CalculateAmountRecriptCash
   ElseIf DocumentType = RECEIPT2_DOCTYPE Then
      Call CalculateAmountRecript
   End If
   m_HasModify = True
End Sub

Private Sub uctlDealer_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlApArMasLookup_Change()
Dim ID As Long
Static OldID As Long
Dim AparMas As CAPARMas
   
   ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
    
   If ID = OldID Then
      Exit Sub
   End If
   If ID > 0 Then
      OldID = ID
      If Area = 1 Then
         Set AparMas = GetObject("CAPARMas", m_CustomerColl, Trim(Str(ID)), False)
         If AparMas Is Nothing Then
            uctlAparMasLookup.MyCombo.ListIndex = -1
            Exit Sub
         End If
      ElseIf Area = 2 Then
         Set AparMas = m_SupplierColl(Trim(Str(ID)))
      End If
      Call m_Adr.SetFieldValue("APAR_MAS_ID", ID)
      Call LoadApArMasAddress(m_Adr, cboAparMasAddress, , True)
      
      Call LoadApArMasAddress(m_Adr, cboBranchAddress, , True)
      
      txtExtraDiscountPercent.Text = AparMas.NORMAL_DISCOUNT
      If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or _
      DocumentType = S_PO_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Then
         txtCredit.Text = AparMas.CREDIT
      End If
      
      FromCustomerLookup = True
      
      Call LoadMaster(cboLocationSale, m_LocationSales, , , MASTER_APARMAS_BRANCH, ID)
      
      If m_LocationSales.Count > 0 Then
         Dim BranchID As Long
         Dim Branch As CMasterRef
         BranchID = cboLocationSale.ItemData(Minus2Zero(1))
'         If AparMas.LABEL_FLAG = "N" Then
            cboLocationSale.Enabled = True
            cboLocationSale.ListIndex = 1
'         Else
'            cboLocationSale.Enabled = False
'         End If
         
         If BranchID > 0 Then
            Set Branch = m_LocationSales(Trim(Str(BranchID)))
            
            If Branch.PARENT_EX_ID > 0 Then
               If AparMas.LABEL_FLAG = "N" Then
                  uctlSale.Enabled = True
                  uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, Branch.PARENT_EX_ID)
               Else
                  uctlSale.Enabled = False
                  uctlSale.MyTextBox.Text = ""
                  uctlSale.MyCombo.ListIndex = -1
               End If
            Else
               uctlSale.MyCombo.ListIndex = -1
            End If
            
            If Branch.DEALER_ID > 0 Then
               uctlDealer.MyCombo.ListIndex = IDToListIndex(uctlDealer.MyCombo, Branch.DEALER_ID)
            Else
               uctlDealer.MyCombo.ListIndex = -1
            End If
            
            If Branch.PARENT_EX_ID3 > 0 Then
               If AparMas.LABEL_FLAG = "N" Then
                  cboBranchAddress.Enabled = True
                  cboBranchAddress.ListIndex = IDToListIndex(cboBranchAddress, Branch.PARENT_EX_ID3)
               Else
                  cboBranchAddress.Enabled = False
                  cboBranchAddress.ListIndex = -1
               End If
            Else
               cboBranchAddress.ListIndex = -1
            End If
         End If
         Set Branch = Nothing
      End If
      FromCustomerLookup = False
   End If
         
 If (DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT2_DOCTYPE) And ID <> 0 Then
      Set AparMas = GetObject("CAPARMas", m_APARCancelOutDocument, Trim(ID), False)
      If AparMas.CANCEL_OUT_DOCUMENT = "Y" Then
         cmdOK.Enabled = False
         cmdPrint.Enabled = False
         cmdSaveConsignment.Enabled = False
         
         m_HasModify = False
         glbErrorLog.LocalErrorMsg = MapText("ลูกค้า" & AparMas.APAR_CODE & " " & AparMas.APAR_NAME & " ถูกยกเลิกการขายเครดิต")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
         
   m_HasModify = True
   If m_BillingDoc.COMMIT_FLAG = "N" Then
      cmdOK.Enabled = True
      cmdPrint.Enabled = True
      cmdSaveConsignment.Enabled = True
   End If
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub
Private Sub GetTotalPriceDebitCredit()
Dim II As CRcpCnDn_Item
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   For Each II In m_BillingDoc.RcpCnDnItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.GetFieldValue("CNDN_AMOUNT")
      End If
   Next II
   
   txtTotalDebt.Text = Format(Sum1, "0.00")
   txtAdditionAmount.Text = Format(0, "0.00")
   txtPaidAmount.Text = Format(Sum1, "0.00")
   txtDebitAmount.Text = Format(0, "0.00")
   txtCreditAmount.Text = Format(0, "0.00")
   
End Sub
Private Sub GetTotalPriceReceipt()
Dim II As CRcpCnDn_Item
Dim BSub As CBillingSubTract
Dim BAdd As CBillingAddition
Dim Tm As CCashTransferItem

Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum5 As Double
Dim Sum6 As Double
Dim Sum7 As Double
Dim Sum8 As Double
Dim Sum9 As Double
Dim Sum10 As Double
Dim Sum11 As Double
Dim Sum12 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   Sum5 = 0
   Sum6 = 0
   
   For Each BSub In m_BillingDoc.BillingSubTracts
      If BSub.Flag <> "D" Then
         Sum6 = Sum6 + BSub.GetFieldValue("ITEM_AMOUNT")
      End If
   Next
   Set BSub = Nothing
   
   
   Sum7 = 0
   Sum8 = 0
   Sum9 = 0
   Sum10 = 0
   For Each Tm In m_BillingDoc.TransferItems
      If Tm.Flag <> "D" Then
         If Tm.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
            Sum7 = Sum7 + Tm.ImportItemEx.GetFieldValue("NET_AMOUNT")
            Sum8 = Sum8 + Tm.ImportItemEx.GetFieldValue("AMOUNT")
            Sum12 = Sum12 + Tm.ImportItemEx.GetFieldValue("FEE_AMOUNT")
         ElseIf Tm.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Then
            Sum7 = Sum7 + Tm.ImportItemEx.GetFieldValue("NET_AMOUNT")
            Sum9 = Sum9 + Tm.ImportItemEx.GetFieldValue("AMOUNT")
            Sum12 = Sum12 + Tm.ImportItemEx.GetFieldValue("FEE_AMOUNT")
         ElseIf Tm.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
            Sum7 = Sum7 + Tm.ImportItem.GetFieldValue("NET_AMOUNT")
            Sum9 = Sum9 + Tm.ImportItemEx.GetFieldValue("AMOUNT")
            Sum12 = Sum12 + Tm.ImportItem.GetFieldValue("FEE_AMOUNT")
         ElseIf Tm.ImportItemEx.GetFieldValue("PAYMENT_TYPE") = BANKTRF_PMT Then
            Sum7 = Sum7 + Tm.ImportItemEx.GetFieldValue("NET_AMOUNT")
            Sum10 = Sum10 + Tm.ImportItemEx.GetFieldValue("AMOUNT")
            Sum12 = Sum12 + Tm.ImportItemEx.GetFieldValue("FEE_AMOUNT")
         End If
      End If
   Next Tm
   
   m_BillingDoc.CASH_PMT = Sum8
   m_BillingDoc.CHEQUE_PMT = Sum9
   m_BillingDoc.BANKTRF_PMT = Sum10
   
   For Each II In m_BillingDoc.RcpCnDnItems
      If II.Flag <> "D" Then
         If II.GetFieldValue("DOC_ID_TYPE") = INVOICE_DOCTYPE Then
            Sum1 = Sum1 + II.GetFieldValue("ITEM_AMOUNT")
            Sum3 = Sum3 + II.GetFieldValue("PAID_AMOUNT")
         ElseIf II.GetFieldValue("DOC_ID_TYPE") = DN_DOCTYPE Then
            Sum4 = Sum4 + II.GetFieldValue("ITEM_AMOUNT")
         ElseIf II.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
            Sum5 = Sum5 + II.GetFieldValue("ITEM_AMOUNT")
         ElseIf II.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Then
            Sum5 = Sum5 + II.GetFieldValue("ITEM_AMOUNT")
         End If
      End If
   Next II
   Set II = Nothing
   
   Sum11 = 0
   
   For Each BAdd In m_BillingDoc.BillingAdditions
      If BAdd.Flag <> "D" Then
         Sum11 = Sum11 + BAdd.GetFieldValue("ITEM_AMOUNT")
      End If
   Next
   Set BAdd = Nothing
   
   txtTotalDebt.Text = Format(Sum1, "0.00")
   txtAdditionAmount.Text = Format(Sum11, "0.00")
   txtPaidAmount.Text = Format(Sum3, "0.00")
   txtDebitAmount.Text = Format(Sum4, "0.00")
   txtCreditAmount.Text = Format(Sum5, "0.00")
   txtSubTractAmount.Text = Format(Sum6, "0.00")
   txtFromCashTran.Text = Format(Sum7, "0.00")
   txtFeeAmount.Text = Format(Sum12, "0.00")
   
   If DocumentType = RECEIPT1_DOCTYPE Then
      Call CalculateAmountRecriptCash
   ElseIf DocumentType = RECEIPT2_DOCTYPE Then
      Call CalculateAmountRecript
   End If
   
End Sub
Private Sub CalculateAmountRecript()
Dim TempAmt As Double
      
   txtAfterSubTract.Text = FormatNumber(Val(txtPaidAmount.Text) - Val(txtSubTractAmount.Text) + Val(txtAdditionAmount.Text), , False)
   txtVatAmountEx.Text = FormatNumber(Val(txtAfterSubTract.Text) * Val(txtVatPercentEx.Text) / 100, , False)
   txtTotalEx.Text = FormatNumber(Val(txtAfterSubTract.Text) + Val(txtVatAmountEx.Text), , False)
   txtAfterDebitCredit.Text = FormatNumber(Val(txtTotalEx.Text) - Val(txtCreditAmount.Text) + Val(txtDebitAmount.Text), , False)
   txtWHAmount.Text = FormatNumber(Val(txtAfterSubTract.Text) * Val(txtWHPercent.Text) / 100, , False)
   txtGetAmount.Text = FormatNumber(Val(txtAfterDebitCredit.Text) - Val(txtWHAmount.Text), , False)
   txtDifRcp.Text = FormatNumber(Val(ReverseFormatNumber(txtGetAmount.Text)) - Val(txtFeeAmount.Text) - Val(txtFromCashTran.Text), , False)
End Sub
Private Sub CalculateAmountRecriptCash()        'กรณีเงินสดจะคิดอะไรจากอันนี้
Dim TempAmt As Double
Dim OldExtraDiscountAmount As Double
   
   TempAmt = FormatNumber(Val(txtTotalPrice.Text) - Val(txtItemDiscount.Text), , False)
      
   txtIncludeDiscount.Text = FormatNumber(TempAmt, , False)
   
   OldExtraDiscountAmount = Val(txtExtraDiscountAmount.Text)
   
   txtExtraDiscountAmount.Text = FormatNumber(TempAmt * Val(txtExtraDiscountPercent.Text) / 100, , False)
   
   If Trim(Str(OldExtraDiscountAmount)) <> Trim(Str(txtExtraDiscountAmount.Text)) Then
      m_HasDiscountModify = True
   End If
      
   txtIncludeDiscountEx.Text = FormatNumber(TempAmt - Val(txtExtraDiscountAmount.Text), , False)
   
   txtVatAmount.Text = FormatNumber((TempAmt - Val(txtExtraDiscountAmount.Text)) * Val(txtVatPercent.Text) / 100, , False)
   
   txtTotal.Text = FormatNumber(TempAmt - Val(txtExtraDiscountAmount.Text) + Val(txtVatAmount.Text), , False)
   
   txtWHAmount.Text = FormatNumber((TempAmt - Val(txtExtraDiscountAmount.Text)) * Val(txtWHPercent.Text) / 100, , False)
   
   txtGetAmount.Text = FormatNumber(Val(txtTotal.Text) - Val(txtWHAmount.Text), , False)
   txtDifRcp.Text = FormatNumber(Val(ReverseFormatNumber(txtGetAmount.Text)) - Val(txtFeeAmount.Text) - Val(txtFromCashTran.Text), , False)
End Sub

Private Sub uctlGetDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSale_Change()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - (2 * GridEX1.Left)
   If SSFrame5.Visible Then
      SSFrame4.Top = ScaleHeight - SSFrame4.Height - 620 - SSFrame5.Height
      SSFrame3.Top = ScaleHeight - SSFrame3.Height - 620 - SSFrame5.Height
   Else
      SSFrame4.Top = ScaleHeight - SSFrame4.Height - 640
      SSFrame3.Top = ScaleHeight - SSFrame3.Height - 640
   End If
   SSFrame5.Width = ScaleWidth
   SSFrame4.Width = ScaleWidth
   SSFrame3.Width = ScaleWidth
   TabStrip1.Width = GridEX1.Width
   If SSFrame3.Visible Then
      GridEX1.Height = SSFrame3.Top - GridEX1.Top - 40
   Else
      GridEX1.Height = SSFrame4.Top - GridEX1.Top - 40
   End If
   SSFrame5.Top = ScaleHeight - 620 - SSFrame5.Height
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdSaveConsignment.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
   cmdSaveConsignment.Left = cmdPrint.Left - cmdSaveConsignment.Width - 50
   SSFrame2.Width = GridEX1.Width
   SSFrame2.Height = GridEX1.Height
   SSFrame2.Left = GridEX1.Left
   SSFrame2.Top = GridEX1.Top
End Sub
Private Sub CreateBillDocCashTranItems()
Dim Ti As CCashTransferItem
Dim IIEx As CCashTran
Dim Ei As CCashTran
Dim II As CCashTran
   
   Set m_BillingDoc.Payments = Nothing
   Set m_BillingDoc.Payments = New Collection
   
   For Each Ti In m_BillingDoc.TransferItems
      Set IIEx = Ti.ImportItemEx
      If Ti.ImportItemEx.Flag = "" Then
         IIEx.Flag = Ti.Flag
      Else
         IIEx.Flag = Ti.ImportItemEx.Flag
      End If
      Call m_BillingDoc.Payments.add(IIEx)
      If IIEx.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
         Set Ei = Ti.ExportItem
         Set II = Ti.ImportItem
         If Ti.ExportItem.Flag = "" Then
            Ei.Flag = Ti.Flag
         Else
            Ei.Flag = Ti.ExportItem.Flag
         End If
         If Ti.ImportItem.Flag = "" Then
            II.Flag = Ti.Flag
         Else
            II.Flag = Ti.ImportItem.Flag
         End If
         
         Call m_BillingDoc.Payments.add(Ei)
         Call m_BillingDoc.Payments.add(II)
      ElseIf IIEx.GetFieldValue("OLD_PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
         Set Ei = Ti.ExportItem
         Set II = Ti.ImportItem
            
         Ei.Flag = "D"
         II.Flag = "D"
         
         Call m_BillingDoc.Payments.add(Ei)
         Call m_BillingDoc.Payments.add(II)
      End If
   Next Ti
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
Private Sub CheckDisCount()
Dim DocItem As CDocItem
Dim SumAmount As Double
   SumAmount = 0
   For Each DocItem In m_BillingDoc.DocItems
      SumAmount = SumAmount + DocItem.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
   Next DocItem
   
   If Trim(Str(SumAmount)) = Trim(Val(txtExtraDiscountAmount.Text)) Then
      Exit Sub
   ElseIf Trim(Str(SumAmount)) < Trim(Val(txtExtraDiscountAmount.Text)) Then
      For Each DocItem In m_BillingDoc.DocItems
          If DocItem.EXTRA_DISCOUNT_AMOUNT_DIF > 0 Then
            Call DocItem.SetFieldValue("EXTRA_DISCOUNT_AMOUNT", DocItem.GetFieldValue("EXTRA_DISCOUNT_AMOUNT") + 0.01)
            DocItem.EXTRA_DISCOUNT_AMOUNT_DIF = DocItem.EXTRA_DISCOUNT_AMOUNT_DIF - 0.01
            Call CheckDisCount
            Exit Sub
         End If
      Next DocItem
   ElseIf Trim(Str(SumAmount)) > Trim(Val(txtExtraDiscountAmount.Text)) Then
      For Each DocItem In m_BillingDoc.DocItems
          If DocItem.EXTRA_DISCOUNT_AMOUNT_DIF < 0 Then
            Call DocItem.SetFieldValue("EXTRA_DISCOUNT_AMOUNT", DocItem.GetFieldValue("EXTRA_DISCOUNT_AMOUNT") - 0.01)
            DocItem.EXTRA_DISCOUNT_AMOUNT_DIF = DocItem.EXTRA_DISCOUNT_AMOUNT_DIF + 0.01
            Call CheckDisCount
            Exit Sub
         End If
      Next DocItem
   End If
End Sub
