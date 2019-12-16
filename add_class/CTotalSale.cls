VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTotalSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EMP_ID As Long
Public SALE_NAME As Long

Public TAGET_PRICE As Double
Public TOTAL_PRICE As Double
Public TOTAL_SELF_PRICE As Double

Public Function Getkey() As String
   Getkey = Trim(Str(EMP_ID))
End Function
