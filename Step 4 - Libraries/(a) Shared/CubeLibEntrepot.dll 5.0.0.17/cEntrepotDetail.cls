VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntrepotDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarStock_ID As Long 'local copy
Private mvarIn_ID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarKey As String 'local copy
'local variable(s) to hold property value(s)
'Private mvarOrigQty As Double 'local copy
Private mvarOrigQty As String 'local copy

'Private mvarOrigGrossWgt As Double 'local copy
Private mvarOrigGrossWgt As String 'local copy

'Private mvarOrigNetWgt As Double 'local copy
Private mvarOrigNetWgt As String 'local copy

'local variable(s) to hold property value(s)
'Private mvarQtyToReserve As Double 'local copy
Private mvarQtyToReserve As String 'local copy

Public Property Let QtyToReserve(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QtyToReserve = 5
    mvarQtyToReserve = vData
End Property


Public Property Get QtyToReserve() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QtyToReserve
    QtyToReserve = mvarQtyToReserve
End Property



Public Property Let OrigNetWgt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewNetWgt = 5
    mvarOrigNetWgt = vData
End Property


Public Property Get OrigNetWgt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewNetWgt
    OrigNetWgt = mvarOrigNetWgt
End Property



Public Property Let OrigGrossWgt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewGrossWgt = 5
    mvarOrigGrossWgt = vData
End Property


Public Property Get OrigGrossWgt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewGrossWgt
    OrigGrossWgt = mvarOrigGrossWgt
End Property






Public Property Let OrigQty(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrigQtyWgt = 5
    mvarOrigQty = vData
End Property


Public Property Get OrigQty() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrigQtyWgt
    OrigQty = mvarOrigQty
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property



Public Property Let In_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.In_ID = 5
    mvarIn_ID = vData
End Property


Public Property Get In_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.In_ID
    In_ID = mvarIn_ID
End Property



Public Property Let Stock_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stock_ID = 5
    mvarStock_ID = vData
End Property


Public Property Get Stock_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stock_ID
    Stock_ID = mvarStock_ID
End Property



