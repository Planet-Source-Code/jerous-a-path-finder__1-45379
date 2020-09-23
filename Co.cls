VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Coordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarX As Integer 'local copy
Private mvarY As Integer 'local copy
Public Property Let Y(ByVal vData As Integer)
    mvarY = vData
End Property
Public Property Get Y() As Integer
    Y = mvarY
End Property
Public Property Let X(ByVal vData As Integer)
    mvarX = vData
End Property
Public Property Get X() As Integer
    X = mvarX
End Property



