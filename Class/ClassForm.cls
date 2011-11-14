VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFormID As String
Private mFormName As String

Public Property Get FormID() As String
    FormID = mFormID
End Property

Public Property Let FormID(Value As String)
    mFormID = Value
End Property

Public Property Get FormName() As String
    FormID = mFormID
End Property

Public Property Let FormName(Value As String)
    mFormName = Value
End Property
