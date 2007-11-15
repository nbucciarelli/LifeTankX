VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x As Single
Public y As Single
Public z As Single


Public Function Len2D() As Single
    Len2D = Sqr(x * x + y * y)
End Function

Public Sub Normalize2D()
Dim fLen As Single

    fLen = Len2D
    
    If fLen <= 0 Then
    '    MyDebug "VecNormalize2D: error fLen <=0"
        Exit Sub
    End If
    
    x = x / fLen
    y = y / fLen
    
End Sub

