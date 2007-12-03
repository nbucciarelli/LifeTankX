VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsD3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_d3ds As D3DService.CService

Private m_bLoaded As Boolean

' ------------------- Public Methods ---------------------------

Public Function Init() As Boolean
On Error GoTo ErrorHandler
    
    MyDebug "[D3D] Init Begin"
    
   'Firstly, create a D3D service:
    Set m_d3ds = g_PluginSite2.object("services\D3DService.Service")
    
    If Not Valid(m_d3ds) Then
        PrintMessage "clsD3D: not able to initialize D3D Service"
        m_bLoaded = False
        GoTo Fin
    End If
    
    m_bLoaded = True
    
    MyDebug "[D3D] Init Complete"

Fin:
    Init = m_bLoaded
    Exit Function
ErrorHandler:
    m_bLoaded = False
    PrintErrorMessage "clsD3D.Init - " & Err.Description
    Resume Fin
End Function


' ------------------- Private Methods ---------------------------

Private Sub Class_Initialize()
    m_bLoaded = False
End Sub

Private Sub Class_Terminate()
    Set m_d3ds = Nothing
End Sub

