VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Description As String
Private m_map As Collection

'#########################################################
'#            Constructor / Destructor
'#########################################################

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    
    Description = "UnnamedCollection"
    Set m_map = New Collection

Fin:
    Exit Sub
ErrorHandler:
    myError "colObjects.Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_map = Nothing
End Sub

'#########################################################
'#                      Methods
'#########################################################

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_map.[_NewEnum]
End Function


Public Function AddObject(objEntity As acObject) As Boolean
On Error GoTo ErrorHandler

    If Not Valid(objEntity) Then
        myError "colObjects.AddObject(objEntity) - Invalid objEntity, ignoring"
        AddObject = False
        GoTo Fin
    End If
    
    If Exists(objEntity.GUID) Then
        myDebug "colObjects.AddObject(objEntity) - Object ID " & objEntity.GUID & " already in collection - removing old one"
        Call m_map.Remove(CStr(objEntity.GUID))
    End If
        
    'add it
    'myDebug "colObject.AddObject: " & objEntity.Name & " : " & objEntity.GUID
    Call m_map.Add(objEntity, CStr(objEntity.GUID))
    AddObject = True
        
Fin:
    Exit Function
ErrorHandler:
    myError "colObjects.AddObject [" & Description & "] - " & Err.Description
    AddObject = False
    Resume Fin
End Function

Public Function Add(ByVal lGUID As Long, Optional ByVal sName As String = "UnnamedObject") As acObject
On Error GoTo ErrorHandler

    If Exists(lGUID) Then
        Set Add = Item(lGUID)
        GoTo Fin
    End If
    
    Dim objEntity As New acObject
    objEntity.GUID = lGUID
    objEntity.Name = sName
    Call AddObject(objEntity)
    Set Add = objEntity

Fin:
    Set objEntity = Nothing
    Exit Function
ErrorHandler:
    myError "colObjects.Add [" & Description & "]"
    Resume Fin
End Function

Public Function Exists(ByVal lGUID As Long, Optional ByRef objOut As acObject) As Boolean
On Error GoTo ErrorHandler

    Set objOut = m_map(CStr(lGUID))
    Exists = True

Fin:
    Exit Function
ErrorHandler:
    Exists = False
    Resume Fin
End Function

Public Function Remove(ByVal lGUID As Long) As Boolean
On Error GoTo NotFound

    Call m_map.Remove(CStr(lGUID))
    Remove = True
    
Fin:
    Exit Function
NotFound:
    Remove = False
    Resume Fin
End Function

Public Function Count() As Long
On Error GoTo ErrorHandler
   Count = m_map.Count
Fin:
    Exit Function
ErrorHandler:
    myError "colObjects.Count: " & Err.Description
    Resume Fin
End Function

Public Function Item(ByVal lGUID As Long) As acObject
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrorHandler
    
    Dim obj As acObject
    
    If Exists(lGUID, obj) Then
        Set Item = obj
    Else
        myError "colObjects.Item[" & Description & "] : Object " & lGUID & " doesn't exist. Skipping"
        Set Item = Nothing
    End If
    
Fin:
    Set obj = Nothing
    Exit Function
ErrorHandler:
    myError "colObjects.Item[" & Description & "]"
    Resume Fin
End Function

Public Function GetObjectsList() As String
On Error GoTo ErrorHandler

    Dim objEntity As acObject
    Dim sList As String
    
    If m_map.Count > 0 Then
        Dim bFirst As Boolean
        bFirst = True
        sList = ""
        For Each objEntity In m_map
            If Not bFirst Then sList = sList & ", "
            sList = sList & objEntity.Name
            bFirst = False
        Next objEntity
    Else
        sList = "<Empty>"
    End If
    
    GetObjectsList = sList
    
Fin:
    Set objEntity = Nothing
    Exit Function
ErrorHandler:
    GetObjectsList = ""
    myError "colObjects.GetObjectsList[" & Description & "] - " & Err.Description
    Resume Fin
End Function

Public Sub DebugList()
On Error GoTo ErrorHandler

    Dim objEntity As acObject
    For Each objEntity In m_map
        myDebug "     " & objEntity.Name & " [Type: " & objEntity.ObjectType & " - Range: " & objEntity.GetRange & "]"
    Next objEntity
    
Fin:
    Set objEntity = Nothing
    Exit Sub
ErrorHandler:
    myError "colObjects.DebugList[" & Description & "] - " & Err.Description
    Resume Fin
End Sub
