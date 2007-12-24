VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Handles the items exceptions list for Salvage/Sell
Option Explicit

Private m_colItems As colObjects

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    MyDebug "clsExceptions.Initialize"
    Set m_colItems = New colObjects
    MyDebug "clsExceptions.Init done"

Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsExceptions.Initialize - " & Err.Description
    Resume Fin
End Sub

Private Sub Class_Terminate()
    Set m_colItems = Nothing
End Sub

Public Property Get Items() As colObjects
    Set Items = m_colItems
End Property

'Load from file
Public Function LoadFromFile() As Boolean
On Error GoTo ErrorHandler

    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As New DOMDocument
    Dim sPath As String
    
    Set m_colItems = New colObjects
    sPath = g_Settings.Profile.FullPath & "\" & FILE_EXCEPTIONS
    
    If xmlDoc.Load(sPath) Then
        For Each xmlNode In xmlDoc.documentElement.childNodes
        
            Dim objItem As acObject
            Dim nodeAttrib As IXMLDOMNode
            Set objItem = New acObject
            
            For Each nodeAttrib In xmlNode.Attributes
                Select Case LCase(nodeAttrib.nodeName)
                    Case "guid"
                        objItem.Guid = CLng(Val(nodeAttrib.Text))
                    Case "icon"
                        objItem.Icon = CLng(Val(nodeAttrib.Text))
                    Case "name"
                        objItem.Name = nodeAttrib.Text
                End Select
            Next nodeAttrib
        
            If objItem.Guid <> 0 Then
                Call m_colItems.addObject(objItem)
                MyDebug "Exceptions : added " & objItem.Name & " (guid: " & objItem.Guid & ")"
            End If
    
            Set objItem = Nothing
        Next
    Else
        PrintErrorMessage "Exceptions.LoadFromFile - Failed to load " & sPath
        GoTo Fin
    End If
    
    MyDebug "Exceptions - " & m_colItems.Count & " items added to list"
    LoadFromFile = True
    
Fin:
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsExceptions.LoadFromFile(" & sPath & ") - " & Err.Description
    LoadFromFile = False
    Resume Fin
End Function

'Load from file
Public Function SaveToFile() As Boolean
On Error GoTo ErrorHandler

    Dim sPath As String
    Dim xmlDoc As New DOMDocument
    Dim xmlRootElem As IXMLDOMElement
    Dim xmlElemItem As IXMLDOMElement

    sPath = g_Settings.Profile.FullPath & "\" & FILE_EXCEPTIONS
    
    ' Creates root element
    Set xmlRootElem = xmlDoc.createElement("Exceptions")
    Call xmlDoc.appendChild(xmlRootElem)
     
    ' Creates Member elements
    Dim objItem As acObject
    For Each objItem In m_colItems
        MyDebug "Exceptions - Saving item: " & objItem.Name
        Set xmlElemItem = xmlDoc.createElement("item")
        Call xmlElemItem.setAttribute("guid", objItem.Guid)
        Call xmlElemItem.setAttribute("name", objItem.Name)
        Call xmlElemItem.setAttribute("icon", objItem.Icon)
        Call xmlRootElem.appendChild(xmlElemItem)
    Next objItem
    
    ' Saves XML data to disk.
    Call xmlDoc.save(sPath)
    MyDebug "Exceptions - " & m_colItems.Count & " items saved to " & sPath
    
    SaveToFile = True
    
Fin:
    Set xmlDoc = Nothing
    Set xmlRootElem = Nothing
    Set xmlElemItem = Nothing
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsExceptions.SaveToFile(" & sPath & ") - " & Err.Description
    SaveToFile = False
    Resume Fin
End Function


Public Function AddItem(objItem As acObject) As Boolean
On Error GoTo ErrorHandler

    If Not Valid(objItem) Then
        PrintErrorMessage "clsExceptions.AddItem - invalid objItem"
        GoTo Fin
    End If
    
    Dim objException As New acObject
    objException.Guid = objItem.Guid
    objException.Name = objItem.Name
    objException.Icon = objItem.Icon
    AddItem = m_colItems.addObject(objException)

Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsExceptions.AddItem - " & Err.Description
    AddItem = False
    Resume Fin
End Function

Public Function RemoveItem(ByVal lObjGuid As Long) As Boolean
On Error GoTo ErrorHandler

    If m_colItems.Remove(lObjGuid) Then
        RemoveItem = True
    Else
        RemoveItem = False
        PrintWarning "clsExceptions.RemoveItem - Failed to remove Item #" & lObjGuid & " from Exceptions list : item may not exist - Ignoring"
    End If
    
Fin:
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsExceptions.RemoveItem - " & Err.Description
    RemoveItem = False
    Resume Fin
End Function
