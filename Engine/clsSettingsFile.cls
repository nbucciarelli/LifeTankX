VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettingsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Extends some functionnality of the DataFile
Private m_cfgFile As DataFile
Private m_bLoaded As Boolean

'#####################################################################################
'#
'#                          CONSTRUCTOR / DESTRUCTOR
'#
'#####################################################################################

Private Sub Class_Initialize()
    Set m_cfgFile = New DataFile
    m_bLoaded = False
End Sub

Private Sub Class_Terminate()
    Set m_cfgFile = Nothing
    m_bLoaded = False
End Sub

'#####################################################################################
'#
'#                              PROPERTIES
'#
'#####################################################################################

Public Property Get File() As DataFile
    Set File = m_cfgFile
End Property

Public Property Get Loaded() As Boolean
    Loaded = m_bLoaded
End Property

'#####################################################################################
'#
'#                              PRIVATE
'#
'#####################################################################################


'#####################################################################################
'#
'#                              FILE CONTROL
'#
'#####################################################################################

Public Function LoadContent(ByVal sPath As String) As Boolean
On Error GoTo ErrorHandler

    'First, reset content
    m_cfgFile.ResetData
    
    'The load the new content
    m_bLoaded = m_cfgFile.Load(sPath)
    
    If Not m_bLoaded Then
        PrintErrorMessage "[LoadContent] Could not load config : " & sPath
    End If
    
Fin:
    LoadContent = m_bLoaded
    Exit Function
ErrorHandler:
    m_bLoaded = False
    PrintErrorMessage "LoadContent - " & Err.Description
    Resume Fin
End Function

Public Sub BeginSave()
On Error GoTo ErrorHandler
    
    Call m_cfgFile.ResetData
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsSettingsFile.BeginSave - " & Err.Description
    Resume Fin
End Sub

Public Function EndSave() As Boolean
On Error GoTo ErrorHandler
    
    If Not m_cfgFile.save Then
        PrintErrorMessage "[Settings] Could not save " & m_cfgFile.FileName
    Else
        MyDebug "[Settings] Saved : " & m_cfgFile.FileName
    End If
    
Fin:
    Exit Function
ErrorHandler:
    EndSave = False
    PrintErrorMessage "clsSettingsFile.EndSave - " & Err.Description
    Resume Fin
End Function


'#####################################################################################
'#
'#                              PUBLIC
'#
'#####################################################################################

Public Function SaveValue(ByVal sId As String, ByVal sValue As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim bRet As Boolean
    Dim dat As New clsDataEntry
    
    If Valid(m_cfgFile) Then
        dat.AddParam TAG_ID, sId
        dat.AddParam TAG_VALUE, sValue
        bRet = m_cfgFile.AddData(dat)
    End If
    
    If Not bRet Then
        PrintErrorMessage "clsSettingsFile.SaveValue - Failed to add new data to " & m_cfgFile.FileName
    End If
    
Fin:
    Set dat = Nothing
    SaveValue = bRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSettingsFile.SaveValue - " & Err.Description
    bRet = False
    Resume Fin
End Function

Public Function GetValue(ByVal sId As String, Optional ByVal vDefaultVal As Variant = 0) As Variant
On Error GoTo ErrorHandler

    Dim vRet As Variant
    Dim dat As clsDataEntry
        
    If m_cfgFile.FindData(TAG_ID, dat, sId) Then
        If dat.ParamExist(TAG_VALUE) Then
            vRet = dat.Param(TAG_VALUE)
        Else
            PrintErrorMessage m_cfgFile.FileName & " -- GetValue : control " & sId & " has no Value field"
            vRet = vDefaultVal
        End If
    Else
        MyDebug m_cfgFile.FileName & " -- GetValue : could not find control: " & sId
        vRet = vDefaultVal
    End If
        
Fin:
    Set dat = Nothing
    GetValue = vRet
    Exit Function
ErrorHandler:
    PrintErrorMessage "clsSettingsFile.GetValue - " & Err.Description
    vRet = vDefaultVal
    Resume Fin
End Function

Public Function SaveTextbox(ByVal txtControl As DecalControls.Edit, ByVal sControlName As String) As Boolean
    If Valid(txtControl) Then
        SaveTextbox = SaveValue(sControlName, txtControl.Text)
    Else
        MyDebug "Invalid Edit control " & sControlName
        SaveTextbox = False
    End If
End Function

Public Function SaveCheckbox(ByVal chkControl As DecalControls.CheckBox, ByVal sControlName As String) As Boolean
    If Valid(chkControl) Then
        SaveCheckbox = SaveValue(sControlName, BoolToInteger(chkControl.Checked))
    Else
        MyDebug "Invalid Checkbox control " & sControlName
        SaveCheckbox = False
    End If
End Function

Public Function SaveChoice(ByVal chControl As DecalControls.Choice, ByVal sControlName As String) As Boolean
    If Valid(chControl) Then
        SaveChoice = SaveValue(sControlName, chControl.Selected)
    Else
        MyDebug "Invalid Choicelist control " & sControlName
        SaveChoice = False
    End If
End Function

Public Function SaveSlider(ByVal sldControl As DecalControls.Slider, ByVal sControlName As String) As Boolean
    If Valid(sldControl) Then
        SaveSlider = SaveValue(sControlName, sldControl.SliderPosition)
    Else
        MyDebug "Invalid Slider control " & sControlName
        SaveSlider = False
    End If
End Function

