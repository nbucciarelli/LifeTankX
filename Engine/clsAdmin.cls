VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'View
Private m_AdminView As DecalPlugins.IView
Private m_sViewShema As String
'Private m_colRouteFiles As New colFileDir

'Controls
Public WithEvents chkAdminDetect As DecalControls.CheckBox
Attribute chkAdminDetect.VB_VarHelpID = -1
Public WithEvents cmdStop As DecalControls.PushButton
Attribute cmdStop.VB_VarHelpID = -1
Public WithEvents chkAdminTell As DecalControls.CheckBox
Attribute chkAdminTell.VB_VarHelpID = -1
Public WithEvents chkAdminIRC As DecalControls.CheckBox
Attribute chkAdminIRC.VB_VarHelpID = -1

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Constructor / Destructor
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private Sub Class_Initialize()
m_sViewShema = Decrypt("ˆ“»ßåê°×¿½²à¶’¥y•n“®ïÙÙÂ…µ·²à¶’¥|œ~e’íÙè·Ê‰v‡×íÕ×¿Î»ÂcÊÂ’”ÂÎ°È«¯›¢¤~‡l¼¨ÛàØèˆ‡~…t”·”k…l¦áçäæºÑlÄµáàÙØˆ‡¹¦Óå³ã¹Ù¾Ã¯å§¶İÃÊ° ¤ëèåèm…¯À¬âéÕØˆ‡n’c’™”k…ltÕèŞè½Ô¸t³äè×İ¯¢n˜¨ÕÚÜ·ºÓÀÆ²Şì·³Ê¯¿¥áñ’”¹Æ¹¹€”ÜØßŒÉ¹½±¶ŞäÙ®Ùnt¯×ßä±mnt·áé­–ƒ‡lË¬ÖíØ±m–„ˆe’áÕİ²ÍÀ‘e£¯’”®Í±·®×İ­–¿×Á¹e’íÕì¿¢n™±ÓÛÜÙk¦°Á¬à™´Ù¿Ê¯È¬áç’£‰…ltc’™”k¡¯Ã±æëßàkÕ¾ÃªÛİ­–Ê¯µ¯µèŞè½Ô¸ÇqÂîãÜÚÀÈ²à›â¬Ò±‘eÕæÔÇ¿Ô¼vcŞŞÖèˆ‡}†s”™äã»¢n…w¦›ë´ÉÀ¼€”°¢–kÍ±½ªÚí­–}™nt·×ñä±m¸ÀÃ³’ºÜÙ½Ùnƒ’™”k…ltc®Üßâ¿×»ÀcâëßÛ´É‰v‡×ÜÑàÔºÈµáåã¢Í±·®Ôèè–kÓ­Á¨¯›ÓÜ¶¦°Á¬àÍÕà·‡lÀ¨Øí­–ƒ‡lÈ²â¶’§}‡lË¬ÖíØ±m–„ˆe’áÕİ²ÍÀ‘e£¯’”®Í±·®×İ­–¿×Á¹e’íÕì¿¢n˜¨æŞÓèk¦°Á¬à™ŸÈ°Ñ¸Çe¡·”k…ltc’™¬×ºÓÀÆ²Ş™àæºÌµ¸€”½Õ×¬ÑÃ±æëßà¾“¼¨ÕäÒãÃ‡lÂ¤ßŞ­–®Í·•§ßâŞ½¨nt¯×ßä±mnt·áé­–€›ntºÛİäÜˆ‡}u”™ØÙ´Ì´È€”ª¦–kÈ´¹¦İŞÔ±mÙ¾É¨”™äÙÃÙ‰v„ÖæÙâk®—cÁîääÀÙnƒ’™”k¡{·²àíâã·£ˆƒ¹ÛŞç²")

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    SaveSettings
    Set chkAdminDetect = Nothing
    Set cmdStop = Nothing
    Set chkAdminTell = Nothing
    Set chkAdminIRC = Nothing
    
    Set m_AdminView = Nothing
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAdmin_Terminate - " & Err.Description
    Resume Fin
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Properties
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Public Property Get Opened() As Boolean
    If Valid(m_AdminView) Then
        Opened = m_AdminView.Activated
    Else
        Opened = False
    End If
End Property

Public Property Get Active() As Boolean
    If Valid(m_AdminView) Then
        Active = True
    Else
        Active = False
    End If
End Property

Public Property Get DetectAdmins() As Boolean
Dim bRet As Boolean

    If Valid(m_AdminView) Then
        If chkAdminDetect.Checked = True Then
            bRet = True
        Else
            bRet = False
        End If
    Else
        bRet = False
    End If
Fin:
    DetectAdmins = bRet
    Exit Property
End Property

Public Property Get DetectAdminTells() As Boolean
Dim bRet As Boolean

    If Valid(m_AdminView) Then
        If chkAdminTell.Checked = True Then
            bRet = True
        Else
            bRet = False
        End If
    Else
        bRet = False
    End If
Fin:
    DetectAdminTells = bRet
    Exit Property
End Property

Public Property Get ReportAdminIRC() As Boolean
Dim bRet As Boolean

    If Valid(m_AdminView) Then
        If chkAdminIRC.Checked = True Then
            bRet = True
        Else
            bRet = False
        End If
    Else
        bRet = False
    End If
Fin:
    ReportAdminIRC = bRet
    Exit Property
End Property
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                               Private
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                           Public
'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


Public Function LoadAdminView() As Boolean
On Error GoTo ErrorHandler

    Dim bRet As Boolean
    
    If AdminAuth = False Then
        Exit Function
    End If
    
    Set m_AdminView = g_PluginSite.LoadView(m_sViewShema)

    If Valid(m_AdminView) Then
    
        Set chkAdminDetect = m_AdminView.Control("chkAdminDetect")
        Set cmdStop = m_AdminView.Control("cmdStop")
        Set chkAdminTell = m_AdminView.Control("chkAdminTell")
        Set chkAdminIRC = m_AdminView.Control("chkAdminIRC")

        LoadSettings
        bRet = True
    Else
        PrintErrorMessage "clsAdmin.LoadEditorView - bad m_AdminView"
        bRet = False
    End If
    
Fin:
    LoadAdminView = bRet
    Exit Function
ErrorHandler:
    bRet = False
    PrintErrorMessage "clsAdmin.LoadAdminView - " & Err.Description
    Resume Fin
End Function

Public Sub ShowAdmin(Optional bAutoPosition As Boolean = True)
On Error GoTo ErrorHandler
    If AdminAuth = False Then
        Exit Sub
    End If
    If m_AdminView Is Nothing Then
        Call LoadAdminView
        
        If bAutoPosition Then
            Dim newPos As Decal.tagRECT
        
            With m_AdminView.Position
                newPos.Bottom = .Bottom
                newPos.Left = .Left
                newPos.Right = .Right
                newPos.Top = .Top
            End With
        
            'newPos.Left = g_Hooks.Area3DWidth - newPos.Right - 50
            'newPos.Top = g_Hooks.Area3DHeight - newPos.Bottom - 50
            newPos.Left = g_Hooks.AC3DRegionRect.Right - g_Hooks.AC3DRegionRect.Left - newPos.Right - 50
            newPos.Top = g_Hooks.AC3DRegionRect.Top - g_Hooks.AC3DRegionRect.Bottom - newPos.Bottom - 50
        
            'update screen pos
            'm_AdminView.Position = newPos
            m_AdminView.Position.Top = newPos.Top
            m_AdminView.Position.Bottom = newPos.Bottom
            m_AdminView.Position.Left = newPos.Left
            m_AdminView.Position.Right = newPos.Right
            'm_AdminView.Position = newPos

        End If
    End If
    
    If Valid(m_AdminView) Then
        Call m_AdminView.Activate
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAdmin.ShowAdmin - " & Err.Description
    Resume Fin
End Sub

Public Sub HideAdmin()
On Error GoTo ErrorHandler

    If Valid(m_AdminView) Then
        Call m_AdminView.Deactivate
    End If
    
Fin:
    Exit Sub
ErrorHandler:
    PrintErrorMessage "clsAdmin.HideAdmin - " & Err.Description
    Resume Fin
End Sub

Private Sub LoadSettings()
    chkAdminDetect.Checked = GetSetting("Lifetank", "Lifetank", "1", CBool("True"))
    chkAdminIRC.Checked = GetSetting("Lifetank", "Lifetank", "2", CBool("True"))
    chkAdminTell.Checked = GetSetting("Lifetank", "Lifetank", "3", CBool("True"))
End Sub

Private Sub SaveSettings()
    SaveSetting "Lifetank", "Lifetank", "1", CStr(chkAdminDetect.Checked)
    SaveSetting "Lifetank", "Lifetank", "2", CStr(chkAdminIRC.Checked)
    SaveSetting "Lifetank", "Lifetank", "2", CStr(chkAdminDetect.Checked)
End Sub
Private Sub cmdStop_Accepted(ByVal nID As Long)
On Error GoTo Error_Handler
    
    'stop admin alert
    If g_AntiBan.AdminDetected Or g_AntiBan.AlarmTriggered Then
        Call g_AntiBan.SetAdminAlert(False)
    End If

Fin:
    Exit Sub
Error_Handler:
    PrintErrorMessage "clsAdmin.cmd_Stop.Accepted - " & Err.Description
    Resume Fin
End Sub

