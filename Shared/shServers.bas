Attribute VB_Name = "shServers"
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                 SHARED MODULE                       [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
'[[                                                     [[
'[[             AC Servers Identifiers                  [[
'[[                                                     [[
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[
Option Explicit

Public Enum eGameServer
    SV_WINTERSEEB = 0
    SV_MORNINGTHAW
    SV_HARVESTGAIN
    SV_FROSTFELL
    SV_SOCLAIM
    SV_THISTLEDOWN
    SV_LEAFCULL
    SV_DARKTIDE
    SV_VERDANTINE
    NUM_AC_SERVERS
    SV_ANY
    SV_NONE
End Enum

Public Function GetServerName(ByVal ServerId As eGameServer) As String
Dim sRet As String
    
    Select Case ServerId
        Case SV_WINTERSEEB
            sRet = "Wintersebb"
            
        Case SV_MORNINGTHAW
            sRet = "Morningthaw"
    
        Case SV_HARVESTGAIN
            sRet = "Harvestgain"
            
        Case SV_FROSTFELL
            sRet = "Frostfell"
            
        Case SV_SOCLAIM
            sRet = "Solclaim"
            
        Case SV_THISTLEDOWN
            sRet = "Thistledown"
    
        Case SV_LEAFCULL
            sRet = "Leafcull"
            
        Case SV_DARKTIDE
            sRet = "Darktide"
            
        Case SV_VERDANTINE
            sRet = "Verdantine"
            
        Case SV_ANY
            sRet = "Any"
        
        Case Else
            sRet = "Unknown Server " & ServerId
            
    End Select

    GetServerName = sRet
End Function

Public Function GetShortServerName(ByVal ServerId As eGameServer) As String
Dim sRet As String
    
    Select Case ServerId
        Case SV_WINTERSEEB
            sRet = "WE"
            
        Case SV_MORNINGTHAW
            sRet = "MT"
    
        Case SV_HARVESTGAIN
            sRet = "HG"
            
        Case SV_FROSTFELL
            sRet = "FF"
            
        Case SV_SOCLAIM
            sRet = "SC"
            
        Case SV_THISTLEDOWN
            sRet = "TD"
    
        Case SV_LEAFCULL
            sRet = "LC"
            
        Case SV_DARKTIDE
            sRet = "DT"
            
        Case SV_VERDANTINE
            sRet = "VT"
            
        Case SV_ANY
            sRet = "ANY"
        
        Case Else
            sRet = "UNKN" & ServerId
            
    End Select

    GetShortServerName = sRet
End Function

Public Function GetVBServerName(ByVal ServerId As eGameServer) As String
Dim sRet As String
    
    Select Case ServerId
        Case SV_WINTERSEEB
            sRet = "SV_WINTERSEEB"
            
        Case SV_MORNINGTHAW
            sRet = "SV_MORNINGTHAW"
    
        Case SV_HARVESTGAIN
            sRet = "SV_HARVESTGAIN"
            
        Case SV_FROSTFELL
            sRet = "SV_FROSTFELL"
            
        Case SV_SOCLAIM
            sRet = "SV_SOCLAIM"
            
        Case SV_THISTLEDOWN
            sRet = "SV_THISTLEDOWN"
    
        Case SV_LEAFCULL
            sRet = "SV_LEAFCULL"
            
        Case SV_DARKTIDE
            sRet = "SV_DARKTIDE"
            
        Case SV_VERDANTINE
            sRet = "SV_VERDANTINE"
            
        Case SV_ANY
            sRet = "SV_ANY"
        
        Case SV_NONE
            sRet = "SV_NONE"
        
        Case Else
            sRet = "Unknown Serverid " & ServerId
            
    End Select

    GetVBServerName = sRet
End Function

Public Function GetServerIdByName(ByVal ServerName As String) As eGameServer
    Dim iRet As eGameServer
    
    ServerName = LCase(ServerName)
    
    Select Case ServerName
        Case "wintersebb"
            iRet = SV_WINTERSEEB
            
        Case "morningthaw"
            iRet = SV_MORNINGTHAW
    
        Case "harvestgain"
            iRet = SV_HARVESTGAIN
            
        Case "frostfell"
            iRet = SV_FROSTFELL
            
        Case "solclaim"
            iRet = SV_SOCLAIM
            
        Case "thistledown"
            iRet = SV_THISTLEDOWN
    
        Case "leafcull"
            iRet = SV_LEAFCULL
            
        Case "darktide"
            iRet = SV_DARKTIDE
            
        Case "verdantine"
            iRet = SV_VERDANTINE
        
        Case Else
            iRet = SV_NONE
            
    End Select
    
    'Assume it's darktide if server name couldn't be determined...
    If iRet = SV_NONE Then 'And InStr(1, ServerName, "darktide") Then
        iRet = SV_DARKTIDE
    End If
    
    GetServerIdByName = iRet
    
End Function
