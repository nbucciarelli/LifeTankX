Attribute VB_Name = "Memlocs"
Option Explicit



'Address=0x004BC250 is for Fellowship recruit

'.text:004BC250 ; ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ S U B R O U T I N E ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
'.text:004BC250
'.text:004BC250
'.text:004BC250 sub_4BC250      proc near               ; CODE XREF: sub_4BC360+39p
'.text:004BC250                                         ; sub_4FFB50+C1p
'.text:004BC250
'.text:004BC250 var_4           = dword ptr -4
'.text:004BC250 arg_0           = dword ptr  4


'------------------------------------------------------------------------------------


'Address=0x004BC160 is for Fellowship dismiss

'.text:004BC160 ; ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ S U B R O U T I N E ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
'.text:004BC160
'.text:004BC160
'.text:004BC160 sub_4BC160      proc near               ; CODE XREF: sub_4BD950+43p
'.text:004BC160                                         ; sub_4FFB50+D8p
'.text:004BC160
'.text:004BC160 var_4           = dword ptr -4
'.text:004BC160 arg_0           = dword ptr  4
'.text:004BC160



'Address=0x004E0E40, This=0x0076381C - June
'Private Declare Sub AttackTarget Lib "Memlocs-L2.dll" (ByVal Address As Long, ByVal ThisPtr As Long, ByVal Height As Long)

'0x58BF00 - June - lifestone recall
'Private Declare Sub LifestoneRecall Lib "Memlocs-L2.dll" (ByVal Address As Long)

'0x58BF90 - June - marketplace recall
'Private Declare Sub MarketplaceRecall Lib "Memlocs-L2.dll" (ByVal Address As Long)

'Public Sub LT_AttackTarget(ByVal iHeight As eAttackHeight)
'On Error GoTo ErrorHandler
'
'    Call AttackTarget(&H4E0E40, &H76381C, iHeight + 1)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "LT_AttackTarget - " & Err.Description
'    Resume Fin
'End Sub


'Public Sub LT_LifestoneRecall()
'On Error GoTo ErrorHandler
'
'    Call LifestoneRecall(&H58BF00)
'
'Fin:
'    Exit Sub
'ErrorHandler:
'    PrintErrorMessage "LT_AttackTarget - " & Err.Description
'    Resume Fin
'End Sub


'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'  (Destination As Long, ByVal pSource As Long, ByVal cb As Long)

'Public Function PointerWidth() As Long
'    Dim pwidth As Long
'    pwidth = Hooks.QueryMemLoc("PointerWidth")
'    Call CopyMemory(PointerWidth, pwidth, 4)
'End Function


'Public Function CopyToMemloc(ByVal aTarget As Long, ByVal aValue As Long) As Long
'    Call CopyMemory(aTarget, aValue, 4)
'End Function


' Fellowship recruit memloc

'text:004BC250 ; ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ S U B R O U T I N E ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
'.text:004BC250
'.text:004BC250
'.text:004BC250 sub_4BC250      proc near               ; CODE XREF: sub_4BC360+39p
'.text:004BC250                                         ; sub_4FFB50+C1p
'.text:004BC250
'.text:004BC250 var_4           = dword ptr -4
'.text:004BC250 arg_0           = dword ptr  4


' Salvaging memloc

'.text:004A5220 ; ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ S U B R O U T I N E ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
'.text:004A5220
'.text:004A5220
'.text:004A5220 sub_4A5220      proc near               ; CODE XREF: sub_548E40+5Ep
'.text:004A5220
'.text:004A5220 var_20          = dword ptr -20h
'.text:004A5220 var_1C          = dword ptr -1Ch
'.text:004A5220 var_10          = dword ptr -10h
'.text:004A5220 var_C           = dword ptr -0Ch
'.text:004A5220 var_8           = dword ptr -8
'.text:004A5220 var_4           = dword ptr -4

