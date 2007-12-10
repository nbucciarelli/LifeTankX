VERSION 5.00
Begin VB.Form frmTimers 
   Caption         =   "Timers"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   720
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   720
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
