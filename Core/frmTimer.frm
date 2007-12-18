VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMacro 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer tmrIrc 
      Interval        =   1000
      Left            =   840
      Top             =   2400
   End
   Begin VB.Timer tmrInputQueue 
      Interval        =   100
      Left            =   2040
      Top             =   240
   End
   Begin VB.Timer tmrRoute 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   240
   End
   Begin VB.Timer tmrTest1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

