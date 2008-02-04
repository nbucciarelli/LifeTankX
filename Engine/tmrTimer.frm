VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "LTEngine Timers"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFellow 
      Interval        =   1000
      Left            =   1440
      Top             =   720
   End
   Begin VB.Timer tmr3Dtext 
      Interval        =   100
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer tmrHealth 
      Interval        =   300
      Left            =   840
      Top             =   720
   End
   Begin VB.Timer tmrTarget 
      Interval        =   500
      Left            =   240
      Top             =   720
   End
   Begin VB.Timer tmrInputQueue 
      Interval        =   100
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer tmrRoute 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
