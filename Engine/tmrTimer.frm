VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "LTEngine Timers"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrInputQueue 
      Interval        =   30
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer tmrRoute 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
