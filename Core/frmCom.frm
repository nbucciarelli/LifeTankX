VERSION 5.00
Begin VB.Form frmCom 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox myInet 
      Height          =   480
      Left            =   960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox wsock 
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
