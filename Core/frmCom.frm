VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
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
   Begin InetCtlsObjects.Inet myInet 
      Left            =   960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   5
      RemotePort      =   443
      URL             =   "https://"
      RequestTimeout  =   5
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
