VERSION 5.00
Begin VB.Form frmRyCamV2b 
   Caption         =   "Statistics"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3825
   Icon            =   "frmRyCamV2b.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstA 
      Height          =   840
      ItemData        =   "frmRyCamV2b.frx":038A
      Left            =   480
      List            =   "frmRyCamV2b.frx":039A
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox lstR 
      Height          =   1620
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox lstS 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Bad Received"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bad Sent"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmRyCamV2b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
