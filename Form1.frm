VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   3
      Left            =   3120
      TabIndex        =   4
      Text            =   "23606"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Text            =   "VA"
      Top             =   480
      Width           =   675
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Text            =   "Newport News"
      Top             =   480
      Width           =   2115
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Text            =   "11835 Cannon Boulevard Suite A-102"
      Top             =   120
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Correct"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Requires Microsoft XML SDK 3.0 available at msdn.microsoft.com.

Private Sub Command1_Click()
    Dim address As String
    Dim city As String
    Dim state As String
    Dim zip As String
    address = txt(0).Text
    city = txt(1).Text
    state = txt(2).Text
    zip = txt(3).Text
    MsgBox AddrCorrect(address, city, state, zip)  ' Correct the address
    txt(0).Text = address
    txt(1).Text = city
    txt(2).Text = state
    txt(3).Text = zip
End Sub
