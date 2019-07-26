VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1 | Text2"
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
kk = Split(Text1, " | ")(1)
Text2.Text = kk
End Sub
