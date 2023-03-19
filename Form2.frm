VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Bot"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   4155
      ItemData        =   "Form2.frx":C922
      Left            =   3360
      List            =   "Form2.frx":C924
      TabIndex        =   5
      Top             =   3000
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form2.frx":C926
      Left            =   3840
      List            =   "Form2.frx":C928
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
