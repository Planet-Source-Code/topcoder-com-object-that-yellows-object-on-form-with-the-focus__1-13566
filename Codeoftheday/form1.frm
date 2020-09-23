VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Code Makes it Easier for Users to remember where they were"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1890
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1455
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1035
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2310
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2745
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Array"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Yel As New clsChgCol
Option Explicit

Private Sub Text1_GotFocus(Index As Integer)
Yel.ChangeCol
End Sub



Private Sub Text2_GotFocus()
Yel.ChangeCol
End Sub



Private Sub Text3_GotFocus()
Yel.ChangeCol
End Sub



Private Sub Text4_GotFocus()
Yel.ChangeCol
End Sub
