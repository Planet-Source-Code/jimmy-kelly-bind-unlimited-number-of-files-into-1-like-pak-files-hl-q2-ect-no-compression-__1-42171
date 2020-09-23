VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "File Encryption"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   330
      Left            =   3570
      TabIndex        =   2
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Text            =   "GenericFile"
      Top             =   315
      Width           =   4425
   End
   Begin VB.Label Label1 
      Caption         =   "Encryption Value:"
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   4425
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form2.Hide

End Sub
