VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstShuffle 
      Height          =   4545
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim start
Dim finished
Dim tProcess
Dim i As Long

    start = Timer
    kRandomize 0, 1200
    finished = Timer
    tProcess = finished - start
    
    MsgBox "Process completed in " & tProcess & " second(s)"
    
    For i = 0 To 55
        lstShuffle.AddItem kRnd
    Next
    
End Sub
