VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Duplicates"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   2550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdRemoveDupesAPI 
      Caption         =   "Remove Dupes API"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdRemoveDupesVB 
      Caption         =   "Remove Dupes VB"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - October 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Private Sub cmdRemoveDupesAPI_Click()
    Dim Tcnt As Long
    Tcnt = GetTickCount 'Start time
    Me.Caption = "Removing..."
    cmdRemoveDupesAPI.Enabled = False
    cmdRemoveDupesVB.Enabled = False
    RemoveDupesFromList List1 'call the function
    Me.Caption = (GetTickCount - Tcnt) / 1000 & " seconds(API)" 'Time elapsed
End Sub
Private Sub cmdRemoveDupesVB_Click()
    Dim Tcnt As Long
    Tcnt = GetTickCount 'Start time
    Me.Caption = "Removing..."
    cmdRemoveDupesAPI.Enabled = False
    cmdRemoveDupesVB.Enabled = False
    RemoveDupes List1 'call the function
    Me.Caption = (GetTickCount - Tcnt) / 1000 & " seconds(VB)" 'Time elapsed
End Sub
Private Sub cmdReset_Click()
    'Refill the listbox with duplicate entries(1500)
    Dim z As Long
    List1.Clear
    For z = 1 To 500
        List1.AddItem "Item " & z
        List1.AddItem "Item " & z
    Next
    For z = 1 To 500
        List1.AddItem "Item " & z
    Next
    cmdRemoveDupesAPI.Enabled = True
    cmdRemoveDupesVB.Enabled = True
    Me.Caption = "Remove Duplicates"
End Sub
Private Sub Form_Load()
    cmdReset_Click 'Fill the listbox
End Sub
