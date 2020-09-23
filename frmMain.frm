VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Zip Distances"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calc Distance"
      Height          =   615
      Left            =   1350
      TabIndex        =   2
      Top             =   60
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Text            =   "11735"
      Top             =   390
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Text            =   "11768"
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sparq's Zip Code Distance Calulcation
'Please Vote.

Private Sub Command1_Click()
    'Set the First Zip Code's Value
    SetCoords Text1, 1
    
    'Set the Second Zip Code's Value
    SetCoords Text2, 2
    
    
    'Display Results
    MsgBox RoundNum(distance(vLat1, vLong1, vLat2, vLong2, "M")) & Chr(9) & " Miles" & vbCrLf & _
           RoundNum(distance(vLat1, vLong1, vLat2, vLong2, "K")) & Chr(9) & " Kilometers" & vbCrLf & _
           RoundNum(distance(vLat1, vLong1, vLat2, vLong2, "N")) & Chr(9) & " Nautical Miles", vbInformation, "Distances:"
End Sub
