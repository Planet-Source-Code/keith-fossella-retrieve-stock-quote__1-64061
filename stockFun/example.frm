VERSION 5.00
Begin VB.Form example 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "getQuote Example"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3585
   Icon            =   "example.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   3060
      Top             =   840
   End
   Begin VB.TextBox txtsymbol 
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Text            =   "Put stock symbol here !"
      Top             =   277
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Stock Fun !"
      Default         =   -1  'True
      Height          =   360
      Left            =   1875
      TabIndex        =   1
      Top             =   697
      Width           =   1020
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   270
      Shape           =   1  'Square
      Top             =   450
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   540
      Shape           =   1  'Square
      Top             =   210
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   120
      Shape           =   1  'Square
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo errorHandler

    Dim tool As clsStockQuote
    Dim tool2 As clsNumToWord
    Dim retVal As String
    Dim msgStr As String
    
    If Trim(txtsymbol.Text) = "" Or txtsymbol.Text = "Put stock symbol here !" Then
        
        MsgBox "Enter a symbol first !!", vbCritical, "DUH!"
        txtsymbol.SetFocus
        txtsymbol.SelStart = 0
        txtsymbol.SelLength = Len(txtsymbol.Text)
        
        Exit Sub
        
    End If
    
    Set tool = New clsStockQuote
    Set tool2 = New clsNumToWord
    
    retVal = tool.getQuote(txtsymbol.Text)
    
    msgStr = UCase(txtsymbol.Text) & vbNewLine & _
            tool.getCompanyName(txtsymbol.Text) & vbNewLine & _
            "$" & retVal & vbNewLine & _
            tool2.numToWord(retVal, True)
    
    MsgBox msgStr, vbOKOnly, "Stock Info"
    
    txtsymbol.SelStart = 0
    txtsymbol.SelLength = Len(txtsymbol.Text)
    
errorHandler:
    
    Set tool = Nothing
    Set tool2 = Nothing
    
End Sub

Private Sub Form_Load()
    
    txtsymbol.SelStart = 0
    txtsymbol.SelLength = Len(txtsymbol.Text)
    
End Sub

Private Sub Timer1_Timer()

    Static whichShape As Long
    Static owhichShape As Long
    
    Randomize
    
    Do Until whichShape <> owhichShape
        DoEvents
        whichShape = Int((3 - 1 + 1) * Rnd + 1)
    Loop
    
    owhichShape = whichShape
    
    Select Case whichShape
        Case 1
            Shape1.ZOrder (0)
        Case 2
            Shape2.ZOrder (0)
        Case 3
            Shape3.ZOrder (0)
    End Select
      
End Sub
