VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developed by: Pablito I. Jolbitado"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Answer"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   300
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BASIC CALCULATOR"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter 2nd number:"
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter 1st number:"
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Command1_Click()
'* Developed by: Pablito I. Jolbitado
'  Date Created: December 8, 2014 Time: 7:30
'  Date Publish: December 8, 2014 Time: 7:45
'  Copyright ® Allright Reserved 2014


' HAHAHA :) sana magustuhan mo :P sample lang ito
On Error GoTo pablito:

If Combo1.Text = "Add" Then
MsgBox "The sum of two numbers ( " & Text1.Text & " + " & Text2.Text & " ) " & " is:  " & Val(Text1.Text) + Val(Text2.Text), 64, ""
ElseIf Combo1.Text = "Minus" Then
MsgBox "The difference of two numbers ( " & Text1.Text & " - " & Text2.Text & " ) " & " is:  " & Val(Text1.Text) - Val(Text2.Text), 64, ""
ElseIf Combo1.Text = "Multiply" Then
MsgBox "The product of two numbers ( " & Text1.Text & " x " & Text2.Text & " ) " & " is:  " & Val(Text1.Text) * Val(Text2.Text), 64, ""
ElseIf Combo1.Text = "Divide" Then
MsgBox "The quetient of two numbers ( " & Text1.Text & " / " & Text2.Text & " ) " & " is:  " & Val(Text1.Text) / Val(Text2.Text), 64, ""
End If
Exit Sub

pablito:
MsgBox "Cannot divide by zero", vbCritical, ""

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
Combo1.AddItem "Add"
Combo1.AddItem "Minus"
Combo1.AddItem "Multiply"
Combo1.AddItem "Divide"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case 46
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case 46
Case Else
KeyAscii = 0
End Select
End Sub
