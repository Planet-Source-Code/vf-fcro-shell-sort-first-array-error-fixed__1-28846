VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VANJA FUCKAR,EMAIL: INGA@VIP.HR"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Shell SORT *Fixed"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Charge"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shell SORT *Original"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bubble SORT"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label7 
      Caption         =   "End:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Start:"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Original By VB ACCELERATOR"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "BUG FIXED BY VANJA FUCKAR"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sorted Array List"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Original Array List"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Do Charge and Shell Sort Original Till Find FIRST ARRAY which does not match."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pit() As String
Dim tmppit() As String
Private Sub Command1_Click()
pit = tmppit
Label1(0) = Time
Label1(1) = ""
DoEvents

BubbleSortStr pit, , True

Label1(1) = Time
DoEvents
List2.Clear
For u = 0 To UBound(pit)
List2.AddItem pit(u)
Next u
End Sub

Private Sub Command2_Click()
ORGorFIX = False
CALC
If pit(0) > pit(1) Then
List2.Selected(0) = True
End If
End Sub
Private Sub CALC()
pit = tmppit
Label1(0) = Time
Label1(1) = ""
DoEvents

ShellSortAny pit

Label1(1) = Time
DoEvents
List2.Clear
For u = 0 To UBound(pit)
List2.AddItem pit(u)
Next u
End Sub
Private Sub Command3_Click()
Charge
List2.Clear
End Sub

Private Sub Command4_Click()
ORGorFIX = True
CALC
End Sub

Private Sub Form_Load()
Charge
End Sub

Private Sub Charge()
ReDim pit(20)
For uu = 0 To UBound(pit)
For u = 0 To 7
Randomize
pit(uu) = pit(uu) & Chr(Int(Rnd * 26) + 65)
Next u
Next uu
tmppit = pit

List1.Clear
For u = 0 To UBound(pit)
List1.AddItem pit(u)
Next u
End Sub



