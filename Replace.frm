VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace color (version 2) by Johannes B!"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   4815
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   8
      ToolTipText     =   "Preview"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdGetPixel 
      Height          =   495
      Left            =   120
      Picture         =   "Replace.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Get color"
      Top             =   1440
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   4815
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   840
      MousePointer    =   2  'Cross
      Picture         =   "Replace.frx":0282
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1080
      Width           =   3060
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click on box to change color!"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "With"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Replace"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API calls
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim BC
Dim A, B As Integer
Dim ColorSelect As Boolean

Private Sub cmdGetPixel_Click()
If ColorSelect = False Then
  ColorSelect = True
 Else
 ColorSelect = False
End If
End Sub

Private Sub Command1_Click()

If Picture2.BackColor = Picture3.BackColor Then
MsgBox "You can't replace a color with the same color!"
Exit Sub
End If
'SHORT CODE!!
On Error Resume Next
A = 0
B = 0

Do
'Get color
BC = GetPixel(Picture1.hdc, A, B)
If BC = Picture2.BackColor Then
'Replace
SetPixel Picture1.hdc, A, B, Picture3.BackColor
End If

'Incrase left
A = A + 1

If A > Picture1.ScaleWidth Then
A = 0
'Incrase top
B = B + 1
Picture1.Refresh
End If
Loop Until B > Picture1.ScaleHeight
End Sub


Private Sub Command2_Click()
On Error GoTo ball
MsgBox "Thanks to Roger Farley for some of the functions! Please vote if you liked the program!"
Do
Form1.Height = Form1.Height - 5
Loop Until Form1.Height < 450
Do

Form1.Width = Form1.Width - 3

Loop Until Form1.Width < 750

End
Exit Sub
ball:
End
Exit Sub
End Sub

Private Sub Form_Load()
  ColorSelect = False
End Sub

Private Sub PBAR_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If ColorSelect = False Then Exit Sub
  Picture2.BackColor = Picture1.Point(x, y)
  ColorSelect = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If ColorSelect = False Then Exit Sub
  Picture4.BackColor = Picture1.Point(x, y)
End Sub

Private Sub Picture2_Click()
On Error GoTo ball
CM.CancelError = True
CM.ShowColor
Picture2.BackColor = CM.Color
Exit Sub
ball:
Exit Sub
End Sub

Private Sub Picture3_Click()
On Error GoTo ball
CM.CancelError = True
CM.ShowColor
Picture3.BackColor = CM.Color
Exit Sub
ball:
Exit Sub
End Sub


