VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Dads retirement program"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   195
      Left            =   2.45745e5
      TabIndex        =   0
      Top             =   0
      Width           =   90
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click me to retire!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4155
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3383
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   6600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ButtonWidth As Integer = 2055
Private Const ButtonHeight As Integer = 615

Private Const Surround As Integer = 500


Private Sub Command1_Click()
 MsgBox "You did it!"
End Sub

Private Sub Form_Load()
    Command1.Width = ButtonWidth
    Command1.Height = ButtonHeight
    
    Label1.Visible = False
    
End Sub

Private Function IsOk(X As Single, Y As Single, Top As Long, Left As Long) As Boolean

    Dim buttonX1 As Long
    buttonX1 = Left - Surround
     Dim buttonX2 As Long
    buttonX2 = Left + ButtonWidth + Surround
    
    
    Dim buttonY1 As Long
    buttonY1 = Top - Surround
     Dim buttonY2 As Long
    buttonY2 = Top + ButtonHeight + Surround
    
    Label1.Caption = "X: " & X & " Y: " & Y & vbNewLine & "x1: " & buttonX1 & " y1 " & buttonY1 & vbNewLine & "x2: " & buttonX2 & " y2 " & buttonY2
    IsOk = X > buttonX1 And X < buttonX2 And Y > buttonY1 And Y < buttonY2
    
End Function


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' check to see if mouse is near button
    '
    If IsOk(X, Y, Command1.Top, Command1.Left) Then
        ' Get the window constraints
        Dim FormWidth As Long
         FormWidth = Form1.Width
        
        Dim FormHeight As Long
         FormHeight = Form1.Height
        Randomize
        
        Dim Ok As Boolean
        Ok = False
        
        Dim NewLeft As Long
        Dim NewTop As Long
        
        Do While Not Ok
            NewLeft = Int((Rnd * (FormWidth - ButtonWidth)) + 1)
            NewTop = Int((Rnd * (FormHeight - ButtonHeight)) + 1)
            
            Ok = Not IsOk(X, Y, NewTop, NewLeft)
        Loop
        
        
        Command1.Left = NewLeft
        Command1.Top = NewTop
    End If
End Sub
