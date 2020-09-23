VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowFocusRect 
      Caption         =   "Show Focus Rect"
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   165
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin Project1.CustomButton CustomButton1 
      Height          =   660
      Left            =   300
      TabIndex        =   1
      Top             =   435
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1164
      Caption         =   "Sample &Caption"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&VB Button Actions"
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   1365
      Width           =   1845
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable/Disable"
      Height          =   495
      Left            =   2445
      TabIndex        =   3
      Top             =   1365
      Width           =   1830
   End
   Begin VB.Label lblFocus 
      Caption         =   "^ Has Focus: False"
      Height          =   165
      Left            =   360
      TabIndex        =   9
      Top             =   1110
      Width           =   1830
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   285
      TabIndex        =   8
      Top             =   1965
      Width           =   4035
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   """Disabled"""
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2250
      TabIndex        =   7
      Top             =   1065
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   """MouseOver"""
      Height          =   255
      Index           =   2
      Left            =   2250
      TabIndex        =   6
      Top             =   750
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   """Down"" or Click State"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2235
      TabIndex        =   5
      Top             =   465
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   """Up"" or Normal State"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2235
      TabIndex        =   4
      Top             =   195
      Width           =   2190
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' rem/unrem and add/remove any debug.print statements needed

Private Sub chkShowFocusRect_Click()
    If chkShowFocusRect.Value Then
        CustomButton1.ShowFocusRect = True
    Else
        CustomButton1.ShowFocusRect = False
    End If
End Sub

Private Sub Command1_Click()
    Debug.Print "command1 clicked"
    Debug.Print "command1 Value = "; Command1.Value
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Debug.Print "command1 keydown keycode,shift "; KeyCode, Shift
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
'    Debug.Print "command1 keypress "; KeyAscii
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print "command1 keyup keycode,shift "; KeyCode, Shift
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "command1 mouse down", Button; Shift; x; y
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "command1 mouse up", Button; Shift; x; y
    Debug.Print "command1 value = "; Command1.Value
End Sub

Private Sub Command2_Click()
    CustomButton1.Enabled = Not CustomButton1.Enabled
    Command1.Enabled = Not Command1.Enabled
End Sub

Private Sub Command1_GotFocus()
    Debug.Print "command1 got focus"
End Sub

Private Sub Command1_LostFocus()
    Debug.Print "command1 lost focus"
End Sub

Private Sub Command3_Click()
    CustomButton1.Default = True
End Sub

Private Sub CustomButton1_Click()
    Debug.Print vbTab; "... uc clicked"
    Debug.Print vbTab; "... uc value = "; CustomButton1.Value
End Sub

Private Sub CustomButton1_DblClick()
    Debug.Print vbTab; "... uc double clicked"
End Sub

Private Sub CustomButton1_GotFocus()
    lblFocus.Caption = "^ Has Focus: True"
    Debug.Print vbTab; "... uc got focus"
End Sub

Private Sub CustomButton1_LostFocus()
    lblFocus.Caption = "^ Has Focus: False"
    Debug.Print vbTab; "... uc lost focus"
End Sub

Private Sub CustomButton1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Debug.Print vbTab; "... uc keydown keycode,shift "; KeyCode; Shift
End Sub

Private Sub CustomButton1_KeyPress(KeyAscii As Integer)
'    Debug.Print vbTab; "... uc keypress keycode "; KeyAscii
End Sub

Private Sub CustomButton1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print vbTab; "... uc keyup keycode,shift "; KeyCode; Shift
End Sub

Private Sub CustomButton1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print vbTab; "... uc mouse down", Button; Shift; x; y
End Sub

Private Sub CustomButton1_MouseEnter()
    Debug.Print ">> uc mouse enter"
End Sub

Private Sub CustomButton1_MouseLeave()
    Debug.Print "<< uc mouse leave"
End Sub

Private Sub CustomButton1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print vbTab; "... uc mouse up", Button; Shift; x; y
    Debug.Print vbTab; "... uc value = "; CustomButton1.Value
End Sub

Private Sub CustomButton2_Click()
Unload Me

End Sub

Private Sub Form_Load()
    chkShowFocusRect.Value = Abs(CustomButton1.ShowFocusRect)
End Sub
