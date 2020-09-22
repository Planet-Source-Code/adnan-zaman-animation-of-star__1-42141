VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "About"
      Top             =   960
      Width           =   1335
   End
   Begin VB.HScrollBar stopsize 
      Height          =   255
      Left            =   120
      Max             =   4000
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.ComboBox list 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0442
      Left            =   120
      List            =   "Form1.frx":0470
      TabIndex        =   8
      Text            =   "2"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.ComboBox opt 
      Height          =   315
      ItemData        =   "Form1.frx":04A4
      Left            =   120
      List            =   "Form1.frx":04AE
      TabIndex        =   6
      Text            =   "Single Colour"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.HScrollBar star 
      Height          =   255
      Left            =   120
      Max             =   15
      TabIndex        =   7
      Top             =   6960
      Width           =   855
   End
   Begin VB.HScrollBar background 
      Height          =   255
      Left            =   120
      Max             =   15
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   480
      Width           =   1335
   End
   Begin VB.HScrollBar delay 
      Height          =   255
      Left            =   120
      Max             =   500
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   120
   End
   Begin VB.HScrollBar change 
      Height          =   255
      Left            =   120
      Max             =   150
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.HScrollBar step 
      Height          =   255
      Left            =   120
      Max             =   800
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   120
      X2              =   1440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Radius Stop Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "This parameter stops the STAR when the size center goes below this value. The larger the value, the quicker the STAR will stop."
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label txtstopsize 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   4560
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   120
      X2              =   1440
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   120
      X2              =   1440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   120
      X2              =   1440
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Colour:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Determines the colour of the STAR."
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   120
      X2              =   1440
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Colours:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Determines how many colours will be used to draw the STAR."
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Colour of Star:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Determines the colour of the STAR."
      Top             =   5955
      Width           =   1335
   End
   Begin VB.Label txtstar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label txtcolour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Background Colour:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Determines the colour of the background."
      Top             =   5085
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Circle Step Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Change this value to change the form of the STAR."
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Delay Time (in 1/1000th second):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Determines how quickly the drawing will be completed."
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label TimeDelay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label RadiusDecrement 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Radius Change Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Change this value to change the form of the STAR."
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label AngleStep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Code By ADNAN ZAMAN (Protik)

Option Explicit

Dim CenterX As Long
Dim CenterY As Long
Dim Radius As Single
Dim Theta As Single

Dim r As Integer


Private Sub DrawArc()

Static OldX As Single
Static OldY As Single

Dim NewX As Single
Dim NewY As Single

NewX = CenterX + Sin(Theta) * Radius
NewY = CenterY + Cos(Theta) * Radius

Form1.Line (OldX, OldY)-(NewX, NewY), QBColor(r)
Form1.DrawMode = vbCopyPen

OldX = NewX
OldY = NewY

Theta = Theta + Val(AngleStep)
Radius = Radius - Val(RadiusDecrement)

Sleep Val(TimeDelay)
DoEvents
                          
End Sub

Public Sub Drawstar()

Form1.DrawMode = vbNop
CenterX = Me.Width / 2
CenterY = Me.Height / 2
Radius = 0.95 * IIf(CenterX < CenterY, CenterX, CenterY)
Theta = 0

Do While Radius > stopsize.Value
 DrawArc
Loop

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
MsgBox "This program has been created by ADNAN ZAMAN.", vbInformation, "About..."
End Sub

Private Sub Form_Load()

step.Value = 236
change.Value = 10
delay.Value = 110
stopsize.Value = 10

background.Value = 0
star.Value = 9

Me.Show

Do While True
 Me.Cls
 Drawstar
Loop

End Sub

Private Sub Timer1_Timer()

On Error Resume Next

If Form1.BackColor = QBColor(0) Then
 Label1.ForeColor = QBColor(15)
 Label2.ForeColor = QBColor(15)
 Label3.ForeColor = QBColor(15)
 Label4.ForeColor = QBColor(15)
 Label5.ForeColor = QBColor(15)
 Label6.ForeColor = QBColor(15)
 Label7.ForeColor = QBColor(15)
 Label8.ForeColor = QBColor(15)
ElseIf Form1.BackColor = QBColor(1) Then
 Label1.ForeColor = QBColor(15)
 Label2.ForeColor = QBColor(15)
 Label3.ForeColor = QBColor(15)
 Label4.ForeColor = QBColor(15)
 Label5.ForeColor = QBColor(15)
 Label6.ForeColor = QBColor(15)
 Label7.ForeColor = QBColor(15)
 Label8.ForeColor = QBColor(15)
Else
 Label1.ForeColor = QBColor(0)
 Label2.ForeColor = QBColor(0)
 Label3.ForeColor = QBColor(0)
 Label4.ForeColor = QBColor(0)
 Label5.ForeColor = QBColor(0)
 Label6.ForeColor = QBColor(0)
 Label7.ForeColor = QBColor(0)
 Label8.ForeColor = QBColor(0)
End If

AngleStep.Caption = step.Value / 100
RadiusDecrement.Caption = change.Value
TimeDelay.Caption = delay.Value
txtstopsize.Caption = stopsize.Value

txtcolour.BackColor = QBColor(background.Value)
Form1.BackColor = txtcolour.BackColor


txtstar.BackColor = QBColor(star.Value)


If opt.Text = "Single Colour" Then
 star.Enabled = True
 txtstar.Enabled = True
 list.Enabled = False
End If

If opt.Text = "Multiple Colours" Then
 star.Enabled = False
 txtstar.Enabled = False
 list.Enabled = True
End If


If opt.Text = "Single Colour" Then
 r = star.Value
Else
 r = r
End If

If opt.Text = "Multiple Colours" Then
 If list.Text = "2" Then
  r = Rnd * 2
 ElseIf list.Text = "3" Then
  r = Rnd * 3
 ElseIf list.Text = "4" Then
  r = Rnd * 4
 ElseIf list.Text = "5" Then
  r = Rnd * 5
 ElseIf list.Text = "6" Then
  r = Rnd * 6
 ElseIf list.Text = "7" Then
  r = Rnd * 7
 ElseIf list.Text = "8" Then
  r = Rnd * 8
 ElseIf list.Text = "9" Then
  r = Rnd * 9
 ElseIf list.Text = "10" Then
  r = Rnd * 10
 ElseIf list.Text = "11" Then
  r = Rnd * 11
 ElseIf list.Text = "12" Then
  r = Rnd * 12
 ElseIf list.Text = "13" Then
  r = Rnd * 13
 ElseIf list.Text = "14" Then
  r = Rnd * 14
 ElseIf list.Text = "15" Then
  r = Rnd * 15
 Else
  r = r
 End If
End If

End Sub
