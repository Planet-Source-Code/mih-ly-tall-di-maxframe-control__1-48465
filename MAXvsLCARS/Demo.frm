VERSION 5.00
Begin VB.Form DemoForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Command6"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame ShipControlsFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000099FF&
      Height          =   6975
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin MaxVsLCARSControls.MaxFrame MaxFrame4 
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Top             =   1080
         Width           =   3015
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trek Generation Regular"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   6724095
         Caption         =   "About..."
         PressedColor    =   10079487
         CapLines        =   2
         State           =   0
         SaveHeight      =   1335
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Does it look like Max or Trek?    You decide!"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   2535
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame3 
         Height          =   345
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   3015
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trek TNG Monitors"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   16751001
         Caption         =   "Enviromental controls"
         PressedColor    =   16764108
         Alignment       =   1
         State           =   0
         SaveHeight      =   3135
         Begin VB.ListBox List1 
            Height          =   840
            ItemData        =   "Demo.frx":0000
            Left            =   240
            List            =   "Demo.frx":0016
            TabIndex        =   10
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Take care!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H006666CC&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Air"
            BeginProperty Font 
               Name            =   "Trek TNG Monitors"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   615
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1935
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame2 
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   3015
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trek TNG Monitors"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   6710988
         Caption         =   "Weapon Controls"
         PressedColor    =   10066431
         Alignment       =   1
         State           =   0
         SaveHeight      =   3735
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   720
            TabIndex        =   8
            Text            =   "Controls"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Caption         =   "WEAPON"
            Height          =   255
            Left            =   1560
            TabIndex        =   7
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "some"
            Height          =   375
            Left            =   960
            TabIndex        =   6
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "for"
            Height          =   375
            Left            =   720
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "It's the place"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "BANG!"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   240
            TabIndex        =   15
            Top             =   3000
            Width           =   2415
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame1 
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3015
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trek TNG Monitors"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   39423
         Caption         =   "Engine Controls"
         Alignment       =   1
         State           =   0
         SaveHeight      =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "Place some engine controls here"
            Height          =   3255
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   2415
         End
      End
      Begin MaxVsLCARSControls.PanelVScroll FramScroll 
         Height          =   6975
         Left            =   3000
         TabIndex        =   18
         Top             =   0
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   12303
         BackColor       =   0
         ForeColor       =   16751001
         Min             =   0
      End
   End
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FRAMESPACING = 4
Dim FrameIndexes() As Integer, ScrollIndex As Integer, LastVal As Integer, CurrVal As Integer
Dim FrameTops() As Integer, MaximHeight As Integer

'Handles all MaxFrame-s in TargetFrame, and the first PanelVScroll it can find in TargetFrame
Private Sub ArrangeMaxFrames(TargetFrame As Frame, LastChanged As String, Optional LastChangedIndex As Integer = 0)
ReDim FrameIndexes(0)
ReDim FrameTops(0)

'Determine the controls to handle
For X = 0 To Me.Controls.Count - 1
   If (TypeOf Me.Controls(X) Is MaxFrame) And (Me.Controls(X).Container Is TargetFrame) Then
      FrameIndexes(UBound(FrameIndexes)) = X
      FrameTops(UBound(FrameTops)) = Me.Controls(X).Top
      ReDim Preserve FrameIndexes(UBound(FrameIndexes) + 1)
      ReDim Preserve FrameTops(UBound(FrameTops) + 1)
   ElseIf (TypeOf Me.Controls(X) Is PanelVScroll) And (Me.Controls(X).Container Is TargetFrame) Then
      ScrollIndex = X
   End If
Next 'X
ReDim Preserve FrameIndexes(UBound(FrameIndexes) - 1)
ReDim Preserve FrameTops(UBound(FrameTops) - 1)

'Sorting
For X = 0 To UBound(FrameTops)
   For Y = UBound(FrameTops) To X + 1 Step -1
      If FrameTops(Y) < FrameTops(X) Then
         temp = FrameTops(X)
         FrameTops(X) = FrameTops(Y)
         FrameTops(Y) = temp
         
         temp = FrameIndexes(X)
         FrameIndexes(X) = FrameIndexes(Y)
         FrameIndexes(Y) = temp
      End If
   Next 'y
Next 'x

'For X = 0 To UBound(FrameIndexes)
'   If X = 0 Then
'      Me.Controls(FrameIndexes(X)).Top = -Me.Controls(ScrollIndex).Value
'   Else
'      Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X - 1)).Top + Me.Controls(FrameIndexes(X - 1)).Height + FRAMESPACING * Screen.TwipsPerPixelY
'   End If
'Next X
'
'MaximHeight = (Me.Controls(FrameIndexes(UBound(FrameIndexes))).Top + Me.Controls(FrameIndexes(UBound(FrameIndexes))).Height + FRAMESPACING * Screen.TwipsPerPixelY) - TargetFrame.Height
'If MaximHeight > 0 Then
'   Me.Controls(ScrollIndex).Max = MaximHeight / Screen.TwipsPerPixelY
'   Me.Controls(ScrollIndex).LargeChange = Me.Controls(ScrollIndex).Max / UBound(FrameIndexes)
'   Me.Controls(ScrollIndex).Visible = True
'Else
'   Me.Controls(ScrollIndex).Visible = False
'   Me.Controls(ScrollIndex).Value = 0
'End If

'Setup maximum length to be scrolled
MaximHeight = 0
For X = 0 To UBound(FrameIndexes)
   MaximHeight = MaximHeight + Me.Controls(FrameIndexes(X)).Height + FRAMESPACING * Screen.TwipsPerPixelY
Next X
MaximHeight = MaximHeight - TargetFrame.Height

'Make visible and setup scrollbar if needed
If MaximHeight > 0 Then
   Me.Controls(ScrollIndex).Max = MaximHeight / Screen.TwipsPerPixelY
   Me.Controls(ScrollIndex).LargeChange = Me.Controls(ScrollIndex).Max / UBound(FrameIndexes)
   Me.Controls(ScrollIndex).Visible = True
Else
   Me.Controls(ScrollIndex).Visible = False
   Me.Controls(ScrollIndex).Value = 0
End If

If Index = 0 Then

'Place frames in order
   For X = 0 To UBound(FrameIndexes)
      If X = 0 Then
         Me.Controls(FrameIndexes(X)).Top = 0
      Else
         Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X - 1)).Top + Me.Controls(FrameIndexes(X - 1)).Height + FRAMESPACING * Screen.TwipsPerPixelY
      End If
   Next X
End If

End Sub

Private Sub Form_Load()
FramScroll.Width = ShipControlsFrame.Width - MaxFrame1.Width
FramScroll.Left = MaxFrame1.Width
End Sub

Private Sub FramScroll_ValueChange()
CurrVal = FramScroll.Value
For X = 0 To UBound(FrameIndexes)
   Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X)).Top - (CurrVal - LastVal) * Screen.TwipsPerPixelY
Next 'x
LastVal = CurrVal
End Sub

Private Sub MaxFrame1_StateChanged()
ArrangeMaxFrames ShipControlsFrame, "maxframe1"
End Sub

Private Sub MaxFrame2_StateChanged()
ArrangeMaxFrames ShipControlsFrame, "maxframe2"
End Sub

Private Sub MaxFrame3_StateChanged()
ArrangeMaxFrames ShipControlsFrame, "maxframe3"
End Sub

Private Sub MaxFrame4_StateChanged()
ArrangeMaxFrames ShipControlsFrame, "maxframe4"
End Sub
