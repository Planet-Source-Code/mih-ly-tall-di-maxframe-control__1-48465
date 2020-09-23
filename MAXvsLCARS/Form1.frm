VERSION 5.00
Object = "*\AMaxCP.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LCARS styled MaxFrames"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Close all"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open all"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame ShipControlsFrame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Fram"
      ForeColor       =   &H00808080&
      Height          =   6975
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin MaxVsLCARSControls.MaxFrame MaxFrame2 
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   13408716
         Caption         =   "Example"
         PressedColor    =   16764159
         State           =   0
         SaveHeight      =   2295
         Begin VB.CommandButton Command1 
            Caption         =   "This is a control inside MaxFrame2. Press this button."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   $"Form1.frx":0000
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   1095
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   2415
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame1 
         Height          =   345
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   39423
         Caption         =   "About MaxFrames..."
         State           =   0
         SaveHeight      =   1815
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "The idea for the MaxFrame control came from discreet's 3D Studio MAX. MAX uses this technique to group it's complex controls."
            ForeColor       =   &H00FF9999&
            Height          =   855
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame3 
         Height          =   345
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   6710988
         Caption         =   "More info:"
         PressedColor    =   10066431
         State           =   0
         SaveHeight      =   1665
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "More info on the handling code and the controls can be found in Form1's code."
            ForeColor       =   &H006699FF&
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2655
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame5 
         Height          =   345
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   6724095
         Caption         =   "Copyright info"
         PressedColor    =   10079487
         State           =   0
         SaveHeight      =   2775
         Begin VB.Label Label5 
            BackColor       =   &H00000000&
            Caption         =   $"Form1.frx":008A
            ForeColor       =   &H00FFCCFF&
            Height          =   975
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            Caption         =   $"Form1.frx":012B
            ForeColor       =   &H00FF9999&
            Height          =   1095
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   2655
         End
      End
      Begin MaxVsLCARSControls.MaxFrame MaxFrame4 
         Height          =   345
         Left            =   0
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   16751001
         Caption         =   "About"
         PressedColor    =   16764108
         State           =   0
         SaveHeight      =   2895
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "About this program..."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000099FF&
            Height          =   1215
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "LCARSButton, VScroll, MaxFrame && handling code by Msi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   975
            Left            =   240
            TabIndex        =   5
            Top             =   1800
            Width           =   2295
         End
      End
      Begin MaxVsLCARSControls.PanelVScroll FramScroll 
         Height          =   6975
         Left            =   2880
         TabIndex        =   1
         Top             =   0
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   12303
         BackColor       =   0
         ForeColor       =   16751001
         Min             =   0
      End
   End
   Begin VB.Label Label8 
      Caption         =   "You can open/close a frames on the right by clicking on it's label."
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A note for the code:
'You will notice that i write the "Next" keywords
'a little bit differently than usual.
'This way ("Next 'X") is a little bit faster than
'if you write the "For" control variable after "Next",
'because VB checks if this variable is the proper
'"For" control variable or not. Don't ask why.

'P.S.:Remember, this code is not well optimised:
'It's up to you if you find the speed of it inadequate.

'Constants to make things look better separated.
'You better leave these as they are.
Private Const FRAMESPACING = 4
Private Const SCROLLBARSPACING = 2

Private FrameIndexes() As Long, FrameTops() As Long, ScrollIndex As Integer, LastVal As Integer
'These are just temporary variables.
Private MaximHeight As Long, Temp As Long, TempHeight As Long, TempState As Byte, CurrVal As Long, X As Integer

'Initiates the FrameIndexes and FrameTops table,
'and sorts them.
'Currently, ArrangeMaxFrames() calls this every time it
'starts. This is unnecessary when you only have one Frame
'object containing MaxFrames.
'In case you have more, you should let ArrangeMaxFrames()
'init the tables every time it starts, or make multiple
'tables, one for each frame.
'Another option is to make a class module, which contains
'the InitFrames(), ArrangeMaxFrames(), InitFrameTable(),
'GetFramOrderIndex() and code to handle scrolling.
'This class should have an internal FrameIndexes and
'FrameTops table, and a "Target" property to set/return
'the Frame to handle. Then all you need to do is to create
'an instance of this class for every Frame containing
'MaxFrames, set the Target property (you should call
'InitMaxFrameTable() here), then call it's Arrange
'function every time it's needed.
'This way handling these frames will be __very__ easy,
'but a little bit slower due to class accessing (i think).
Private Sub InitMaxFrameTable(TargetFrame As Frame)
ReDim FrameIndexes(0)
ReDim FrameTops(0)
ScrollIndex = -1

'Determine the controls to handle...
'This iterates through the controls of Form1.
For X = 0 To Me.Controls.Count - 1
   If (TypeOf Me.Controls(X) Is MaxFrame) And (Me.Controls(X).Container Is TargetFrame) Then
   'This is a MaxFrame in TargetFrame.
      FrameIndexes(UBound(FrameIndexes)) = X
      FrameTops(UBound(FrameTops)) = Me.Controls(X).Top
      ReDim Preserve FrameIndexes(UBound(FrameIndexes) + 1)
      ReDim Preserve FrameTops(UBound(FrameTops) + 1)
   End If
   If ScrollIndex = -1 Then
   'We haven't picked a PanelVScroll control yet.
      If (TypeOf Me.Controls(X) Is PanelVScroll) And (Me.Controls(X).Container Is TargetFrame) Then
      'This is a PanelVScroll control in TargetFrame.
         ScrollIndex = X
      End If
   End If
Next 'X
ReDim Preserve FrameIndexes(UBound(FrameIndexes) - 1)
ReDim Preserve FrameTops(UBound(FrameTops) - 1)

'Sorting.
'This is a simple and slooow Bubble sort routine.
'You can change this if speed is the essence, but
'remember, you don't only need to sort the FrameTops array,
'the FramIndexes array must be sorted according to the
'order of the FrameTops array.
'I hope it's clear...
For X = 0 To UBound(FrameTops)
   For Y = UBound(FrameTops) To X + 1 Step -1
      If FrameTops(Y) < FrameTops(X) Then
         Temp = FrameTops(X)
         FrameTops(X) = FrameTops(Y)
         FrameTops(Y) = Temp
         
         Temp = FrameIndexes(X)
         FrameIndexes(X) = FrameIndexes(Y)
         FrameIndexes(Y) = Temp
      End If
   Next 'Y
Next 'X
End Sub

'The main function of the program.
'This is the way it places the MaxFrames:

'If you do not pass any FrameIndexes index("LastChanged")
'to it, the code starts placing MaxFrames from the top
'of TargetFrame.
'Things go complicated when you use the "LastChanged"
'argument. When this points to an opened MaxFrame,
'this means that the user have just opened it,
'so ArrangeMaxFrames() tries to place it to the
'top of TargetFrame. If it's not possible 'cause this
'would create an empty space below the MaxFrames,
'ArrangeMaxFrames() places this MaxFrame to the closest
'place possible.
'Second case: LastChanged points to a closed MaxFrame.
'If you close a MaxFrame, ArrangeMaxFrames() will place
'the MaxFrames below the LastChanged MaxFrame. But
'it tries not to create empty spaces, so if the above is
'not possible, it calculates the optimal location,
'and places the LastChanged MaxFrame there, and below this
'the other MaxFrames.
Private Sub ArrangeMaxFrames(TargetFrame As Frame, Optional LastChanged As Integer = -1)
InitMaxFrameTable TargetFrame

'Setup maximum length to be scrolled
MaximHeight = 0
For X = 0 To UBound(FrameIndexes)
   MaximHeight = MaximHeight + Me.Controls(FrameIndexes(X)).Height + FRAMESPACING * Screen.TwipsPerPixelY
Next 'X
MaximHeight = MaximHeight - TargetFrame.Height

'Make visible and setup scrollbar if needed
If MaximHeight > 0 Then
   Me.Controls(ScrollIndex).Max = MaximHeight / Screen.TwipsPerPixelY
   Me.Controls(ScrollIndex).LargeChange = Me.Controls(ScrollIndex).Max / UBound(FrameIndexes)
   LastVal = 0
   Me.Controls(ScrollIndex).Visible = True
Else
   Me.Controls(ScrollIndex).Visible = False
   Me.Controls(ScrollIndex).Value = 0
End If

If (LastChanged > -1) And (MaximHeight > 0) Then
   If Me.Controls(FrameIndexes(LastChanged)).State = 0 Then
      Temp = Me.Controls(FrameIndexes(LastChanged)).Top
      TempState = 0
   End If
End If

'Place frames in order
For X = 0 To UBound(FrameIndexes)
   If X = 0 Then
      Me.Controls(FrameIndexes(0)).Top = 0
   Else
      Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X - 1)).Top + Me.Controls(FrameIndexes(X - 1)).Height + FRAMESPACING * Screen.TwipsPerPixelY
   End If
Next 'X

If (LastChanged > -1) And (MaximHeight > 0) Then
   If Me.Controls(FrameIndexes(LastChanged)).State = 1 Then
      Temp = Me.Controls(FrameIndexes(LastChanged)).Top
      TempState = 1
   End If
End If

If (LastChanged > -1) And (MaximHeight > 0) Then
   TempHeight = 0
   For X = LastChanged To UBound(FrameIndexes)
      TempHeight = TempHeight + Me.Controls(FrameIndexes(X)).Height + FRAMESPACING * Screen.TwipsPerPixelY
   Next 'X
      
   If TempState = 1 Then
   'An opened MaxFrame
      If TempHeight < TargetFrame.Height Then
         Temp = Temp - (TargetFrame.Height - TempHeight)
      End If
      
      For X = 0 To UBound(FrameIndexes)
         Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X)).Top - Temp
      Next 'X
   Else
   'A closed MaxFrame
      Temp = Me.Controls(FrameIndexes(LastChanged)).Top - Temp
      
      If Me.Controls(FrameIndexes(LastChanged)).Top - Temp + TempHeight < TargetFrame.Height Then
         Temp = Temp - TargetFrame.Height + Me.Controls(FrameIndexes(LastChanged)).Top - Temp + TempHeight
      End If
      
      For X = 0 To UBound(FrameIndexes)
         Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X)).Top - Temp
      Next 'X
   End If
   
   'Finally, set the scrollbar's value to reflect the
   'current position of the MaxFrames.
   Me.Controls(ScrollIndex).Value = Temp / Screen.TwipsPerPixelY
   LastVal = Me.Controls(ScrollIndex).Value
End If

End Sub

Private Sub Command1_Click()
MsgBox "See? Like a frame."
End Sub

Private Sub Command2_Click()
MaxFrame1.OpenFrame
MaxFrame2.OpenFrame
MaxFrame3.OpenFrame
MaxFrame4.OpenFrame
MaxFrame5.OpenFrame
ArrangeMaxFrames ShipControlsFrame
End Sub

Private Sub Command3_Click()
MaxFrame1.CloseFrame
MaxFrame2.CloseFrame
MaxFrame3.CloseFrame
MaxFrame4.CloseFrame
MaxFrame5.CloseFrame
ArrangeMaxFrames ShipControlsFrame
End Sub

'This is just a simple routine to setup the scrollbar of
'TargetFrame. It searches for the widest MaxFrame in
'TargetFrame, and places the scrollbar accordingly.
'It also sets the spacing for the MaxFrames.
Private Sub InitFrame(TargetFrame As Frame)
Dim tMaxFrameIndex As Integer, tMaxFrameMaxWidth As Long, tVScrollIndex As Integer
tVScrollIndex = -1
For X = 0 To Me.Controls.Count - 1
   If (TypeOf Me.Controls(X) Is MaxFrame) And (Me.Controls(X).Container Is TargetFrame) Then
      If tMaxFrameMaxWidth < Me.Controls(X).Width Then
         tMaxFrameIndex = X
         tMaxFrameMaxWidth = Me.Controls(X).Width
      End If
   End If
   If tVScrollIndex = -1 Then
      If (TypeOf Me.Controls(X) Is PanelVScroll) And (Me.Controls(X).Container Is TargetFrame) Then
         tVScrollIndex = X
      End If
   End If
Next 'X

Me.Controls(tVScrollIndex).Width = TargetFrame.Width - tMaxFrameMaxWidth - SCROLLBARSPACING * Screen.TwipsPerPixelX
Me.Controls(tVScrollIndex).Left = tMaxFrameMaxWidth + SCROLLBARSPACING * Screen.TwipsPerPixelX
ArrangeMaxFrames ShipControlsFrame
End Sub

Private Sub Form_Load()
InitFrame ShipControlsFrame
End Sub

Private Sub FramScroll_ValueChange()
CurrVal = FramScroll.Value
For X = 0 To UBound(FrameIndexes)
   Me.Controls(FrameIndexes(X)).Top = Me.Controls(FrameIndexes(X)).Top - (CurrVal - LastVal) * Screen.TwipsPerPixelY
Next 'X
LastVal = CurrVal
End Sub

Private Sub MaxFrame1_StateChanged()
ArrangeMaxFrames ShipControlsFrame, GetFramOrderIndex("maxframe1")
End Sub

Private Sub MaxFrame2_StateChanged()
ArrangeMaxFrames ShipControlsFrame, GetFramOrderIndex("maxframe2")
End Sub

Private Sub MaxFrame3_StateChanged()
ArrangeMaxFrames ShipControlsFrame, GetFramOrderIndex("maxframe3")
End Sub

Private Sub MaxFrame4_StateChanged()
ArrangeMaxFrames ShipControlsFrame, GetFramOrderIndex("maxframe4")
End Sub

Private Sub MaxFrame5_StateChanged()
ArrangeMaxFrames ShipControlsFrame, GetFramOrderIndex("maxframe5")
End Sub

'This runs trough the FrameIndexes array, and searches
'for a MaxFrame with the given name(and optionally with
'the given index), then returns it's index in the
'FrameIndexes array.
Private Function GetFramOrderIndex(ByVal FramName As String, Optional FramIndex As Integer = -1) As Integer
FramName = LCase(FramName)
If FramIndex = -1 Then
   For X = 0 To UBound(FrameIndexes)
      If LCase(Me.Controls(FrameIndexes(X)).Name) = FramName Then
         GetFramOrderIndex = X
         Exit Function
      End If
   Next 'X
   GetFramOrderIndex = -1
Else
   For X = 0 To UBound(FrameIndexes)
      If (LCase(Me.Controls(FrameIndexes(X)).Name) = FramName) And (Me.Controls(FrameIndexes(X)).Index = FramIndex) Then
         GetFramOrderIndex = X
         Exit Function
      End If
   Next 'X
   GetFramOrderIndex = -1
End If
End Function
