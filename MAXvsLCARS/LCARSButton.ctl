VERSION 5.00
Begin VB.UserControl LCARSButton 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   KeyPreview      =   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   3885
   ToolboxBitmap   =   "LCARSButton.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Image ButtonIcon 
      Height          =   375
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Line RCapLine 
      BorderColor     =   &H8000000F&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   480
      Y2              =   0
   End
   Begin VB.Line LCapLine 
      BorderColor     =   &H8000000F&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   0
   End
   Begin VB.Label CapLab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LCARSButton"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin VB.Shape REnd 
      BorderColor     =   &H000099FF&
      FillColor       =   &H000099FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape LEnd 
      BorderColor     =   &H000099FF&
      FillColor       =   &H000099FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape LabelBack 
      BorderColor     =   &H000099FF&
      FillColor       =   &H000099FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "LCARSButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum ButtonStyleEnum
   LBNoRound = 0
   LBLeftRound = 1
   LBRightRound = 2
   LBBothRound = 3
End Enum

Public Enum CapLineStyleEnum
   LBNoCapline = 0
   LBLeftCapline = 1
   LBRightCapline = 2
   LBBothCapline = 3
End Enum

Public Enum TextHAlignEnum
   LBLeftAlign = 0
   LBCenterAlign = 1
   LBRightAlign = 2
End Enum

Public Enum TextVAlignEnum
   LBTopAlign = 0
   LBCenterAlign = 1
   LBBottomAlign = 2
End Enum

'Default Property Values:
Const m_def_PressedColor = &H33CCFF
Const m_def_Style = 3
Const m_def_CapLines = 0
Const m_def_Alignment = 0
Const m_def_VAlignment = 1
'Property Variables:
Private m_PressedColor As Long
Private m_ButtonColor As Long
Private m_Style As ButtonStyleEnum
Private m_CapLines As CapLineStyleEnum
Private m_Alignment As TextHAlignEnum
Private m_VAlignment As TextVAlignEnum
Private m_MaxSpace As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor = New_BackColor
   RCapLine.BorderColor = New_BackColor
   LCapLine.BorderColor = New_BackColor
   
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapLab,CapLab,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
   ForeColor = CapLab.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   CapLab.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get MaxSpace() As Boolean
   MaxSpace = m_MaxSpace
End Property

Public Property Let MaxSpace(ByVal New_MaxSpace As Boolean)
   m_MaxSpace = New_MaxSpace
   PosCaption
   PosIcon
   PropertyChanged "MaxSpace"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapLab,CapLab,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = CapLab.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set CapLab.Font = New_Font
   PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
   UserControl_Resize
   UserControl.Refresh
End Sub

Private Sub CapLab_Click()
   RaiseEvent Click
End Sub

Private Sub CapLab_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub CapLab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X + CapLab.Left, Y + CapLab.Top)
LabelBack.FillColor = m_PressedColor
LEnd.FillColor = m_PressedColor
REnd.FillColor = m_PressedColor

LabelBack.BorderColor = m_PressedColor
LEnd.BorderColor = m_PressedColor
REnd.BorderColor = m_PressedColor

End Sub

Private Sub CapLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X + CapLab.Left, Y + CapLab.Top)
End Sub

Private Sub CapLab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X + CapLab.Left, Y + CapLab.Top)
LabelBack.FillColor = m_ButtonColor
LEnd.FillColor = m_ButtonColor
REnd.FillColor = m_ButtonColor

LabelBack.BorderColor = m_ButtonColor
LEnd.BorderColor = m_ButtonColor
REnd.BorderColor = m_ButtonColor

End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
LabelBack.FillColor = m_PressedColor
LEnd.FillColor = m_PressedColor
REnd.FillColor = m_PressedColor

LabelBack.BorderColor = m_PressedColor
LEnd.BorderColor = m_PressedColor
REnd.BorderColor = m_PressedColor

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
LabelBack.FillColor = m_ButtonColor
LEnd.FillColor = m_ButtonColor
REnd.FillColor = m_ButtonColor

LabelBack.BorderColor = m_ButtonColor
LEnd.BorderColor = m_ButtonColor
REnd.BorderColor = m_ButtonColor

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PressedColor() As OLE_COLOR
   PressedColor = m_PressedColor
End Property

Public Property Let PressedColor(ByVal New_PressedColor As OLE_COLOR)
   m_PressedColor = New_PressedColor
   PropertyChanged "PressedColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapLab,CapLab,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
   Caption = CapLab.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   CapLab.Caption() = New_Caption
   PosCaption
   PosIcon
   PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapLab,CapLab,-1,Alignment
Public Property Get Alignment() As TextHAlignEnum
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
   Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As TextHAlignEnum)
   m_Alignment = New_Alignment
   PosCaption
   PosIcon
   PropertyChanged "Alignment"
End Property

Public Property Get VAlignment() As TextVAlignEnum
   VAlignment = m_VAlignment
End Property

Public Property Let VAlignment(ByVal New_VAlignment As TextVAlignEnum)
   m_VAlignment = New_VAlignment
   PosCaption
   PosIcon
   PropertyChanged "VAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Style() As ButtonStyleEnum
   Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ButtonStyleEnum)
m_Style = New_Style
UpdateStyle
PropertyChanged "Style"
End Property

Private Sub UpdateStyle()
Select Case m_Style
Case 0
   LEnd.Visible = False
   REnd.Visible = False
Case 1
   LEnd.Visible = True
   REnd.Visible = False
Case 2
   LEnd.Visible = False
   REnd.Visible = True
Case 3
   LEnd.Visible = True
   REnd.Visible = True
End Select

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
   Set Picture = ButtonIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set ButtonIcon.Picture = New_Picture
   PosIcon
   PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LEnd,LEnd,-1,FillColor
Public Property Get ButtonColor() As OLE_COLOR
Attribute ButtonColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
   ButtonColor = m_ButtonColor
End Property

Public Property Let ButtonColor(ByVal New_ButtonColor As OLE_COLOR)
   LEnd.FillColor = New_ButtonColor
   REnd.FillColor = New_ButtonColor
   LabelBack.FillColor = New_ButtonColor
   LEnd.BorderColor = New_ButtonColor
   REnd.BorderColor = New_ButtonColor
   LabelBack.BorderColor = New_ButtonColor
   
   m_ButtonColor = New_ButtonColor
   PropertyChanged "ButtonColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get CapLines() As CapLineStyleEnum
   CapLines = m_CapLines
End Property

Public Property Let CapLines(ByVal New_CapLines As CapLineStyleEnum)
m_CapLines = New_CapLines
UpdateCaplines

PosCaption
PosIcon

PropertyChanged "CapLines"
End Property

Private Sub UpdateCaplines()
Select Case m_CapLines
Case 0
   LCapLine.Visible = False
   RCapLine.Visible = False
Case 1
   LCapLine.Visible = True
   RCapLine.Visible = False
Case 2
   LCapLine.Visible = False
   RCapLine.Visible = True
Case 3
   LCapLine.Visible = True
   RCapLine.Visible = True
End Select

End Sub

Private Sub UserControl_Initialize()
ButtonIcon.Width = 16 * Screen.TwipsPerPixelX
ButtonIcon.Height = 16 * Screen.TwipsPerPixelY

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_PressedColor = m_def_PressedColor
   m_Style = m_def_Style
   Set m_Image = LoadPicture("")
   m_CapLines = m_def_CapLines
   m_Alignment = m_def_Alignment
   m_VAlignment = m_def_VAlignment
   m_ButtonColor = &H99FF&
  
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   CapLab.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_MaxSpace = PropBag.ReadProperty("MaxSpace", False)
   Set CapLab.Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_PressedColor = PropBag.ReadProperty("PressedColor", m_def_PressedColor)
   CapLab.Caption = PropBag.ReadProperty("Caption", "LCARSButton")
   m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
   m_VAlignment = PropBag.ReadProperty("VAlignment", m_def_VAlignment)
   m_Style = PropBag.ReadProperty("Style", m_def_Style)
   Set ButtonIcon.Picture = PropBag.ReadProperty("Picture", Nothing)
   m_ButtonColor = PropBag.ReadProperty("ButtonColor", &H99FF&)
   m_CapLines = PropBag.ReadProperty("CapLines", m_def_CapLines)
   
   LEnd.FillColor = m_ButtonColor
   REnd.FillColor = m_ButtonColor
   LabelBack.FillColor = m_ButtonColor
   LEnd.BorderColor = m_ButtonColor
   REnd.BorderColor = m_ButtonColor
   LabelBack.BorderColor = m_ButtonColor

   UserControl.BackColor = Me.BackColor
   RCapLine.BorderColor = Me.BackColor
   LCapLine.BorderColor = Me.BackColor
   
   UpdateStyle
   UpdateCaplines
   PosCaption
   PosIcon
   
End Sub

Private Sub UserControl_Resize()
If UserControl.Height > UserControl.Width Then
   Exit Sub
End If
LEnd.Height = UserControl.Height
LEnd.Width = LEnd.Height
REnd.Height = LEnd.Height
REnd.Width = LEnd.Height

REnd.Left = UserControl.Width - REnd.Width

LabelBack.Left = LEnd.Width / 2
LabelBack.Height = LEnd.Height
LabelBack.Width = UserControl.Width - REnd.Width

LCapLine.X1 = LabelBack.Left - 2 * Screen.TwipsPerPixelX
LCapLine.X2 = LabelBack.Left - 2 * Screen.TwipsPerPixelX
LCapLine.Y1 = 0
LCapLine.Y2 = UserControl.Height

RCapLine.X1 = LabelBack.Left + LabelBack.Width - 2 * Screen.TwipsPerPixelX
RCapLine.X2 = LabelBack.Left + LabelBack.Width - 2 * Screen.TwipsPerPixelX
RCapLine.Y1 = 0
RCapLine.Y2 = LCapLine.Y2

PosCaption
PosIcon

End Sub

'A poscaption után kell meghívni!
Private Sub PosIcon()
If ButtonIcon.Picture = 0 Then Exit Sub
ButtonIcon.Top = CapLab.Top + (CapLab.Height - ButtonIcon.Height) / 2

On Error GoTo EEE
Select Case m_Alignment
Case 0
   ButtonIcon.Left = LabelBack.Left
   If ((m_CapLines = 0) Or (m_CapLines = 2)) And (MaxSpace) Then ButtonIcon.Left = ButtonIcon.Left - Sqr((LEnd.Height / 2) ^ 2 - ((UserControl.Height - ButtonIcon.Top * 2 + Screen.TwipsPerPixelY * 2) / 2) ^ 2)
   CapLab.Left = ButtonIcon.Left + ButtonIcon.Width + Screen.TwipsPerPixelX * 3
Case 1
   ButtonIcon.Left = LabelBack.Left
Case 2
   ButtonIcon.Left = LabelBack.Left
End Select
EEE:

End Sub

Private Sub PosCaption()
Select Case m_Alignment
Case 0
   CapLab.Left = LabelBack.Left
Case 1
   CapLab.Left = LabelBack.Left + (LabelBack.Width - CapLab.Width) / 2
Case 2
   CapLab.Left = LabelBack.Left + LabelBack.Width - CapLab.Width
End Select


On Error GoTo EEE
Select Case m_VAlignment
Case 0
   CapLab.Top = Screen.TwipsPerPixelY * 3
   If (m_Alignment = 0) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 2)) Then CapLab.Left = CapLab.Left - Sqr((LEnd.Height / 2) ^ 2 - ((UserControl.Height - Screen.TwipsPerPixelY * 6 + Screen.TwipsPerPixelY * 2) / 2) ^ 2)
   If (m_Alignment = 2) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 1)) Then CapLab.Left = CapLab.Left + Sqr((LEnd.Height / 2) ^ 2 - ((UserControl.Height - Screen.TwipsPerPixelY * 6 + Screen.TwipsPerPixelY * 2) / 2) ^ 2)
Case 1
   CapLab.Top = (UserControl.Height - CapLab.Height) / 2
   If (m_Alignment = 0) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 2)) Then CapLab.Left = CapLab.Left - Sqr((LEnd.Height / 2) ^ 2 - ((CapLab.Height + Screen.TwipsPerPixelY * 15) / 2) ^ 2)
   If (m_Alignment = 2) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 1)) Then CapLab.Left = CapLab.Left + Sqr((LEnd.Height / 2) ^ 2 - ((CapLab.Height + Screen.TwipsPerPixelY * 15) / 2) ^ 2)
Case 2
   CapLab.Top = UserControl.Height - CapLab.Height - Screen.TwipsPerPixelY * 3
   If (m_Alignment = 0) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 2)) Then CapLab.Left = CapLab.Left - Sqr((LEnd.Height / 2) ^ 2 - ((UserControl.Height - Screen.TwipsPerPixelY * 6 + Screen.TwipsPerPixelY * 2) / 2) ^ 2)
   If (m_Alignment = 2) And (m_MaxSpace) And ((m_CapLines = 0) Or (m_CapLines = 1)) Then CapLab.Left = CapLab.Left + Sqr((LEnd.Height / 2) ^ 2 - ((UserControl.Height - Screen.TwipsPerPixelY * 6 + Screen.TwipsPerPixelY * 2) / 2) ^ 2)
End Select
EEE:

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
   Call PropBag.WriteProperty("ForeColor", CapLab.ForeColor, &HFFFFFF)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("MaxSpace", m_MaxSpace, False)
   Call PropBag.WriteProperty("Font", CapLab.Font, Ambient.Font)
   Call PropBag.WriteProperty("PressedColor", m_PressedColor, m_def_PressedColor)
   Call PropBag.WriteProperty("Caption", CapLab.Caption, "LCARSButton")
   Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
   Call PropBag.WriteProperty("VAlignment", m_VAlignment, m_def_VAlignment)
   Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
   Call PropBag.WriteProperty("Picture", ButtonIcon.Picture, Nothing)
   Call PropBag.WriteProperty("ButtonColor", m_ButtonColor, &H99FF&)
   Call PropBag.WriteProperty("CapLines", m_CapLines, m_def_CapLines)
End Sub

