VERSION 5.00
Begin VB.UserControl MaxFrame 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ControlContainer=   -1  'True
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   ToolboxBitmap   =   "MaxCP.ctx":0000
   Begin MaxVsLCARSControls.LCARSButton LCARSButton1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LCARSMaxFrame"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000099FF&
      BorderWidth     =   2
      Height          =   1935
      Left            =   15
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "MaxFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum StateEnum
   MaxClosed = 0
   MaxOpen = 1
End Enum
'Default Property Values:
Const m_def_State = 1
Const m_def_SaveHeight = 1
'Property Variables:
Dim m_State As StateEnum
Dim SaveHeight As Long
'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
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
Event StateChanged()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   LCARSButton1.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,PressedColor
Public Property Get PressedColor() As OLE_COLOR
   PressedColor = LCARSButton1.PressedColor
End Property

Public Property Let PressedColor(ByVal New_PressedColor As OLE_COLOR)
   LCARSButton1.PressedColor() = New_PressedColor
   PropertyChanged "PressedColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
   ForeColor = LCARSButton1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   LCARSButton1.ForeColor() = New_ForeColor
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = LCARSButton1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set LCARSButton1.Font = New_Font
   PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
   UserControl.Refresh
End Sub

Private Sub LCARSButton1_Click()
If m_State = 1 Then
   SaveHeight = UserControl.Height
   UserControl.Height = 23 * Screen.TwipsPerPixelY '19+2+2
   m_State = 0
Else
   UserControl.Height = SaveHeight
   m_State = 1
End If

RaiseEvent StateChanged
PropertyChanged "State"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
   BorderColor = Shape1.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   Shape1.BorderColor() = New_BorderColor
   LCARSButton1.ButtonColor = New_BorderColor
   PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,Picture
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
   Set Image = LCARSButton1.Picture
End Property

Public Property Set Image(ByVal New_Image As Picture)
   Set LCARSButton1.Picture = New_Image
   PropertyChanged "Image"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
   Caption = LCARSButton1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   LCARSButton1.Caption() = New_Caption
   PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,VAlignment
Public Property Get VAlignment() As TextVAlignEnum
   VAlignment = LCARSButton1.VAlignment
End Property

Public Property Let VAlignment(ByVal New_VAlignment As TextVAlignEnum)
   LCARSButton1.VAlignment() = New_VAlignment
   PropertyChanged "VAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,MaxSpace
Public Property Get MaxSpace() As Boolean
   MaxSpace = LCARSButton1.MaxSpace
End Property

Public Property Let MaxSpace(ByVal New_MaxSpace As Boolean)
   LCARSButton1.MaxSpace() = New_MaxSpace
   PropertyChanged "MaxSpace"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,Alignment
Public Property Get Alignment() As TextHAlignEnum
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
   Alignment = LCARSButton1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As TextHAlignEnum)
   LCARSButton1.Alignment() = New_Alignment
   PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,CapLines
Public Property Get CapLines() As CapLineStyleEnum
   CapLines = LCARSButton1.CapLines
End Property

Public Property Let CapLines(ByVal New_CapLines As CapLineStyleEnum)
   LCARSButton1.CapLines() = New_CapLines
   PropertyChanged "CapLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LCARSButton1,LCARSButton1,-1,Style
Public Property Get Style() As ButtonStyleEnum
   Style = LCARSButton1.Style
End Property

Public Property Let Style(ByVal New_Style As ButtonStyleEnum)
   LCARSButton1.Style() = New_Style
   PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub CloseFrame()
If m_State = 1 Then
   SaveHeight = UserControl.Height
   UserControl.Height = 23 * Screen.TwipsPerPixelY '19+2+2
   m_State = 0
   RaiseEvent StateChanged
   PropertyChanged "State"
End If
End Sub

Public Sub ToggleFrame()
If m_State = 1 Then
   SaveHeight = UserControl.Height
   UserControl.Height = 23 * Screen.TwipsPerPixelY '19+2+2
   m_State = 0
   RaiseEvent StateChanged
   PropertyChanged "State"
Else
   UserControl.Height = SaveHeight
   m_State = 1
   RaiseEvent StateChanged
   PropertyChanged "State"
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub OpenFrame()
If m_State = 0 Then
   UserControl.Height = SaveHeight
   m_State = 1
   RaiseEvent StateChanged
   PropertyChanged "State"
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,1
Public Property Get State() As StateEnum
   State = m_State
End Property

Public Property Let State(New_State As StateEnum)

If New_State = 0 Then
   SaveHeight = UserControl.Height
   UserControl.Height = 23 * Screen.TwipsPerPixelY '19+2+2
Else
   If SaveHeight = 0 Then Exit Property
   UserControl.Height = SaveHeight
End If

m_State = New_State
RaiseEvent StateChanged
PropertyChanged "State"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_State = m_def_State
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   UserControl.BackColor = PropBag.ReadProperty("BackColor", 0)
   LCARSButton1.BackColor = UserControl.BackColor
   LCARSButton1.ForeColor = PropBag.ReadProperty("ForeColor", &H0&)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   Set LCARSButton1.Font = PropBag.ReadProperty("Font", Ambient.Font)
   Shape1.BorderColor = PropBag.ReadProperty("BorderColor", &H80FF&)
   LCARSButton1.ButtonColor = Shape1.BorderColor
   Set Picture = PropBag.ReadProperty("Image", Nothing)
   LCARSButton1.Caption = PropBag.ReadProperty("Caption", "LCARSMaxFrame")
   LCARSButton1.PressedColor = PropBag.ReadProperty("PressedColor", &H33CCFF)
   LCARSButton1.VAlignment = PropBag.ReadProperty("VAlignment", 1)
   LCARSButton1.MaxSpace = PropBag.ReadProperty("MaxSpace", False)
   LCARSButton1.Alignment = PropBag.ReadProperty("Alignment", 0)
   LCARSButton1.CapLines = PropBag.ReadProperty("CapLines", 0)
   LCARSButton1.Style = PropBag.ReadProperty("Style", 3)
   m_State = PropBag.ReadProperty("State", m_def_State)
   SaveHeight = PropBag.ReadProperty("SaveHeight", m_def_SaveHeight)
End Sub

Private Sub UserControl_Resize()
If UserControl.Height <> 23 * Screen.TwipsPerPixelY Then SaveHeight = UserControl.Height
Shape1.Width = UserControl.Width / Screen.TwipsPerPixelX - 1
Shape1.Height = UserControl.Height / Screen.TwipsPerPixelY - 8
LCARSButton1.Width = UserControl.Width / Screen.TwipsPerPixelX - 16
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, 0)
   Call PropBag.WriteProperty("ForeColor", LCARSButton1.ForeColor, &H0&)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Font", LCARSButton1.Font, Ambient.Font)
   Call PropBag.WriteProperty("BorderColor", Shape1.BorderColor, &H80FF&)
   Call PropBag.WriteProperty("Image", Picture, Nothing)
   Call PropBag.WriteProperty("Caption", LCARSButton1.Caption, "LCARSMaxFrame")
   Call PropBag.WriteProperty("PressedColor", LCARSButton1.PressedColor, &H33CCFF)
   Call PropBag.WriteProperty("VAlignment", LCARSButton1.VAlignment, 1)
   Call PropBag.WriteProperty("MaxSpace", LCARSButton1.MaxSpace, False)
   Call PropBag.WriteProperty("Alignment", LCARSButton1.Alignment, 0)
   Call PropBag.WriteProperty("CapLines", LCARSButton1.CapLines, 0)
   Call PropBag.WriteProperty("Style", LCARSButton1.Style, 3)
   Call PropBag.WriteProperty("State", m_State, m_def_State)
   Call PropBag.WriteProperty("SaveHeight", SaveHeight, m_def_SaveHeight)
End Sub

