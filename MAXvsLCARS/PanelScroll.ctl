VERSION 5.00
Begin VB.UserControl PanelVScroll 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   135
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   9
   ToolboxBitmap   =   "PanelScroll.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Label Dragger 
      BackColor       =   &H00808080&
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "PanelVScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Max = 100
Const m_def_Min = -100
Const m_def_LargeChange = 5
Const m_def_Value = 0
'Hungarian comment, the translation of the important
'part can be found below.

'--------------------------------------------------
'A használható terület, amibõl a value származik,
'a usercontrol tetejétõl tart a (usercontrol.height-dragger.height)-ig.

'A windowsos scrollbar úgy mûxik, hogy a rendelkezésre álló
'területet felosztja (max-min) db. részre, a csúszka ezen
'részek között ugrál. Ha egy rész hossza túl kicsi lenne((usercontrol.height-dragger.height)/(max-min)),
'akkor a csúszka mérete egy beállított minimum lesz.

'A PanelVScroll mûködése: mivel csak ablak-scrollozásra
'van használva, ezért megpróbálja megtartani az 1:1 pixel arányt,
'és ennek megfelõen méretezi a csúszkát. Ennek is van egy minimumja,
'ez mondjuk a scrollbar magasságának 0,2-szerese.
'Ezt elérve a scrollbar csak a Pixel-Elmozdulás:Value hányadost változtatja,
'nem csökkenti tovább a csúszka magasságát.
'--------------------------------------------------

'The way PanelVScroll works: because it's only used to
'scroll controls in a window, it tries to stick to
'1:1 pixel ratio and sizes the slider accordingly.
'The slider has a minimum height (actually, its
'0.2*PanelVScroll.Height). If the slider reaches it's
'minimum, the scrollbar only changes the
'Movement-in-Pixels:Value ratio, rather than reducing
'the slider's height more.

'Property Variables:
Dim m_Max As Long
Dim m_Min As Long
Dim m_LargeChange As Long
Dim m_Value As Long

Dim Ratio As Double
Dim oldY As Long
'Event Declarations:
Event ValueChange()
Event ValueChangeByCode()
Event ValueChangeByMaxMin()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Dragger,Dragger,-1,BackColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   ForeColor = Dragger.BackColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   Dragger.BackColor() = New_ForeColor
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
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
   Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
   m_Max = New_Max
   ResizeDragger
   PropertyChanged "Max"
   RaiseEvent ValueChangeByMaxMin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
   Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
   m_Min = New_Min
   ResizeDragger
   PropertyChanged "Min"
   RaiseEvent ValueChangeByMaxMin
End Property

Public Property Get LargeChange() As Long
   LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
   m_LargeChange = New_LargeChange
   PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
m_Value = New_Value
Dragger.Top = (m_Value - m_Min) / Ratio
PropertyChanged "Value"
RaiseEvent ValueChangeByCode
End Property

'Ratio=Control : Value
Private Sub ResizeDragger()
If (m_Max - m_Min) <= (UserControl.ScaleHeight * 0.8) Then 'Nincs ratio állítás
   Ratio = 1
   Dragger.Height = UserControl.ScaleHeight - (m_Max - m_Min)
Else
   Ratio = (m_Max - m_Min) / (UserControl.ScaleHeight * 0.8)
   Dragger.Height = UserControl.ScaleHeight * 0.2
End If

Dragger.Top = 0
m_Value = m_Min
End Sub

Private Sub Dragger_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
oldY = Y
End Sub

Private Sub Dragger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = 1) And (Dragger.Top + (Y - oldY) / Screen.TwipsPerPixelY >= 0) And (Dragger.Top + (Y - oldY) / Screen.TwipsPerPixelY <= UserControl.ScaleHeight - Dragger.Height) Then
    Dragger.Top = Dragger.Top + (Y - oldY) / Screen.TwipsPerPixelY
    m_Value = CLng(Dragger.Top * Ratio) + m_Min
    PropertyChanged "Value"
    RaiseEvent ValueChange
ElseIf (Button = 1) And (Dragger.Top + (Y - oldY) / Screen.TwipsPerPixelY < 0) Then
    Dragger.Top = 0
    m_Value = m_Min
    PropertyChanged "Value"
    RaiseEvent ValueChange
ElseIf (Button = 1) And (Dragger.Top + (Y - oldY) / Screen.TwipsPerPixelY > UserControl.ScaleHeight - Dragger.Height) Then
    Dragger.Top = UserControl.ScaleHeight - Dragger.Height
    m_Value = m_Max
    PropertyChanged "Value"
    RaiseEvent ValueChange
End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Max = m_def_Max
   m_Min = m_def_Min
   m_LargeChange = m_def_LargeChange
   m_Value = m_def_Value
   ResizeDragger
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= Dragger.Top Then
   If m_Value - m_LargeChange >= m_Min Then
      m_Value = m_Value - m_LargeChange
   Else
      m_Value = m_Min
   End If
Else
   If m_Value + m_LargeChange <= m_Max Then
      m_Value = m_Value + m_LargeChange
   Else
      m_Value = m_Max
   End If
End If
Dragger.Top = (m_Value - m_Min) / Ratio
PropertyChanged "Value"
RaiseEvent ValueChange
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   Dragger.BackColor = PropBag.ReadProperty("ForeColor", &H808080)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   m_Max = PropBag.ReadProperty("Max", m_def_Max)
   m_Min = PropBag.ReadProperty("Min", m_def_Min)
   m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
   m_Value = PropBag.ReadProperty("Value", m_def_Value)
   ResizeDragger
   
   Dragger.Top = (m_Value - m_Min) / Ratio

End Sub

Private Sub UserControl_Resize()
Dragger.Width = UserControl.Width
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
   Call PropBag.WriteProperty("ForeColor", Dragger.BackColor, &H808080)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
   Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
   Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
   Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

