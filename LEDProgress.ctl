VERSION 5.00
Begin VB.UserControl LEDProgress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ToolboxBitmap   =   "LEDProgress.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2895
   End
End
Attribute VB_Name = "LEDProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_Horizontal = 0
Const m_def_Yellow = 70
Const m_def_Red = 90
Const m_def_Value = 50
'Const m_def_Value = 50
'Property Variables:
Dim m_Max As Integer
Dim m_Min As Integer
Dim m_Horizontal As Boolean
Dim m_Yellow As Integer
Dim m_Red As Integer
Dim m_Value As Integer
Public Enum BStyle
    i3D = 1
    iFlat = 0
End Enum
'Dim m_Value As Integer
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get Value() As Integer
'    Value = m_Value
'End Property
'
'Public Property Let Value(ByVal New_Value As Integer)
'    m_Value = New_Value
'    PropertyChanged "Value"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Value = m_def_Value
    m_Value = m_def_Value
    m_Yellow = m_def_Yellow
    m_Red = m_def_Red
    m_Horizontal = m_def_Horizontal
    m_Max = m_def_Max
    m_Min = m_def_Min
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Yellow = PropBag.ReadProperty("Yellow", m_def_Yellow)
    m_Red = PropBag.ReadProperty("Red", m_def_Red)
    m_Horizontal = PropBag.ReadProperty("Horizontal", m_def_Horizontal)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
End Sub

Private Sub UserControl_Resize()
    SetBar
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Yellow", m_Yellow, m_def_Yellow)
    Call PropBag.WriteProperty("Red", m_Red, m_def_Red)
    Call PropBag.WriteProperty("Horizontal", m_Horizontal, m_def_Horizontal)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As BStyle
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As BStyle)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
    SetBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,50
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    PropertyChanged "Value"
    SetBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,70
Public Property Get Yellow() As Integer
    Yellow = m_Yellow
End Property

Public Property Let Yellow(ByVal New_Yellow As Integer)
    m_Yellow = New_Yellow
    PropertyChanged "Yellow"
    SetBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,90
Public Property Get Red() As Integer
    Red = m_Red
End Property

Public Property Let Red(ByVal New_Red As Integer)
    m_Red = New_Red
    PropertyChanged "Red"
    SetBar
End Property

Private Sub SetBar()
    Dim i As Integer
    Dim intValue As Integer
    Dim intYellow As Integer
    Dim intGreen As Integer
    Dim intRed As Integer
    Dim intMin As Integer
    DrawWidth = 2
    If m_Max = 0 Then
        m_Max = 100
    End If
    intValue = Round(m_Value / (m_Max - m_Min) * 100)
    intYellow = Round(m_Yellow / (m_Max - m_Min) * 100)
    intRed = Round(m_Red / (m_Max - m_Min) * 100)
    intGreen = Round(m_Green / (m_Max - m_Min) * 100)
    If m_Horizontal Then
        For i = 0 To ScaleWidth Step 2
            If i <= Round(intValue / 100 * ScaleWidth) Then
                If i < Round(intYellow / 100 * ScaleWidth) Then
                    Line (i, ScaleTop)-(i, ScaleTop + ScaleWidth), RGB(0, 255, 0)
                    Line (i + 1, ScaleTop)-(i + 1, ScaleTop + ScaleWidth), RGB(50, 50, 50)
                End If
                If i >= Round(intYellow / 100 * ScaleWidth) And i < Round(intRed / 100 * ScaleWidth) Then
                    Line (i, ScaleTop)-(i, ScaleTop + ScaleWidth), RGB(255, 255, 0)
                    Line (i + 1, ScaleTop)-(i + 1, ScaleTop + ScaleWidth), RGB(50, 50, 50)
                End If
                If i >= Round(intRed / 100 * ScaleWidth) Then
                    Line (i, ScaleTop)-(i, ScaleTop + ScaleWidth), RGB(255, 0, 0)
                    Line (i + 1, ScaleTop)-(i + 1, ScaleTop + ScaleWidth), RGB(50, 50, 50)
                End If
            Else
                Line (i, ScaleTop)-(i, ScaleTop + ScaleWidth), RGB(127, 127, 127)
                Line (i + 1, ScaleTop)-(i + 1, ScaleTop + ScaleWidth), RGB(50, 50, 50)
            End If
        Next i
    Else
        For i = 0 To ScaleHeight Step 2
            Line (ScaleLeft, i)-(ScaleLeft + ScaleWidth, i), RGB(127, 127, 127)
            Line (ScaleLeft, i + 1)-(ScaleLeft + ScaleWidth, i + 1), RGB(50, 50, 50)
            If i <= Round(intValue / 100 * ScaleHeight) Then
                If i < Round(intYellow / 100 * ScaleHeight) Then
                    Line (ScaleLeft, i)-(ScaleLeft + ScaleWidth, i), RGB(0, 255, 0)
                    Line (ScaleLeft, i + 1)-(ScaleLeft + ScaleWidth, i + 1), RGB(50, 50, 50)
                End If
                If i >= Round(intYellow / 100 * ScaleHeight) And i < Round(intRed / 100 * ScaleHeight) Then
                    Line (ScaleLeft, i)-(ScaleLeft + ScaleWidth, i), RGB(255, 255, 0)
                    Line (ScaleLeft, i + 1)-(ScaleLeft + ScaleWidth, i + 1), RGB(50, 50, 50)
                End If
                If i >= Round(intRed / 100 * ScaleHeight) Then
                    Line (ScaleLeft, i)-(ScaleLeft + ScaleWidth, i), RGB(255, 0, 0)
                    Line (ScaleLeft, i + 1)-(ScaleLeft + ScaleWidth, i + 1), RGB(50, 50, 50)
                End If
            Else
                Line (ScaleLeft, i)-(ScaleLeft + ScaleWidth, i), RGB(127, 127, 127)
                Line (ScaleLeft, i + 1)-(ScaleLeft + ScaleWidth, i + 1), RGB(50, 50, 50)
            End If
        Next i
    End If
    Label1(0).Caption = intValue & "%"
    Label1(1).Caption = intValue & "%"
    Label1(0).Left = ScaleLeft + 5
    Label1(0).Width = ScaleWidth
    Label1(1).Left = ScaleLeft
    Label1(1).Width = ScaleWidth
    Label1(0).Top = Round((ScaleHeight / 2) - (Label1(0).Height / 2)) + 4
    Label1(1).Top = Round((ScaleHeight / 2) - (Label1(1).Height / 2))
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Horizontal() As Boolean
    Horizontal = m_Horizontal
End Property

Public Property Let Horizontal(ByVal New_Horizontal As Boolean)
    m_Horizontal = New_Horizontal
    PropertyChanged "Horizontal"
    SetBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

