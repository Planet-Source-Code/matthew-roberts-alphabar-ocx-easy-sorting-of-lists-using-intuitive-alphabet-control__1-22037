VERSION 5.00
Begin VB.UserControl ctlAlphaBar 
   Appearance      =   0  'Flat
   ClientHeight    =   7755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   7755
   ScaleWidth      =   660
   ToolboxBitmap   =   "ctlAlphaBar.ctx":0000
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   26
      Left            =   105
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   25
      Left            =   105
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7155
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   24
      Left            =   105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6870
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   23
      Left            =   105
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6585
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   22
      Left            =   105
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6300
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   21
      Left            =   105
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Click to filter by letter"
      Top             =   6015
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   20
      Left            =   105
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5730
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   19
      Left            =   105
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5445
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   18
      Left            =   105
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   17
      Left            =   105
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4875
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   16
      Left            =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4590
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   15
      Left            =   105
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4305
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   14
      Left            =   105
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4020
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   13
      Left            =   105
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3735
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   12
      Left            =   105
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3450
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   11
      Left            =   105
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3165
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   10
      Left            =   105
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   9
      Left            =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2595
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   8
      Left            =   105
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   7
      Left            =   105
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2025
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   6
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Filter by first letter"
      Top             =   1740
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   5
      Left            =   105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1455
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   4
      Left            =   105
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1170
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   3
      Left            =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   885
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   2
      Left            =   105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   1
      Left            =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Width           =   375
   End
   Begin VB.CommandButton cmdAlphaBar 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HelpContextID   =   5
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   375
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H80000006&
      Height          =   7740
      Left            =   75
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "ctlAlphaBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum alpOrient
    Horizontal = 1
    Vertical = 0
End Enum



'Default Property Values:
Const m_def_ControlID = "0"
Const m_def_Orientation = 1
Const m_def_MouseTracking = 0
Const m_def_AutoResize = 0
Const m_def_ShowToolTip = 0
'Property Variables:
Dim m_ControlID As String
Dim m_Orientation As alpOrient
Dim m_MouseTracking As Boolean
Dim m_AutoResize As Boolean
Dim m_ShowToolTip As Variant
'Event Declarations:
Event Click(Letter As String)
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event RightClick()
Attribute RightClick.VB_Description = "Occurs when the user clicks the right mouse button."



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
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Private Sub cmdAlphaBar_Click(Index As Integer)
    If Index = 26 Then
        RaiseEvent Click("*")
    Else
        RaiseEvent Click(Chr(Index + 65))
    End If
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

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,0
Public Property Get ControlID() As String
Attribute ControlID.VB_Description = "Unique Identifier for the control (Read Only)"
    ControlID = m_ControlID
End Property

Public Property Let ControlID(ByVal New_ControlID As String)
    If Ambient.UserMode Then Err.Raise 382
    m_ControlID = New_ControlID
    PropertyChanged "ControlID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Orientation() As alpOrient
Attribute Orientation.VB_Description = "Indicates whether the control should be drawn horizontally or vertically."
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As alpOrient)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    DrawAlphaBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MouseTracking() As Boolean
    MouseTracking = m_MouseTracking
End Property

Public Property Let MouseTracking(ByVal New_MouseTracking As Boolean)
    m_MouseTracking = New_MouseTracking
    PropertyChanged "MouseTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoResize() As Boolean
Attribute AutoResize.VB_Description = "Control will keep the same proportions relative to the parent object when the parent object is resized."
    AutoResize = m_AutoResize
End Property

Public Property Let AutoResize(ByVal New_AutoResize As Boolean)
    m_AutoResize = New_AutoResize
    PropertyChanged "AutoResize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ShowToolTip() As Variant
Attribute ShowToolTip.VB_Description = "Indicates whether the Tool Tip Text should be shown."
    ShowToolTip = m_ShowToolTip
End Property

Public Property Let ShowToolTip(ByVal New_ShowToolTip As Variant)
    m_ShowToolTip = New_ShowToolTip
    PropertyChanged "ShowToolTip"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_ControlID = m_def_ControlID
    m_Orientation = m_def_Orientation
    m_MouseTracking = m_def_MouseTracking
    m_AutoResize = m_def_AutoResize
    m_ShowToolTip = m_def_ShowToolTip
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_ControlID = PropBag.ReadProperty("ControlID", m_def_ControlID)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_MouseTracking = PropBag.ReadProperty("MouseTracking", m_def_MouseTracking)
    m_AutoResize = PropBag.ReadProperty("AutoResize", m_def_AutoResize)
    m_ShowToolTip = PropBag.ReadProperty("ShowToolTip", m_def_ShowToolTip)
End Sub

Private Sub UserControl_Resize()
    DrawAlphaBar
End Sub

Private Sub UserControl_Show()
    DrawAlphaBar
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ControlID", m_ControlID, m_def_ControlID)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("MouseTracking", m_MouseTracking, m_def_MouseTracking)
    Call PropBag.WriteProperty("AutoResize", m_AutoResize, m_def_AutoResize)
    Call PropBag.WriteProperty("ShowToolTip", m_ShowToolTip, m_def_ShowToolTip)
End Sub

Private Sub DrawAlphaBar()
   Dim intButtonSize As Integer
   
   If Orientation = Vertical Then
        
        With cmdAlphaBar(0)
            .Height = (Height - 150) / 27
            .Width = Width - 10
            .Left = 10
            .Top = 10
            
            If cmdAlphaBar(intButtonSize).Height / 35 > 5 Then
                .FontSize = Int(cmdAlphaBar(intButtonSize).Height / 35)
            End If
        End With
    
       For intButtonSize = 1 To 26
            
            With cmdAlphaBar(intButtonSize)
                .Height = (Height - 150) / 27
                .Width = Width - 10
                .Left = cmdAlphaBar(0).Left
                .Top = cmdAlphaBar(intButtonSize - 1).Top + cmdAlphaBar(intButtonSize - 1).Height + 5
               If cmdAlphaBar(intButtonSize).Height / 35 > 5 Then
                   .FontSize = Int(cmdAlphaBar(intButtonSize).Height / 35)
                End If
            End With
            
       Next intButtonSize
    
    
    Else        'Horizontal Orientation
       
       With cmdAlphaBar(0)
            If Height > 20 And Width > Width / 27 Then
                .Width = Int((UserControl.Width - 20) / 27)
                .Height = Height - 20
                If .Height / 35 > 5 Then
                    .FontSize = .Width / 35
                End If
                .Left = 10
                .Top = 10
            End If
        End With

       For intButtonSize = 1 To 26
           If Height > 20 And Width / 27 > 0 Then
                With cmdAlphaBar(intButtonSize)
                    .Height = Height - 20
                    .Top = cmdAlphaBar(0).Top
                    .Width = (Width - 20) / 27
                    .Left = cmdAlphaBar(intButtonSize - 1).Left + cmdAlphaBar(intButtonSize - 1).Width
                   If cmdAlphaBar(intButtonSize).Width / 35 > 5 Then
                       .FontSize = Int(cmdAlphaBar(intButtonSize).Width / 35)
                    End If
                End With
            End If
       Next intButtonSize
        
    End If

    With shpBorder
        .Width = Width
        .Height = Height
        .Top = 0
        .Left = 0
   End With
   

    
    
End Sub

