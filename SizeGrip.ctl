VERSION 5.00
Begin VB.UserControl SizeGrip 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   FillStyle       =   0  'Solid
   ScaleHeight     =   630
   ScaleWidth      =   975
   ToolboxBitmap   =   "SizeGrip.ctx":0000
   Begin VB.Label lblGrip 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblGrip 
      BackStyle       =   0  'Transparent
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "SizeGrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       SizeGrip - Grip Control that simulates the Statusbar Control Grip
'
'   Product Name:
'       SizeGrip.ctl
'
'   Compatability:
'       Windows: 98, ME, NT4, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'       Adapted from the following online article(s):
'       http://vb.mvps.org/articles/ap199906.pdf
'       http://vbnet.mvps.org/index.html?code/helpers/iswinversion.htm
'       http://www.ftponline.com/archives/premier/mgznarch/vbpj/1998/07jul98/fb0798.pdf
'
'   Legal Copyright & Trademarks (Current Implementation):
'       Copyright © 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this software.
'
'-  Modification(s) History:
'       17Mar05 - Initial build of the SizeGrip Control
'       19Mar05 - Added OS Detection support and ability to change the Grip Style,
'                 X & Y Offsets, ShadowOffset, and Radius of the Grip Shape. The
'                 Control now supports, XP Square Grips, XP Circle Grips, and the
'                 Classic Windows style Grips.
'       20Mar05 - Added event monitoring of the host form's resize event handler to
'                 permit the control to automatically reposition itself in the lower
'                 right corner of the container. In the 17Mar05 implementation, this
'                 "moving" was handled in the parent form's resize event structure.
'       22Mar05 - Fixed a bug which did not correctly place the control on the form
'                 when the Forms ScaleMode was not set to vbTwips.
'
'   Force Declarations
Option Explicit

'   Type Declarations (Adapted from VBNet, Randy Birch)
Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: string providing arbitrary additional information
End Type

'   Custom Control Shape Enumerations
Public Enum ucShape
    ucSquare = 0
    ucCircle = 1
    ucBars = 2
End Enum

'   API Declarations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
'   API Message Constants
Private Const HTBOTTOMRIGHT = 17
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const WM_NCLBUTTONDOWN = &HA1

'   Local Declarations
Private m_Radius        As Long
Private m_ShadowOffset  As Long
Private m_Shape         As ucShape
Private m_XOffset       As Long
Private m_YOffset       As Long
Private Ctrl            As Control

'   Local Events Declarations
Private WithEvents ParentForm As Form
Attribute ParentForm.VB_VarHelpID = -1

Public Property Get GripShape() As ucShape
    GripShape = m_Shape
End Property

Public Property Let GripShape(lShape As ucShape)
    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lShape <> m_Shape Then
        m_Shape = lShape
        PropertyChanged "GripShape"
        UserControl_Repaint
    End If
End Property

Public Function IsWinXP() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Adapted from: VBnet, Randy Birch, All Rights Reserved.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And _
            (OSV.dwBuildNumber >= 2600)
    End If
End Function

Private Sub lblGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '   Relase any events captured previously
    ReleaseCapture
    '   Send a message that we are resizing the form
    SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub

Private Sub ParentForm_Resize()
    '   The routine, effecivtly captures all of the parent forms resizing
    '   events and allows our control to properly place itself with out the
    '   the need for additional external moving calls in the parent forms
    '   resize event handler.
    Dim pOldMode            As Long
    Dim cOldMode            As Long
    Dim ParentContainer     As Object
    
    '   Handle errors quielty...
    On Error Resume Next
        
    For Each Ctrl In ParentForm.Controls
        '   Is this control a Sizegrip or somthing else?
        If TypeOf Ctrl Is SizeGrip Then
            '   Store the Old Scale mode
            pOldMode = ParentForm.ScaleMode
            cOldMode = UserControl.ScaleMode
            '   Set the mode to Twips
            ParentForm.ScaleMode = vbTwips
            UserControl.ScaleMode = vbTwips
            If GripShape <> ucBars Then
                '   This is for XP Styled Grips with XP Form Borders
                Ctrl.Left = ParentForm.ScaleWidth - UserControl.ScaleWidth + 20
                Ctrl.Top = ParentForm.ScaleHeight - UserControl.ScaleHeight + 20
            Else
                If IsWinXP Then
                    '   If we are showing this with XP boarders then the padding is
                    '   not scaled correctly so we need to adjust this a touch....
                    Ctrl.Left = ParentForm.ScaleWidth - UserControl.ScaleWidth - 20
                    Ctrl.Top = ParentForm.ScaleHeight - UserControl.ScaleHeight - 20
                Else
                    '   Must be a windows classic form, so do not adjsut the padding...
                    Ctrl.Left = ParentForm.ScaleWidth - UserControl.ScaleWidth
                    Ctrl.Top = ParentForm.ScaleHeight - UserControl.ScaleHeight
                    Debug.Print Ctrl.Top, Ctrl.Left
                End If
            End If
            '   Set the Scale mode back
            ParentForm.ScaleMode = pOldMode
            UserControl.ScaleMode = cOldMode
        End If
    Next
End Sub

Public Property Get Radius() As Long
    Radius = m_Radius
End Property

Public Property Let Radius(lValue As Long)
    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lValue <> m_Radius Then
        m_Radius = lValue
        PropertyChanged "Radius"
        UserControl_Repaint
    End If
End Property

Public Property Get ShadowOffset() As Long
    ShadowOffset = m_ShadowOffset
End Property

Public Property Let ShadowOffset(lValue As Long)
    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lValue <> m_ShadowOffset Then
        m_ShadowOffset = lValue
        PropertyChanged "ShadowOffset"
        UserControl_Repaint
    End If
End Property

Public Sub Refresh()
    '   Force a resize event to make sure that the control is
    '   place correclty, otherwise the padding may be off...
    ParentForm_Resize
    '   Now repaint the control face
    UserControl_Repaint
End Sub

Private Sub UserControl_Initialize()
    '   Force a resize event to make sure the control ends
    '   up in the correct postion on the parent form...
    ParentForm_Resize
End Sub

Private Sub UserControl_InitProperties()
    '   Set the appropraite type of grip and the properties
    If IsWinXP Then
        XOffset = 40
        YOffset = 40
        Radius = 15
        ShadowOffset = 17
        GripShape = ucCircle
    Else
        XOffset = 70
        YOffset = 70
        ShadowOffset = 20
        GripShape = ucBars
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '   Relase any events captured previously
    ReleaseCapture
    '   Send a message that we are resizing the form
    SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub

' Read & Write the few properties this control can cache
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        m_Radius = .ReadProperty("Radius", 15)
        m_YOffset = .ReadProperty("YOffset", 40)
        m_XOffset = .ReadProperty("XOffset", 40)
        m_Shape = .ReadProperty("GripShape", ucCircle)
        m_ShadowOffset = .ReadProperty("ShadowOffset", 17)
        If Ambient.UserMode = False Then Exit Sub
        '   Reference the parent form and start recieving events
        Set ParentForm = UserControl.Parent
    End With
End Sub

Private Sub UserControl_Repaint()
    '   Custom reoutine, to paint/repaint the shapes on the
    '   screen to represent the Grip Style selected...
    Dim fnt             As New StdFont
    Dim i               As Long
    Dim j               As Long
    Dim BoxSz           As Long
    Dim X1              As Long
    Dim Y1              As Long
    Dim X2              As Long
    Dim Y2              As Long

    With UserControl
        '   Clear any previous drawings
        .Cls
        '   Check to see if we are running XP or Classic
        'If (Not IsWinXP) Or (GripShape = ucBars) Then
        If (GripShape = ucBars) Then
            '   Create a new font
            With fnt
                .Name = "Marlett"
                .Bold = False
                .Size = 12
            End With
            '   Set the grip labels with the respective properties
            With .lblGrip(0)
                Set Font = fnt
                .AutoSize = True
                .Caption = "o"
                .ForeColor = vb3DHighlight
                .MousePointer = vbSizeNWSE
                .Left = 50
                .Top = 50
                .ZOrder
            End With
            With .lblGrip(1)
                Set Font = fnt
                .AutoSize = True
                .Caption = "p"
                .ForeColor = vb3DShadow
                .MousePointer = vbSizeNWSE
                .Left = 50
                .Top = 50
                .ZOrder
            End With
            '   Free up memory
            Set fnt = Nothing
            '   Show our Win98-Win2K GripLabels
            .lblGrip(0).Visible = True
            .lblGrip(1).Visible = True
        Else
            '   Hide our Win98-Win2K GripLabels
            .lblGrip(0).Visible = False
            .lblGrip(1).Visible = False
            '   Make sure to refresh the control with changes..
            .AutoRedraw = True
            '   Set the mouse pointer
            .MousePointer = vbSizeNWSE
            '   Grip images are set up in a 25 element array, with only
            '   5, 12, 20, 14, 22, and 24 that are plotted...
            For i = 1 To 5 Step 2
                For j = 1 To 5 Step 2
                    If (i * j <> 1) And (i * j <> 3) And (i * j <> 10) Then
                        If GripShape = ucSquare Then
                            '   Position box sizes for the control
                            BoxSz = 36
                            '   Compute the upper left corner(s) of the squares
                            X1 = BoxSz * i + XOffset + 35
                            Y1 = BoxSz * j + YOffset + 35
                            '   Compute the lower right corner(s) of the squares
                            X2 = BoxSz * 2 * (i / 2 + 1) - BoxSz + XOffset
                            Y2 = BoxSz * j + BoxSz + YOffset
                            '   Create a highlight that is offset
                            Line (X1 + ShadowOffset, Y1 + ShadowOffset)-(X2 + ShadowOffset - 15, Y2 + ShadowOffset - 15), vb3DHighlight, BF
                            '   We are using &HC0C0C0 as 'vb3DShadow is too dark
                            Line (X1, Y1)-(X2 - ShadowOffset, Y2 - ShadowOffset), vb3DShadow, BF
                        Else
                            'XOffset = 40
                            'YOffset = 40
                            Radius = 15
                            'ShadowOffset = 17
                            '   Compute the centers of the circles
                            X1 = j * XOffset + XOffset + 20
                            Y1 = i * YOffset + YOffset + 20
                            '   Create a highlight that is offset
                            .FillColor = vb3DHighlight
                            Circle (X1 + ShadowOffset, Y1 + ShadowOffset), Radius, vb3DHighlight
                            .FillColor = vb3DShadow
                            Circle (X1, Y1), Radius, vb3DShadow
                        End If
                    End If
                Next j
            Next i
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    '   Prevent resizing...
    With UserControl
        .Width = 345
        .Height = 345
    End With
End Sub

Private Sub UserControl_Show()
    UserControl_Repaint
End Sub

Private Sub UserControl_Terminate()
    '   Make sure to clean up...
    Set ParentForm = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "Radius", m_Radius, 15
        .WriteProperty "YOffset", m_YOffset, 40
        .WriteProperty "XOffset", m_XOffset, 40
        .WriteProperty "GripShape", m_Shape, ucCircle
        .WriteProperty "ShadowOffset", m_ShadowOffset, 17
    End With
End Sub

Public Property Get XOffset() As Long
    XOffset = m_XOffset
End Property

Public Property Let XOffset(lValue As Long)
    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lValue <> m_XOffset Then
        m_XOffset = lValue
        PropertyChanged "XOffset"
        UserControl_Repaint
    End If
End Property

Public Property Get YOffset() As Long
    YOffset = m_YOffset
End Property

Public Property Let YOffset(lValue As Long)
    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lValue <> m_YOffset Then
        m_YOffset = lValue
        PropertyChanged "YOffset"
        UserControl_Repaint
    End If
End Property


