VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "SizeGrip Control Example"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   1.27
   ScaleMode       =   0  'User
   ScaleWidth      =   6.959
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Styles 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2175
   End
   Begin Project1.SizeGrip SizeGrip1 
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'       Adapted from the following online article:
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
'       17Mar05 - Initial test harness for SizeGrip Control
'
'   Force Declarations
Option Explicit

Private Sub Form_Load()
    With Me
        .Styles.AddItem "XP Squares Grip"
        .Styles.AddItem "XP Circles Grip"
        .Styles.AddItem "Windows Classic Grip"
        .Styles.ListIndex = 1
    End With
End Sub

Private Sub Styles_Click()
    With Me
        '   We changed the styles, so update the control, and force a
        '   refresh event to reposition the controls....
        '   (This is not needed unless one changes the Grip Style at runtime)
        .SizeGrip1.GripShape = .Styles.ListIndex
        '   Force a refresh in the control event handler
        .SizeGrip1.Refresh
    End With
End Sub
