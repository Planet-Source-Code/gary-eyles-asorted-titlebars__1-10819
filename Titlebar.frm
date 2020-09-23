VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Animated titlebar"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7455
   Icon            =   "Titlebar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   1920
      Style           =   1  'Checkbox
      TabIndex        =   19
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Size down"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Size default"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Size up"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Animate buttons"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Modal"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
      TickStyle       =   1
      TickFrequency   =   10
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Simple"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6840
      Top             =   2400
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Menu visible"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4800
      Picture         =   "Titlebar.frx":164A
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4560
      Picture         =   "Titlebar.frx":D68C
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Close"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Restore"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Sysbut 
      Caption         =   "Minimize"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   3720
      Width           =   4455
      Begin VB.CommandButton sMinimize 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sRestore 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   240
      End
      Begin VB.CommandButton sClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Change the size of the titlebar, this will appear above menu items."
      Height          =   1215
      Left            =   6000
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "System button alpha transparency"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   2400
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsInFocus As Boolean
Dim xad As Long
Dim TheAlpha As Long
Dim AlphaAdd As Long
Dim OldItem As Integer

Private Const constalpha = 10

Private WithEvents pTitlebar As TitlebarCustom
Attribute pTitlebar.VB_VarHelpID = -1

Private Sub Check1_Click()
Dim cc As Object
'Search's through all the objects finding
'all the menus, then makes them visible or
'hides them depending on if the check box
'is checked or not.
If Check1.Value = 1 Then
    For Each cc In Me
        If TypeOf cc Is Menu Then
            cc.Visible = True
        End If
    Next
Else
    For Each cc In Me
        If TypeOf cc Is Menu Then
            cc.Visible = False
        End If
    Next
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Slider1.Enabled = False
Else
    Slider1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Dim snew As Form
Set snew = New Form2
Load snew
snew.Show 0, Me
End Sub

Private Sub Command2_Click()
Form2.Show 1, Me
End Sub

Private Sub Command3_Click()
If pTitlebar.DifferentHeight = 0 Then
    pTitlebar.DifferentHeight = -1
End If

pTitlebar.DifferentHeight = pTitlebar.DifferentHeight + 1
pTitlebar.Refresh True

On Error Resume Next
Dim tControls As Control
For Each tControls In Me
    tControls.Top = tControls.Top + 1
Next

Form_Resize
End Sub

Private Sub Command4_Click()
If pTitlebar.DifferentHeight = 0 Then
    pTitlebar.DifferentHeight = -1
End If

pTitlebar.DifferentHeight = pTitlebar.DifferentHeight - 1
pTitlebar.Refresh True

On Error Resume Next
Dim tControls As Control
For Each tControls In Me
    tControls.Top = tControls.Top - 1
Next

Form_Resize
End Sub

Private Sub Command5_Click()
pTitlebar.DifferentHeight = 0

On Error Resume Next
Dim tControls As Control
Dim tmpX As Long
tmpX = Check1.Top
For Each tControls In Me
    If tControls.hWnd <> TitleBar.hWnd Then
        tControls.Top = tControls.Top - tmpX + 10
    End If
Next

pTitlebar.Refresh True

Form_Resize
End Sub

Private Sub Form_Load()
Set pTitlebar = New TitlebarCustom

pTitlebar.TitleBar Me, TitleBar, True
pTitlebar.SetButton pCloseButton, sClose
pTitlebar.SetButton pRestoreButton, sRestore
pTitlebar.SetButton pMinimizeButton, sMinimize
pTitlebar.HasAnIcon = True
pTitlebar.Alpha = 150

Slider1.Value = pTitlebar.Alpha

List1.AddItem "Standard Titlebar"
List1.AddItem "Horizontal Gradient"
List1.AddItem "Special Horizontal Gradient"
List1.AddItem "Vertical Gradient"
List1.AddItem "Special gradient"
List1.AddItem "Tiled"
List1.AddItem "Animated"
List1.Selected(0) = True
End Sub

Private Sub Form_Resize()
Slider1.Left = 0
Slider1.Width = ScaleWidth
Slider1.Top = ScaleHeight - Slider1.Height - 10
Label3.Top = Slider1.Top - Label3.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
pTitlebar.UnTitlebar

Dim ccFrm As Form
For Each ccFrm In Forms
     Unload ccFrm
Next
End Sub

Private Sub List1_ItemCheck(Item As Integer)
Dim cI As Integer

If Item = OldItem Then
    List1.Selected(Item) = True
    Exit Sub
End If

For cI = 0 To List1.ListCount - 1
    If cI <> Item Then
        List1.Selected(cI) = False
    End If
Next

OldItem = Item
End Sub

Private Sub Slider1_Click()
pTitlebar.Alpha = Slider1.Value
pTitlebar.Refresh
End Sub

Private Sub Slider1_Scroll()
pTitlebar.Alpha = Slider1.Value
pTitlebar.Refresh
End Sub

Private Sub Sysbut_Click(index As Integer)
'Disables or Enables one of
'the titlebar buttons
If index = 0 Then
    If sClose.Enabled Then
        sClose.Enabled = False
    Else
        sClose.Enabled = True
    End If
ElseIf index = 1 Then
    If sRestore.Enabled Then
        sRestore.Enabled = False
    Else
        sRestore.Enabled = True
    End If
ElseIf index = 2 Then
    If sMinimize.Enabled Then
        sMinimize.Enabled = False
    Else
        sMinimize.Enabled = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
'This animates the titlebar
'Simply delete if you don't
'wont it animated
xad = xad + 5
If xad > Picture1.ScaleHeight Then
    xad = 0
End If

TheAlpha = TheAlpha + AlphaAdd
If TheAlpha >= 255 Then
    AlphaAdd = -constalpha
    TheAlpha = 255
ElseIf TheAlpha <= 0 Then
    AlphaAdd = constalpha
    TheAlpha = 0
End If

If Check4.Value = 1 Then
    pTitlebar.Alpha = TheAlpha
Else
    pTitlebar.Alpha = Slider1.Value
End If

pTitlebar.Refresh
End Sub

Public Sub pTitlebar_DrawTitlebar()
Dim TheIcon As Long
Dim xx, yy As Long
TheIcon = Me.Icon

If List1.List(OldItem) = "Standard Titlebar" Then
    pTitlebar.DrawDefaultCaption True, True, True
        
    DrawIconEx TitleBar.hDC, 1, 1, TheIcon, TitleBar.ScaleHeight - 2, TitleBar.ScaleHeight - 2, ByVal 0&, ByVal 0&, &H3
    'Put more drawing command here if you want something
    'extra on the default titlebar. e.g. Like a picture.
    
    TitleBar.Refresh
    Exit Sub
End If

If List1.List(OldItem) = "Animated" Or List1.List(OldItem) = "Tiled" Then
    With TitleBar
    'Clear the titlebar
    .Cls

    'Depending whether the form is in focus
    'or not, tile one of the following pictures
    If pTitlebar.Focus Then
        'Form is in focus so tile the redish picture
        'from Picture1
        pTitlebar.TilePicture Picture1, IIf(List1.List(OldItem) = "Animated", xad, 0)
    Else
        'Form isn't in focus so tile the redish picture
        'from Picture2
        pTitlebar.TilePicture Picture2, IIf(List1.List(OldItem) = "Animated", xad, 0)
    End If

    'Draw the forms icons in the top left
    DrawIconEx .hDC, 1, 1, TheIcon, .ScaleHeight - 2, .ScaleHeight - 2, ByVal 0&, ByVal 0&, &H3
    pTitlebar.DrawTextEx Me.Caption, .ScaleHeight + 5, 0, sMinimize.Left, .ScaleHeight

    'Refresh the titlebar in order for the
    'final result to show
    .Refresh
    End With
    Exit Sub
End If

Dim FsColor As Long
Dim LsColor As Long

If pTitlebar.Focus Then
    FsColor = QBColor(12)
    LsColor = QBColor(14)
Else
    FsColor = QBColor(8)
    LsColor = QBColor(7)
End If

Dim sLeft As Long
sLeft = pTitlebar.DrawTextEx(Me.Caption, TitleBar.ScaleHeight + 5, 0, sMinimize.Left, TitleBar.ScaleHeight, True)

TitleBar.ForeColor = QBColor(9)
If List1.List(OldItem) = "Special gradient" Then
    TitleBar.ForeColor = QBColor(15)
    pTitlebar.TriangleFill TitleBar, QBColor(9), QBColor(14), QBColor(10), QBColor(12)
End If
If List1.List(OldItem) = "Horizontal Gradient" Then
    pTitlebar.GradientFill TitleBar, FsColor, LsColor
End If
If List1.List(OldItem) = "Special Horizontal Gradient" Then
    If pTitlebar.Focus Then
        FsColor = QBColor(14)
        LsColor = QBColor(12)
    End If
    pTitlebar.GradientFill TitleBar, FsColor, LsColor, sLeft, 25
End If
If List1.List(OldItem) = "Vertical Gradient" Then
    pTitlebar.GradientFill TitleBar, FsColor, LsColor, , , True
End If

    DrawIconEx TitleBar.hDC, 1, 1, TheIcon, TitleBar.ScaleHeight - 2, TitleBar.ScaleHeight - 2, ByVal 0&, ByVal 0&, &H3
    pTitlebar.DrawTextEx Me.Caption, TitleBar.ScaleHeight + 5, 0, sMinimize.Left, TitleBar.ScaleHeight
    TitleBar.Refresh

TitleBar.ForeColor = QBColor(9)
End Sub
