VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "C to VB"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frOptions 
      Caption         =   "Options"
      Height          =   1455
      Left            =   7140
      TabIndex        =   9
      Top             =   2130
      Width           =   975
      Begin VB.TextBox txtLibraryName 
         Height          =   285
         Left            =   30
         TabIndex        =   13
         Text            =   "libname"
         Top             =   1110
         Width           =   915
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Private"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Width           =   795
      End
      Begin VB.OptionButton optOptions 
         Caption         =   "Public"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.Label lblLib 
         AutoSize        =   -1  'True
         Caption         =   "DLL name"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   870
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   435
      Left            =   750
      TabIndex        =   16
      Top             =   1230
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   5850
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   435
      Left            =   7140
      TabIndex        =   14
      Top             =   1650
      Width           =   975
   End
   Begin VB.OptionButton optWhat 
      Caption         =   "User Input"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   7
      Top             =   1770
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame frMain 
      Height          =   4365
      Left            =   180
      TabIndex        =   8
      Top             =   1770
      Width           =   6945
      Begin VB.PictureBox Lower 
         Height          =   1785
         Left            =   3630
         ScaleHeight     =   1725
         ScaleWidth      =   3045
         TabIndex        =   21
         Top             =   390
         Width           =   3105
         Begin VB.TextBox txtBas 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   0
            Width           =   3315
         End
      End
      Begin VB.PictureBox Upper 
         Height          =   1635
         Left            =   300
         ScaleHeight     =   1575
         ScaleWidth      =   3225
         TabIndex        =   19
         Top             =   360
         Width           =   3285
         Begin VB.TextBox txtC 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "VB Declaration"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   18
         Top             =   4320
         Width           =   1065
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "C Declaration"
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   17
         Top             =   4800
         Width           =   960
      End
      Begin VB.Label Sep 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   50
         Left            =   210
         MousePointer    =   7  'Size N S
         TabIndex        =   23
         Top             =   2430
         Width           =   105
      End
   End
   Begin VB.OptionButton optWhat 
      Caption         =   "File"
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   60
      Width           =   585
   End
   Begin VB.Frame frFile 
      Enabled         =   0   'False
      Height          =   915
      Left            =   270
      TabIndex        =   1
      Top             =   60
      Width           =   2925
      Begin VB.CommandButton cmdSelectFiles 
         Caption         =   "Select"
         Height          =   585
         Left            =   2220
         TabIndex        =   6
         Top             =   270
         Width           =   645
      End
      Begin VB.TextBox txtBASfile 
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtHFile 
         Height          =   285
         Left            =   690
         TabIndex        =   3
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   ".bas file:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   ".h file:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    C2VB  Converts C style definitions to VB
'    Copyright (C) 2000  Kimon Andreou (kimon@mindless.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit

'Local variables
Private hFilename As String
Private basFilename As String
Private pcntSep As Double

'Clears the text from both C and VB text boxes
Private Sub cmdClear_Click()
txtC.Text = ""
txtBas.Text = ""
End Sub

'Copies the contents of the VB text box to the Clipboard
Private Sub cmdCopy_Click()
Clipboard.SetText txtBas.Text   'You can also set pictures, check the help file.
End Sub

'Calls the Common Dialog control and retrieves the filenames selected by
'the user.
'If you don't like the limitations set by this control, you should check
'out The Common Controls Replacement Project (CCRP) home page at
' http://www.mvps.org/ccrp
'
Private Sub GetFilenames(strHfile As String, strBASfile As String)

'Set the properties and call it
With cdlg
    .Filter = "C header files|*.h|All files|*.*"
    .DialogTitle = "Choose .h file"
    .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    .ShowOpen
    
    'Okay, we got the .h file, now to get the VB file.
    strHfile = .FileName
    .FileName = ""
    .Filter = "VB modules|*.bas|All files|*.*"
    .DefaultExt = ".bas"
    .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    .DialogTitle = "Select target file"
    .ShowSave
    strBASfile = .FileName
End With
    
End Sub

'Do all the work required prior to processing.
Private Sub cmdProcess_Click()
Dim dummy As String

'Lets the user know something is going on.
Me.MousePointer = vbHourglass

'Get the library name (a.k.a. DLL name)
txtLibraryName.Text = Trim(txtLibraryName.Text)
LibraryName = IIf(txtLibraryName.Text = "", "libname", txtLibraryName.Text)

'Ok, are we working with files or with user input?
If frFile.Enabled Then  'files
    hFilename = txtHFile.Text
    basFilename = txtBASfile.Text
    
    'make sure we have files to work with
    If (txtHFile.Text = "") And (txtBASfile.Text = "") Then
        GetFilenames hFilename, basFilename
        If (hFilename = "") Or (basFilename = "") Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If txtHFile.Text = "" Then
            MsgBox "You have to provide an input file", vbExclamation
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        If txtBASfile.Text = "" Then
            MsgBox "You have to provide an output file", vbExclamation
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        'See if it exists.  Although Dir() is not perfect, it works.
        'If you want to optimize this piece of code, try the API function
        'FindFirstFile()
        If Dir(txtHFile.Text) = "" Then
            MsgBox "File: " & txtHFile.Text & " does not exist!", vbExclamation
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    txtHFile.Text = hFilename
    txtBASfile.Text = basFilename
    
    'Process the files
    Process hFilename, , True, basFilename
    MsgBox "Conversion complete", vbOKOnly, "Done"
Else    'User input
    Process txtC.Text, dummy
    txtBas.Text = dummy
    txtBas.SelStart = Len(txtBas.Text)
End If
'Restore the mouse pointer.
Me.MousePointer = vbDefault
End Sub

'Retrieve the source and target files
Private Sub cmdSelectFiles_Click()
GetFilenames hFilename, basFilename
If (hFilename = "") Or (basFilename = "") Then Exit Sub
txtHFile.Text = hFilename
txtBASfile.Text = basFilename
End Sub

Private Sub Form_Load()
Dim Args As String

'Initialize the local and global variables.
IsPublic = True
LibraryName = "libname"

'See if we got any command line arguments
Args = Command()
If Args <> "" Then              'if we did then...
    Dim fparts() As String
    
    txtHFile.Text = Args
    fparts = Split(Args, ".")
    fparts(UBound(fparts)) = "bas"
    Args = Join(fparts, ".")
    txtBASfile.Text = Args
    frFile.Enabled = True
    frMain.Enabled = False
    optWhat(0).Value = True
    hFilename = txtHFile.Text
    basFilename = txtBASfile.Text
End If
    
pcntSep = 0.5       'Set the ratio for the splitter
End Sub

'Here is the magic for the form...
Private Sub Form_Resize()
Dim frmWidth As Long
Dim frmHeight As Long

'if the form was minimized, ignore this stuff
If Me.WindowState = vbMinimized Then Exit Sub

'Store the "perfect" width in a variable, easier to access.
frmWidth = Me.ScaleWidth - cmdProcess.Width - 200

'Set .Top and .Left for buttons and options
frOptions.Top = 0
frOptions.Left = frmWidth + 100
cmdProcess.Left = frmWidth + 100
cmdProcess.Top = frOptions.Top + frOptions.Height + 100
cmdClear.Left = cmdProcess.Left
cmdClear.Top = cmdProcess.Top + cmdProcess.Height + 50
cmdCopy.Left = cmdProcess.Left
cmdCopy.Top = cmdClear.Top + cmdClear.Height + 50

'Set coordinates for the 2 main frames
With frFile
    .Left = 0
    .Top = 0
    .Width = frmWidth
End With

frmHeight = Me.ScaleHeight - frFile.Height
With frMain
    .Left = 0
    .Top = frFile.Top + frFile.Height
    .Width = frmWidth
    .Height = IIf(frmHeight < 0, 0, frmHeight)
End With
frmWidth = (frMain.Width - 100)

'Set the separator's position on the form (actually in the frame)
Sep.Left = 50
Sep.Width = frMain.Width - 100
'This will make sure that the textboxes will occupy the same ratio of the
'frame whatever the size of the frame itself.
Sep.Top = IIf(frMain.Height < lblLabels(0).Height * 5, Sep.Top, frMain.Height * pcntSep)

'Set the label positions
lblLabels(0).Top = 250
lblLabels(0).Left = 50
lblLabels(0).Width = frmWidth
lblLabels(1).Top = Sep.Top + Sep.Height
lblLabels(1).Left = 50
lblLabels(1).Width = frmWidth


'This is where the textboxes (which are contained in pictureboxes) get resized.
Upper.Left = 50
Lower.Left = 50
Upper.Top = lblLabels(0).Top + lblLabels(0).Height
Lower.Top = lblLabels(1).Top + lblLabels(1).Height
Upper.Width = frmWidth
Lower.Width = frmWidth

'Call the pane resize function
Sep_MouseMove vbLeftButton, 0, 0, 0

'This takes care of the upper frame that holds the filenames
txtHFile.Width = IIf(txtC.Width < (lblLabels(3).Width + lblLabels(3).Left + _
    cmdSelectFiles.Width), 0, txtC.Width - lblLabels(3).Width - lblLabels(3).Left - _
    cmdSelectFiles.Width)
txtBASfile.Width = txtHFile.Width

cmdSelectFiles.Left = txtHFile.Left + txtHFile.Width + 50
cmdSelectFiles.Top = txtHFile.Top

optWhat(0).Left = frFile.Left + 200
optWhat(1).Left = optWhat(0).Left
optWhat(0).Top = frFile.Top
optWhat(1).Top = frMain.Top
End Sub

'Sets the size of the VB textbox
Private Sub Lower_Resize()
txtBas.Top = 0
txtBas.Left = 0
txtBas.Width = Lower.ScaleWidth
txtBas.Height = Lower.ScaleHeight
End Sub

Private Sub optOptions_Click(Index As Integer)
IsPublic = optOptions(0).Value
End Sub

Private Sub optWhat_Click(Index As Integer)
frFile.Enabled = optWhat(0).Value
frMain.Enabled = optWhat(1).Value

If Not Me.Visible Then Exit Sub

If frFile.Enabled Then
    txtHFile.SetFocus
Else
    txtC.SetFocus
End If
 
End Sub

'This is where the "splitting" occurs.
'
'I am using picture boxes for the panes.  One for each one.
'The reason I chose to use picture boxes, is because I can manipulate them
'the same way I can a form.  Most importantly, it's got a resize event
'which can be used to resize its contained controls.
Private Sub Sep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Exit Sub     'We only care about the left button

'Make sure that the separator doesn't get out of bounds
Sep.Top = IIf(Sep.Top + Y < lblLabels(0).Top + lblLabels(0).Height + 450, lblLabels(0).Top + lblLabels(0).Height + 450, Sep.Top + Y)
Sep.Top = IIf(Sep.Top > frMain.Height - lblLabels(1).Height - Sep.Height - 470, frMain.Height - lblLabels(1).Height - Sep.Height - 470, Sep.Top)

'Check the size of the frame  (see if we fit)
If frMain.Height < lblLabels(0).Height * 5 Then Exit Sub

'Resize the top part
Upper.Height = Sep.Top - Upper.Top - 50

'Set the vertical coordinate of the label box "VB Declaration"
lblLabels(1).Top = Sep.Top + Sep.Height

'Place and resize the lower part
Lower.Top = lblLabels(1).Top + lblLabels(1).Height
Lower.Height = frMain.Height - Lower.Top - 100

'Store the new ratio
pcntSep = Sep.Top / frMain.Height

End Sub

'Resize the C textbox
Private Sub Upper_Resize()
txtC.Top = 0
txtC.Left = 0
txtC.Width = Upper.ScaleWidth
txtC.Height = Upper.ScaleHeight
End Sub
