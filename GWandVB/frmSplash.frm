VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "www.IPDG3.com Splash Form"
   ClientHeight    =   3960
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIPDG3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Picture         =   "frmSplash.frx":1CCA
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   676
      TabIndex        =   1
      Top             =   120
      Width           =   10200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9960
      Top             =   2280
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   8655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Width           =   8715
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Tag             =   "Version"
      Top             =   2280
      Width           =   10155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   3000
      Width           =   9615
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Platform"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   360
      TabIndex        =   2
      Tag             =   "Platform"
      Top             =   2640
      Width           =   9900
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program was rewritten and redesigned by George Goehring
' Date: 05/10/2001
'
' Interactive PsyberTechnology Developers Group
' George Goehring
' Developing Products to fit all your computer needs...
' Giving you the tools to build future business today...
' For more information please visit our website for details...
' www.ipdg3.com - info@ipdg3.com - Voice: 630.236.5584
' Aurora, IL. USA
'
' Source Code page www.ipdg3.com/sourcecode.html
' Online Chat www.ipdg3.com/chatroom.html
' Forums http://pub52.ezboard.com/binteractivepsybertechnologydevelopersgroup
'
' If you would like our Digital Brochure Download it here
' http://www.ipdg3.com/files/ipdg3_db.zip
'
' See gwvb01.txt for more info on the Original Author...

Option Explicit

Dim miUnloadTime As Integer

Dim msngCount As Single

Private Sub Form_Load()
    ' Set progress to 0%
    
    lblPlatform.Caption = "Platform: Windows 95 / 98 / 2000 / NT"
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor
    lblProduct.Caption = "Product Name: " & App.ProductName
    On Error GoTo Error

    UpdateStatus picStatus, 0

    Timer1.Enabled = True

    On Error GoTo 0
    
Form_Load:
    Exit Sub

Error:
    MsgBox Err.Description & " Trouble Opening File to get a working copy for your system please email us at info@ipdg3.com", vbExclamation, "Eroor Opening File"
    Timer1.Enabled = True

End Sub

Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
   
    
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
    DoEvents
    
End Sub

Private Sub Timer1_Timer()

    If miUnloadTime = 1000 Then
      UpdateStatus picStatus, 1
      Timer1.Enabled = False
      Unload Me
      frmSendMail.Show
    Else
      msngCount = msngCount + 0.005
      UpdateStatus picStatus, msngCount
      picStatus.Refresh
      miUnloadTime = miUnloadTime + Timer1.Interval
    End If

End Sub

