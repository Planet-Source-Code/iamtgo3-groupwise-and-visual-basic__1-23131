VERSION 5.00
Object = "{2537A820-4A74-11CF-A794-00AA002AFE9E}#1.0#0"; "GWNCC1.OCX"
Begin VB.Form frmSendMail 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Email and Attachment"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstAttachedFiles 
      Height          =   1035
      ItemData        =   "frmSendMail.frx":0CCA
      Left            =   120
      List            =   "frmSendMail.frx":0CCC
      TabIndex        =   11
      Top             =   5280
      Width           =   6375
   End
   Begin VB.CheckBox chkInternalAddress 
      Caption         =   "Internal Email Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6660
      Width           =   6600
      Begin VB.CommandButton cmdAttachFile 
         Caption         =   "&Attach File"
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearAttachment 
         Caption         =   "Clear &Attachment"
         Height          =   300
         Left            =   3600
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
   End
   Begin GWNCCLib.GWncc GWncc1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   529
      _StockProps     =   69
      BackColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMessage 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attached File :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6600
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Message :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Message To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSendMail"
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

Public objGroupWise As Object
Public objAccount As Object
Public mstAttachFile As String

Private Sub cmdAttachFile_Click()
   
    On Error GoTo Error
   
    frmAttachFile.Show vbModal, Me
    
    cmdAttachFile.Enabled = False
    
    On Error GoTo 0
    
cmdAttachFile_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [cmdAttachFile]"

End Sub

Private Sub cmdClearAttachment_Click()

    lstAttachedFiles.Clear
    cmdAttachFile.Enabled = True

End Sub

Private Sub cmdClose_Click()
    
    End

End Sub

Private Sub cmdSend_Click()

Dim objMessages As Object
Dim objMessage As Object
Dim objMailBox As Object
Dim objRecipients As Object
Dim objRecipient As Object
Dim objAttachment As Object
Dim objAttachments As Object
Dim GWCount As Integer
Dim objMessageSent As Variant

    On Error GoTo Error
          
    Set objMailBox = objAccount.MailBox
    Set objMessages = objMailBox.Messages
    Set objMessage = objMessages.Add("GW.MESSAGE.MAIL", "Draft")
    Set objRecipients = objMessage.Recipients
    Set objAttachments = objMessage.Attachments

    mstAttachFile = lstAttachedFiles.List(0)
    
    If mstAttachFile <> "" Then
      Set objAttachment = objAttachments.Add(mstAttachFile)
    End If
    
    For GWCount = 1 To GWncc1.Count
      If chkInternalAddress.Value = True Then
        Set objRecipient = objRecipients.Add(GWncc1.EMailAddress(GWCount - 1), "NGW", "egwTo")
      Else
        Set objRecipient = objRecipients.Add(GWncc1.EMailAddress(GWCount - 1))
      End If
    Next GWCount
  
    For GWCount = 1 To objMessage.Recipients.Count
      objMessage.Recipients.Resolve
      If objRecipient.Resolved = "egwNotResolved" Then
        MsgBox "Recipient not resolved. Try again."
        GWncc1.SetFocus
        Exit Sub
     End If
    Next GWCount
     
    objMessage.Subject = txtSubject.Text
    objMessage.BodyText = txtMessage.Text
      
    Set objMessageSent = objMessage.Send
   
    MsgBox "Your message has been sent Successfully..." & vbCrLf & vbCrLf & _
           "To : " & objRecipient & vbCrLf & vbCrLf & _
           "Subject : " & txtSubject.Text, , "Message Sent Successfully"
            
    txtSubject.Text = ""
    txtMessage.Text = ""
    GWncc1.Clear
    lstAttachedFiles.Clear
    picButtons.SetFocus
    
    On Error GoTo 0

cmdSend_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [cmdSend]"
    
End Sub

Private Sub Form_Load()
Dim hMenu As Long
Dim menuItemCount As Long
      
    On Error GoTo Error
   
    Set objGroupWise = CreateObject("NovellGroupWareSession")
    Set objAccount = objGroupWise.Login

    'Obtain the handle to the form's system menu
    hMenu = GetSystemMenu(Me.hwnd, 0)
    
    If hMenu Then
    
      'Obtain the number of items in the menu
      menuItemCount = GetMenuItemCount(hMenu)
      
      'Remove the system menu Close menu item.
      'The menu item is 0-based, so the last
      'item on the menu is menuItemCount - 1
      Call RemoveMenu(hMenu, menuItemCount - 1, _
                      MF_REMOVE Or MF_BYPOSITION)
     
      'Remove the system menu separator line
      Call RemoveMenu(hMenu, menuItemCount - 2, _
                      MF_REMOVE Or MF_BYPOSITION)
      
      'Force a redraw of the menu. This
      'refreshes the titlebar, dimming the X
      Call DrawMenuBar(Me.hwnd)
     
    End If
        
    On Error GoTo 0

Form_Load_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [Form_Load]"
 
End Sub
