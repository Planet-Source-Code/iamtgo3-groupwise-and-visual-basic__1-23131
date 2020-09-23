VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAttachFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attach File"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   8040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1770
      Width           =   8040
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   7815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File To Attach :"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmAttachFile"
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

Private Sub cmdAdd_Click()
       
    On Error GoTo Error
    
    If txtFile.Text = "" Then
      frmSendMail.mstAttachFile = ""
      MsgBox "You have not yet specified a file to attach...", vbInformation, "No File Attached"
      txtFile.SetFocus
    Else
      frmSendMail.lstAttachedFiles.AddItem txtFile.Text
      Unload Me
    End If
        
    On Error GoTo 0
    
cmdAdd_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [cmdAdd]"

End Sub

Private Sub cmdBrowse_Click()
       
    On Error GoTo Error
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.Flags = cdlOFNFileMustExist
    
    On Error Resume Next
    
    CommonDialog1.ShowOpen
    
    If Not Err Then
        txtFile.Text = CommonDialog1.FileName
    End If
    
    txtFile.SetFocus
        
    On Error GoTo 0

cmdBrowse_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [cmdBrowse]"

End Sub

Private Sub cmdClose_Click()
    
    frmSendMail.mstAttachFile = ""
    
    Unload Me

End Sub

Private Sub Form_Load()
       
    On Error GoTo Error

    frmSendMail.mstAttachFile = ""

    On Error GoTo 0

Form_Load_Exit:
    Exit Sub

Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [Form_Load]"

End Sub
