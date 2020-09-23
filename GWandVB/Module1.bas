Attribute VB_Name = "Module1"
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

' Disables "X" Close button on title bar
Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000

Public Declare Function DrawMenuBar Lib "user32" _
      (ByVal hwnd As Long) As Long
        
Public Declare Function GetMenuItemCount Lib "user32" _
      (ByVal hMenu As Long) As Long
        
Public Declare Function GetSystemMenu Lib "user32" _
      (ByVal hwnd As Long, _
       ByVal bRevert As Long) As Long
         
Public Declare Function RemoveMenu Lib "user32" _
      (ByVal hMenu As Long, _
       ByVal nPosition As Long, _
       ByVal wFlags As Long) As Long '--end block--'

Sub Main()
    
    frmSplash.Show
    frmSplash.Refresh
 
End Sub
