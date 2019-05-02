'MacroName:LCCallOpenShelf
'MacroDescription:LC Call version 2.0
   
   'LC Call number
'Created by Miriam Gloger 
'Copy Call number from 050 to Call number and Item record Version 1.2

Option Explicit

Sub Main
  Dim CS as object
  Set CS = CreateObject("Connex.Client")  
  Dim sLccall as String
  Dim sCall as String 
  Dim sCutter as String 
  Dim nPlace1 as Integer
  Dim nPlace2 as Integer
  Dim sUslhg as String
  Dim sMapdiv as String
  Dim sMagshelf as String
  Dim sMapshelf
  Dim sData as String
  Dim nBool as Integer
  Dim sInitials as String
  Dim z as Integer


   sInitials = "MFG"  
   If CS.ItemType = 0 Or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 or CS.ItemType = 19 or CS.ItemType = 26 or CS.ItemType = 31 or CS.ItemType = 33 or CS.ItemType = 35 Then
         
   nBool = CS.GetField("050",1, sLccall)
   If nBool = True Then
     
      sLccall = Mid(sLccall, 6)
      nPlace1 = InStr(sLccall, Chr(223))
      sCall = Rtrim(Left(sLccall, nPlace1-1))
      nPlace2=InStr(sLccall, Chr(223) & "b")
      sCutter=Ltrim(Mid(sLccall, nPlace2+2))
      'set location
      sUslhg = ("*R-USLHG ")   
      sMagshelf = ("MAGG1")  
      'sMapdiv = ("*R-Map Div.")   
      'sMapshelf = ("MAPP1")  
      CS.SetField 1, "85201" & "ßk " & sUslhg & "ßh " & sCall & "ßi" & sCutter
      'CS.SetField 1, "85201" & "ßk " & sMapdiv & "ßh " & sCall & "ßi" & sCutter & "ßc +"
      CS.Reformat     

      CS.SetField 1, "901  CAT/RL ßb " & sInitials
      CS.SetField 1, "945  .b" 
      CS.SetField 1, "946  m"
      CS.SetField 1, "949  *b2=a;recs=oclcgw;" 
      If CS.GetField("949", 2, sData) = True Then
           
    
     'CS.AddField 2,"949 1" & "ßz 85201" & "ßf " & sUslhg & "ßa " & sCall & "ßb" & sCutter &"ßi XXXXXXXXXXXXXX ßl " & sMagshelf & " ßs b ßt 002 ßh 033 ßo 1 ßv" & "CATRL/" & sInitials

               
     End If
     CS.Reformat
     MsgBox "Please verify the Library of Congress Classification"
     End if
 
    Else
      MsgBox "This Macro can only be used in Bibliographic records"
      End If
      
    End Sub
