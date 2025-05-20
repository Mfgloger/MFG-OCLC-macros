'MacroName:Unicode detector
'MacroDescription:Find unicode hiding as text in  245, 246, 520 fields


Sub Main

   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")
 
   Dim sNameTitle As String
   Dim sAddTitle As String
   Dim sNote As String
   Dim sTagt As String
   Dim sTagu As String
   Dim sTagn As String
   
   
   sTagt = "245"   
   sTagu = "246"
   sTagn = "520"
   CS.GetFieldUnicode sTagt, 2, sNameTitle
   If InStr(sNameTitle, "&#") <> 0 Then
   MsgBox "Field " & Left(sNameTitle, 3) & " includes a possibly incorrectly coded character (see &#x code). Please replace with ALA diacritics: '" & Mid(sNameTitle, 6) &  "'" 

Else 
   MsgBox "All characters in title are ALA -- Good to go!"
   End If

      
   CS.GetFieldUnicode sTagu, 2, sAddTitle
   If InStr(sAddTitle, "&#") <> 0 Then
   MsgBox "Field " & Left(sAddTitle,3) & " includes a possibly incorrectly coded character (see &#x code). Please replace with ALA diacritics: '" & Mid(sNameTitle, 6) &  "'"        

Else   
   MsgBox "All characters in added title are ALA -- Good to go!"
   End If
   
      
   CS.GetFieldUnicode sTagn, 1, sNote
   If InStr(sAddTitle, "&#") <> 0 Then
   MsgBox "Field " & Left(sNote,3) & " includes a possibly incorrectly coded character (see &#x code). Please replace with ALA diacritics: '" & Mid(sNameTitle, 6) &  "'"        

Else   
   MsgBox "All characters in contents are ALA -- Good to go!"
   End If
   
    
   
End Sub
