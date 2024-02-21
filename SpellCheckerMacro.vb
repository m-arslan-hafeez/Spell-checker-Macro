'main design file

Sub SpellCheckerMacro()
'
' SpellCheckerMacro
'
'
' Prepared By Muhammad Arslan Hafeez (Software Engineer)
' Contact: arslanhafeez1211@gmail.com
'
  Dim limit As range
  
  Set docSource = ActiveDocument
  
  Set docNew = Documents.Add
  
  For Each limit In docSource.SpellingErrors
  
    limit.Font.Color = wdColorRed
    
    limit.Font.Bold = True
    
    docNew.range.InsertAfter limit.Text
    
  Next
  
'
' This macro will check the spell in current document.
' If found word with misspelled it will highlight in red.
' Misspeld words will be save in second unsaved file.
'
  
End Sub
  
