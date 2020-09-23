<div align="center">

## Spell Checker \(uses MS Word\)


</div>

### Description

This code uses OLE Automation to allow VB to open an instance of MS Word if the user has it on their system and spell check the contents of a text box. It could easily be modified to work with any control that has text on it. I would recommend better error control than the On Error statement listed here.
 
### More Info
 
Just create the text box and command button listed in the code.

The code automatically replaces the original text of the message box with the corrected text.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[T\. Heinel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/t-heinel.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/t-heinel-spell-checker-uses-ms-word__1-1904/archive/master.zip)





### Source Code

```
Private Sub cmdSpellCheck_Click()
  'On Error Resume Next 'Best to un-comment this while testing
  Dim objMsWord As Word.Application
  Dim strTemp As String
  Set objMsWord = CreateObject("Word.Application")
  objMsWord.WordBasic.FileNew
  objMsWord.WordBasic.Insert txtMessage.Text
  objMsWord.WordBasic.ToolsSpelling
  objMsWord.WordBasic.EditSelectAll
  objMsWord.WordBasic.SetDocumentVar "MyVar", objMsWord.WordBasic.Selection
  objMsWord.Visible = False ' Mostly prevents Word from being shown
  strTemp = objMsWord.WordBasic.GetDocumentVar("MyVar")
  txtMessage.Text = Left(strTemp, Len(strTemp) - 1)
  objMsWord.Documents.Close (0) ' Close file without saving
  objMsWord.Quit         ' Exit Word
  Set objMsWord = Nothing    ' Clear object memory
  frmMain.SetFocus        ' Return focus to Main form
End Sub
```

