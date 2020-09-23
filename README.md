<div align="center">

## Open website with default web browser


</div>

### Description

Open your default web browser with any url.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Boiko](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-boiko.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-boiko-open-website-with-default-web-browser__1-37987/archive/master.zip)





### Source Code

```
Create a new module and put this code into it <p>
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long<p>
Public Const conSwNormal = 1
Put a command button on the form and put this code into it Private Sub Command1_Click()
ShellExecute hwnd, "open", "http://www.google.com", vbNullString, vbNullString, conSwNormal
End Sub
```

