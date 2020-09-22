<div align="center">

## Change Windows Caption


</div>

### Description

This Program Changes the Caption of > ALMOST < any windows program!
 
### More Info
 
Surprizingly Easy... Just Type the text of the Windows Program that you wish to change then type the text you want to change it too!

There is no Side Effects to This Program


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dillon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dillon.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dillon-change-windows-caption__1-25986/archive/master.zip)

### API Declarations

```
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
```


### Source Code

```
'make 2 text boxes
'Name them Text1 - For The Windows Caption
'And Text2 - For the New WIndows Caption
'Make 1 Button
'Name it Command1
Private Sub Command1_Click()
Dim temp As Long
temp = FindWindow(vbNullString, Text1.Text)
SetWindowText temp, Text2.Text
End Sub
```

