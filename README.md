<div align="center">

## PathExists%


</div>

### Description

Determine if a directory exists. It is a variation of the Planet-Source-Code function I borrowed called 'FileExists that always works'.
 
### More Info
 
A path name to be checked (as string)

Directory length are size 0. But just even though checking to see if the file is zero length works, it isn't good enough. Some files can be 0 length also. Using GetAttr function with the mask vbDirectory ensures that what you're looking at is indeed a directory.

true if path if valid


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marshall Youngblood](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marshall-youngblood.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marshall-youngblood-pathexists__1-3184/archive/master.zip)

### API Declarations

API not used


### Source Code

```

Function PathExists(FullPath as string) as Boolean
'based on function borrowed from Planet Source Safe
  Dim blnDirectory As Boolean
  On Error Resume Next
  If FileLen(FullPath) = 0& Then
    If Err = 0 Then
      blnDirectory = (GetAttr(FullPath) And vbDirectory)
      If blnDirectory Then PathExists = True
    End If
  End If
End Function
```

