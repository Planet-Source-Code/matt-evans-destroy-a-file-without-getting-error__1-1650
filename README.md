<div align="center">

## \*\*\* Destroy a file without getting error\! \*\*\*


</div>

### Description

This DOES use the kill function, but when you use this it actually opens the file you want to destroy, cleans it out, then deletes it so you don't get any error because of sensitive data!
 
### More Info
 
File To Delete

Make sure not to delete win.com, win.ini, config.sys, himem.sys, autoexec.bat or any of those files, lol ;D

No more file


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Evans](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-evans.md)
**Level**          |Unknown
**User Rating**    |4.0 (36 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-evans-destroy-a-file-without-getting-error__1-1650/archive/master.zip)





### Source Code

```
Sub DestroyFile(sFileName As String)
  Dim Block1 As String, Block2 As String, Blocks As Long
  Dim hFileHandle As Integer, iLoop As Long, offset As Long
  'Create two buffers with a specified 'wipe-out' characters
  Const BLOCKSIZE = 4096
  Block1 = String(BLOCKSIZE, "X")
  Block2 = String(BLOCKSIZE, " ")
  'Overwrite the file contents with the wipe-out characters
  hFileHandle = FreeFile
  Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
    For iLoop = 1 To Blocks
      offset = Seek(hFileHandle)
      Put hFileHandle, , Block1
      Put hFileHandle, offset, Block2
    Next iLoop
  Close hFileHandle
  'Now you can delete the file, which contains no sensitive data
  Kill sFileName
End Sub
```

