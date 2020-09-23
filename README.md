<div align="center">

## GetPathSize


</div>

### Description

Very fast recursive function to calculate the size of a directory (folder). This code is simple and built for speed. This code does NOT use the FileSystemObject because it is NOT installed on all PCs, therefore you cannot write code using it (unless you include the scrrun.dll - Microsoft Scripting Runtime with your application). **Update - I added the search options for System, Hidden, and Read-Only files so the result will truely match the same number of bytes that is displayed in Windows Explorer properties.
 
### More Info
 
sPathName - The path to the directory

The size (number of bytes) as a Double


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert E\. Phelps](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-e-phelps.md)
**Level**          |Intermediate
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-e-phelps-getpathsize__1-25803/archive/master.zip)





### Source Code

```
Public Function GetPathSize(ByRef sPathName As String) As Double
 Dim sFileName As String
 Dim dSize As Double
 Dim asFileName() As String
 Dim i As Long
 ' Enumerate DirNames and FileNames
 If StrComp(Right$(sPathName, 1), "\", vbBinaryCompare) <> 0 Then sPathName = sPathName & "\"
 sFileName = Dir$(sPathName, vbDirectory + vbHidden + vbSystem + vbReadOnly)
 Do While Len(sFileName) > 0
  If StrComp(sFileName, ".", vbBinaryCompare) <> 0 And StrComp(sFileName, "..", vbBinaryCompare) <> 0 Then
   ReDim Preserve asFileName(i)
   asFileName(i) = sPathName & sFileName
   i = i + 1
  End If
  sFileName = Dir
 Loop
 If i > 0 Then
  For i = 0 To UBound(asFileName)
   If (GetAttr(asFileName(i)) And vbDirectory) = vbDirectory Then
    ' Add dir size
    dSize = dSize + GetPathSize(asFileName(i))
   Else
    ' Add file size
    dSize = dSize + FileLen(asFileName(i))
   End If
  Next
 End If
 GetPathSize = dSize
End Function
```

