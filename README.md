<div align="center">

## Match\(\) \- Check any string for any number of characters\.


</div>

### Description

Simple function to validate string contents. Compares a given string to a list of illegal values and evaluates whether or not it contains any. Very fast and easy. Can also be used as a string search function with a little modification.
 
### More Info
 
strSource - String you are checking

strCompare - String of illegal characters

To run:

If Match (strMyFile, "~!@#$%^&*()+`{}[]?><,/") Then

MsgBox "File contains illegal characters!", vbExclaimation

End If

.

Boolean - True = Source string contains an illegal character.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-match-check-any-string-for-any-number-of-characters__1-14371/archive/master.zip)





### Source Code

```

Public Function Match(strSource As String, strCompare As String) As Boolean
Dim lngCheck As Long
 For lngCheck = 1 To Len(strCompare)
 If InStr(strSource, Mid(strCompare, lngCheck, 1)) Then
 Match = True
 Exit Function
 End If
 Next lngCheck
End Function
```

