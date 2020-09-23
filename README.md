<div align="center">

## Caption Scroller \- IMPROVED\!


</div>

### Description

This Code will Scroll the title

of your window(s) like Winamp

To the left OR to the right.
 
### More Info
 
Just create a Timer called

TIMER1 and plaste this code

in the TIMER1_TIMER sub.

The Variable "Direction"

needs to contain one of

two values:"Right" or

"Left"

You May need to change the

INTERVAL Property of the timer

as it defaults to Zero.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sparq](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sparq.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sparq-caption-scroller-improved__1-7791/archive/master.zip)





### Source Code

```
'Code provided by Alpha Media Inc.
'http://www.alphamedia.net
'Makers of Pink Notes Plus!
'http://www.pinknotesplus.com
Private Sub Timer1_Timer()
 Dim String2 As String
 Dim String1 As String
 If Direction = "Left" Then
  String2 = Left$(Caption, 1)
  String1 = Right$(Caption, Len(Caption) - 1)
 ElseIf Direction = "Right" Then
  String1 = Right$(Caption, 1)
  String2 = Left$(Caption, Len(Caption) - 1)
 End If
 Caption = String1 & String2
End Sub
```

