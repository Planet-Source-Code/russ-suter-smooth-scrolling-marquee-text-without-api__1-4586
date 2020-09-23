<div align="center">

## Smooth scrolling marquee text \(without API\)


</div>

### Description

With this code you can scroll text smoothly in any direction without needing any API calls. All you need is a PictureBox control, a Label control, and a timer.

The code is fully commented and has examples for scrolling in the four main directions.
 
### More Info
 
Create a new project. On the main form create a PictureBox control named "pbScrollBox", a Label control named "lblText", and a Timer named "tmrScroll".

Make sure the PictureBox control's AutoRedraw property is set to True (to prevent flicker), and the ScaleMode property is set to Pixels.

The PictureBox control's ForeColor property determines the color of the text.

Make sure the Label control's AutoSize property is set to True. Also make sure the Label control is contained within the PictureBox control.

Make sure the PictureBox control and Label control's Font property are set identically.

The code gets its geometry information from the label size.

If the Font properties don't match and/or the AutoSize property is not set, the text might not wrap properly.

If you have other timers running concurrently, you might get an occasional stutter while the other timer processes its code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Russ Suter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/russ-suter.md)
**Level**          |Unknown
**User Rating**    |4.3 (170 globes from 40 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/russ-suter-smooth-scrolling-marquee-text-without-api__1-4586/archive/master.zip)





### Source Code

```
Option Explicit
Private TheX as Long
Private TheY as Long
' I have included commented lines to scroll in any of the four main directions.
' You can uncomment the appropriate lines for your needs.
' If you uncomment both the left to right and bottom to top, for example,
' you get diagonal scrolling.
' I found that a timer interval of 50 milliseconds works well in most cases.
' Windows 95/98 machines don't get any faster from that point but NT machines do.
' Playing with the Timer's interval property as well as adjusting the number of
' pixels to step by will eventually satisfy your needs.
Private Sub Form_Load()
  lblText.Caption = "Insert your credits here..." ' Set the text to be shown
  ' Use this line of code if you want to scroll right to left
  TheX = pbScrollBox.ScaleWidth ' Set the starting point (off the right edge)
  ' Use this line of code if you want to scroll left to right
'  TheX = 0 - lblText.Width ' Set the starting point (off the left edge)
  ' Use this line of code if you want to scroll bottom to top
'  TheY = pbScrollBox.ScaleHeight ' Set the starting point (off the bottom edge)
  ' Use this line of code if you want to scroll top to bottom
'  TheY = 0 - lblText.Height ' Set the starting point (off the top edge)
End Sub
Private Sub tmrScroll_Timer()
  pbScrollBox.Cls ' so we don't get text trails
  ' Scroll from right to left
  If TheX <= 0 - lblText.Width Then
    TheX = pbScrollBox.ScaleWidth
  Else
    TheX = TheX - 1 ' larger number means faster scrolling
  End If
  ' uncomment the following lines to scroll from left to right
'  If TheX >= pbScrollBox.ScaleWidth Then
'    TheX = 0 - lblText.Width
'  Else
'    TheX = TheX + 1
'  End If
  ' uncomment the following lines to scroll from bottom to top
'  If TheY <= 0 - lblText.Height Then
'    TheY = pbScrollBox.ScaleHeight
'  Else
'    TheY = TheY - 1
'  End If
  ' uncomment the following lines to scroll from top to bottom
'  If TheY >= pbScrollBox.ScaleHeight Then
'    TheY = 0 - lblText.Height
'  Else
'    TheY = TheY + 1
'  End If
  ' set the text position and print the text
  pbScrollBox.CurrentX = TheX
  pbScrollBox.CurrentY = TheY
  pbScrollBox.Print lblText.Caption
End Sub
```

