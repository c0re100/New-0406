<HTML><HEAD><TITLE>社區時間</TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<SCRIPT language=JScript>
<!--
if (window.top.frames.length!=0 && window.top.frames[0].ShowNoButtons!=null)
    window.top.frames[0].ShowNoButtons();
//-->
</SCRIPT>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY link=#0033cc bgColor=#000000 leftMargin=20>
<p align="center">
<FONT color=white size=2>
<OBJECT id=DAControl hspace=20 
classid=CLSID:B6FFC24C-7E13-11D0-9B47-00C04FC2F51D align=left width=220 
height=220>
  <param name="OpaqueForHitDetect" value="1">
  <param name="UpdateInterval" value=".033"></OBJECT>
</FONT><FONT ir
face="Verdana, Arial, Helvetica" color=white size=2>
<P>  </P>
</FONT>
<SCRIPT language=VBScript>
<!--
  Set m = DAControl.PixelLibrary
  pi = 3.14159265359

  Sub window_onLoad
    'Get the current time and break it down into hours, minutes, and seconds.
    a = time
    min1 = minute(time)
    hr1 = hour(time)
    sec1 = second(time)
    Set xPos = m.Mul(m.DANumber(150), m.Cos(m.Mul(m.LocalTime,m.DANumber(0.3))))
    Set yPos = m.Mul(m.DANumber(35), m.Cos(m.Mul(m.LocalTime,m.DANumber(0.5))))

    'Create the final clock image.
    Set clock =  m.Overlay(hands(hr1,min1,sec1),face())

    'Display the clock.
    DAControl.Image = clock

    'Set the background in case of a non-windowless browser (such as IE3).
    DAControl.BackgroundImage = m.SolidColorImage(m.Blue)

    'Start the animation.
    DAControl.Start
  End Sub
  
  Function face()
    'Create the background for the clock.
    Set fs1 = m.Font("Verdana",14,m.Yellow).Bold()
	Set lineStyle = m.DefaultLineStyle.Color(m.Black)
    Set fillStyle = m.SolidColorImage(m.ColorRGB(64,64,255))

    Set txtPath1 = m.StringPathAnim(m.DAString(""), fs1)
    Set textpcs1 = txtPath1.fill(lineStyle, fillStyle).Transform(m.Translate2(-10,-30))

    Set txtPath2 = m.StringPathAnim(m.DAString(""), fs1)
    Set textpcs2 = txtPath2.fill(lineStyle, fillStyle)

    Set txtPath3 = m.StringPathAnim(m.DAString(""), fs1)
    Set textpcs3 = txtPath3.fill(lineStyle, fillStyle).Transform(m.Translate2(10,30))

    Set fgColor = m.Red
    Set bgColor = m.Blue

    Set bgFill= m.RadialGradientRegularPoly(fgColor,bgColor,50,2)

	Set bgFill = bgFill.Transform(m.Scale2Uniform(0.055))

    Set background = m.Oval(200,200).Fill(m.DefaultLineStyle,bgFill)

    Set text = m.Overlay(textpcs1,m.Overlay(textpcs2,textpcs3))

    'Create the numbers for the clock.
     Set fs2 = m.Font("Verdana",12,m.White).Bold()
     For i = 1 To 12
       Set vec = m.Vector2(82.5,0).Transform(m.Rotate2(-pi/6*(i-3)))
       Set text = m.Overlay(m.StringImage(i,fs2).Transform(m.Compose2(m.Translate2Vector(vec), _
	     m.Scale2Uniform(1.5))),text)
     Next 

    Set text = text.Transform(m.Translate2(1,9))

    'Put the numbers on top of the background.
     Set face = m.Overlay(text,background)

  End Function

  Function hands(hr,min,sec)
    'Create the hour, minute and second hands of the clock.
    Set bvr60 = m.DANumber(60)
    Set secFromMidnight = m.Add(m.DANumber(hr*3600+min*60+sec),m.LocalTime)
    Set secBvr = m.Mod(secFromMidnight,bvr60)
    Set minBvr = m.Mod(m.Div(secFromMidnight,bvr60),bvr60)
    Set hrBvr = m.Mod(m.Div(secFromMidnight,m.DANumber(3600)),m.DANumber(12))
    
	ptsSec = Array( -5, -2, 90, -2, 90, 2, -5, 2 )
    ptsMin = Array(-5, -3, 80, -3, 80, 3, -5, 3 )
    ptsHr = Array(-5, -5, 65, -3, 65, 3, -5, 3 )

    Set temp1 = m.Mul(m.DANumber(-pi/30),m.Sub(secBvr,m.DANumber(15)))
    Set temp2 = m.Mul(m.DANumber(-pi/30),m.Sub(minBvr,m.DANumber(15)))
    Set temp3 = m.Mul(m.DANumber(-pi/6),m.Sub(hrBvr,m.DANumber(3)))

    Set imgSec = m.PolyLine(ptsSec).Fill(m.DefaultLineStyle,m.SolidColorImage(m.White))
    Set imgSec = imgSec.TransForm(m.Rotate2Anim(temp1))

    Set imgMin = m.PolyLine(ptsMin).Fill(m.DefaultLineStyle,m.SolidColorImage(m.Red))
    Set imgMin = imgMin.TransForm(m.Rotate2Anim(temp2))

    Set imgHr = m.PolyLine(ptsHr).Fill(m.DefaultLineStyle,m.SolidColorImage(m.Magenta))
    Set imgHr = imgHr.TransForm(m.Rotate2Anim(temp3))
	
    Set hands = m.Overlay(imgSec,m.Overlay(imgMin,imgHr))

  End Function
-->
</SCRIPT>
</BODY></HTML>
 