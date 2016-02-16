<%'紅光字
function a(fn1)
if info(2)<2000 then
	Response.Write "<script language=JavaScript>{alert('操作只有2000級以上官府才可以操作！');}</script>"
	Response.End
end if
a="<style type=text/css>.ww{line-height:240%}</style><div align='center'><span class=ww><font color=ffffff size=6 face=標楷體 style=width:100%;filter:glow(color=FF0000)><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></span></div>"

//a="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>