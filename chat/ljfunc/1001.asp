<%'綠光字
function b(fn1)
if info(2)<3000 then
	Response.Write "<script language=JavaScript>{alert('放大漢字操作只有3000級以上官府才可以操作！');}</script>"
	Response.End
end if
b="<style type=text/css>.ww{line-height:240%}</style><div align='center'><span class=ww><font color=ffffff size=6 face=標楷體 style=width:100%;filter:glow(color=00FF33)><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></span></div>"

//b="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>