<%'放大藍字
function fangda(fn1)
if info(2)<9 then
	Response.Write "<script language=JavaScript>{alert('放大漢字操作只有9級以上官府才可以操作！');}</script>"
	Response.End
end if
fangda="<style type=text/css>.ww{line-height:240%}</style><div align='center'><span class=ww><font size=+4 color=blue><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></span></div>"

//fangda="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>