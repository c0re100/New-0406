<%'�����r 
function e(fn1)
if info(2)<5000 then
	Response.Write "<script language=JavaScript>{alert('�ާ@�u��5000�ťH�W�x���~�i�H�ާ@�I');}</script>"
	Response.End
end if
e="<style type=text/css>.ww{line-height:240%}</style><div align='center'><span class=ww><font style=width:100%;filter:glow(color=#8080ff,strength=50)><center><b>" & fn1 & "</center></font><font size=2>[" & info(0) & "]</font>" & "</b></font></span></div>"

//e="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>