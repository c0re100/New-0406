<%'清場
function kakaka(fn1)
if info(2)<20000000 then
	Response.Write "<script language=JavaScript>{alert('操作只有20000000級以上官府才可以操作！');}</script>"
	Response.End
end if
kakaka="<script>parent.f3.location.href='AUTOKICK.ASP';</script>"
//fangda="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>