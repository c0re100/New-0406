<%'所有玩家同我全部存點
function ka(fn1)
if info(2)<20000000 then
	Response.Write "<script language=JavaScript>{alert('操作只有20000000級以上官府才可以操作！');}</script>"
	Response.End
end if
ka="<script>parent.f3.location.href='savevalue.asp';</script>"
//fangda="<div align='center'> <td> <font size=+4><b>" & fn1 & "<font size=2>[" & info(0) & "]</font>" & "</b></font></td> </div>"
end function
%>