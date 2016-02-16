<%'公告
function gong(fn1)
if info(2)<9 then
	Response.Write "<script language=JavaScript>{alert('你不是官府中人，所以不能操作！');}</script>"
	Response.End
end if
fn1=trim(fn1)
if info(5)<>"官府" and info(5)<>"站長" then
gong="<table bgcolor=red cellpadding=4 cellspacing=4 align=center><tr><td style='border:2px dashed white;'><table bgcolor=red cellspacing=1 cellpadding=4><tr bgcolor=FFFFFF><td colspan=2><div align=center><font color=000000 size=4>【官府公告"&info(0)&"】</font></td></tr><tr bgcolor=FFFF00><td><font color=000000>"&fn1&"</font></div></td></tr></table></td></tr></table>"
'gong="<table bgcolor=FFFF00 width=70% align=center border=1 cellspacing=0 cellpadding='0'><tr><td bgcolor=00FF00 align=center>官 府 公 告</td><tr><td><div align='center'><font size=-1>"&fn1&info(0)&"[便衣聊管]</font></div></td></tr></table>"
else
if info(5)="站長" then
gong="<table bgcolor=red cellpadding=4 cellspacing=4 align=center><tr><td style='border:2px dashed white;'><table bgcolor=red cellspacing=1 cellpadding=4><tr bgcolor=FFFFFF><td colspan=2><div align=center><font color=000000 size=4>【官府公告"&info(0)&"】</font></td></tr><tr bgcolor=FFFF00><td><font color=000000>"&fn1&"</font></div></td></tr></table></td></tr></table>"
else
gong="<table bgcolor=red cellpadding=4 cellspacing=4 align=center><tr><td style='border:2px dashed white;'><table bgcolor=red cellspacing=1 cellpadding=4><tr bgcolor=FFFFFF><td colspan=2><div align=center><font color=000000 size=4>【官府公告"&info(0)&"】</font></td></tr><tr bgcolor=FFFF00><td><font color=000000>"&fn1&"</font></div></td></tr></table></td></tr></table>"
end if
end if
end function
%>
