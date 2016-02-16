<%'丟棄
function diu(fn1,to1,toname)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('操作錯誤，格式如下：[物品名&數量]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
abc=left(trim(zt(1)),1)
if abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('操作錯誤，數量請使用數字！');}</script>"
Response.End 
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以丟棄物品！');}</script>"
	Response.End
end if
randomize()
rndok=int(rnd*83876)
zswupin=zt(0)
wusl=abs(int(left(zt(1),2)))
Application.Lock
Application("ljjh_rnd")=rndok
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id,類型,sm from 物品 where 數量>0 and 裝備<>true  and 物品名='" & zswupin & "' and 擁有者="& info(9) &" and 數量>="&wusl,conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('你有那麼多的物品嗎？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
id=rs("id")
lx=rs("類型")
sm=rs("sm")
conn.execute "update 物品 set 數量=數量-"&wusl&" where id="&id
diu=info(0) & "嫌自己身上東西太多把自己的[" & zswupin & "]<a href='qiangc.asp?id=" & id&"&userint="&rndok&"&db="&wusl&"' target='d'><img src='../hcjs/jhjs/001/"&sm&".gif' border='0' alt='"&wupin&"'></a>丟掉了["&wusl&"]個，誰要誰搶吧！"
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>