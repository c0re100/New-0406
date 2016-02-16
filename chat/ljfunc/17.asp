<%'贈送
function zen(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('贈送對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
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
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以贈送！');}</script>"
	Response.End
end if
zswupin=zt(0)
wusl=abs(int(left(zt(1),2)))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id,物品名,類型,內力,體力,攻擊,防御,堅固度,銀兩,等級,說明,sm from 物品 where 數量>0 and 裝備<>true and 物品名='" & zswupin & "' and 擁有者="& info(9) &" and 數量>="&wusl,conn
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有那麼多的物品嗎？');}</script>"
	Response.End
end if
id=rs("id")
wupin=rs("物品名")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
gj=rs("攻擊")
fy=rs("防御")
jgd=rs("堅固度")
dj=rs("等級")
yin=rs("銀兩")
say=rs("說明")
sm=rs("sm")
conn.execute "update 物品 set 數量=數量-"&wusl&" where id="&id
rs.close
rs.open "select id from 用戶 where 姓名='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select id from 物品 where 數量>0 and 物品名='" & zswupin & "' and 擁有者=" & idd ,conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品 (物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,等級,說明,sm) values ('"&wupin&"','"&idd&"','"&lx&"',"&jgd&","&gj&","&fy&","&nl&","&tl&","&wusl&","&yin&","&dj&",'"&say&"',"&sm&")"
else
	id=rs("id")
conn.execute "update 物品 set 數量=數量+"&wusl&" where id="&id
end if
	zen=info(0) & "把自己的[" & zswupin & "]<img src='../hcjs/jhjs/001/"&sm&".gif'>贈送給了" & toname & "["&wusl&"]個，"&toname&"很是感謝！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>