<%
function qiangjielin(fn1,to1,toname)
if toname="大家" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('搶劫令對像有錯，請看仔細了！！');}</script>"
Response.End
exit function
end if
if info(3)<10 then
Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以使用搶劫令！');}</script>"
Response.End
end if
if info(2)>6 then
Response.Write "<script language=JavaScript>{alert('官府人員不可操作！');}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力 FROM 用戶 WHERE id="&info(9),conn
if rs("法力")<10000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得10000點啊！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select 門派,存款 FROM 用戶 WHERE 姓名='"&toname&"'",conn
money=int(rs("存款")/5)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
if rs("門派")="官府"  then
Response.Write "<script language=JavaScript>{alert('失敗，你不能對高級管理員或站長特封的人員操作!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE 物品名='搶劫令' and 數量>0 and 擁有者="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有搶劫令嗎？');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 銀兩=200000,數量=數量-1 where id="& id &" and 擁有者="&info(9)
conn.execute "update 用戶 set 法力=法力-10000 where id="&info(9)
conn.execute "update 用戶 set 銀兩=銀兩+"&moneyc&" WHERE  id="&info(9)
conn.execute "update 用戶 set 存款=存款-"&money&" where 姓名='"&toname&"'"
qiangjielin=info(0) & "拿出法器<bgsound src=wav/Phant012.wav loop=1>搶劫令,對著<font color=red>"&towho&"</font>大吼一聲:搶劫  ,從"&towho&"那搶得存款"&moneyc&"兩,"&info(0)&"消耗法力3000點" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>