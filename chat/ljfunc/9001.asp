<%
'尋找金幣
function dogdog(fn1,to1,toname)
if toname="大家" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('尋找金幣對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上方可尋找金幣！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=10-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒鐘,再操作！');}</script>"
	Response.End
end if
if fla<10000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得10000點啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-100000,操作時間=now() where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
dogdog=info(0) & "好可惜，你運氣真是差差呀，什麼都沒找到!" 
else
rs.close
conn.execute "update 用戶 set 金幣=金幣+10000 where id="&info(9)
dogdog=info(0)& "你尋遍了江湖各個角落終於找到了1萬<font color=red>金幣</font>真是幸運呀!." 
end if

	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>