<%
'靈劍參禪
function canchan()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 防御加,魅力,等級,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("魅力")<100  then
Response.Write "<script language=JavaScript>{alert('你的魅力不夠，需要魅力100,才能在山洞中參禪！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("防御加")<rs("等級")*1000  then
conn.execute "update 用戶 set 防御加=防御加+5,魅力=魅力-100,操作時間=now() where id="&info(9)
canchan=info(0)& "在山洞獨自參禪,微風陣陣，突有所悟，失去<font color=red>100</font>點魅力,防御上限提升<font color=red>+5</font>點,學無止境，總算有所收獲啊!"
else
	Response.Write "<script language=JavaScript>{alert('現在你的上限滿了，等升了級再練吧');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
set conn=nothing
end function
%>