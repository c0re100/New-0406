<%
'靈劍修養
function xiuyang()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 魅力,知質,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
Response.Write "<script language=JavaScript>{alert('你的魅力不夠，需要魅力100,才能修養！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 知質=知質+10,魅力=魅力-100,操作時間=now() where id="&info(9)
xiuyang=info(0)& "在療養院修養，看遍天下奇書收獲不少，失去<font color=red>100</font>點魅力,知質提升<font color=red>+10</font>點,真是聰明人啊!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>