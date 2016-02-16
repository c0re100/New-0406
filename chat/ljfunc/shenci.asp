<%
'靈劍神思
function shenci()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 攻擊加,道德,等級,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("道德")<100  then
Response.Write "<script language=JavaScript>{alert('你的道德不夠，需要道德100,才能去思過崖神思！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("攻擊加")<rs("等級")*1000  then
conn.execute "update 用戶 set 攻擊加=攻擊加+5,道德=道德-100,操作時間=now() where id="&info(9)
shenci=info(0) & "在思過崖神思,一陣胡思亂想後，突有所悟，失去<font color=red>100</font>點道德,攻擊上限提升<font color=red>+5</font>點,學無止境，總算有所收獲啊!"
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
	conn.close
	set conn=nothing
	
end function
%>