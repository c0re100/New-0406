<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 等級 FROM 用戶 WHERE id=" & info(9),conn
if rs("等級")<85 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的等級不夠煆造此武器！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=30 and 物品名='藍寶石'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的藍寶石不夠30個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id1=rs("id")
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=180 and 物品名='礦石'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的礦石不夠180個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id2=rs("id")
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=180 and 物品名='特種金屬'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的特種金屬不夠180個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id3=rs("id")
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=180 and 物品名='黑石'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的黑石不夠180個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id4=rs("id")
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=180 and 物品名='水晶'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的水晶不夠180個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id5=rs("id")
rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>=180 and 物品名='鐵片'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的鐵片不夠180個不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id6=rs("id")
conn.Execute "update 物品 set 數量=數量-30 WHERE id="&id1
conn.Execute "update 物品 set 數量=數量-180 WHERE id="&id2
conn.Execute "update 物品 set 數量=數量-180 WHERE id="&id3
conn.Execute "update 物品 set 數量=數量-180 WHERE id="&id4
conn.Execute "update 物品 set 數量=數量-180 WHERE id="&id5
conn.Execute "update 物品 set 數量=數量-180 WHERE id="&id6
rs.close
rs.open "select id from 物品 where 物品名='青頭魔劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,sm,數量,等級) values('青頭魔劍',"&info(9)&",'右手',0,0,0,1800000,500000,2003209,2003209,1,85)")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("青頭魔劍煆造完成！")
history.back()
</script>
