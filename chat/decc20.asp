<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
myname=info(0)
if info(0)<>"江湖總站" then
	Response.Write "<script Language=Javascript>alert('提示：你來這裡作什麼，想死呀！');window.close();</script>"
	response.end
end if
if info(2)<10 then 
	Response.Write "<script Language=Javascript>alert('提示：你來這裡作什麼，想死呀！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from 物品 where id=" & id & " and 數量>=0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：沒有你要增加的物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if 
conn.execute  "Update 物品 Set 數量=數量+20 WHERE id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：增加物品ID"& id &"成功');location.href = 'javascript:history.go(-1)';</script>"
%>
 