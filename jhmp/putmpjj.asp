<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<title>門派基金</title>
<body leftmargin="0" topmargin="0" bgcolor="#660000">
<%if info(0)="" then Response.Redirect "../error.asp?id=210"
my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then
 	Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,銀兩 from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if trim(rs("門派"))="遊俠" or trim(rs("門派"))="無" or trim(rs("門派"))="未知"  then
	Response.Write "<script Language=Javascript>alert('你是遊俠，還沒有自己的門派，你想搞什麼呀？');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if rs("銀兩")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的錢好像不夠呀！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-"& money &",道德=道德+"&int(money/500)&",門派基金=門派基金+"&money &",操作時間=now() where id="&info(9)
conn.execute "update 門派 set 幫派基金=幫派基金+"& money &" where 門派='" & info(5) &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('您向你的門派中捐獻了"& money &"兩，道德上漲："&int(money/500)&"點幫中的弟子對你感激不盡！');location.href = 'javascript:history.back()';</script>"
%>
