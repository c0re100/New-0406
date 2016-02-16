<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>2000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('你想作什麼？！');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間,銀兩 from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	s=8-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起系統限制，請等["&s&"分鐘]再操作！');}</script>"
	Response.End
end if	
if rs("銀兩")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('您的銀兩不夠呀，雖然是師徒，不過沒有錢是萬萬不行的！');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
conn.execute "update 用戶 set 武功=武功+"&int(money/1000)*30&",銀兩=銀兩-"& money &",操作時間=now() where id="&info(9)
rs.close
rs.open "select 等級,武功,武功加 from 用戶 where id="&info(9),conn
wgj=(rs("等級")*1280+3800)+rs("武功加")
if rs("武功")>wgj then
conn.execute "update 用戶 set 武功=" & wgj & ",操作時間=now() where id="&info(9)
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('你與師傅在密室中苦心學習！終於使自己的武功大進,武功：+"& ((money/1000)*30) &"點，花費銀兩：-"&money&"兩！');location.href = 'javascript:history.back()';;</script>"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">