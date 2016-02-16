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
yilao=request.form("yilao")
if yilao<>"無" then
'校驗用戶
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,體力,體力加,銀兩 from 用戶 where id="&info(9),conn
tlj=(rs("等級")*1500+5260)+rs("體力加")
if rs("體力")>tlj then
Response.Write "<script language=JavaScript>{alert('你的體力已經滿了不需要再補！');location.href = 'diyu2.asp';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
If Rs.Bof OR Rs.Eof Then
mess="你不能進行治療"
else
sm=rs("體力")
select case yilao
case "一級煉火"
bl=1.5
cd=10000-sm
case "二級煉火"
bl=1
cd=10000-sm
case "三味真火"
bl=0.5
cd=10000-sm
end select
fy=int(cd/bl)
if fy>rs("銀兩") then
mess="你已經窮的連救命錢都沒有了。還想煉火哦？！"
else
conn.execute "update 用戶 set 體力=10000,銀兩=銀兩-" & fy & " where id="&info(9)
mess="經過地獄的無情之火煆燒，你花了" & fy & "兩銀子生命恢復到10000點"
end if
end if
else
mess="你需要煉獄了"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'diyu.asp';}</script>"
response.end%>
