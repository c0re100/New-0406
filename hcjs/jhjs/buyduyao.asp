<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../images/7.jpg">
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
sl=int(abs(Request.form("sl")))
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 類型,物品名,銀兩,內力,體力,攻擊,防御,堅固度,sm,等級 FROM 物品買 where ID=" & ID,conn
if rs("類型")<>"毒藥" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wu=rs("物品名")
yin=rs("銀兩")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
sm=rs("sm")
gj=rs("攻擊")
fy=rs("防御")
jgd=rs("堅固度")
dj=rs("等級")
if info(4)>1 then
	yin=int(yin/2)
end if
rs.close
rs.open "select 銀兩,操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if yin*sl>rs("銀兩") then
	Response.Write "<script Language=Javascript>alert('提示：購買不成功，原因：你的銀兩不夠了');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",操作時間=now() where id="&info(9)
rs.close
'物品
rs.open "select id from 物品 where 物品名='" & wu & "' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
	sql="insert into 物品(物品名,擁有者,類型,內力,體力,攻擊,防御,堅固度,數量,銀兩,說明,sm,等級,會員) values ('"&wu&"',"&info(9)&",'"&lx&"',"&nl&","&tl&","&gj&","&fy&","&jgd&","&sl&","&int(yin*2)&",'無',"&sm&","&dj&","&aaao&")"
	conn.execute sql
else
	id=rs("id")
	sql="update 物品 set 數量=數量+" & sl & ",會員="&aaao&" where id="&id
	conn.execute sql
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if info(4)>1 then
	Response.Write "<script Language=Javascript>alert('提示：2級以上會員購買物品打5折,購買"&wu&sl&"個成功！');location.href = 'javascript:history.go(-1)';</script>"
response.end
else
	Response.Write "<script Language=Javascript>alert('提示：購買"&wu&"物品"&sl&"個成功,如果您是2級以上會員可以打5折！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
%>
