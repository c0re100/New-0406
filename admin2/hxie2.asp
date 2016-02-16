<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,會員費,內力,體力,攻擊,防御,會員等級,操作時間,內力加,體力加,攻擊加,防御加 from 用戶 where id="&info(9),conn
nlj=(rs("等級")*640+2000)+rs("內力加")
nla=rs("內力")
tlj=(rs("等級")*1500+5260)+rs("體力加")
tla=rs("體力")
gjj=(rs("等級")*1200+4500)+rs("攻擊加")
gja=rs("攻擊")
fyj=(rs("等級")*1120+3000)+Rs("防御加")
fya=rs("防御")
hy=rs("會員等級")
hf=rs("會員費")
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'hxie.asp';}</script>"
	Response.End
end if
if hy=0 then
Response.Write "<script language=JavaScript>{alert('你不是會員請回吧！');location.href = '../help/info.asp?type=2&name=會員加入辦法';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
else
cid=Request.QueryString("id")
select case cid
Case "1"
if nla>=nlj then
conn.execute("update 用戶 set 體力=體力-1000,內力="& nlj &",操作時間=now()  where id="&info(9))
else
conn.execute("update 用戶 set 內力=內力+1000,體力=體力-1000,操作時間=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('你用1000的體力換了1000的內力');location.href='hxie.asp';</script>"
Response.End
Case "2"
if tla>=tlj then
conn.execute("update 用戶 set 內力=內力-1000,體力="& tlj &",操作時間=now()  where id="&info(9))
else
conn.execute("update 用戶 set 內力=內力-1000,體力=體力+1000,操作時間=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('你用1000的內力換了1000的體力');location.href='hxie.asp';</script>"
Response.End
Case "3"
if fya>=fyj then
conn.execute("update 用戶 set 攻擊=攻擊-1000,防御="& fyj &",操作時間=now()  where id="&info(9))
else
conn.execute("update 用戶 set 攻擊=攻擊-1000,防御=防御+1000,操作時間=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('你用1000的攻擊換了1000的防御');location.href='hxie.asp';</script>"
Response.End
Case "4"
if gja>=gjj then
conn.execute("update 用戶 set 防御=防御-1000,攻擊="& gjj &",操作時間=now()  where id="&info(9))
else
conn.execute("update 用戶 set 防御=防御-1000,攻擊=攻擊+1000,操作時間=now()   where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('你用1000的防御換了1000的攻擊');location.href='hxie.asp';</script>"
Response.End
Case "5"
if rs("會員費")>=1 then
jb=int(abs(hf)*100)
conn.execute("update 用戶 set 會員費=0,金幣="&jb&",操作時間=now()  where id="&info(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('會費轉換成金幣成功，請注意查收!');location.href='hxie.asp';</script>"
Response.End
else

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('你還有會費嗎？？？');location.href='hxie.asp';</script>"
Response.End
end if
Case else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('錯誤類型');location.href='hxie.asp';</script>"
Response.End
end select
end if%>
