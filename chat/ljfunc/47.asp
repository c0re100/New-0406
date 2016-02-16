<%'求婚
function qiuhun(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('求婚對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if

if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以求婚！');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('為了維護江湖的公正，官府人員不可以結婚！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 配偶,道德,魅力,銀兩 FROM 用戶 WHERE id="&info(9),conn
if rs("配偶")<>"無" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('江湖是一夫一妻制，你想作什麼！');}</script>"
	Response.End
end if
if rs("道德")<300 or rs("魅力")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('道德與魅力不夠300，人家看不上你的！');}</script>"
	Response.End
end if
if rs("銀兩")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('沒有5萬塊錢，不能登記結婚的！');}</script>"
	Response.End
end if
rs.close
rs.open "select 性別,配偶,等級,門派 FROM 用戶 WHERE 姓名='"&toname&"'" ,conn
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('為了維護江湖管理、公平競爭,官府禁止結婚！');}</script>"
	Response.End
end if
if rs("性別")=info(1) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('完了吧，江湖不許同性戀的！');}</script>"
	Response.End
end if
if rs("配偶")<>"無" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你想作什麼，第三者插足呀！');}</script>"
	Response.End
end if
if rs("等級")<=1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('取一個江湖1級的人，沒面子吧！');}</script>"
	Response.End
end if

conn.execute "update 用戶 set 銀兩=銀兩-50000 where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_hunyin")=toname
if info(1)="男" then
qiuhun="["&info(0)&"]向{"&toname&"}求婚：<img src='img/29.gif'>"&fn1&"  <input type=button value='我願意' onClick=javascript:;disabled=1;window.open('jiehun.asp?id="&info(9)&"&yn=1','d')><input type=button value='不願意' onClick=javascript:;disabled=1;window.open('jiehun.asp?id="&info(9)&"&yn=0','d')>"
else

qiuhun="["&info(0)&"]向{"&toname&"}求婚：<img src='img/girl.gif'>"&fn1&"  <input type=button value='我願意' onClick=javascript:;window.open('jiehun.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'><input type=button value='不願意' onClick=javascript:;window.open('jiehun.asp?id="&info(9)&"&yn=0','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end if
end function
%>
