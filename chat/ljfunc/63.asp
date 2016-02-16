<%'軍官
function junguan()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,職業,道德,武功,武功加,會員等級,操作時間 FROM 用戶 WHERE id="&info(9),conn
if left(rs("職業"),3)<>"水真人" and left(rs("職業"),3)<>"水天師" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的職業還不是水真人、水天師怎麼恢復武功！');}</script>"
Response.End
end if 
if rs("道德")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的道德小於1000，進入不了火雲洞！');}</script>"
Response.End
end if
if rs("等級")<40 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的等級小於40，進入不了火雲洞！');}</script>"
Response.End
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("會員等級")>0 then 
 hy=1
else
 hy=0
 end if
wgj=(rs("等級")*1280+3800)+rs("武功加")
wg=rs("武功")
if rs("武功")>=wgj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你武功很高了啊，不需要再提升了啊！');}</script>"
Response.End
end if
if hy=1 then 
 butl=wgj-wg
 blsq=10000
 else
 butl=wgj-wg
 blsq=50000
end if

conn.execute "update 用戶 set 武功=武功+"&butl&",道德=道德-"&blsq&",操作時間=now() where id="&info(9)
if rs("武功")>wgj then
conn.execute "update 用戶 set 武功=" & wgj & " where id="&info(9)
end if
if rs("道德")<0 then
conn.execute "update 用戶 set 道德=100 where id="&info(9)
end if

junguan="軍官"&info(0)& "為了提升自己的武功，不惜冒著走火入魔的危險進入火雲洞，終於成功恢復武功"& butl &"點!"&from1&"道德降低"&blsq&"！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
