<%'美容
function meirong()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 體力,等級,職業,魅力,魅力加,會員等級,操作時間 FROM 用戶 WHERE id="&info(9),conn
if left(rs("職業"),3)<>"水道士" and left(rs("職業"),3)<>"水法師" and left(rs("職業"),3)<>"水真人" and left(rs("職業"),3)<>"水天師" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的職業還不是水道士、水法師、水真人、水天師！');}</script>"
Response.End
end if 
if rs("體力")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的體力小於1000，不能美容！');}</script>"
Response.End
end if
if rs("等級")<15 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的等級小於15，怎麼美容！');}</script>"
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
ml=rs("魅力")
mlj=(rs("等級")*1120+2100)+rs("魅力加")
if rs("魅力")>=mlj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你現在好好的，不需要美容啊！');}</script>"
Response.End
end if
if hy=1 then 
 butl=mlj-ml
 blsq=10000
 else
 butl=mlj-ml
 blsq=30000
end if

conn.execute "update 用戶 set 魅力=魅力+"&butl&",體力=體力-"&blsq&",操作時間=now() where id="&info(9)
if rs("魅力")>mlj then
conn.execute "update 用戶 set 魅力=" & mlj & " where id="&info(9)
end if
if rs("體力")<0 then
conn.execute "update 用戶 set 體力=100 where id="&info(9)
end if
meirong="美容師"&info(0)& "施展絕技給自己擺弄了一番，使"& info(0) &"的魅力迅速恢復"& butl &"點!"&info(0)&"體力降低"&blsq&"！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
