<%'醫生
function yiliao()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 內力,職業,體力,會員等級,等級,體力加,內力加,操作時間 FROM 用戶 WHERE id="&info(9),conn
nl=rs("內力")
tl=rs("體力")
if left(rs("職業"),2)<>"道士" and left(rs("職業"),3)<>"水道士" and left(rs("職業"),3)<>"水法師" and left(rs("職業"),3)<>"水真人" and left(rs("職業"),3)<>"水天師" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的職業還不是道士類怎麼治療！');}</script>"
Response.End
end if 
if rs("內力")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的內力小於1000，無法治療！');}</script>"
Response.End
end if
if rs("等級")<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的等級小於20，無法治療！');}</script>"
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
tl=rs("體力")
tlj=(rs("等級")*1500+5260)+rs("體力加")
nlj=(rs("等級")*640+2000)+rs("內力加")
if rs("體力")>=tlj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你現在身體好好的，不需要治療啊！');}</script>"
Response.End
end if
if hy=1 then 
 butl=tlj-tl
 blsq=1000
 else
 butl=tlj-tl
 blsq=5000
end if

conn.execute "update 用戶 set 體力=體力+"&butl&",內力=內力-"&blsq&",操作時間=now() where id="&info(9)
if tl>tlj then
conn.execute "update 用戶 set 體力=" & tlj & " where id="&info(9)
end if
if nl<0 then
 conn.execute "update 用戶 set 內力=100 where id="&info(9)
end if

yiliao="神醫"&info(0)& "對自己使用了治療術，"&info(0)&"的體力迅速恢復"& butl &"點!"&info(0)&"損耗"&blsq&"的內力！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
