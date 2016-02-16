<%
'皌腳ホ
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select 单 FROM ノめ WHERE id="&info(9),conn
if rs("单")<10 then
Response.Write "<script language=JavaScript>{alert('惠璶10驹矮单');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 珇  WHERE 珇='腳ホ'   and 计秖>0 and 局Τ="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('Τ厚屡腳ホ盾');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 珇  WHERE  珇='厚腳ホ'  and 计秖>0 and 局Τ="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('Τ厚屡腳ホ盾');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM 珇  WHERE  珇='屡腳ホ'  and 计秖>0 and 局Τ="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('Τ厚屡腳ホ盾');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 珇 set 计秖=计秖-1 where 珇='腳ホ'  and 局Τ="&info(9)
conn.execute "update 珇 set 计秖=计秖-1 where 珇='厚腳ホ'  and 局Τ="&info(9)
conn.execute "update 珇 set 计秖=计秖-1 where 珇='屡腳ホ'  and 局Τ="&info(9)
rs.close
rs.open "select id FROM 珇  WHERE 珇='臸苝ホ'  and 计秖>0 and 局Τ="&info(9)
If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 珇(珇,局Τ,摸,ず,砰,计秖,蝗ㄢ,弧,sm,单,ю阑,ň眘,穦) values ('臸苝ホ',"&info(9)&",'猭竟',0,0,1,200000,1100,1100,0,0,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 珇 set 计秖=计秖+1,穦="&aaao&" where id="& id &" and 局Τ="&info(9)
	end if
peibashi=info(0) & "厚屡贺腳ホ贺腳ホ挡癬笵▇ど癬<img src='img/look52.gif'>贺腳ホてΘ臸苝ホ." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>