<%'武學修煉
function wuxiu(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
Response.End 
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if

randomize()
rnd1=int(rnd*800)+100
rnd2=int(rnd*800)+120
rs.close
rs.open "select 修煉級 FROM 武功 WHERE   門派='" & info(5) & "' and 武功='"&fn1&"'",conn
if    rs.eof and   rs.bof then 
Response.Write "<script language=JavaScript>{alert('你派沒有["&fn1&"]這樣的武功呀！');}</script>"
Response.End
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
xiu=rs("修煉級")
if xiu<20000*xiu  then
cha=20000*info(2)-xiu
conn.execute "update 武功 set 修煉級=修煉級+1 where 武功='"&fn1&"' and 門派='"&info(5)&"'"
conn.execute "update 用戶 set 體力=體力-"&rnd1&",內力=內力-"&rnd2&",操作時間=now() where id="&info(9)
wuxiu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行["&info(5)&"]["&fn1&"]的修煉,體力失去<font color=red>-"&rnd1&"</font>點,內力失去<font color=red>-"&rnd2&"</font>點,["&fn1&"]有所小成，還差"&cha&"點火候!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
exit function
else
conn.execute "update 武功 set 修煉級=1,級別=級別+1 where 武功='"&fn1&"' and 門派='"&info(5)&"'"
wuxiu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行["&info(5)&"]["&fn1&"]的修煉,經過一番修煉["&fn1&"]武學提升1級!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
exit function
end if
end  function

%>