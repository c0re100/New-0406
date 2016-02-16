<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><title>職業轉換</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">
<%my=info(0)
ziye=trim(request.form("jiu"))
if InStr(ziye,"or")<>0 or InStr(ziye,"'")<>0 or InStr(ziye,"`")<>0 or InStr(ziye,"=")<>0 or InStr(ziye,"-")<>0 or InStr(ziye,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 職業,等級,銀兩 from 用戶 where id="&info(9),conn
if left(rs("職業"),2)=ziye or left(rs("職業"),3)=ziye or left(rs("職業"),4)=ziye then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你原來就是這個職業，還轉什麼職!');location.href = 'changework.asp';}</script>"
	response.end
end if
dj=rs("等級")
if rs("銀兩")<1000000  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職需要100萬轉職費!');location.href = 'changework.asp';}</script>"
	response.end
end if
if dj<50 and ziye="銅盔戰士"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職銅盔戰士戰鬥等級必須大等於50級!');location.href = 'changework.asp';}</script>"
	response.end
end if 

if ziye="銀鎧戰士"  then 
if dj<250 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職銀鎧戰士戰鬥等級必須大等於250級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職銀鎧戰士需要流星5個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if 

if ziye="金甲戰士"  then 
if dj<400 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職金甲戰士戰鬥等級必須大等於400級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職金甲戰士需要流星淚5個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if


if ziye="神龍戰士"  then 
if dj<650 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職神龍戰士戰鬥等級必須大等於650級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=30 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職神龍戰士需要流星30個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-30 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if

if dj<80 and ziye="百戰勇士"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職百戰勇士戰鬥等級必須大等於80級!');location.href = 'changework.asp';}</script>"
	response.end

end if

if ziye="虎威勇士"  then 
if dj<280 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職虎威勇士戰鬥等級必須大等於280級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職虎威勇士需要流星5個!');location.href = 'changework.asp';}</script>"	
Response.End 			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if
if ziye="擒龍勇士"  then 
if dj<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職擒龍勇士戰鬥等級必須大等於500級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職擒龍勇士需要流星淚5個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if
if ziye="金剛勇士"  then 
if dj<750 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職金剛勇士戰鬥等級必須大等於750級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=45 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職金剛勇士需要流星45個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-45 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if
if dj<90 and ziye="水道士"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水道士戰鬥等級必須大等於90級!');location.href = 'changework.asp';}</script>"
	response.end
end if
if ziye="水法師"  then 
if dj<320 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水法師戰鬥等級必須大等於320級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水法師需要流星5個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if
if ziye="水真人"  then 
if dj<550 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水真人戰鬥等級必須大等於550級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 數量>=5 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水真人需要流星淚5個!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-5 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if
if ziye="水天師"  then 
if dj<780 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水天師戰鬥等級必須大等於780級!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from 物品 where 物品名='流星' and 數量>=65 and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('轉職水真人需要流星65個!');location.href = 'changework.asp';}</script>"
Response.End 		
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-65 where id="& id &" and 擁有者="&info(9)
				
		        end if
end if
end if

if ziye="戰士" or ziye="銅盔戰士" or ziye="銀鎧戰士" or ziye="金甲戰士" or ziye="神龍戰士" or ziye="勇士" or ziye="百戰勇士" or ziye="虎威勇士" or ziye="擒龍勇士" or ziye="金剛勇士" or ziye="道士" or ziye="水道士" or ziye="水法師"  or ziye="水真人" or ziye="水天師" then
conn.execute "update 用戶 set 銀兩=銀兩-1000000,職業='"& ziye &"' where id="&info(9),conn


else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
Response.End 
end if
rs.close
	set rs=nothing
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您職業轉換成了："& ziye &"工作，點擊確定關閉當前窗口！！');window.close();}</script>"
Response.End 

%>