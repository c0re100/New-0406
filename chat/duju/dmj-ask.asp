<%
Function dmjask(fn1,from1,to1)
chatroomsn=session("nowinroom")

if to1="大家" or to1=info(0) then 
  response.write "<script Language='Javascript'>alert('你不可以選擇大家或自已作為對手!');</script>"
  response.end
end if

f=Minute(time()) 

'------------------------新格式------------------------
ARR_FN=Split(fn1,"|")
dim ub_Err,xia_class,xiamoney,xia_fir,xia_sql

if ubound(ARR_FN)<>1 Then 
  response.write "<script Language='Javascript'>alert('[操作錯誤]命令格式如下：\n /打麻將 下注數目|銀兩[或：金幣，法力] \n\n[示例]：\n /打麻將 1000|銀兩\n /打麻將 1000|金幣\n /打麻將 1000|法力!');</script>"
  response.end
end if

xia_class=ARR_FN(1)
xiamoney=ARR_FN(0)

select case xia_class
	case "金幣"
		xia_fir="1"
		xia_sql="金幣"
	case "銀兩"
		xia_fir="2"
		xia_sql="銀兩"
	case "法力"
		xia_fir="3"
		xia_sql="法力"
	case else
  response.write "<script Language='Javascript'>alert('[操作錯誤]命令格式如下：\n /打麻將 下注數目|銀兩[或：金幣，法力] \n\n[示例]：\n /打麻將 1000|銀兩\n /打麻將 1000|金幣\n /打麻將 1000|法力!');</script>"
  response.end
end select

If not isnumeric(xiamoney) Then
  response.write "<script Language='Javascript'>alert('[操作錯誤]命令格式如下：\n /打麻將 下注數目|銀兩[或：金幣，法力] \n\n[示例]：\n /打麻將 1000|銀兩\n /打麻將 1000|金幣\n /打麻將 1000|法力!');</script>"
  response.end
end if

xiamoney=abs(int(xiamoney))

if (xia_fir="1")and(xiamoney<10 or xiamoney>1000) then 
response.write "<script Language='Javascript'>alert('最少押：10" & xia_class & "，最多1000個金幣!');</script>"
  response.end
end if

if (xia_fir="2")and(xiamoney<1000000 or xiamoney>10000000) then 
response.write "<script Language='Javascript'>alert('最少押：1000000" & xia_class & "，最多1000萬銀兩!');</script>"
  response.end
end if
if (xia_fir="3")and(xiamoney<1000 or xiamoney>100000) then 
response.write "<script Language='Javascript'>alert('最少押：1000" & xia_class & "，最多10萬法力!');</script>"
  response.end
end if

'------------------------新格式------------------------
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'------------------------新格式------------------------
sql="select " & xia_sql & " from 用戶 where 姓名='" &info(0)&"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")

if (xiamoney * 2)>yin1 then 
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('至少要有下注的兩倍以上的賭本,可是你身上沒有這麼多" & xia_class & "!');</script>"
  response.end
end if

sql="select " & xia_sql & " from 用戶 where 姓名='" &to1&"'"
set rs=conn.execute(sql)
yin2=rs(""&xia_sql&"")
if (xiamoney * 2)>yin2 then 
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('至少要有下注的兩倍以上的賭本,可是對方身上沒有這麼多" & xia_class & "!');</script>"
  response.end
end if

'------------------------新格式------------------------
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/DMJ.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& info(0) & "'"

Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('您還有牌局沒有結束，不能開局!');</script>"
  response.end
else
  rs.close

dmjask="<b><font color=green>["&info(0)&"]</font></b>對<b><font color=black>["&to1&"]</font></b>說：打麻將嗎？賭注為<b><font color=red>"& xiamoney & xia_class &"</font></b>，<img src='duju/dpk/1.GIF'><input type=button value='願意' onclick=javascript:window.open('duju/dmjok.asp?id="&myid&"&from1="&info(9) & "&to1="&to1&"&yn=1','a','width=380,height=210');this.disabled=1 style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=button223><img src='duju/dpk/2.gif'><input type=button value='拒絕' onclick=javascript:window.open('duju/dmjok.asp?id="&myid&"&from1="&info(9) & "&to1="&to1&"&yn=0','a','width=380,height=210');this.disabled=1 style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=button223>"

sql="insert into dmj(ufrom,uto,duz) values ('"& info(0) & "$下注','"& to1 & "$下注', " & xia_fir & xiamoney & ")"
Set Rs=conn.Execute(sql)

conn.close
end if
end function
%>
