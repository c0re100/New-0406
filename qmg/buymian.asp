<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
sql="SELECT 等級 FROM 用戶 WHERE 姓名='" & name & "'"
Set Rs=conn.Execute(sql)
jishu=rs("等級")
sql="select 美味度 from 巧面 where 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "你沒有面啊，是不是想騙我，快去面館作面吧。"
location.href = "qmg.htm"
</script><%
conn.close
response.end
else
mw=rs("美味度")
tl=mw*8
yin=jishu+mw*8+100
sql="select id from 賣面 where 時限=true"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="insert into 賣面(擁有者,體力,售價) values ('"& name &"',"& tl &","& yin &")"
conn.execute(sql)
else
id =rs("id")
sql="update 賣面 set 時限=false,擁有者='"& name &"',體力="& tl &",售價="& yin &" where id="&id&""
conn.execute(sql)
end if	
conn.execute("update 巧面 set 時限=true where 擁有者='" & name & "' and 時限=false" )
conn.execute("update 用戶 set 銀兩=銀兩+"& yin &" where 姓名='"&name&"'")
Response.write"你把你親手做的面賣給了巧面館，得到銀子" & yin& "兩"
set rs=nothing
conn.close
set conn=nothing
end if%> 