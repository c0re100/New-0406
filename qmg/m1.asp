<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from 原料 where 類型='湯料' and 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "你沒有湯料，怎麼做面啊，快到菜市場去買吧。"
location.href = "caichang.asp"
</script><%conn.close
response.end
else
randomize timer
mw=2*rnd*rs("美味度")
sql="select * from 巧面 where  擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="select * from 巧面 where  時限=true"
if rs.eof or rs.bof then		
sql="insert into 巧面 (擁有者,美味度) values ('"& name &"'," & mw & ")"
rs=conn.execute(sql)
else
sql="update 巧面 set 時限=false,美味度="&mw&",擁有者='" & name & "' where 時限=true"
rs=conn.execute(sql)
end if
sql="update 原料 set 時限=true where 類型='湯料' and 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
%><script language=vbscript>
MsgBox "你把雞湯放進鍋裡，一碗好面就快在眼前了。"
location.href = "m2.htm"
</script><%else
sql="update 巧面 set 時限=true where 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
%>
<script language=vbscript>MsgBox "你把以前沒有做好的面倒掉，回到廚房重新來過。"
location.href = "ctl_work_noodle.htm"
</script><%
set rs=nothing
conn.close
set conn=nothing
end if	
end if
%> 