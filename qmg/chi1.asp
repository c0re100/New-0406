<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
my=info(0)
sql="select 銀兩 from 用戶 where 姓名='" & my & "'"
set rs=conn.execute(sql)
xianjin=rs("銀兩")
sql="SELECT 售價,體力,擁有者 FROM 賣面 where ID=" & id & " and 時限=false"
Set rs=conn.Execute(sql)
if not(rs.eof and rs.bof) then
yin=rs("售價")
tl=rs("體力")
yy=rs("擁有者")
if yin <= xianjin then
sql="update 用戶 set 銀兩=銀兩-" & yin & ",體力=體力+" & tl & "   where  姓名='" & my & "'"
conn.execute(sql)
conn.execute("update 賣面 set 時限=true where ID="&id)
response.write "你喫了巧面館師傅"& yy &"做的面，增加體力"&tl
else
response.write "不能喫面，原因：你的銀兩不夠了"
conn.close
response.end
end if
end if
set rs=nothing
conn.close
set conn=nothing
%> 