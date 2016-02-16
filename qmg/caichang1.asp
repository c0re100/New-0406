<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=request("id")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 物品名,銀兩,類型,說明,功能,美味度 FROM 菜場 where ID=" & id
Set Rs=conn.Execute(sql)
wu=rs("物品名")
yin=rs("銀兩")
lx=rs("類型")
sm=rs("說明")
gn=rs("功能")
mw=rs("美味度")
sql="select 銀兩 from 用戶 where 姓名='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("銀兩") then
sql="update 用戶 set 銀兩=銀兩-" & yin & "  where 姓名='" & my & "'"
rs=conn.execute(sql)
sql="select * from 原料 where 物品名='" & wu & "' and 時限=false and 擁有者='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="select * from 原料 where 物品名='" & wu & "' and 時限=true"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="insert into 原料(物品名,擁有者,類型,說明,功能,美味度,銀兩) values ('"& wu &"','"& my &"','"& lx &"','" & sm &"','"& gn &"',"& mw &","& yin &")"
rs=conn.execute(sql)
conn.close
Response.Redirect "caichang.asp"
else
id2=rs("id")
sql="update 原料 set 擁有者='"&my&"',時限=false where id=" & id2
rs=conn.execute(sql)
conn.close
Response.Redirect "caichang.asp"
end if
else
response.write "由於你已購買了這個物品，所以不能再買！"
conn.close
response.end
				
end if
else
response.write "不能買東西，原因：你的銀兩不夠了"
conn.close
response.end
end if
rs.close
set rs=nothing
set conn=nothing
%> 