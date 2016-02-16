<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
If Session("usepro") = true Then
'id=request("id")
win=request("win")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if win<>0 then
rs.open "select id from 用戶 where id="&info(9),conn
conn.execute "update 用戶 set 內力=內力+1000,體力=體力+1000,攻擊=攻擊+1000,防御=防御+1000,銀兩=銀兩+2000  where id="&info(9)
rs.close
rs.open "select 等級,攻擊加,防御加,內力加,體力加,防御,攻擊,體力,內力 from 用戶 where id="&info(9),conn
gjj=rs("攻擊加")
fyj=rs("防御加")
nlj=(rs("等級")*640+2000)+rs("內力加")
tlj=(rs("等級")*1500+5260)+rs("體力加")
if rs("防御")>fyj then
conn.execute "update 用戶 set 防御=" & fyj & "  where id="&info(9)
end if
if rs("攻擊")>gjj then
conn.execute "update 用戶 set 攻擊=" & gjj & "  where id="&info(9)
end if
if rs("體力")>tlj then
conn.execute "update 用戶 set 體力=" & tlj & "  where id="&info(9)
end if
if rs("內力")>nlj then
conn.execute "update 用戶 set 內力=" & nlj & " where id="&info(9)
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Session("usepro")= false 
Response.Write "<script Language=Javascript>alert('大俠，你打擊倭寇有功，盟主獎內力、體力、攻擊、防御各1000點，銀兩2000兩！');</script>"
response.end

%>
<script language="vbscript">
window.close()
</script>	  
 <%
 else 

conn.execute "update 用戶 set 內力=內力-300,體力=體力-300,銀兩=銀兩-20 where id="&info(9)
Session("usepro")= false 
rs.open "select 內力,體力,狀態 from 用戶 where id="&info(9),conn
if rs("內力")<1 then conn.execute "update 用戶 set 內力=100 where id="&info(9)
if rs("體力")<0 or rs("狀態")="死" then Response.Redirect "../exit.asp"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('菜鳥，平時不練功，現在打敗了，還敢來見我，罰你體力和內力300，銀兩20！');</script>"
response.end
%>
<script language="vbscript">
window.close()
</script>	
<%
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Session("usepro")= false 
else 
response.write "請通過正常途徑來打倭寇." 
response.end 
end if
%>
