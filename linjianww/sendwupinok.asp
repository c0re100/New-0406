<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

if info(4)=0 then 
aaao=0
else
aaao=1
end if

if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=trim(request.form("a"))
b=trim(request.form("b"))
c=trim(request.form("c"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM 用戶 where 姓名='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，沒有該用戶！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
idd=rs("id")
rs.close
rs.open "select id from 物品 where 物品名='流星' and 擁有者=" & idd,conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,會員) values ('流星',"&idd&",'物品',0,1000,0,0,0,"&b&",10000000,111111,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量+"&b&",會員="&aaao&" where id="& id &" and 擁有者="&idd
				
		        end if
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 擁有者=" & idd,conn
if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,會員) values ('流星淚',"&idd&",'物品',0,1000,0,0,0,"&c&",10000000,2003,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="& id &" and 擁有者="&idd
				
		        end if
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','流星、流星淚發送操作')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
