<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhzb")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(Request.form("wpsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=lcase(trim(request.querystring("action")))
if InStr(action,"or")<>0 or InStr(action,"=")<>0 or InStr(action,"`")<>0 or InStr(action,"'")<>0 or InStr(action," ")<>0 or InStr(action," ")<>0 or InStr(action,"'")<>0 or InStr(action,chr(34))<>0 or InStr(action,"\")<>0 or InStr(action,",")<>0 or InStr(action,"<")<>0 or InStr(action,">")<>0 then Response.Redirect "../../error.asp?id=54"
if action="" then 
		Response.Write "<script Language=Javascript>alert('提示：操作錯誤,指定物品名！！');location.href = 'javascript:history.go(-1)'</script>"
		response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
rs.close
rs.open "select * from 物品 where 類型='藥品' and 物品名='" & action & "' and 擁有者=" & info(9) & " and 數量>="&sl,conn
	if rs.eof or rs.bof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：你的物品數量不足！');location.href = 'javascript:history.go(-1)'</script>"
			response.end
	end if
id=rs("ID")
nei=int(rs("內力"))*sl
ti=int(rs("體力"))*sl
conn.execute "update 物品 set 數量=數量-"& sl &" where id=" & id
conn.execute "delete * from 物品 where 數量<=0"
conn.execute "update 用戶 set 內力=內力+" & nei & ", 體力=體力+" & ti & ",操作時間=now() where id="&info(9)
rs.close
rs.open "SELECT 等級,內力,體力,內力加,體力加,操作時間 FROM 用戶 WHERE id="&info(9),conn
nlj=(rs("等級")*640+2000)+rs("內力加")
if rs("內力")>nlj then
conn.execute "update 用戶 set 內力=" & nlj & ",操作時間=now() where id="&info(9)
mess="你喫用了藥物"& sl &"個，大於現在等級的內力值，所以內力為：" & nlj & "請不要再服用！"
end if
tlj=(rs("等級")*1500+5260)+rs("體力加")
if rs("體力")>tlj then
conn.execute "update 用戶 set 體力=" & tlj & ",操作時間=now() where id="&info(9)
mess="你喫用了藥物"& sl &"個，大於現在等級的體力值，所以體力為：" & tlj & "請不要再服用！"
else
mess="你喫用了藥物"& sl &"個，增加內力" & nei & "，增加體力" & ti & "！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>江湖提示</b></font></td>
</tr>
<tr align="center">
<td>
<table width="260">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
</td>
</tr>
</table>
</td>
</tr>
<tr align="center">
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回裝備一覽</a></td>
</tr>
</table>

</body>