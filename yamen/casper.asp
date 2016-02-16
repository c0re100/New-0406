<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then%>
	<script language=vbscript>
		alert ("你輸入的認證碼不正確，應該輸入:<%=regjm%>")
		location.href = "javascript:history.back()"
	</script>
	<%response.end
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then  Response.End
next
for each element in Request.QueryString
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.QueryString(element),"'")<>0 or instr(Request.QueryString(element)," ")<>0 or instr(Request.QueryString(element),"|")<>0 then  Response.End
next
wbname=""
wbpf=0
ip=Request.ServerVariables("REMOTE_ADDR")
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../wb/ljjmwb.asa")
Conn.Open connstr
rs.open "SELECT barname FROM bar WHERE ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=1
end if
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=request.form("name")
pass=trim(request.form("pass"))
if name="" or pass="" then
		%><script language=vbscript>
			alert "是不是不想雪恨了？連大名和進門口令都不報？"
			location.href = "javascript:history.back()"
		</script><%
	response.end
end if
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
rs.open "select 狀態,體力,密碼,lastkick,操作時間,會員等級 from 用戶 where 姓名='" & name & "' and 密碼='" & pass & "' ",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	%>
	<script language=vbscript>
	alert "有沒有搞錯？哪有這個人啊？"
	location.href = "javascript:history.back()"
	</script>
	<%response.end
else
lastkick=rs("lastkick")
	if (rs("狀態")<>"死") and (rs("體力")>-100) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing%>
		<script language=vbscript>
			alert "有沒有搞錯？這個人還沒死啊？"
			location.href = "javascript:history.back()"
		</script>
		<%response.end
	else
sj=DateDiff("n",rs("操作時間"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=1-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分鐘再復活！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
			if trim(pass)=rs("密碼") then
if wbpf<>0 then
	conn.execute("update 用戶 set 狀態='正常',體力=10000,保護=true where 姓名='"&name&"'")
	Response.Write "<script Language=Javascript>alert('你在加盟網吧["&wbname&"上網，復活狀態不會變！]');</script>"
else
				if rs("會員等級")>1 then
					conn.execute("update 用戶 set 狀態='正常',體力=1000,保護=true where 姓名='"&name&"'")%>
					<script language=vbscript>
					alert "OK,你是靈劍江湖會員，所以你復活了什麼也沒有丟，但是，我們還是不希望你再來了!"
					</script><%
				else
					conn.execute("update 用戶 set 狀態='正常',體力=1000,內力=100,保護=true,銀兩=100 where 姓名='"&name&"'")%>
					<script language=vbscript>
					alert "OK,你現在已經復活了，不要再來了啊,如果是江湖會員或在加盟網吧上網將不會有任何損失!"
					</script><%
				end if
end if%><head>
						<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
						<meta name="ProgId" content="FrontPage.Editor.Document">
						<title>復活成功</title></head>
					
<body background="../linjianww/0064.jpg">
<div align="center"><p><font size="2">恭喜大俠你成功的復生了，站長幫助你保護了！記住，我讓你復生<br>
					你一定要報仇一定要。。。。。</font></p>
					<p><input type="button" value="關閉窗口" onclick="window.close()"
					name="button">
		<%else%><script language=vbscript>
		alert "密碼不對啊，會不會記錯了？"
		location.href = "javascript:history.back()"
		</script><%
		 end if
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【系統】</font>["&name&"]大俠為了對["&lastkick&"]展開報仇行動，發揮本身潛能提前復活了!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
%></p></div> 