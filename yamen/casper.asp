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
		alert ("�A��J���{�ҽX�����T�A���ӿ�J:<%=regjm%>")
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
			alert "�O���O���Q����F�H�s�j�W�M�i���f�O�������H"
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
rs.open "select ���A,��O,�K�X,lastkick,�ާ@�ɶ�,�|������ from �Τ� where �m�W='" & name & "' and �K�X='" & pass & "' ",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	%>
	<script language=vbscript>
	alert "���S���d���H�����o�ӤH�ڡH"
	location.href = "javascript:history.back()"
	</script>
	<%response.end
else
lastkick=rs("lastkick")
	if (rs("���A")<>"��") and (rs("��O")>-100) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing%>
		<script language=vbscript>
			alert "���S���d���H�o�ӤH�٨S���ڡH"
			location.href = "javascript:history.back()"
		</script>
		<%response.end
	else
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=1-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]�����A�_���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
			if trim(pass)=rs("�K�X") then
if wbpf<>0 then
	conn.execute("update �Τ� set ���A='���`',��O=10000,�O�@=true where �m�W='"&name&"'")
	Response.Write "<script Language=Javascript>alert('�A�b�[�����a["&wbname&"�W���A�_�����A���|�ܡI]');</script>"
else
				if rs("�|������")>1 then
					conn.execute("update �Τ� set ���A='���`',��O=1000,�O�@=true where �m�W='"&name&"'")%>
					<script language=vbscript>
					alert "OK,�A�O�F�C����|���A�ҥH�A�_���F����]�S����A���O�A�ڭ��٬O���Ʊ�A�A�ӤF!"
					</script><%
				else
					conn.execute("update �Τ� set ���A='���`',��O=1000,���O=100,�O�@=true,�Ȩ�=100 where �m�W='"&name&"'")%>
					<script language=vbscript>
					alert "OK,�A�{�b�w�g�_���F�A���n�A�ӤF��,�p�G�O����|���Φb�[�����a�W���N���|������l��!"
					</script><%
				end if
end if%><head>
						<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
						<meta name="ProgId" content="FrontPage.Editor.Document">
						<title>�_�����\</title></head>
					
<body background="../linjianww/0064.jpg">
<div align="center"><p><font size="2">���ߤj�L�A���\���_�ͤF�A�������U�A�O�@�F�I�O��A�����A�_��<br>
					�A�@�w�n�����@�w�n�C�C�C�C�C</font></p>
					<p><input type="button" value="�������f" onclick="window.close()"
					name="button">
		<%else%><script language=vbscript>
		alert "�K�X����ڡA�|���|�O���F�H"
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�t�Ρj</font>["&name&"]�j�L���F��["&lastkick&"]�i�}������ʡA�o��������ണ�e�_���F!" 
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
%></p></div> 