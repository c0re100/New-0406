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
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(Request.form("wpsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=lcase(trim(request.querystring("action")))
if InStr(action,"or")<>0 or InStr(action,"=")<>0 or InStr(action,"`")<>0 or InStr(action,"'")<>0 or InStr(action," ")<>0 or InStr(action," ")<>0 or InStr(action,"'")<>0 or InStr(action,chr(34))<>0 or InStr(action,"\")<>0 or InStr(action,",")<>0 or InStr(action,"<")<>0 or InStr(action,">")<>0 then Response.Redirect "../../error.asp?id=54"
if action="" then 
		Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���~,���w���~�W�I�I');location.href = 'javascript:history.go(-1)'</script>"
		response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
rs.close
rs.open "select * from ���~ where ����='�ī~' and ���~�W='" & action & "' and �֦���=" & info(9) & " and �ƶq>="&sl,conn
	if rs.eof or rs.bof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('���ܡG�A�����~�ƶq�����I');location.href = 'javascript:history.go(-1)'</script>"
			response.end
	end if
id=rs("ID")
nei=int(rs("���O"))*sl
ti=int(rs("��O"))*sl
conn.execute "update ���~ set �ƶq=�ƶq-"& sl &" where id=" & id
conn.execute "delete * from ���~ where �ƶq<=0"
conn.execute "update �Τ� set ���O=���O+" & nei & ", ��O=��O+" & ti & ",�ާ@�ɶ�=now() where id="&info(9)
rs.close
rs.open "SELECT ����,���O,��O,���O�[,��O�[,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
nlj=(rs("����")*640+2000)+rs("���O�[")
if rs("���O")>nlj then
conn.execute "update �Τ� set ���O=" & nlj & ",�ާ@�ɶ�=now() where id="&info(9)
mess="�A��ΤF�Ī�"& sl &"�ӡA�j��{�b���Ū����O�ȡA�ҥH���O���G" & nlj & "�Ф��n�A�A�ΡI"
end if
tlj=(rs("����")*1500+5260)+rs("��O�[")
if rs("��O")>tlj then
conn.execute "update �Τ� set ��O=" & tlj & ",�ާ@�ɶ�=now() where id="&info(9)
mess="�A��ΤF�Ī�"& sl &"�ӡA�j��{�b���Ū���O�ȡA�ҥH��O���G" & tlj & "�Ф��n�A�A�ΡI"
else
mess="�A��ΤF�Ī�"& sl &"�ӡA�W�[���O" & nei & "�A�W�[��O" & ti & "�I"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����</title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>���򴣥�</b></font></td>
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
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='��^';return true;" onmouseout="window.status='';return true;">��^�˳Ƥ@��</a></td>
</tr>
</table>

</body>