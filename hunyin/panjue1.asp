<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(5)="�x��" and info(2)>=7  then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name1=request("name")
fg=request("fg")
my=request("my")
rs.open "select �m�W1 from ���B where �m�W1='" & name1 & "'and �m�W2='" & my &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ҭn�P�M���B���H["&my&"]���s�b�I�I');history.go(-1);</script>"
	response.end
end if
rs.close
rs.open "select ID from �Τ� where �m�W='" & name1 & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ҭn�P�M���B���H["&name1&"]���s�b�I�I');history.go(-1);</script>"
	response.end
end if
conn.execute "delete * from ���B where �m�W1='" & name1 & "' or �m�W2='" & name1 & "' "
message="" & my & "�P" & name1 & "�g�x���P�M���B�L�ġA�Ѱ����B�n�O�I"
conn.close
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�����j'"& message &"'</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock%>
<body bgcolor="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=Yuanou.asp>��^</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
<%end if%> 