<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
if info(2)<10 then	Response.Redirect "../error.asp?id=439"
%>
<!--#include file="jhb.asp"-->
<%
Randomize
sj=int(rnd*26)+28
sql= "select * from �Ѳ� where sid="&sj        
set rs=conn.execute(sql)  
if (rs("��e����")-rs("�}�L����"))/rs("�}�L����")>=0.3 then 
ming=rs("���~")        
sql="update �Ѳ� set ��e����=��e����*0.8 where sid="&sj
conn.execute sql           
mess=ming&"���Ѳ��D�j��߰�A�Ѳ�����U��20%"
elseif (rs("��e����")-rs("�}�L����"))/rs("�}�L����")<=-0.3 then
ming=rs("���~")
sql="update �Ѳ� set ��e����=��e����*1.2 where sid="&sj
conn.execute sql              
mess=ming&"���Ѳ�Ĳ���ϼu�A����W��20%"
else     
ming=rs("���~")     
Randomize
sj1=int(10*rnd)+1
if sj1>4 then
sql="update �Ѳ� set ��e����=��e����*"&(1-(sj1-1)/100)&" where sid="&sj
conn.execute sql 
mess=ming&"���ƪ����ڰk�]�A�Ѳ�����U�^"&(sj1-1)&"%"
else
sql="update �Ѳ� set ��e����=��e����*"&(1+sj1/100)&" where sid="&sj
conn.execute sql 
mess=ming&"�ͷN�����A�Ѳ�����W��"&sj1&"%"
end if
end if     
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
	sd(196)="FFFDAF"
	sd(197)="FFFDAF"
	sd(198)="��"
	sd(199)="<font color=red>�i�ѥ��j"& mess &"</font>"
	sd(200)=session("nowinroom")
	Application("ljjh_sd")=sd
	Application.UnLock
Response.Redirect "index.asp"

%>

