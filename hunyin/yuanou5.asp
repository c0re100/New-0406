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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �t�� from �Τ� where �m�W='" & info(0) & "'",conn
peier=rs("�t��")
if rs("�t��")="�L" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�٨S���B������H');window.close();</script>"
	response.end
end if
rs.close
rs.open "select id from �Τ� where �m�W='" & peier & "'",conn
peierid=rs("id")
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ѤF�a�@���y�P�\�F�I');window.close();</script>"
	response.end			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="& id &" and �֦���="&info(9)
				
		        end if
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���=" & peierid,conn
if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�y�P�\',"&peierid&",'���~',0,1000,0,0,0,1,10000000,111111,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&peierid
				
		        end if
conn.execute "update �Τ� set �t��='�L',���B�ɶ�=date(),�p��='�L' where �m�W='" & info(0) & "'"
conn.execute "update �Τ� set �t��='�L',���B�ɶ�=date(),�p��='�L' where �m�W='" & peier & "'"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
message="[" & info(0) & "]�M[" & peier & "]�P���}���A�ŧi���B�I"
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
sd(199)="<font color=FFFDAF>�i���B�����j"& message &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<title>���B</title>

<STYLE type=text/css>A:hover {
	CURSOR: resize
}
A {
	TEXT-DECORATION: none
}
select       { background-color: #FFFFFF; font-size: 9pt; border-left: medium solid #999900; 
              border-right: medium solid #FFCC66; 
               border-top: medium solid #999900; 
               border-bottom: medium solid #FFCC66  }
</STYLE>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body text="#00FFFF" bgcolor="#FFFFFF" link="#99FF33" vlink="#FFFFFF" alink="#FF0000" leftmargin="0" topmargin="0">
<table width="691" border="0" cellspacing="0" cellpadding="0" height="119">
  <tr>
    <td background="lj1.gif"> 
      <p><img src="111.gif" width="77" height="111" align="left"><font color="#FF0000">�դU���ýt�w�g�Ѱ��F�A�Ʊ�դU�H��n�n�D��ۤv����Q�A�O�A<br>
        �ӧ�ڤF�C</font></p>
      <p><a href="#" onClick="window.close()">�n���C</a> </p>
    </td>
  </tr>
</table>
<div align="center"></div>

</body>

</html>

 