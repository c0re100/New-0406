<%
Response.Expires=0
Response.Buffer=True
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
action=Request("action")
if action="true" then
if instr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("SERVER_NAME"))<>0 then
port=Request.Form("port")
what=Request.Form("type")
submit=trim(Request.Form("submit"))
num=trim(Request.Form("num"))
On Error Resume Next
num=Clng(num)
if err.number>0 then
  err.clear
  response.write "<script language='javascript'>alert('�п�J�Ʀr�I�I�I');history.back(-1);</script>"
  response.end
end if
if num<1 or num=empty then num=1
if num>10 then num=10
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open Application("ljjh_usermdb")
Set RS=Conn.Execute("select �֦���,����,�ɶ� from ������~ where �֦���='"& name &"'")
if RS.eof and RS.bof then
  Conn.Execute("insert into ������~(�֦���) values('"& name &"')")
  times=8
else
times=rs("����")
  if now()-rs("�ɶ�")>1 then
    conn.execute("update ������~ set ����=����+800,�ɶ�=now() where �֦���='"& name &"'")
    times=times+800
  end if
end if
RS.close
if times<1 then
mess="�A����ʤO�S�F�I�I�I"
else
Rs=Conn1.execute("select �Ȩ� from �Τ� where �m�W='"& name &"'")
yin=Rs("�Ȩ�")
Sql="select * from ����ɥN where ��f='"& port &"'"
Rs=Conn.Execute(Sql)
jia1=Rs("����") : jia2=Rs("�H��") : jia3=Rs("�]�_") : jia4=Rs("����")
  if submit="�R" then
    Select Case what
      Case "����"
          if num*jia1>yin then
            mess="�A���������F�I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�-"& num*jia1 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,����=����+"& num &" where �֦���='"& name &"'")
            Conn.Execute("update ����ɥN set ����=����+1 where ��f='"& port &"'")
            mess="���\�ʶR�I�I�I"
          end if
      Case "�H��"
          if num*jia2>yin then
            mess="�A���������F�I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�-"& num*jia2 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,�H��=�H��+"& num &" where �֦���='"& name &"'")
            Conn.Execute("update ����ɥN set �H��=�H��+2 where ��f='"& port &"'")
            mess="���\�ʶR�I�I�I"
          end if
      Case "�]�_"
          if num*jia3>yin then
            mess="�A���������F�I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�-"& num*jia3 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,�]�_=�]�_+"& num &" where �֦���='"& name &"'")
            Conn.Execute("update ����ɥN set �]�_=�]�_+3 where ��f='"& port &"'")
            mess="���\�ʶR�I�I�I"
          end if
      Case "����"
          if num*jia4>yin then
            mess="�A���������F�I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�-"& num*jia4 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,����=����+"& num &" where �֦���='"& name &"'")
            Conn.Execute("update ����ɥN set ����=����+1 where ��f='"& port &"'")
            mess="���\�ʶR�I�I�I"
          end if
     End Select
  else
    Select Case what
      Case "����"
          rs=Conn.Execute("select ���� from ������~ where �֦���='"& name &"'")
          if rs("����")<num then
            mess="�A�S����h�����I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�+"& num*jia1 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,����=����-"& num &" where �֦���='"& name &"'")
            if jia1>1 then
               Conn.Execute("update ����ɥN set ����=����-1 where ��f='"& port &"'")
            end if
            mess="���\��X�I�I�I"
          end if
      Case "�H��"
          rs=Conn.Execute("select �H�� from ������~ where �֦���='"& name &"'")
          if rs("�H��")<num then
            mess="�A�S����h�H�ѡI�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�+"& num*jia2 &" where �m�W='"& my &"'")
            Conn.Execute("update ������~ set ����=����-1,�H��=�H��-"& num &" where �֦���='"& name &"'")
            if jia1>2 then
               Conn.Execute("update ����ɥN set �H��=�H��-2 where ��f='"& port &"'")
            end if     
            mess="���\��X�I�I�I"
          end if
      Case "�]�_"
          rs=Conn.Execute("select �]�_ from ������~ where �֦���='"& name &"'")
          if rs("�]�_")<num then
            mess="�A�S����h�]�_�I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�+"& num*jia3 &" where �m�W='"& my &"'")
            Conn.Execute("update ������~ set ����=����-1,�]�_=�]�_-"& num &" where �֦���='"& name &"'")
            if jia1>3 then
               Conn.Execute("update ����ɥN set �]�_=�]�_-3 where ��f='"& port &"'")
            end if
            mess="���\��X�I�I�I"
          end if
      Case "����"
          rs=Conn.Execute("select ���� from ������~ where �֦���='"& name &"'")
          if rs("����")<num then
            mess="�A�S����h�����I�I�I"
          else
            Conn1.Execute("update �Τ� set �Ȩ�=�Ȩ�+"& num*jia4 &" where �m�W='"& name &"'")
            Conn.Execute("update ������~ set ����=����-1,����=����-"& num &" where �֦���='"& name &"'")
            if jia1>1 then
               Conn.Execute("update ����ɥN set ����=����-1 where ��f='"& port &"'")
            end if
            mess="���\��X�I�I�I"
          end if
     End Select
  end if
end if
set Rs=nothing
Conn.close : set Conn=nothing
Conn1.close : set Conn1=nothing
response.write "<script language='javascript'>alert('"& mess &"');history.back(-1);</script>"
else
response.write "�Ф��n�@���I�I�I"
end if
else
port=Request.Querystring("port")
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
Sql="select * from ����ɥN where ��f='"& port &"'"
Set Rs=Conn.Execute(Sql)
%>
<html>
<head>
<title><%=port%>��ӷ|</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
input {  font-family: "�s�ө���"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<table width="90%" border="1" cellspacing="2" cellpadding="2" align="center">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td colspan="4" height="22"><%=port%>��ӷ| <br>
      <br>
      (�@���ܦh�ʶR10���)<br>
    </td>
  </tr>
  <tr align="center" bgcolor="#660000"> 
    <td height="20"><font color="#FFFFFF">����</font></td>
    <td height="20"><font color="#FFFFFF">�H��</font></td>
    <td height="20"><font color="#FFFFFF">�]�_</font></td>
    <td height="20"><font color="#FFFFFF">����</font></td>
    <%do while not Rs.eof%> </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="20"><%=Rs("����")%>��/���</td>
    <td height="20"><%=Rs("�H��")%>��/���</td>
    <td height="20"><%=Rs("�]�_")%>��/���</td>
    <td height="20"><%=Rs("����")%>��/���</td>
  </tr>
  <tr align="center">
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" �R " OnClick=check()>
        <input type="submit" name="submit" value=" �� " OnClick=check()>
        <input type="hidden" name="type" value="����">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" �R " OnClick=check()>
        <input type="submit" name="submit" value=" �� " OnClick=check()>
        <input type="hidden" name="type" value="�H��">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" �R " OnClick=check()>
        <input type="submit" name="submit" value=" �� " OnClick=check()>
        <input type="hidden" name="type" value="�]�_">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" �R " OnClick=check()>
        <input type="submit" name="submit" value=" �� " OnClick=check()>
        <input type="hidden" name="type" value="����">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
  </tr>
  <%
Rs.movenext
Loop
Rs.close : set Rs=nothing
Conn.close : set Conn=nothing
end if
%> 
</table>
</body>
</html>