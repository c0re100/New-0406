<%
Response.Expires=0
Response.Buffer=True
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
action=Request.Querystring("action")
if action="search" then
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Randomize timer
r=int(rnd*10)
Set Conn1=Server.CreateObject("ADODB.Connection")
Connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn1.Open Connstr1
set rs=conn1.execute("select ����,�ɶ� from ������~ where �֦���='"& name &"'")
if rs.eof and rs.bof then
  Conn1.Execute("insert into ������~(�֦���) values('"& name &"')")
times=8
else
times=rs("����")
  if now()-rs("�ɶ�")>1 then
    conn1.execute("update ������~ set ����=����+8,�ɶ�=now() where �֦���='"& name &"'")
    times=times+8
  end if
end if
        if times<1 then
            mess="�A����ʤO�S�F�I�I�I"
        else
            conn1.execute("update ������~ set ����=����-1 where �֦���='"& name &"'")
            if r=5 then
              conn.execute("update �Τ� set �Ȩ�=�Ȩ�+100000 where �m�W='"& name &"'")
              mess="�o�{�F�_��10�U��I�I�I"
            else
              mess="�@�L����I�I�I"
            end if
        end if
set rs=nothing
conn1.close : set conn1=nothing
conn.close : set conn=nothing
response.write "<script language='javascript'>alert('"& mess &"');window.close();</script>"
else
%>
<html>
<head>
<title>����ɥN</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
input {  font-family: "�s�ө���"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong>�������ɥN</strong></big><br>
      <br>
      (�C�ѥi��<font face="Verdana">8</font>���̥ʦ����ͷN) 
      <input type="button" Onclick="MM_openBrWindow('see.asp','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="�ӫ~�@�l"> 
      ��������դ��A�o�{���~�Чi�D�ڡI�I�I <br> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000">  
    <td width="25%"><font color="#FFFFFF">�j�s��</font></td> 
    <td width="25%"><font color="#FFFFFF">�C�{��</font></td> 
    <td width="25%"><font color="#FFFFFF">��i��</font></td> 
    <td width="25%"><font color="#FFFFFF">�u�w��</font></td> 
  </tr> 
  <tr align="center">  
    <td width="25%">  
      <input type="button" Onclick="MM_openBrWindow('jys.asp?port=�j�s','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="�����"> 
      <input type="button" Onclick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="�����j��"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=�C�{','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="�����" name="button"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="�����j��" name="button"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=��i','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="�����" name="button2"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="�����j��" name="button2"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=�u�w','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="�����" name="button3"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="�����j��" name="button3"> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center"> 
  <tr align="center">  
    <td height="22"><br>
      �F�C���� <br>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>
