<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�򥻩R�O</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt}
</style>
</head>
<body bgcolor="#660000" leftmargin="0" topmargin="0">
<center>
  <table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" >
    <tr bgcolor="#EEFFEE"> 
      <td colspan="4"> 
        <div align="center"><font color="#FF6633">�����ѩR�O</font><br>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM �ʧ@",conn%>
        </div>
      </td>
    </tr>
    <tr bgcolor="#FF6633"> 
      <td width=25%> 
        <div align="center">�R�O</div>
      </td>
      <td width=27%> 
        <div align="center">�ۤv</div>
      </td>
      <td width=25%> 
        <div align="center">�j�a</div>
      </td>
      <td width=23%> 
        <div align="center">�O�H</div>
      </td>
    </tr>
    <%do while not rs.eof and not rs.bof
%>
    <tr bgcolor="#f7f7f7"> 
      <td valign=top width="25%"> 
        <div align="center"><%=rs("�R�O")%></div>
      </td>
      <td valign=top width="27%"><a href="javascript:s('//<%=replace(rs("�ۤv"),"##","")%>')">�ۤv</a></td>
      <td valign=top width="25%"><a href="javascript:s('//<%=replace(rs("�j�a"),"##","")%>')">�j�a</a></td>
      <td valign=top width="23%"><a href="javascript:s('//<%=replace(rs("�O�H"),"##","")%>')">�O�H</a></td>
    </tr>
    <%
rs.movenext
l=l+1
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  </table>
</center>
</body>
</html> 
<html><script language="JavaScript">                                                                  </script></html>