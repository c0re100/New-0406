<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �w���I�� from �Τ� where id="&info(9),conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
if rs("�w���I��")<1000 then
	mess="�A�{�b�٨S�����l�A�u��:<font color=red>"&rs("�w���I��")&"</font>�I"
else
	dousl=int(rs("�w���I��")/1000)
	conn.execute "update �Τ� set �w���I��=�w���I��-"& dousl*1000 &",�|���O=�|���O+"& 2*dousl &" where id="&info(9)
	mess="�樧�l�G"&dousl&"�ӡA�|���O�W�[�G"&dousl*2&"���A�ٳ�:"&rs("�w���I��")-dousl*1000&"�I�I"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>�樧�l</title>
<body leftmargin="0" topmargin="0" bgcolor="#003366" text="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="259" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=#FFCC55><%=info(0)%></font><font color="#CCFF33">�w����{�ۥѥ����樧</font></b></font></div>
      <table width="280" align="center">
        <tr> 
            <td> 
          <tr> 
            
          <td valign="top" height="11" > 
            <div align="center">���l²��</div>
            </td>
          </tr>
          <tr> 
            
          <td valign="top" height="150" > 
            <p>���l�O�ѧA�w�I����U�ӤF,�P�w�I�t��C�C�w��1000�I�A�p����t�Φ۰ʵ��A�@�Ө��A��A��ܼɨ���A���ˤO������3���A�i�H����1�p�ɮɶ��C�p�G�A���Q�ϥΡA�i�H�b�o�̽�F�C�@�Ө��l=2���|���O�ΡA�A�i�H�ν樧�l���|���O�ʶR�u���|���צ����d���I</p>
<%=mess%><br><br><br>
            </td>
          </tr>
        </table>

</td>
</tr>
</table>
<div align="center"> </div>
</body>
</html>



