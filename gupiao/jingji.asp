<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>�Ѳ��g���H�줽��</title>
<style>
input,body,select,td{font-size:14;line-height:160%}
</style>
</head>

<body bgcolor="#990000">
<p align=center style='font-size:16;color:yellow'>�w��<%=info(0)%>���{�Ѳ��g���H���줽�ǡI
<p align=center style='font-size:14;color:blue'><a href=index.asp>��^�Ѳ�����</a></p>
<!--#include file="jhb.asp"-->
<%
sql="select �g���H from �Ȥ� where �b��='"&info(0)&"'"
set rs=conn.execute(sql)
if not rs.eof then
jjr=rs(0)
%>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>�z���Ѳ��g�٤H<%=jjr%>�b�����z�A��</p><br>
<a href=cha.asp>�d�ߤ�����</a>
<a href=cun.asp>�s���i�Ѳ��b��</a><br>
<a href=qu.asp>�q�Ѳ��b�ᴣ��</a>
<br> 
<p align=center><a href=index.asp>���}</a></p>
</table></table>  
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<p align=center style='font-size:16;color:yellow'>�ӽЪѲ��b��/�󴫸g�٤H 
<form method=POST action='jingji2.asp' id=form2 name=form2>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
        <table width=100%>
          <tr> 
            <td width="47%">�b���G 
              <input type=text name=name size=10 value="<%=info(0)%>" maxlength="0">
            </td>
            <td width="53%">�g�٤H�G 
              <select name=jjr size=1>
                <option value="��Pip">��Pip</option>
                <option value="�d�ޮ��">�d�ޮ��</option>
                <option value="�L�ת��T">�L�ת��T</option>
                <option value="���ҤH">���ҤH</option>
                <option value="���g�B">���g�B</option>
                <option value="�ϦʸU">�ϦʸU</option>
                <option value="������">������</option>
                <option value="�F���ڴ�">�F���ڴ�</option>
                <option value="�Ԥӭ�">�Ԥӭ�</option>
                <option value="�L����">�L����</option>
                <option value="����">����</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td colspan=2 align=center> 
              <input type=submit value=�T�w id=submit2 name=submit2>
              <input type=reset value="����" id=reset2 name=reset2>
            </td>
          </tr>
          <tr> 
            <td colspan=2 style='font-size:9pt'> 
              <hr>
              1.�ߤ�ɡA������Ĥ@�ӱK�X�F�ڭ̱N�P�ɬ��z�w�Ƥ@��Ѳ��g���H�A���U�z�����Ѳ�����A�C���R��N����������B��1%�@���S�ҡC<br>
              2.�ק�K�X�ɡA�Ĥ@�ӱK�X���±K�X�A�ĤG�ӱK�X���s�K�X�I�ѩ󥻨t�Ϊ��ΥΤ�[�K����k�A���|�X�{�Q�H�������Ʊ��A�ҥH�K�X�i�H�����ק�I 
              <br>
              3.�ߤ�ɱb���i�H���N��g�A���O�C�Ӧ���H�h���i�H�ӽФ@�ӪѲ��b���I</td>
          </tr>
            <td width="47%"> 
          </table>
</td></tr>
</table>
</form>
</body>
</html>
