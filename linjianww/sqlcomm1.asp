<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(2)<10  then Response.Redirect "manerr.asp?id=100"
%>
<html>
<head>
<title>sql���O�t��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.sqlstr.value;
revisedTitle = addTitle;
document.PostTopic.sqlstr.value=revisedTitle;
document.PostTopic.sqlstr.focus();
return; }
</script>

<body bgcolor="#FFFFFF" background="../IMAGES/b115.jpg" text="#FFFFFF" topmargin="0">
<blockquote>
  <p align="center"> <br>
  </p>
</blockquote>
<%if sjjh_name=Application("sjjh_user") then%>
<form method="POST" action="sqlcommok.asp" name="PostTopic">
<div align="center">
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">�ֳt���O</OPTION> 
<OPTION value="update �Τ� set �Ȩ�=0,�s��=0 where (�Ȩ�+�s��)>2000000000">[�M���겣�j��20���Τ�겣��0]</OPTION>
<OPTION value="delete * from �ާ@�O��" >[�R���Ҧ��ާ@�O��]</OPTION>
<OPTION value="delete * from �H�R" >[�R���Ҧ��H�R�O��]</OPTION>
<OPTION value="delete * from �޲z�O��" >[�R���Ҧ��޲z�O��]</OPTION>
<OPTION value="delete * from ��O��" >[�R���Ҧ���O���O��]</OPTION>
<OPTION value="delete * from �ӭ�" >[�R���Ҧ��ӭްO��]</OPTION>
<OPTION value="delete * from �X���" >[�R���Ҧ��X��ްO��]</OPTION>
<OPTION value="delete * from ���B" >[�R���Ҧ����B�O��]</OPTION>
<OPTION value="delete * from �ϥΧޯ� where aszcc=0" >[�R���D�|�����[�ۦ�]</OPTION>
<OPTION value="delete * from �ϥΧޯ� where ����<100" >[�R�����[���Ťp��100]</OPTION>
<OPTION value="delete * from ���~ where �|��=0" >[�R���Ҧ��D�|�����~]</OPTION>
<OPTION value="delete * from ������� where �|��=0" >[�R���Ҧ��D�|���O�I�d]</OPTION>
<OPTION value="delete * from �Τ� where �|������=0" >[�R���Ҧ��D�|��]</OPTION>
</SELECT><br>
    <br>
    �п�J���O�G 
    <input type="text" name="sqlstr" size="50" maxlength="280">
    <br>
    <input type="submit" value="����" name="B1" class="p9">
    <input type="reset" name="Submit" value="�M��">
  </div>
  </form>
 <%else%>
  <div align="center">�z�D���v�����A�����T��ϥΡI</div>
 <%end if%>
</body>
</html>