<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")

%>

<!--#include file="data1.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("nick"),"�p�ϦW�r") 
%>
<%
IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request.form("nick")
if sheepname="" or len(sheepname)="" then
%>
<script language="Vbscript">
msgbox"�p�Ī��W�r��g�����T�I",0,"����"
history.back
</script>
<%
else
rs.open"select dellifeday from rules",conn,1,1
if rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�S���W�h�s�b�I');location.href = 'javascript:history.go(-1)';}</script>"
else
dellifeday=rs("dellifeday")
rs.close
rs.open"select logintoday,feeddate,life,hungry,clean,sheephappy,sheephealth from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
if rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z���O�o�Ӥp�Ī��D�H�I');location.href = 'javascript:history.go(-1)';}</script>"
else
logintoday=rs("logintoday")
feeddate=rs("feeddate")
life=rs("life")
hungry=rs("hungry")
clean=rs("clean")
sheephappy=rs("sheephappy")
sheephealth=rs("sheephealth")
rs.close
if datediff("d",logintoday,date)<>0 then
tempdatetoday=date
conn.execute"update sheep set logintoday='"&tempdatetoday&"' where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
temptime=datediff("d",feeddate,date)
templife=life-dellifeday*temptime
temphungry=hungry-dellifeday*temptime/2
if temphungry<0 then
temphungry=0
end if
tempclean=clean-dellifeday*temptime/2
if tempclean<0 then
tempclean=0
end if
tempsheephappy=sheephappy-dellifeday*temptime/2
if tempsheephappy<0 then
tempsheephappy=0
end if
tempsheephealth=sheephealth-dellifeday*temptime/2
if tempsheephealth<0 then
tempsheephealth=0
end if
'conn.execute"update sheep set clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"',hungry='"&temphungry&"' where sheepname='"&sheepname&"' and id="&tempid&" and user='"&info(0)&"'"
conn.execute"update sheep set life='"&templife&"',hungry='"&temphungry&"',clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"' where sheepname='"&sheepname&"' and user='"&info(0)&"'"
end if
rs.open"select life,clean,sheephappy,sheephealth,hungry from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
if rs("life")<=0 or rs("clean")<=0 or rs("sheephappy")<=0 or rs("sheephealth")<=0 or rs("hungry")<=0 then
rs.close
conn.execute"delete * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'"
conn.close
 Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �t�� FROM �Τ� WHERE �m�W='" & info(0) & "'"
rs.open sql,conn,1,3
peier=rs("�t��")
conn.execute("update �Τ� set �Ȩ�=�Ȩ�-1000000,�p��='�L',�ĦW='�L'  where �m�W='"&info(0)&"'")
conn.execute("update �Τ� set �Ȩ�=�Ȩ�-1000000,�p��='�L',�ĦW='�L'  where �m�W='"&peier&"'")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
msgbox"�ѩ�z�S����߷��U�p�ġI�z���p�Ĥw�g���F�A���s�A�i�@�ӧa�C",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select life,sheep001,sheep002,sheepname,sheepdate,hungry,clean,sheephealth,sheepxb,sheephappy,wugong,leinl,gongji,feedsheepday,milk from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
old=int(now()-rs("sheepdate"))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">


<HTML><HEAD><TITLE>�F�C����-����|��|</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="jh/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=80%>
  <TBODY> 
  <TR>
    <TD background=jh/history_top_bg.gif vAlign=top width=100%> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="100%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">&nbsp;</TD>
          <TD height=25 vAlign=top width="33%">
            <DIV align=center><font color="#000000">����|��|</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="36%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=jh/history_table_bg.gif height=158 vAlign=top width="100%"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=85%>
        <TBODY> 
        <TR>
          <TD vAlign=top width="100%"><BR>
            <TABLE align=center background=jh/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=table1 width="99%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b>�i���򴣥ܡj</b></font><br>
                    �o��O�޾i�p�Ī��a��A�A�p�ĴN��b�o��A�ݵۦo�L�~�L�{���ͬ��ۡA�Aı�o�b����̤��b�t�W�C���L�A�C�ѱo�ӳo���U�o�@�A���M�o�|�������@<br>
                  </P>
                  <table border="0" width="97%" cellspacing="1" cellpadding="0"
                height="330">
                    <tr> 
                      <td class="p9" colspan="4" height="18">
                        <div align="center">
                          <p>�Ĥl�L�R�G<%=rs("sheep001")%> <br>
                            �Ĥl�L���G<%=rs("sheep002")%></p>
                          <p> 
                            <%if rs("sheepxb")="�k" then %>
                            �_����l:<%=rs("sheepname")%> <font size="2"><img src="baby3.gif" width="144" height="133" align="top"> 
                            <%else%>
                            �_���k��G <%=rs("sheepname")%> <img src="girl15.gif" width="45" height="57" align="top"> 
                            <%end if%>
                            </font></p>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td class="p9" width="25%" height="18">��-�p�ĥͩR�ȡG</td>
                      <td class="p2" width="25%" height="18"> 
                        <%
if rs("life")>80 then
typelife="���d"
end if
if rs("life")<=80 and rs("life")>60 then
typelife="�}�n"
end if
if rs("life")<=60 and rs("life")>40 then
typelife="��z"
end if
if rs("life")<=40 and rs("life")>20 then
typelife="�ͯf"
end if
if rs("life")<=20 and rs("life")>0 then
typelife="�f�M"
end if
if rs("life")<=0 then
typelife="���`"
end if
response.write typelife
                      %>
                      </td>
                      <td class="p2" width="25%" height="18"><span class="p9">��-�p�Ī��~��</span>�G</td>
                      <td class="p2" width="25%" height="18"><%=old%></td>
                    </tr>
                    <tr> 
                      <td class="p3" width="25%" height="18"><span class="p9">��-�Ⱦj�{��</span>�G</td>
                      <td class="p3" width="25%" height="18"><%=rs("hungry")%></td>
                      <td class="p3" width="25%" height="18"><span class="p9">��-�M��{��</span>�G</td>
                      <td class="p3" width="25%" height="18"><%=rs("clean")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">��-���d�{��</span>�G</td>
                      <td class="p2" width="25%" height="18"><%=rs("sheephealth")%></td>
                      <td class="p9" width="25%" height="18">��-�ֵּ{�סG</td>
                      <td class="p2" width="25%" height="18"><%=rs("sheephappy")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">��-�Z�\�{��</span>�G</td>
                      <td class="p2" width="25%" height="18"><%=rs("wugong")%></td>
                      <td class="p9" width="25%" height="18">��-���O�{�סG</td>
                      <td class="p2" width="25%" height="18"><%=rs("leinl")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">��-�����{��</span>�G</td>
                      <td class="p2" width="25%" height="18"><%=rs("gongji")%></td>
                      <td class="p9" width="25%" height="18">&nbsp;</td>
                      <td class="p2" width="25%" height="18">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="17"><font color="#FF0000">��-���a�����q�i�G</font></td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="71">��-�A�q�L���P����k���U�A���p�ķ|�o�줣�P���u�@�q�A�A�u�����ɶ����n�ָg��A�C���@�߲ӭP���[��A���p�ĨC�Ѫ����p�A�ׯ��A�̫Ĥl���i������N����C�A���p�ĨC�ѥi�޲z8��(�@�P��),���U�p�Ĥ��өP���H�W�N�i�H��t��|���u�åB�i�H�a�ۧA�̪��Ĥl��������C<br>
                        ��-�`�G�C���ӮƤp�Ī�O50���A��p�Ĥl���|���Ȥ�����@���Ȭ�0�ɡA�p�ķ|���h�C</td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepeat.asp">
                          <p align="center"> 
                            <input type="submit" value="�ޥ�" 
                        name="B1" style="font-family: �s�ө���; font-size: 12px">
                            <font              
                        color="#E18C59">�Ⱦj+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepclean.asp">
                          <p align="center"> 
                            <input type="submit" value="�~��" 
                        name="B12" style="font-family: �s�ө���; font-size: 12px">
                            <font              
                        color="#E18C59">�M��+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepsun.asp">
                          <p align="center"> 
                            <input type="submit" value="����" 
                        name="B12" style="font-family: �s�ө���; font-size: 12px">
                            <font              
                        color="#E18C59">���d+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheeppei.asp">
                          <p align="center"> 
                            <input type="submit" value="����" 
                        name="B12" style="font-family: �s�ө���; font-size: 12px">
                            <font              
                        color="#E18C59">�ּ�+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepwu.asp">
                          <p align="center"> 
                            <input type="submit" value="�m�Z" 
                        name="B1" style="font-family: �s�ө���; font-size: 12px">
                            <font color="#E18C59">�Z�\+50</font> 
                            <input type="hidden"              
                        name="num2" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepnl.asp">
                          <p align="center"> 
                            <input type="submit" value="�R��" 
                        name="B1" style="font-family: �s�ө���; font-size: 12px">
                            <font color="#E18C59">���O+50</font> 
                            <input type="hidden"              
                        name="num22" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepgj.asp">
                          <p align="center"> 
                            <input type="submit" value="�Գ�" 
                        name="B1" style="font-family: �s�ө���; font-size: 12px">
                            <font color="#E18C59">����+50</font> 
                            <input type="hidden"              
                        name="num222" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38">&nbsp;</td>
                    </tr>
                  </table>
                  <div align="center">
                    <p>&nbsp;</p>
                    <p><a href="javascript:window.close()"><font color="#FFFFFF">�������f</font></a> 
                    </p>
                  </div>
                  <P class=p9> 
                    <%if rs("feedsheepday")>=3 then%>
                    <br>
                  </P>
                  <table border="0" width="470" cellspacing="1" cellpadding="0" 
            height="60">
                    <tr> 
                      <td class="p9" width="100%"><b><font color="#FF0000">���ߡI���ѱz���p�Ħb�t��|���u�F�A���u�ɶ���</font></b><%=rs("milk")%><b><font color="#FF0000">�Ӥu�@���</font></b><br>
                        <font color="#000000"><b><font color="#FF0000">�֨�</font></b></font><a href="indexsheep.asp"><font color="#FFFFFF">�t��|</font></a><b><font color="#FF0000">�h�u��a�I</font></b></td>
                    </tr>
                  </table>
                  <P class=p9 align="center"><a href="javascript:window.close()"><font color="#FFFFFF">�������f</font></a> 
                  </P>
                  <P align=center class="p9">&nbsp;</P>
                  </TD></TR></TBODY></TABLE>
            <P align=center>&nbsp; </P>
          </TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="100%"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE>
<%rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if                  
                
%> </BODY></HTML>
<%                 
            
end if                  
end if                  
end if                  
end if                  
end if               
%> 

