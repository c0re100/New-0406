<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
  s=now()-session("tao_time")
  if int(s*1000000)<30 then
  response.write "�Ф��n��r�A�y���I�A����W���֤H�O�I"
  response.end
  end if
%>
<!--#include file="data.asp"-->
<%
  randomize timer
  t=int(rnd*100)
    if t<30 then
    session("tao_time")=now()
  randomize timer
  jin=int(rnd*100)
  sql="select * from �^���� where �֦���='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �^����(�֦���,��,��,��) values('" & info(0) & "'," & jin & ",0,0)"
       conn.execute(sql)
       else
       sql="update �^���� set ��=��+" & jin & " where �֦���='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update �q�� set �`�I=�`�I+1,�`�y=�`�y+" & jin& " where �q��='��'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "���߱z!��o" & <%=jin%> & "�I���q!"
  location.href="index.asp"
</script>
<%
    elseif t<65 and t>=30 then
  session("tao_time")=now()
  randomize timer
  yin=int(rnd*200)
  sql="select * from �^���� where �֦���='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �^����(�֦���,��,��,��) values('" & info(0) & "',0," & yin & ",0)"
       conn.execute(sql)
       else
       sql="update �^���� set ��=��+" & yin & " where �֦���='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update �q�� set �`�I=�`�I+1,�`�y=�`�y+" & yin& " where �q��='��'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "���߱z!��o" & <%=yin%> & "�I���q!"
  location.href="index.asp"
</script>
<%
    else
  session("tao_time")=now()
  randomize timer
  tong=int(rnd*300)
  sql="select * from �^���� where �֦���='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into �^����(�֦���,��,��,��) values('" & info(0) & "',0,0," & tong & ")"
       conn.execute(sql)
       else
       sql="update �^���� set ��=��+" & tong & " where �֦���='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update �q�� set �`�I=�`�I+1,�`�y=�`�y+" & tong & " where �q��='��'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "���߱z!��o" & <%=tong%> & "�I���q!"
  location.href="index.asp"
</script>
<%
    end if
%>