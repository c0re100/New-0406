
<%money = clng(request.form("money"))
my=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="select * from �Τ� where �m�W='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
            MsgBox "�A���O���򤤤H!"
          location.href = "../jh.asp"
       </script><%
	conn.close
	response.end
else%>
<!--#include file="data1.asp"--><%
randomize()
         waittime=(rnd()*10000000)
         for i=0 to waittime
         next
sql="select * from �Ы� where ��D='" & my & "' or  ��Q='" & my & "'"
set rs=conn.execute(sql)
yin=rs("�Ȩ�")
if money>yin then%>
<script language=vbscript>
                         MsgBox "���~�I�z�����w�̨S������h�Ȩ�C"
                         location.href = "xiaowu5.asp"
                    </script>
	<%elseif money>50000 or money<1 then%>
       <script language=vbscript>
                         MsgBox "���~�I�z�C���̦h�����5�U�C"
                         location.href = "xiaowu5.asp"
                    </script><%
		else
       sql="update �Ы� set �Ȩ�=�Ȩ�-'" & money & "' where ��D='" & my & "' or  ��Q='" & my & "'"
       rs=conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
       sql="update �Τ� set �Ȩ�=�Ȩ�+'" & money & "' where �m�W='" & my & "'"
       rs=conn.execute(sql)
       conn.close
       if yin<0 then
       sql="update �Τ� set �Ȩ�=�Ȩ�+'" & yin & "' where �m�W='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       end if
	Response.Redirect "xiaowu5.asp"
			conn.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>
