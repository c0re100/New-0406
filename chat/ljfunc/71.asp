<%'�Z�ǭ׷�
function wuxiu(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
Response.End 
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if

randomize()
rnd1=int(rnd*800)+100
rnd2=int(rnd*800)+120
rs.close
rs.open "select �׷ү� FROM �Z�\ WHERE   ����='" & info(5) & "' and �Z�\='"&fn1&"'",conn
if    rs.eof and   rs.bof then 
Response.Write "<script language=JavaScript>{alert('�A���S��["&fn1&"]�o�˪��Z�\�r�I');}</script>"
Response.End
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
xiu=rs("�׷ү�")
if xiu<20000*xiu  then
cha=20000*info(2)-xiu
conn.execute "update �Z�\ set �׷ү�=�׷ү�+1 where �Z�\='"&fn1&"' and ����='"&info(5)&"'"
conn.execute "update �Τ� set ��O=��O-"&rnd1&",���O=���O-"&rnd2&",�ާ@�ɶ�=now() where id="&info(9)
wuxiu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��["&info(5)&"]["&fn1&"]���׷�,��O���h<font color=red>-"&rnd1&"</font>�I,���O���h<font color=red>-"&rnd2&"</font>�I,["&fn1&"]���Ҥp���A�ٮt"&cha&"�I����!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
exit function
else
conn.execute "update �Z�\ set �׷ү�=1,�ŧO=�ŧO+1 where �Z�\='"&fn1&"' and ����='"&info(5)&"'"
wuxiu=info(0) & "<bgsound src=wav/dz.wav loop=1>�i��["&info(5)&"]["&fn1&"]���׷�,�g�L�@�f�׷�["&fn1&"]�Z�Ǵ���1��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
exit function
end if
end  function

%>