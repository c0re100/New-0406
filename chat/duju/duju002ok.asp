<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
Response.Write "<script Language=Javascript>alert('�n��սХ��h�j�U�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
'f=Minute(time())
'if f>30 then
'	Response.Write "<script language=JavaScript>{alert('�{�b�����աI��ծɶ����C�p�ɪ��e30����,�Ҧp�G17:00�}�l17:30����!!');}</script>"
'	Response.End 
'end if
num=trim(request.form("num"))
for t=1 to len(num)
abc=mid(num,t,1)
if abc<>"0" and abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�U�`�ШϥμƦr�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
next
yinliang=abs(int(num))
if Request.Form("submit")="�j" then
tt="�j"
end if
if Request.Form("submit")="�p" then
tt="�p"
end if
if Request.Form("submit")="��" then
tt="��"
end if
if Request.Form("submit")="��" then
tt="��"
end if
select case tt
	case "�j"
		duboimg="<img src='../jhimg/da.gif'>"
	case "�p"
		duboimg="<img src='../jhimg/xiao.gif'>"
	case "��"
		duboimg="<img src='../jhimg/dan.gif'>"
	case "��"
		duboimg="<img src='../jhimg/shuang.gif'>"
case else
	Response.Write "<script language=JavaScript>{alert('��ާ@���G�j�B�p�B��B���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end select
	if yinliang<100000 or yinliang>2000000 then
		Response.Write "<script language=JavaScript>{alert('�̤֩�G10�U��A�̦h200�U�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
'�}�l�P�_
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select �Ȩ� FROM �Τ� WHERE id="&info(9),conn
	if rs("�Ȩ�")<yinliang then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A���o��h�����ܡH�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where �Ȩ�<>0")
	durs=tmprs("�ƶq")
set tmprs=nothing
	if durs>=10 then
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�@["&durs&"]�`���b���ݶ}���A�y��A�U�`�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.open "select top 1 �m�W FROM �b�u��� WHERE ����='���a'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�{�b�S�����a�A����i��b�u��աI�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM �b�u��� WHERE �m�W='"& info(0) &"'",conn
	if not(rs.eof) or not(rs.bof) then
		if rs("����")="���a" then
			temp=info(0)&"�A�{�b�O���a�r�A�A�n�d����r!"
		else
			temp=info(0)&"�A��["&rs("��j�p")&"]"&rs("�Ȩ�")&"���۶}�a�I"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&yinliang&" where id="&info(9)
	conn.execute "insert into �b�u���(�m�W,����,�Ȩ�,��j�p,�ɶ�) values ('"&info(0) &"','���a',"&yinliang&",'"&tt&"',now())"	
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where ��j�p='�j'")
	dars=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where ��j�p='�p'")
	xiaors=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where ��j�p='��'")
	danrs=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where ��j�p='��'")
	shuangrs=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where �Ȩ�<>0")
	durs=tmprs("�ƶq")


'�}���F
if durs>=10 then
	Randomize
	m1 = Int(6 * Rnd + 1)
	Randomize
	m2 = Int(6 * Rnd + 1)
	Randomize
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'�d����a
rs.open "select �m�W FROM �b�u��� WHERE ����='���a'",conn
zhuangjia=rs("�m�W")
rs.close

'�\�l�B�z
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select �Ȩ� FROM �b�u��� WHERE �Ȩ�<>0 and ����<>'���a'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("�Ȩ�")
	rs.movenext
	loop
	rs.close
	conn.execute "update �Τ� set �s��=�s��+"&yingyin&" where �m�W='"& zhuangjia &"'"
	xiazhu="���a�}�G<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�p�G<font color=red>"& sjdubo &"</font>�I!���a�}�X�\�l  �q���I���a�G<font color=red>"&zhuangjia&"</font>Ĺ�o�G<font color=red>"&yingyin&"</font>��I"
		Application.Lock
	Application("ljjh_dubozhang")=""
Application.UnLock
set rs=nothing
		conn.close
		set conn=nothing
		call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�⦨�\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if

'��l�Ƽƾ�
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'�}�����B�z
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select �Ȩ�,�m�W FROM �b�u��� WHERE ��j�p='��'",conn
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("�Ȩ�")
		shuangname=shuangname&rs("�m�W")&" "
		conn.execute "update �Τ� set �s��=�s��+"&rs("�Ȩ�")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set �s��=�s��-"& rs("�Ȩ�") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �b�u��� where  ��j�p='��'"
else
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select �Ȩ�,�m�W FROM �b�u��� WHERE ��j�p='��'",conn
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("�Ȩ�")
		danname=danname&rs("�m�W")&" "
		conn.execute "update �Τ� set �s��=�s��+"&rs("�Ȩ�")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set �s��=�s��-"& rs("�Ȩ�") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �b�u��� where  ��j�p='��'"
end if


'��}�j�p�B�z
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select �Ȩ�,�m�W FROM �b�u��� WHERE ��j�p='�p'",conn
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("�Ȩ�")
		xiaoname=xiaoname&rs("�m�W")&" "
		conn.execute "update �Τ� set �s��=�s��+"&rs("�Ȩ�")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set �s��=�s��-"& rs("�Ȩ�") &" where �m�W='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "delete * from �b�u��� where  ��j�p='�p'"

else
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select �Ȩ�,�m�W FROM �b�u��� WHERE ��j�p='�j'",conn
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("�Ȩ�")
		daname=daname&rs("�m�W")&" "
		conn.execute "update �Τ� set �s��=�s��+"&rs("�Ȩ�")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set �s��=�s��-"& rs("�Ȩ�") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �b�u��� where  ��j�p='�j'"
end if

'��ѤU�骺�Τ�Ȩ�
	rs.open "select �Ȩ� FROM �b�u��� WHERE �Ȩ�<>0 and ����<>'���a'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("�Ȩ�")
	rs.movenext
	loop
	rs.close
	conn.execute "update �Τ� set �s��=�s��+"&yingyin&" where �m�W='"& zhuangjia &"'"
	conn.execute "delete * from �b�u��� where  �m�W<>''"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
'shuangyinliang=0
'shuangname=""
'danyinliang=0
'danname=""
'xiaoyinliang=0
'xiaoname=""
'dayinliang=0
'daname=""
	duboname=shuangname&danname&xiaoname&daname

	
	xiazhu="���a�}�G<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�p�G<font color=red>"& sjdubo &"</font>�I!"&danshuang&daxiao&"�`�U�`�G<font color=red>"&zong&"</font>��A���a�G<font color=red>["&zhuangjia&"]</font>���J�G<font color=red>"&yingyin &"</font>��,�ߥX�G<font color=red>"&pei&"</font>��A�X�p�G<font color=red>"&yingyin-pei&"</font>��,�@���G<font color=red>"&duboname&"</font>���a���B�㤤�I"
Application.Lock
	Application("ljjh_dubozhang")=""
Application.UnLock
	set rs=nothing
	conn.close
	set conn=nothing
	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�⦨�\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End

end if
xiazhu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�q�ۤv���y�]�̮��X�@�i�䲼,�o�N�a��:�o�O:<font color=red>"& yinliang &"</font>�⪺�F�C����Ȧ�䲼<img src='pic/Dz01.gif'>�A�ڶR"& duboimg &"�I�A�@�w�����I�{�b�U�`�p�U�G��j�G<font color=red>"& dars &"</font>�ӡA��p:<font color=red>"& xiaors &"</font>��,���G<font color=red>"&danrs&"</font>��,����:<font color=red>"&shuangrs&"</font>�ӡI�ٮt:<font color=red>"&(10-durs)&"</font>�Ӷ}���I"

	set rs=nothing	
	conn.close
	set conn=nothing


	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�⦨�\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End

sub say(xiazhu)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)=info(0)
sd(195)=info(0)
sd(199)="<font color=green>�i�Ȩ�䧽�j</font><font color=#FFC01F>"& xiazhu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
%>