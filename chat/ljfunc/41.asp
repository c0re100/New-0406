<%'�dip
function seeip(to1,toname)
if info(2)<7 then
	Response.Write "<script language=JavaScript>{alert('�޲z�ݭn7�Ū��~�i�H�d��ip���I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�dip�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'�d�Τ�ip
rs.open "select lastip,regip FROM �Τ� WHERE  �m�W='"&toname&"'",conn
ip1=rs("lastip")   '�̫�ip
ip2=rs("regip")    '�`�Uip
sip=split(rs("lastip"),".")
sip1=split(rs("regip"),".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'�d�̫�ip���a��
rs.close
rs.open "select ��a,���� FROM �a�} WHERE ip1<="& num &" and ip2>="&num,conn
if rs.eof or rs.bof then
	country="�Ȭw"
	city="����"
else
	country=rs("��a")
	city=rs("����")
end if
if country="" then country="����"
if city="" then city="����"
last="�a��:"& country &" ����:"& city
rs.close
'�d�`�Uip���a��
num=cint(sip1(0))*256*256*256+cint(sip1(1))*256*256+cint(sip1(2))*256+cint(sip1(3))-1
rs.open "select ��a,���� FROM �a�} WHERE ip1<="& num &" and ip2>="&num,conn
if rs.eof or rs.bof then
	country="�Ȭw"
	city="����"
else
	country=rs("��a")
	city=rs("����")
end if
if country="" then country="����"
if city="" then city="����"
reg="�a��:"& country &" ����:"& city
seeip="[�dip]"& toname &"���{�bip�a�}��:"&ip1&",�D�w���G"&last&"  �`�Uip�a�}��:"&ip2 &"�D�w���G"&reg
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>