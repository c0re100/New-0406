<%
Set Conn = Server.CreateObject("ADODB.Connection")
jiutiandata="DBQ="+server.mappath("../../news.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open jiutiandata
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4
Const adUseServer = 2
Const adUsenews = 3
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
pagesize=10
Public Arr_News_Type(12,2)
Arr_News_Type(1,1)="�F�C�޲z����"
Arr_News_Type(1,2)="0"
Arr_News_Type(2,1)="�F�C�|����k"
Arr_News_Type(2,2)="0"
Arr_News_Type(3,1)="�F�C�������i"
Arr_News_Type(3,2)="0"
Arr_News_Type(5,1)="�F�C�������U"
Arr_News_Type(5,2)="1"
Arr_News_Type(6,1)="�J���g"
Arr_News_Type(6,2)="10"
Arr_News_Type(7,1)="�o�]�g"
Arr_News_Type(7,2)="10"
Arr_News_Type(8,1)="���a�g"
Arr_News_Type(8,2)="10"
Arr_News_Type(9,1)="�m�Z�g"
Arr_News_Type(9,2)="10"
Arr_News_Type(10,1)="��ѽg"
Arr_News_Type(10,2)="10"
Arr_News_Type(11,1)="�ʪZ�g"
Arr_News_Type(11,2)="10"
Arr_News_Type(12,1)="�޲z�g"
Arr_News_Type(12,2)="10"
%>