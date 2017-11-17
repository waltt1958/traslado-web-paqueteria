<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>CIERRE HASTA AYER</title>
</HEAD>
<BR>
<BR>
<H3 align= "right">Hoy es: <%=date%></H3>

<body bgcolor="#FF00FF" >

<%

Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

conectarOEP.execute ("UPDATE imposicionCI SET imposicionCI.CERRADO_CI = 'FORZADO', imposicionCI.FECHA_CIERRE_CI = Now() WHERE imposicionCI.CERRADO_CI= 'NO' AND imposicionCI.FECHA_CARGA_CI < DATE()")

conectarOEP.execute ("UPDATE lectPAQ SET lectPAQ.FECHA_CIERRE_OEP = Now(), lectPAQ.CERRADO_OEP = 'FORZADO' WHERE lectPAQ.CERRADO_OEP='NO' AND lectPAQ.FECHA_CARGA_oep < date()")

conectarOEP.close


%>


<BR>
<BR>
<BR>
<BR>
<BR>
<BR>


<center><h1><p align="center"><b><font size="12"> TERMINO EL PROCESO DE CIERRE FORZADO DE LOS PAQUETES QUE NO FUERON LEIDOS EN EL CI O EN PAQUETERIA</font size></b></p> </h1></center>
</body>

</SCRIPT>
</HTML>