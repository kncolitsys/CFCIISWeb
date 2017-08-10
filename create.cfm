<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<cfset iisObj = createObject("component", "com.iisweb").init() />
<cfdump var="#variables.iisObj.create('temp1.knivis.net','C:\Websites\knivis\knivis.com\www','209.50.147.102','temp1.knivis.net')#">

</body>
</html>
