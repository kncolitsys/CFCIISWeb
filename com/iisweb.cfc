<!---
		iisAdmin.cfc
		-----------------------------
		Written and maintained by:
			Lance Knight
			http://www.knivis.com/
			lknight@knivis.com
		-----------------------------
		Facilitates creation of iisweb via access to the command line with CFExecute.
				Uses technique described here:
		http://technet2.microsoft.com/windowsserver/en/library/6b672523-789a-4523-8f27-f745802db40b1033.mspx?mfr=true
		-----------------------------
		Usage: 
		<cfset iis = createObject("component", "com.iisweb").init(cscriptpath, vscriptpath) />
		<cfset result = iis.create('name','path','ipaddress','hostheader') />
		<cfset result = iis.query('blank for list || sitename || Metabase Path') />
		<cfset result = iis.delete('sitename || Metabase Path') />
		<cfset result = iis.start('sitename || Metabase Path') />
		<cfset result = iis.stop('sitename || Metabase Path') />
		<cfset result = iis.pause('sitename || Metabase Path') />
		-----------------------------
		Notes:
		 - Full documentation will reside on the wiki: http://iisvdir.riaforge.org/wiki
		 - Does not currently understand system variables like %SystemRoot%, so the true fully qualified path is 
		   required (C:\winnt\system32\...)
		 - At time of writing, for windows server 2000, CScriptPath should be: C:\winnt\system32\cscript.exe
		 - At time of writing, for windows server 2003, CScriptPath should be: C:\windows\system32\cscript.exe
		 - init() returns this
		 - query('blank for list || sitename || Metabase Path') returns a query object () DISCRIPTION ,HOSTHEADER ,IDENTIFIER ,IPNUMBER ,PORT ,STATE 
		 - create('name','path','ipaddress','hostheader','port') returns a Struct object
		 - delete('sitename || Metabase Path') returns a Struct object
		 - start('sitename || Metabase Path') returns a Struct object
		 - stop('sitename || Metabase Path') returns a Struct object
		 - pause('sitename || Metabase Path') pause returns a Struct object
		 - All other functions return true/false for success
		-----------------------------
		History:
		Date        Developer   	Notes
		==========  ============	==========================================
		03 01 2009  Lance Knight    Created
		03 07 2009  Lance Knight	Fixed space in name  bug

--->
<cfcomponent displayname="iisAdmin" output="no">
    <cfset variables.instance = StructNew()/>
    <cfset variables.instance.cscriptPath = "" />
    <cfset variables.instance.vscriptPath = "" />
	
    <!--- structure used to return multiple bits of information to the caller --->
	<cfset variables.instance.returnStruct = StructNew() />
	<cfset variables.instance.returnStruct.success = false />
	<cfset variables.instance.returnStruct.msg = "" />
	<cfset variables.instance.returnStruct.detail = "" />
    <!--- function init --->
 	<cffunction name="init" access="public" output="no" hint="initialize component data" returntype="any">
		<cfargument name="cscriptPath" required="no" default="c:\windows\system32\cscript.exe" type="string" />
		<cfargument name="vscriptPath" required="no" default="c:\windows\system32\" type="string" />
		<cfset variables.instance.cscriptPath = arguments.cscriptPath />
		<cfset variables.instance.vscriptPath = arguments.vscriptPath />
		<cfreturn this />
 	</cffunction>
     <!--- function query --->
	<cffunction name="query" output="no" access="public" hint="return a query of existing site" returntype="query">
        <cfargument name="queryString" required="no" type="string" default="" />
		<cfset var result = QueryNew("discription,identifier,state,ipnumber,port,hostheader", "varchar,varchar,varchar,varchar,varchar,varchar") />
		<cfset var strResult = "" />
		<cfset var row = "" />
		<cfset var discription = "" />
		<cfset var args = "" />
		<cfset var dividerLoc = 0 />
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
		
        <cfset args = variables.instance.vscriptPath & "iisweb.vbs /query " & queryString>
		<cftry>
		<cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		   <cfcatch type="any">
		        <!--- unknown error --->
		        <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>
		<cfif Find('not found', strResult)>
		<cfreturn result />
		<cfabort>
		</cfif>
		
		<cfset dividerLoc = FindNoCase("==", strResult) + 78 /><!--- returns 78 consecutive equal signs before first row of data --->
		<cfset strResult = right(strResult, len(strResult) - dividerLoc) />
		<cfloop list="#strResult#" delimiters="#chr(10)##chr(13)#" index="row">
		<cfif len(row)><!--- ignore empty rows --->
			<cfset dividerLoc = FindNoCase("(", row)-1 />
			<cfif dividerLoc gt 1>
			<cfset discription = trim(left(row, dividerLoc)) />
			<cfset row = trim(right(row, len(row) - dividerLoc)) />
			<cfset rowArray = ListToArray(row ," ")>
			<cfset QueryAddRow(result, 1) />
			<cfset QuerySetCell(result, "discription" , discription) />
			<cfset identifier = replace(replace(rowArray[1],'(',''),')','')>
			<cfset QuerySetCell(result, "identifier" , trim(identifier) ) />
			<cfset QuerySetCell(result, "state" , trim(rowArray[2])) />
			<cfset QuerySetCell(result, "ipnumber" , trim(rowArray[3])) />
			<cfset QuerySetCell(result, "port" , trim(rowArray[4])) />
			<cfset QuerySetCell(result, "hostheader" , trim(rowArray[5])) />
			</cfif>
		</cfif>
		</cfloop>
		
		<cfreturn result />
	</cffunction>
    <!--- function create --->
	<cffunction name="create" output="no" access="public" hint="create a site in iis" returntype="struct">
		<cfargument name="Name" type="string" required="yes">	
		<cfargument name="Path" type="string" required="yes">
        <cfargument name="IPAddress" type="string" required="yes">
		<cfargument name="HostHeader" type="string" required="yes">
		<cfargument name="Port" type="numeric" default="80">
		<cfargument name="dontstart" type="boolean" default="false">
	
		<cfset var args = "" />
		
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
		
        <cfset args = variables.instance.vscriptPath & 'iisweb.vbs /create ' & arguments.Path & ' "'& arguments.Name &'" /b ' & arguments.Port & ' /i ' & arguments.IPAddress & ' /d ' & arguments.HostHeader />
		<cfif arguments.dontstart eq true>
			<cfset args = args & ' /dontstart'>
		</cfif>
		 <cftry>
		 <cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		 <cfcatch type="any">
		     <!--- unknown error --->
		     <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	      </cfcatch>
		  </cftry>
		 <cfset variables.instance.returnStruct.detail = strResult />
		<cfif findNoCase("/? for help", strResult) gt 0>
			<cfset variables.instance.returnStruct.success = false />
			<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail.">
		<cfelseif findNoCase("Done", strResult) gt 0>
			<cfset variables.instance.returnStruct.success = true />
			<cfset variables.instance.returnStruct.msg = "create" />
				<cfset dividerLoc = FindNoCase("Done.", strResult) + 4/>
				<cfset strResult = trim(right(strResult, len(strResult) - dividerLoc)) />
				<cfset variables.instance.returnStruct.detail = strResult />
				<cfloop list="#strResult#" delimiters="#chr(10)##chr(13)#" index="row">
					<cfif len(row)><!--- ignore empty rows --->
					<cfset var = ListToArray(row ,'=')>
						<cfif trim(var[1]) eq 'Server'>
						<cfset variables.instance.returnStruct.Server = trim(var[2]) /> 
						<cfelseif trim(var[1]) eq 'Site Name'>
						<cfset variables.instance.returnStruct.SiteName = trim(var[2]) />
						<cfelseif trim(var[1]) eq 'Metabase Path'>
						<cfset variables.instance.returnStruct.MetabasePath = trim(var[2]) /> 
						<cfelseif trim(var[1]) eq 'IP'>
						<cfset variables.instance.returnStruct.IP = trim(var[2]) />
						<cfelseif trim(var[1]) eq 'Host'>
						<cfset variables.instance.returnStruct.Host = trim(var[2]) /> 
						<cfelseif trim(var[1]) eq 'Root'>
						<cfset variables.instance.returnStruct.Root = trim(var[2]) /> 
						<cfelseif trim(var[1]) eq 'App Pool'>
						<cfset variables.instance.returnStruct.AppPool = trim(var[2]) /> 
						<cfelseif trim(var[1]) eq 'Status'>
						<cfset variables.instance.returnStruct.Status = trim(var[2]) />
						</cfif>
					</cfif>
				</cfloop>
		</cfif>	
		<cfreturn variables.instance.returnStruct />
	</cffunction>
    <!--- function delete --->
	<cffunction name="delete" output="no" access="public" hint="delete a site in iis" returntype="struct">
		<cfargument name="Name" type="string" required="yes">	
		<cfset var args = "" />
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
		
        <cfset args = variables.instance.vscriptPath & 'iisweb.vbs /delete' & ' "'& arguments.Name &'"' />
		<cftry>
		<cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		   <cfcatch type="any">
		        <!--- unknown error --->
		        <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>
		<cfset variables.instance.returnStruct.detail = strResult />
	    <cfif findNoCase("has been DELETED", strResult)>
		    <cfset variables.instance.returnStruct.success = true />
		    <cfset variables.instance.returnStruct.msg = "" />
	    <cfelseif findNoCase("/?", strResult)>
	    	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail." />
		<cfelseif findNoCase("not found", strResult)>
		   	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Error. Not found." />
	    </cfif> 
		<cfreturn variables.instance.returnStruct />
	</cffunction>
		 <!--- function start --->
	<cffunction name="start" output="no" access="public" hint="starts a site in iis" returntype="struct">
	<cfargument name="Name" type="string" required="yes">	
		<cfset var args = "" />
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
        <cfset args = variables.instance.vscriptPath & 'iisweb.vbs /start' & ' "'& arguments.Name &'"' />
		<cftry>
		<cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		   <cfcatch type="any">
		        <!--- unknown error --->
		        <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>

		<cfset variables.instance.returnStruct.detail = strResult />
	    <cfif findNoCase("started", strResult)>
		    <cfset variables.instance.returnStruct.success = true />
		    <cfset variables.instance.returnStruct.msg = "" />
	    <cfelseif findNoCase("/?", strResult)>
	    	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail." />
		<cfelseif findNoCase("not found", strResult)>
		   	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Error. Not found." />
	    </cfif> 
		
	<cfreturn variables.instance.returnStruct  />
	</cffunction>
	 <!--- function stop --->
	<cffunction name="stop" output="no" access="public" hint="stops a site in iis" returntype="struct">
		<cfargument name="Name" type="string" required="yes">	
		<cfset var args = "" />
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
        <cfset args = variables.instance.vscriptPath & 'iisweb.vbs /stop' & ' "'& arguments.Name &'"' />
		<cftry>
		<cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		   <cfcatch type="any">
		        <!--- unknown error --->
		        <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>

		<cfset variables.instance.returnStruct.detail = strResult />
	    <cfif findNoCase("STOPPED", strResult)>
		    <cfset variables.instance.returnStruct.success = true />
		    <cfset variables.instance.returnStruct.msg = "" />
	    <cfelseif findNoCase("/?", strResult)>
	    	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail." />
		<cfelseif findNoCase("not found", strResult)>
		   	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Error. Not found." />
	    </cfif> 
		
	<cfreturn variables.instance.returnStruct  />
	</cffunction>
	    <!--- function pause --->
	<cffunction name="pause" output="no" access="public" hint="pauses a site in iis" returntype="struct">
		<cfargument name="Name" type="string" required="yes">	
		<cfset var args = "" />
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vscriptPath)>
        	<cfthrow message= "Error: You must first set the VBS script path. This is usually 'c:\windows\system32\'. Process aborted." />
		</cfif>
        <cfset args = variables.instance.vscriptPath & 'iisweb.vbs /pause' & ' "'& arguments.Name &'"' />
		<cftry>
		<cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
		   <cfcatch type="any">
		        <!--- unknown error --->
		        <cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>
		<cfset variables.instance.returnStruct.detail = strResult />
	    <cfif findNoCase("PAUSED", strResult)>
		    <cfset variables.instance.returnStruct.success = true />
		    <cfset variables.instance.returnStruct.msg = "" />
	    <cfelseif findNoCase("/?", strResult)>
	    	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail." />
		<cfelseif findNoCase("not found", strResult)>
		   	<cfset variables.instance.returnStruct.success = false />
	    	<cfset variables.instance.returnStruct.msg = "Error. Not found." />
	    </cfif> 
	<cfreturn variables.instance.returnStruct />
	</cffunction>
</cfcomponent>