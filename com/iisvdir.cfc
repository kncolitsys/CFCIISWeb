<!---
		iisvdir.cfc
		-----------------------------
		Written and maintained by:
			Adam Tuttle
			http://tuttletree.com/nerdblog/
			adam.tuttle@perficient.com
		-----------------------------
		Facilitates creation of IIS Virtual Directories via access to the command line with CFExecute.
		Uses technique described here:
		http://technet2.microsoft.com/windowsserver/en/library/6b672523-789a-4523-8f27-f745802db40b1033.mspx?mfr=true
		-----------------------------
		Usage: 
		<cfset iis = createObject("component", "iisvdir").init(cscriptpath, sitename) />
		<cfset result = iis.create(sitepath, vdirname, vdirpath) />
		-----------------------------
		Notes:
		 - Full documentation will reside on the wiki: http://iisvdir.riaforge.org/wiki
		 - Does not currently understand system variables like %SystemRoot%, so the true fully qualified path is 
		   required (C:\winnt\system32\...)
		 - At time of writing, for windows server 2000, CScriptPath should be: C:\winnt\system32\cscript.exe
		 - At time of writing, for windows server 2003, CScriptPath should be: C:\windows\system32\cscript.exe
		 - init() returns this
		 - List() returns a query object with 2 string columns: name, path
		 - All other functions return true/false for success
		 - Soon, all other functions will return variables.instance.returnStruct, which will include success flag, 
		   err#, err message, and details, which will contain all of the output from the execute(s). 
		-----------------------------
		History:
		Date        Developer   Notes
		==========  ==========  ==========================================
		2008-04-03  ATuttle     Created
		2008-04-08	ATuttle		Updated to include list, edit, delete functions
		2008-04-09	ATuttle		Updated created, edit, delete functions to return struct instead of bool, to allow
								more detailed data to be returned.
--->
<cfcomponent displayname="iisvdir" output="no">
	
    <cfset variables.instance = StructNew()/>
    <cfset variables.instance.cscriptPath = "" />
    <cfset variables.instance.siteName = "" />
    <cfset variables.instance.sitePath = "" />
    <cfset variables.instance.vdirName = "" />
    <cfset variables.instance.vdirPath = "" />

	<!--- structure used to return multiple bits of information to the caller --->
	<cfset variables.instance.returnStruct = StructNew() />
	<cfset variables.instance.returnStruct.success = false />
	<cfset variables.instance.returnStruct.msg = "" />
	<cfset variables.instance.returnStruct.detail = "" />
    
    <cffunction name="init" access="public" output="no" hint="I initialize component data" returntype="any">
    	<cfargument name="cscriptPath" required="no" default="c:\windows\system32\cscript.exe" type="string" />
        <cfargument name="siteName" required="yes" type="string" />
        <cfset variables.instance.cscriptPath = arguments.cscriptPath />
        <cfset variables.instance.siteName = arguments.siteName />
        <!--- add leading slash for site path, if it's not there. --->
        <cfif len(variables.instance.sitePath) and left(variables.instance.sitePath, 1) neq "/">
        	<cfset variables.instance.sitePath = "/" & variables.instance.sitePath />
        </cfif>
       <cfreturn this />
    </cffunction>
    
    <cffunction name="create" output="no" access="public" hint="I create the virtual directory" returntype="struct">
        <cfargument name="sitePath" required="no" type="string" default="/" />
        <cfargument name="vdirName" required="yes" type="string" />
        <cfargument name="vdirPath" required="yes" type="string" />
    	<cfset var result = "" />
        <cfset var args = "" />
        <cfset variables.instance.sitePath = arguments.sitePath />
        <cfset variables.instance.vdirName = arguments.vdirName />
        <cfset variables.instance.vdirPath = arguments.vdirPath />
		<!--- vdir name must not have leading slash --->
		<cfif left(variables.instance.vdirName, 1) eq "/">
			<cfset variables.instance.vdirName = right(variables.instance.vdirName, len(variables.instance.vdirName) - 1) />
		</cfif>
		<!--- trap missing data errors --->
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.siteName)>
        	<cfthrow message="Error: You must first set the Site Name that you want to add the Virtual Directory for. Ex: 'Default Web Site'. Process aborted." />
		</cfif>
        <cfif not len(variables.instance.vdirName)>
        	<cfthrow message="Error: You must first set the name of the new Virtual Directory to be created. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.vdirPath)>
        	<cfthrow message="Error: You must first set the path of the new Virtual Directory to be created. Ex: 'c:\inetpub\wwwroot\mydir'. Process aborted." />
        </cfif>
        <cfset args = Replace(variables.instance.cscriptPath, "cscript.exe", "iisvdir.vbs") />
        <cfset args = args & " /create " & variables.instance.siteName & variables.instance.sitePath />
        <cfset args = args & " " & variables.instance.vdirName & " " & variables.instance.vdirPath />
		<cftry>
	        <cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="result" />
	        <cfcatch type="any">
	        	<cfthrow message="#cfcatch.message#<br/>#cfcatch.detail#">
	        </cfcatch>
	    </cftry>
		<cfset variables.instance.returnStruct.detail = result />
		<cfif findNoCase("already exists", result) gt 0>
			<cfset variables.instance.returnStruct.success = false />
			<cfset variables.instance.returnStruct.msg = "An IIS virtual directory with this name ('#arguments.vdirName#') already exists.">
		<cfelseif findNoCase("/? for help", result) gt 0>
			<cfset variables.instance.returnStruct.success = false />
			<cfset variables.instance.returnStruct.msg = "Unrecognized error. See detail.">
		<cfelseif findNoCase("Metabase Path =", result) gt 0>
			<cfset variables.instance.returnStruct.success = true />
			<cfset variables.instance.returnStruct.msg = "" />
		</cfif>
	    <cfreturn variables.instance.returnStruct />
    </cffunction>
    
	<cffunction name="list" output="no" access="public" hint="I return a query of existing virtual directories" returntype="query">
        <cfargument name="sitePath" required="no" type="string" default="" />
		<cfset var result = QueryNew("name,path", "varchar,varchar") />
		<cfset var strResult = "" />
		<cfset var row = "" />
		<cfset var args = "" />
		<cfset var junk = ArrayNew(1) />
		<cfset var dividerLoc = 0 />
        <cfset variables.instance.sitePath = arguments.sitePath />
		<!--- trap missing data errors --->
		<cfif not len(variables.instance.cscriptPath)>
        	<cfthrow message="Error: You must first set the Cscript path. This is usually 'c:\windows\system32\cscript.exe'. Process aborted." />
        </cfif>
        <cfif not len(variables.instance.siteName)>
        	<cfthrow message="Error: You must first set the Site Name that you want to add the Virtual Directory for. Ex: 'Default Web Site'. Process aborted." />
		</cfif>
        <cfset args = Replace(variables.instance.cscriptPath, "cscript.exe", "iisvdir.vbs") />
        <cfset args = args & " /query " & variables.instance.siteName & variables.instance.sitePath />
		<cftry>
	        <cfexecute name="#variables.instance.cscriptPath#" arguments="#args#" timeout="30" variable="strResult" />
	        <cfcatch type="any">
		        <!--- unknown error --->
	        	<cfthrow message="#cfcatch.message#<br>#cfcatch.detail#" />
	        </cfcatch>
	    </cftry>
	    <cfif FindNoCase("/query /? for help", strResult)>
		    <!--- error connecting to IIS --->
	    	<cfthrow message="#strResult#" />
	    </cfif>
	    <cfif FindNoCase("No virtual sub-directories", strResult)>
	    	<!--- no virtual directories exist, return empty query --->
	    	<cfreturn result />
	    </cfif>
		<cfset dividerLoc = FindNoCase("==", strResult) + 78 /><!--- returns 78 consecutive equal signs before first row of data --->
		<cfset strResult = right(strResult, len(strResult) - dividerLoc) />
		<cfloop list="#strResult#" delimiters="#chr(10)##chr(13)#" index="row">
			<cfif len(row)><!--- ignore empty rows --->
				<cfset dividerLoc = FindNoCase(" ", row) />
				<cfset QueryAddRow(result, 1) />
				<cfset QuerySetCell(result, "name", trim(left(row, dividerLoc))) />
				<cfset row = trim(right(row, len(row) - dividerLoc)) />
				<cfset QuerySetCell(result, "path", row) />
			</cfif>
		</cfloop>
		<cfreturn result />
	</cffunction>

	<cffunction name="delete" output="no" access="public" hint="I delete an existing virtual directory" returntype="struct">
		<cfargument name="sitePath" type="string" required="false" default="/" />
		<cfargument name="vdirName" type="string" required="true" />
		<cfset var args = "" />
		<cfset var path = "" />
		<cfset var result = false />
		<cfset var strResult = "" />
	    <cfset variables.instance.returnStruct.success = false />
        <cfset variables.instance.sitePath = arguments.sitePath />
        <cfset variables.instance.vdirName = arguments.vdirName />
		<!--- trap missing data errors --->
        <cfif not len(variables.instance.vdirName)>
        	<cfthrow message="Error: You must first set the name of the new Virtual Directory to be created. Process aborted." />
        </cfif>
		<!--- for delete, leading slash is needed for vdirName, because it is part of the path token, not its own token --->
		<cfif left(variables.instance.vdirName, 1) neq "/"><cfset variables.instance.vdirName = "/" & variables.instance.vdirName /></cfif>
		<cfif left(variables.instance.sitePath, 1) neq "/"><cfset variables.instance.sitePath = "/" & variables.instance.sitePath /></cfif>
		<cfset path = variables.instance.siteName & variables.instance.sitePath & variables.instance.vdirName />
		<cfset path = ReplaceNoCase(path, "//", "/", "ALL") />
        <cfset args = Replace(variables.instance.cscriptPath, "cscript.exe", "iisvdir.vbs") />
        <cfset args = args & " /delete " & path />
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
	    </cfif> 
		<cfreturn variables.instance.returnStruct />
	</cffunction>
	
	<cffunction name="edit" output="no" access="public" hint="I edit an existing virtual directory" returntype="struct">
        <cfargument name="sitePath" required="no" type="string" default="/" />
        <cfargument name="oldVdirName" required="yes" type="string" />
        <cfargument name="newVdirName" required="yes" type="string" />
        <cfargument name="newVdirPath" required="yes" type="string" />
		<cfset var delResult = delete(sitePath:arguments.sitePath, vdirName:arguments.oldVdirName) />
		<cfset var createResult = "" />
		<cfset var saveDetail = "" />
		<cfif delResult.success eq false>
			<cfreturn delResult />
		</cfif>
		<cfset saveDetail = delResult.detail />
		<cfset createResult = create(sitePath:arguments.sitePath, vdirName:arguments.newVdirName, vdirPath:arguments.newVdirPath) />
		<cfset createResult.detail = saveDetail & chr(10) & chr(13) & "---------------------------------" & createResult.detail />
		<cfreturn createResult />
	</cffunction>
	
</cfcomponent>