<#
.NOTES
	Name: PFtoHTML.ps1
	Original Author: Chris Heilman
	Requires: Exchange Management Shell (Exchange Server 2010) and administrator rights on the Exchange server and Public Folders.
	Version: 1.0 -- 11/02/2017

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
	
	
.SYNOPSIS
	A HTML output for your Legacy Public Folders

.DESCRIPTION
	A HTML output for your Legacy Public Folders
	
	
.EXAMPLE
	.\PFtoHTML.ps1

	What is your Public Folder server?: <ServerName>

	<Generates Report>

#>


#================
#Parameters
#================

$server = Read-Host "What is your Public Folder server?"

$folder = "\"

#================
#Command
#================

Get-Date | Select-Object Date | convertTo-HTML -head $a -Title "Public Folder Report" | Out-File c:report.htm

#================
#Variable addition
#================

$a= "<script src=c:sorttable.js type=text/javascript></script>"

$a = $a + "<style>"

$a = $a + "BODY{background-color:gray;}"

$a = $a + "TABLE.sortable thead {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"

$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:yellow}"

$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:white}"

$a = $a + "</style>"

#================
#Output
#================

Get-PublicFolderStatistics -Server $server| Select-Object admindisplayname, creationtime, Lastmodificationtime, itemcount, totalitemsize, servername | where-object {$_.Lastmodificationtime -lt "11/1/2017 12:00:00 AM"} | convertTo-HTML  -Body "<H2>Public Folders</H2>"   |  Out-File c:report.htm -append

Get-PublicFolder -Recurse -Server $server -Identity $folder | Select-Object Name, MailEnabled, HasSubFolders, IssueWarningQuota, MaxItemSize  | convertTo-HTML -head $a -Title "Public Folder Report" -Body "<H2>Public Folder Information</H2>" |  Out-File c:report.htm -append

(Get-Content C:report.htm) | Foreach-Object {$_ -replace "<table>", "<table class =sortable>"} | Set-Content C:report.htm

(Get-Content C:report.htm) | Foreach-Object {if($_ -like "*TotalItemSize*"){$_ -replace "t</th><th>", "t</th><th>"} else{$_}} | Set-Content C:report.htm

(Get-Content C:report.htm) | Foreach-Object {if($_ -like "*ItemCount*"){$_ -replace "ModificationTime</th><th>", "ModificationTime</th><th>"} else{$_}} | Set-Content C:report.htm

(Get-Content C:report.htm) | Foreach-Object {if($_ -like "*TotalItemSize*"){$_ -replace "count</th><th>", "count</th><th>"} else{$_}} | Set-Content C:report.htm

#================
#Displaying Report
#================

.\report.htm