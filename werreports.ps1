Set-ExecutionPolicy Unrestricted
cls
#################################
# Model
class User{    

    [string]$Name

    #User hält eine Liste von report-Objekten in einer ArrayList
    $Reports = [System.Collections.ArrayList]::New()
    [Computer]$pc
   
    #Konstruktormehode die beim Erzeugen von User aufgerufen wird
    User([string]$uname)
    {
        $this.Name = $uname
        $this.pc = [Computer]::new()
    }
}
class Computer{

    [string]$macadresse
    [string]$osname
    [string]$hostname

    Computer()
    {
	    $mac = get-wmiobject -class "Win32_NetworkAdapterConfiguration" | Where {$_.IpEnabled -Match "True"} | Select MacAddress
	    $this.macadresse = $mac[0].MacAddress
	    $this.osname = (Get-WmiObject Win32_OperatingSystem).Name
	    $this.hostname = (Get-WmiObject Win32_OperatingSystem).CSName
    }
}

[Computer]$pc = [Computer]::new()

class Report{

    [string]$ReportID
    [int]$ReportType
    [string]$EventType
    [string]$EventTime
    [string]$BucketID
    [string]$Appname

    Report($repid, $repType, $evType, $evTime, $buckId, $appnam)
    {
        $this.ReportID = $repid
        $this.ReportType = $repType
        $this.EventType = $evType
        $this.EventTime = $evTime
        $this.BucketID = $buckId
        $this.Appname = $appnam
    }

}

######################################################
######################################################
# Alte Welt
# Funktionen müssen die Daten besorgen und in die 
#Objekte des Models reinschreiben
function GetUsers($Benutzer)
{

	$temp =  Get-ChildItem "C:\Users" | Select-Object Name
    foreach($element in $temp)
    {
        [User]$t = [User]::new($element.name)

        #$users += $t
        $Benutzer.add($t)
    }
}

function GetWERPath($_user)
{


	$_ordner = "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportArchive";
	
	[array]$werdata = $NULL;
	
	if((Test-Path -path $_ordner) -eq $true) # Testen ob es den Ordner gibt
	{
		$_ordner = Get-ChildItem $_ordner | Where-Object {$_.mode -match "d"} ;# Holt die Ordner aus dem Verzeichnis
		if($_ordner -ne $null) # Prüfen ob es unterordner gibt !!!!!!!!
		{
			foreach($_reportarch in $_ordner)
			{
				$_tempreportarch = $_reportarch.Name; # Namen wieder Temporär selektieren
				$_ordner2 = "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportArchive\$_tempreportarch";
				if((Test-Path $_ordner2) -eq $true)# Testen ob es den Ordner gibt
				{
					$_unterordner = Get-ChildItem $_ordner2 | Where-Object {$_.name -eq "Report.wer"}; # nur Dateien welche Report.wer heisen werden angezeigt
					foreach($_error in $_unterordner)
					{
						$werdata += "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportArchive\$_tempreportarch\$_error";
					}
				}
			}
		}
	}
	
	$_ordner = "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportQueue";
	if((Test-Path $_ordner) -eq $true) # Testen ob es den Ordner gibt
	{
		$_ordner = Get-ChildItem $_ordner | Where-Object {$_.mode -match "d"};
		if($_ordner -ne $null) # Prüfen ob es unterordner gibt !!!!!!!!
		{
			foreach($_reportarch in $_ordner)
			{
				$_tempreportarch = $_reportarch.Name; # Namen wieder Temporär selektieren
				$_ordner2 = "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportQueue\$_tempreportarch";
				if((Test-Path $_ordner2) -eq $true)# Testen ob es den Ordner gibt
				{
					$_unterordner = Get-ChildItem $_ordner2 | Where-Object {$_.name -eq "Report.wer"}; # nur Dateien welche Report.wer heisen werden angezeigt
					foreach($_error in $_unterordner)
					{
						$werdata += "C:\Users\$_user\AppData\Local\Microsoft\Windows\WER\ReportQueue\$_tempreportarch\$_error";
					}
				}
			}
		}
	}
	return $werdata;
}

function GetReportInnerData($_path,$_methode)
{
	$reportid= Select-String -Encoding Unicode -Path $_path -AllMatch "$_methode" | select line # sucht nach dem Stichpunkten welche in Methode übegeben werden im Path welcher über &_path kommt
	if ($reportid.Line -ne $null)
	{
		$reportidResult = $reportid.Line.Split("=");			 
	 	$reportid = $reportidResult[1];
		return $reportid;
	 	#### $reportid enthält die report ID 	
	}
	else
	{
        #Wa macht diese zeile ??
		$reportidResult = $reportid[1].Line.Split("=");
		$reportid = $reportidResult[1];
		return $reportid;
		#### $reportid enthält die report ID
	}
}

#copy by ref $Benutzer hat einen neuen Namen, nämlich $Users
#alles was mit Users gemacht wird, wirkt sich auf Benutzer aus
#$Users ist die ArrayList, die unten übergebenn wurde
function GetReportData($Benutzer)
{
    
	#ArrayList wird von der GetUsers-Funktion mit Objektwen der Klasse User gefüllt 
	GetUsers $Benutzer
    #Write-Host $Benutzer
	foreach($_user in $Benutzer) # $_users hier stehen alle benuter vom Rechner und werden nach $_user geschrieben in der Schleife
	{
        #Liste der WER-Dateien pro user
		$paths = GetWERPath $_user.Name;
		
        foreach($_path in $paths) # $_users hier stehen alle benuter vom Rechner und werden nach $_user geschrieben in der Schleife
		{
			if($_path -ne $null)
			{                                   #Pfad zur WER_Datei  u. gesuchter Key
				$_reportid = GetReportInnerData $_path "ReportIdentifier";
				$_reporttype = GetReportInnerData $_path "ReportType";
				$_eventtype = GetReportInnerData $_path "EventType"; # z.B AppCrash
				$_eventtime = GetReportInnerData $_path "EventTime";
				$_bucketid = GetReportInnerData $_path "Response.BucketId";
				$_appname = GetReportInnerData $_path "AppName"; 
           
                [Report]$rep = [Report]::new($_reportid, $_reporttype, $_eventtype, $_eventtime, $_bucketid, $_appname)
                $_user.Reports += $rep
            }
		}
	}
}

###################################
###################################
#MySQL-Funktionen
[void] [System.Reflection.Assembly]::Loadfrom("c:\Program Files (x86)\MySQL\Connector.NET 6.9\Assemblies\v4.0\\MySql.Data.dll");
$connstring = "server=localhost;uid=watson;pwd='watson';Database=watson_11fi3"
$con = new-object Mysql.Data.MySqlClient.MySqlConnection;
$con.connectionstring = $connstring

#Commandobject
$global:command = new-Object mySql.Data.MySqlClient.MySqlCommand;
$command.connection = $con
                                                       #Variable !!!
$sql="insert into computer(MAC, OSName, Hostname) values ('macadresse','bla', 'blub');" 

function setUser($Benutzer)
{
    foreach($user in $Benutzer)
    {

        $name = $user.Name

        $sql = "replace into User(Anmeldename) values('$name');"
        $command.commandText = $sql
        try {
            $command.executeNonQuery()
        }
        
        catch {
        
            Write-Debug "Fehler beim Schreiben des Users"
        } 
    }
}

function setComputer([Computer]$comp)
{
    $macad = $comp.macadresse
    $os = $comp.osname
    $hostname = $comp.hostname


    $sql = "replace into computer(MAC, OSNAME, Hostname) values ('$macad', '$os', '$hostname');"
   
    $command.commandText = $sql
        
    try {
            $command.executeNonQuery()
        }
        catch {
            Write-Debug "Fehler beim Schreiben des Computers"
        }
}

function setReports($Benutzer)
{
    foreach($_user in $Benutzer)
    {
        $reps = $_user.Reports
        foreach($report in $reps)
        {
            $rid = $report.ReportID
            $rtype = $report.ReportType
            $evtime = $Report.EventTime
            #$evType = $report.EventType
            $b_id = $report.BucketID
            $App = $report.Appname

            $name = $_user.Name
            $macadresse = $_user.PC.macadresse

            $sql = "insert into report(ReportID, Appname, EventTime, BucketID, ReportType, User, Computer) "
            $sql += "values('$rid', '$App', '$evtime','$b_id','$rtype', '$name', '$macadresse');"

            $command.CommandText = $sql

             try {
                $command.executeNonQuery()
            }
            catch {
                throw $_.Exception.Message
             Write-Debug "Fehler beim Schreiben der Reports"
            }         
   # [string]$ReportID
   # [int]$ReportType
   # [string]$EventType
   # [string]$EventTime
   # [string]$BucketID
   # [string]$Appname
        }
    }
}

function SaveData($Benutzer)
{
    $con.open()
    #User-Tabelle füllen
    setUser $Benutzer
    setComputer $Benutzer[0].pc 
    setReports $Benutzer
    $con.close()
}

function remove_werreports($Users)
{
    foreach($user in $Users)
    {
    $filepath = "C:\Users\" + $user.Name + "\AppData\Local\Microsoft\Windows\Wer\*.*"
    if(Test-Path $filepath)
        {
        remove-item -path $filepath -recurse -force
        }
    }
}

function getReportsFromDatabase
{

    $sql = "select reportID, Appname from report;"


    $command.CommandText = $sql
    $con.Open();

    $reader = $command.ExecuteReader()

    while ($reader.Read()) 
        {
        write-host $reader["reportID"] $reader["Appname"];
        }
    $con.Close();
}

###################################
###################################
#Ablaufsteuerung
$Benutzer = [System.Collections.ArrayList]::New()

#GetReportData $Benutzer

try{
    #SaveData $Benutzer

}
catch{
    $_.Exception.Message
   
}

#remove_werreports $Benutzer
Write-Host "Fertig"

getReportsFromDatabase
