# Gets the locations and values for the specified field(s) 

###################
#region Parameters#
param(		
	[string] $SiteCollectionUrl = "http://spdevader",
	[string] $Fields = "AT,Author"
)
#endregion

##################
#region Utilities#
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -EA SilentlyContinue) -eq $null) 
{
   Write-Host "Caricamento SharePoint cmdlets..." -ForegroundColor Magenta
   Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue
}
Start-SPAssignment -Global

class Logger
{
	[string] $FilePath	
	[string] $FileName
	[string] $SecondaryFilePath
	
	hidden Initializer() { $this.Initializer((Split-Path $Script:MyInvocation.MyCommand.Path), $Script:MyInvocation.MyCommand.Name) }
	hidden Initializer([string]$path, [string]$name) 
	{		
		$name = "$name.txt"
		$this.FilePath = $path + "\" + $name
		$this.FileName = $name
	}
	
	Logger() { $this.Initializer() }	
	Logger([string]$fileName) { $this.Initializer((Split-Path $Script:MyInvocation.MyCommand.Path), $fileName) }
    Logger([string]$fileName, [bool]$overwrite) { $this.Initializer((Split-Path $Script:MyInvocation.MyCommand.Path), $fileName); "" | Out-File $this.FilePath -Encoding string -NoNewline; }
    
    SimpleWrite([string]$message)
    {
        $message | Out-File $this.FilePath -Append -Encoding string
    }

	WriteLog([string]$message)
	{
		$this.WriteLog($message, [System.ConsoleColor]::Gray)
	}
	
	WriteLog([string]$message, [System.ConsoleColor]$customColor)
	{
		$timeStamp = Get-Date -UFormat "%Y/%m/%d-%H:%M:%S"	
		Write-Host $message -ForegroundColor $customColor
		"[$timeStamp] " + $message | Out-File $this.FilePath -Append -Encoding string
		if($this.SecondaryFilePath) 
		{ 
			"[$timeStamp] " + $message | Out-File $this.SecondaryFilePath -Append -Encoding string 
		}
	}
}

# Example: $stage1Progress = [Progress]::new($site.AllWebs.Count, $true); $stage1Progress.WriteProgress("Analyzing web $webUrl")
class Progress
{
	[double] $Step
	[double] $CurrentIndex
	[double] $CurrentCount
	[double] $MaxCount
	[int] $ProgressID

	hidden Initializer() { $this.Initializer(100) }
	hidden Initializer([double]$maxCount) { $this.Initializer($maxCount, $false) }
	hidden Initializer([double]$maxCount, [bool]$showMultipleProgress) 
	{
		$this.MaxCount = $maxCount
		$this.Step = [Math]::Pow($maxCount / 100, -1)
		$this.CurrentIndex = [Math]::Round($this.Step, 2) + 1
		$this.CurrentCount = 1
		if($showMultipleProgress) { $this.ProgressID = [System.Random]::new().Next() } else { $this.ProgressID = 1 }
	}
	
	Progress() { $this.Initializer(100) }	
    Progress([double]$maxCount) { $this.Initializer($maxCount) }
	Progress([double]$maxCount, [bool]$showMultipleProgress) { $this.Initializer($maxCount, $showMultipleProgress) }
	
    WriteProgress() { $this.WriteProgress("Computing..") }
	
    WriteProgress([string]$activityMessage)
	{	
		$status = "Status: ($($this.CurrentCount)/$($this.MaxCount)) $($this.CurrentIndex)% Complete:"
		if($this.CurrentIndex -gt 100) { $this.CurrentIndex = 100 }
		Write-Progress -Id $this.ProgressID -Activity $activityMessage -Status $status -PercentComplete $this.CurrentIndex;
		$this.CalculateProgress()
	}
	
	hidden CalculateProgress()
	{
		if($this.CurrentIndex + $this.Step -lt 100) 
		{ 
			$this.CurrentIndex = [Math]::Round($this.CurrentIndex + $this.Step, 2) 
		} 
		else 
		{ 
			$this.CurrentIndex = 100 
		}
		$this.CurrentCount = [Math]::Ceiling([Math]::Round($this.CurrentIndex / $this.Step))
	}
}
#endregion

##################
#region Main functions
function RecursiveSearch([Microsoft.SharePoint.SPWeb]$web)
{
    
}
#endregion

#############
#region Exec#

try
{
	$Log = [Logger]::new("log")	
    $report = [Logger]::new("report", $true)
    $report.SimpleWrite("Field,Web,List,ID,Value")    
	$site = Get-SPSite $SiteCollectionUrl
	$webs = $site.AllWebs
    $stage1Progress = [Progress]::new($webs.Count, $true);    		
    $fields = $Fields.split(",")

    if($fields.Length -eq 0) 
    { 
        $Log.WriteLog("No fields specified.")
        return
    }

	foreach($web in $webs)
	{
        $stage1Progress.WriteProgress("Analyzing web $($web.Url)")
        $Log.WriteLog("Analyzing web $($web.Url)")	        
        foreach($field in $fields)
        {
	        $query = [Microsoft.SharePoint.SPQuery]::new()
	        $query.Query = "<Where><IsNotNull><FieldRef Name='$field' /></IsNotNull></Where>"            
            $listsWithField = $web.Lists | ? { $_.Fields.ContainsField("$field") }
            foreach($list in $listsWithField)
            {
                $Log.WriteLog("Analyzing list $($list.Title) for field $field")
                $items = $list.GetItems($query)    
                foreach($item in $items)
                {
                    $line = "$field,$($web.Url),$($list.Title),$($item['ID']),$($item[$field])"
                    $report.SimpleWrite($line)
                }
            } 
        }                       
	}	
}
Catch
{	
	Write-Host ""
	Throw $_.Exception
	    $Log.WriteLog("The following error occured: $($_.Exception.Message)", ([System.ConsoleColor]::Red))
}

Write-Host ""
Write-Host ""
#$Log.WriteLog("WSP Deployment script END")
Write-Host ""

Stop-SPAssignment -Global
#endregion