# Searches all the documents in the specified Site Collection

###################
#region Parameters#
param(		
	[string] $SiteCollectionUrl = "http://spdevader",
	[string] $Fields = "Created,Author,Modified,Editor"
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

#############
#region Main Process#

try
{
	$Log = [Logger]::new("log")	
    $report = [Logger]::new("report", $true)
	$headers = "Area,Subsite,ListTitle,ListType,ItemID,Title,ContentTypeName,Link"
	if($Fields.Length -gt 0)
	{
		$headers = "$headers,$Fields"
	}
    $report.SimpleWrite($headers)    
	
	$site = Get-SPSite $SiteCollectionUrl
	$webs = $site.AllWebs
    $websProgressTracker = [Progress]::new($webs.Count, $true);    		
    $customFields = $Fields.split(",")

	foreach($web in $webs)
	{
        $websProgressTracker.WriteProgress("Analyzing web [$($web.Url)]")
        $Log.WriteLog("Analyzing web $($web.Url)")	        		
		$lists = $web.Lists | Where-Object {$_.BaseType -eq ([Microsoft.SharePoint.SPBaseType]::GenericList) -or $_.BaseType -eq ([Microsoft.SharePoint.SPBaseType]::Survey)}
		$listsProgressTracker = [Progress]::new($lists.Count, $false)		
		foreach($list in $lists)
		{				
			$listsProgressTracker.WriteProgress("Analyzing list [$($list.Title)]")
			$Log.WriteLog("Analyzing list $($list.Title)")
			$items = $list.Items    
			foreach($item in $items)
			{
				$csvLine = "`"$($item.Web.Title)`",$($item.Web.ServerRelativeUrl),`"$($list.Title)`",`"$($list.BaseType)`",$($item.ID),`"$($item.Title)`",`"$($item.ContentType.Name)`",`"$("$($web.Url)/$($list.RootFolder.Url)/DispForm.aspx?ID=$($item.ID)")`""
				foreach($customField in $customFields) {
					$fieldValue = $item["$customField"]
					if($fieldValue -ne $null)
					{
						$fieldValue = $fieldValue.ToString().Replace('"', '""').Replace("`n", "").Replace("`n", "")
					}
					$csvLine = "$csvLine,`"$fieldValue`""
				}				
				$report.SimpleWrite($csvLine)
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
Write-Host ""

Stop-SPAssignment -Global
#endregion