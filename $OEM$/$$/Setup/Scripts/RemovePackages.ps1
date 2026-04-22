$selectors = @(
	'Microsoft.Microsoft3DViewer';
	'Microsoft.BingSearch';
	'Clipchamp.Clipchamp';
	'Microsoft.549981C3F5F10';
	'Microsoft.Windows.DevHome';
	'MicrosoftCorporationII.MicrosoftFamily';
	'Microsoft.WindowsFeedbackHub';
	'Microsoft.GetHelp';
	'Microsoft.Getstarted';
	'microsoft.windowscommunicationsapps';
	'Microsoft.WindowsMaps';
	'Microsoft.MixedReality.Portal';
	'Microsoft.BingNews';
	'Microsoft.MicrosoftOfficeHub';
	'Microsoft.Office.OneNote';
	'Microsoft.OutlookForWindows';
	'Microsoft.MSPaint';
	'Microsoft.People';
	'Microsoft.PowerAutomateDesktop';
	'MicrosoftCorporationII.QuickAssist';
	'Microsoft.SkypeApp';
	'Microsoft.MicrosoftSolitaireCollection';
	'Microsoft.MicrosoftStickyNotes';
	'Microsoft.Todos';
	'Microsoft.WindowsSoundRecorder';
	'Microsoft.Wallet';
	'Microsoft.BingWeather';
	'Microsoft.Xbox.TCUI';
	'Microsoft.XboxApp';
	'Microsoft.XboxGameOverlay';
	'Microsoft.XboxGamingOverlay';
	'Microsoft.XboxIdentityProvider';
	'Microsoft.XboxSpeechToTextOverlay';
	'Microsoft.GamingApp';
	'Microsoft.YourPhone';
	'Microsoft.ZuneVideo';
);
$getCommand = {
  Get-AppxProvisionedPackage -Online;
};
$filterCommand = {
  $_.DisplayName -eq $selector;
};
$removeCommand = {
  [CmdletBinding()]
  param(
    [Parameter( Mandatory, ValueFromPipeline )]
    $InputObject
  );
  process {
    $InputObject | Remove-AppxProvisionedPackage -AllUsers -Online -ErrorAction 'Continue';
  }
};
$type = 'Package';
$logfile = 'C:\Windows\Setup\Scripts\RemovePackages.log';
& {
	$installed = & $getCommand;
	foreach( $selector in $selectors ) {
		$result = [ordered] @{
			Selector = $selector;
		};
		$found = $installed | Where-Object -FilterScript $filterCommand;
		if( $found ) {
			$result.Output = $found | & $removeCommand;
			if( $? ) {
				$result.Message = "$type removed.";
			} else {
				$result.Message = "$type not removed.";
				$result.Error = $Error[0];
			}
		} else {
			$result.Message = "$type not installed.";
		}
		$result | ConvertTo-Json -Depth 3 -Compress;
	}
} *>&1 >> $logfile;