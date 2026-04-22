$scripts = @(
	{
		reg.exe add "HKU\DefaultUser\Software\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f;
	};
	{
		Remove-ItemProperty -LiteralPath 'Registry::HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDriveSetup' -Force -ErrorAction 'Continue';
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f;
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d 0 /f;
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Hidden" /t REG_DWORD /d 1 /f;
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarAl /t REG_DWORD /d 0 /f;
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Policies\Microsoft\Windows\Explorer" /v DisableSearchBoxSuggestions /t REG_DWORD /d 1 /f;
	};
	{
		# Disable screensaver and screen lock for default user profile
		reg.exe add "HKU\DefaultUser\Control Panel\Desktop" /v ScreenSaveActive /t REG_SZ /d "0" /f;
		reg.exe add "HKU\DefaultUser\Control Panel\Desktop" /v ScreenSaverIsSecure /t REG_SZ /d "0" /f;
		reg.exe add "HKU\DefaultUser\Control Panel\Desktop" /v ScreenSaveTimeOut /t REG_SZ /d "0" /f;
		reg.exe add "HKU\DefaultUser\Control Panel\Desktop" /v "SCRNSAVE.EXE" /t REG_SZ /d "" /f;
	};
	{
		reg.exe add "HKU\DefaultUser\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v "UnattendedSetup" /t REG_SZ /d "powershell.exe -WindowStyle Normal -NoProfile -ExecutionPolicy Bypass -Command \""Get-Content -LiteralPath 'C:\Windows\Setup\Scripts\UserOnce.ps1' -Raw | Invoke-Expression;\""" /f;
	};
);

& {
  [float] $complete = 0;
  [float] $increment = 100 / $scripts.Count;
  foreach( $script in $scripts ) {
    Write-Progress -Activity 'Running scripts to modify the default user’’s registry hive. Do not close this window.' -PercentComplete $complete;
    '*** Will now execute command «{0}».' -f $(
      $str = $script.ToString().Trim() -replace '\s+', ' ';
      $max = 100;
      if( $str.Length -le $max ) {
        $str;
      } else {
        $str.Substring( 0, $max - 1 ) + '…';
      }
    );
    $start = [datetime]::Now;
    & $script;
    '*** Finished executing command after {0:0} ms.' -f [datetime]::Now.Subtract( $start ).TotalMilliseconds;
    "`r`n" * 3;
    $complete += $increment;
  }
} *>&1 >> "C:\Windows\Setup\Scripts\DefaultUser.log";