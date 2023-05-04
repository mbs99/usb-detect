#Requires -version 2.0
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Register-WmiEvent -Class win32_VolumeChangeEvent -SourceIdentifier volumeChange
write-host (get-date -format s) " Beginning script..."
do {
    $newEvent = Wait-Event -SourceIdentifier volumeChange
    $eventType = $newEvent.SourceEventArgs.NewEvent.EventType
    $eventTypeName = switch ($eventType) {
        1 { "Configuration changed" }
        2 { "Device arrival" }
        3 { "Device removal" }
        4 { "docking" }
    }
    write-host (get-date -format s) " Event detected = " $eventTypeName
    if ($eventType -eq 2) {
        $driveLetter = $newEvent.SourceEventArgs.NewEvent.DriveName
        $driveLabel = ([wmi]"Win32_LogicalDisk='$driveLetter'").VolumeName
        write-host (get-date -format s) " Drive name = " $driveLetter
        write-host (get-date -format s) " Drive label = " $driveLabel
        # Execute process if drive matches specified condition(s)
        if ($driveLabel -eq 'Backup') {
            
            $Result = [System.Windows.Forms.MessageBox]::Show("Backup starten?", "Backup", 1)
 
            If ($Result -eq "Yes") {
                write-host (get-date -format s) " Starting task in 3 seconds..."
                start-sleep -seconds 3
                $cwd = $driveLetter + "\backup"
                start-process -FilePath rsyncStart.bat -WorkingDirectory $cwd
            }
            else {
                # noop
            }
        }
    }
    Remove-Event -SourceIdentifier volumeChange
} while (1 -eq 1) #Loop until next event
Unregister-Event -SourceIdentifier volumeChange

