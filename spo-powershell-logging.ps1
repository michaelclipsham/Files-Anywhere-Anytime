#-----------------------------------------------------
# Add Audit Log
#-----------------------------------------------------

function AddAuditLog() {
	param(
		[Parameter(Mandatory)]
		$RunId,
		[Parameter(Mandatory)]
		$shortName,
		[Parameter(Mandatory)]
		$code,
		[Parameter(Mandatory)]
		$eventMessage
	)
	try {
		# Write to log file
		$values = @{
			"RunID"           = $RunId;
			"SchoolShortName" = $shortName;
			"SchoolCode"      = $code;
			"EventMessage"    = $eventMessage;
			"EventTime"       = [datetimeoffset]::Now.ToString("yyyy-MM-dd HHmm");
		}
		$logInfo = $values | Out-String
		Write-Information $logInfo
	}
	catch {}
}