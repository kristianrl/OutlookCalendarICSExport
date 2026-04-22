# @License: None (Large amounts of Stackoverflow code, credit belongs to the code segment's respective authors where it can be identified)

Write-Output 'Beginning Outlook Calendar Export'
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
$Appointments = $folder.Items
$Appointments.Sort("[Start]")

$sb = [System.Text.StringBuilder]::new()
# Fill in ICS/iCalendar properties based on RFC 5545
[void]$sb.AppendLine("BEGIN:VCALENDAR")
[void]$sb.AppendLine("VERSION:2.0")
[void]$sb.AppendLine("PRODID:-//CHCC HIT//PowerShell ICS Migration Tool//EN")
[void]$sb.AppendLine("CALSCALE:GREGORIAN")
[void]$sb.AppendLine("METHOD:PUBLISH")

foreach($appt in ($Appointments)){
    #Write-Output 'Loading appointment...'
    #combining multiple calendar appointments is as simple as multiple BEGIN:VEVENT [...] END:VEVENT in the .ics (iCalendar) text document
    $x = $appt | Select-Object -Property *
    
    [void]$sb.AppendLine("BEGIN:VEVENT")
    [void]$sb.AppendLine("UID:" + [guid]::NewGuid())
    [void]$sb.AppendLine("DTSTAMP:" + "{0:yyyyMMddTHHmmss}" -f ([datetime]::UtcNow) + "Z")
    [void]$sb.AppendLine("DTSTART:" + "{0:yyyyMMddTHHmmss}" -f ($x.Start.ToUniversalTime()) + "Z")
    [void]$sb.AppendLine("DTEND:" + "{0:yyyyMMddTHHmmss}" -f ($x.End.ToUniversalTime()) + "Z")
    [void]$sb.AppendLine("CREATED:" + "{0:yyyyMMddTHHmmss}" -f ($x.CreationTime.ToUniversalTime()) + "Z")
    [void]$sb.AppendLine("LAST-MODIFIED:" + "{0:yyyyMMddTHHmmss}" -f ($x.LastModificationTime.ToUniversalTime()) + "Z")
    # [void]$sb.AppendLine("ORGANIZER;CN=" + $x.Organizer + ":mailto:" + $x.Organizer)
    
    # Recover the Event Response
    $ParamResponseStatus = "NEEDS-ACTION"
    if($x.ResponseStatus -eq "2"){$ParamResponseStatus = "TENTATIVE"}
    if($x.ResponseStatus -eq "3"){$ParamResponseStatus = "CONFIRMED"}
    
    <#
    # Build recipient dictionary for email lookups
    $recipient_dict = @{}
    $recipient = $appt.Recipients
    foreach($r in ($recipient)){
                $recipient_dict.Add($r.Name,$r.Address)
                #echo $r.Name
                #echo $r.Address
                #$recipients_list.append($r)
                }
    try{
        foreach($attendee in ($x.RequiredAttendees.Split(";"))){
            if($recipient_dict.Get_Item($attendee) -like '*@*') {
                [void]$sb.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"REQ-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $recipient_dict.Get_Item($attendee) + ";X-NUM-GUESTS=0:mailto:" + $recipient_dict.Get_Item($attendee))
            }
            else{
                $attendeeStr = "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"REQ-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $attendee.Trim() + ";X-NUM-GUESTS=0:mailto:" + $attendee.Trim()
                if(-Not ($attendeeStr -like '*CN=;*')){[void]$sb.AppendLine($attendeeStr.TrimStart())}
            }         
        }
    }
    catch{}   
    #echo "Optional Attendees"
    #echo $x.OptionalAttendees.Length
    #if($x.OptionalAttendees.Length > 0){$recipient_dict.Get_Item($attendee)
    
    try{
        foreach($attendee in ($x.OptionalAttendees.Split(";"))){
            #echo $attendee
            if($recipient_dict.Get_Item($attendee) -like '*@*') {
                #echo $recipient_dict.Get_Item($attendee)
                [void]$sb.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"OPT-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $recipient_dict.Get_Item($attendee).Trim() + ";X-NUM-GUESTS=0:mailto:" + $recipient_dict.Get_Item($attendee).Trim())
            }
            else{
                $attendeeStr = "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"OPT-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $attendee.Trim() + ";X-NUM-GUESTS=0:mailto:" + $attendee.Trim()
                if(-Not ($attendeeStr -like '*CN=;*')){[void]$sb.AppendLine($attendeeStr.TrimStart())}
            }
        }
    }


    catch{}   
    #>
    
    [void]$sb.AppendLine("SUMMARY:" + $x.Subject)
    #[void]$sb.AppendLine("DESCRIPTION:" + $x.Body)
    [void]$sb.AppendLine("LOCATION:" + $x.Location)
    [void]$sb.AppendLine("STATUS:" + $ParamResponseStatus)
    [void]$sb.AppendLine("TRANSP:TRANSPARENT")
    [void]$sb.AppendLine("SEQUENCE:0")
    [void]$sb.AppendLine("END:VEVENT")

    Write-Output $x.Start.ToUniversalTime().ToString() - $x.Subject.ToString()
    Write-Output "*** Appointment event saved."
}
#Once we’ve defined our event, we close out the “objects”.	
[void]$sb.AppendLine("END:VCALENDAR")

Write-Output 'Saving Appointment to .ics file!'

$fileName = "My_Outlook_Calendar.ics"
$sb.ToString() | Out-File $fileName -Encoding utf8

Write-Output "Successful Outlook Calendar Export! Please upload My_Outlook_Calendar.ics to Outlook 365 using Add Calendar wizard."