function New-EMail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$To,
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$BodyHtml,
        [Switch]$Show
    )

    begin {
        $outlook = New-Object -ComObject Outlook.Application 
    }

    process {
        $mail = $outlook.CreateItem(0)

        $to | Foreach-Object {
            $Mail.Recipients.Add($_) | Out-Null
        }

        $Mail.Subject = $subject
        $Mail.HTMLBody = $BodyHtml
        $Mail.Save()

        if ( $Show ) {
            $Mail.Display();
        }
        $Mail
    }
}