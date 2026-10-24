# SMTP Toolkit 1.0.1 - https://ihavenoram.com - https://github.com/ihavenoram/smtp-toolkit
# MIT licence, (c) 2026 ihavenoram.com. If you publish something built on this, please credit ihavenoram.com.
<#
    SMTP Toolkit - on-site mail sending diagnostics (WinForms)
    ----------------------------------------------------------
    A single-file tool for Windows PowerShell 5.1, built to be launched inside a
    ScreenConnect Backstage session (often as SYSTEM) when a line-of-business
    application "cannot send email any more".

    It exists because of a real case: a practice management app configured for
    SMTP2GO on port 465 timed out on every send, while DNS, all eight SMTP2GO
    ports, implicit TLS on 465 and STARTTLS on 2525 were all provably healthy.
    The app used System.Net.Mail.SmtpClient, which only ever speaks STARTTLS -
    so on an implicit-TLS port it waits for a banner that the server will never
    send until the client offers a TLS ClientHello. Deadlock, then timeout.

    That class of fault is invisible to a port test, which is why this tool
    layers its checks:

      * can we reach it            - DNS, the whole port matrix, latency
      * is the transport sound     - implicit TLS, STARTTLS, protocol versions,
                                     certificate chain, interception detection
      * will it accept us          - AUTH, then MAIL FROM / RCPT TO with no
                                     delivery, which proves the sender domain
                                     is verified without sending anything
      * will the APPLICATION work  - a faithful System.Net.Mail.SmtpClient
                                     reproduction, plus a port x "Use SSL"
                                     matrix, because that is the actual question

    Everything measured lands in one JSON report (schema smtp-diagnostic v1)
    for a later HTML report builder. Credentials are typed in for the session
    and never reach the JSON, the transcript, or any file on disk.

    Launch via Launch-Smtp-Toolkit.cmd (forces -STA, required for WinForms).
    No admin needed; a few environment checks report less without it and say so.
#>
[CmdletBinding()]
param(
    # The console is hidden in-process by handle. -WindowStyle Hidden on the
    # launcher is NOT equivalent: it hides the form as well.
    [switch]$NoHideConsole,

    # Internal. Loads the engine functions and stops before the GUI, so a
    # background runspace can rebuild identical definitions from this same file.
    [switch]$EngineOnly
)

#region ENGINE
# ===========================================================================
#  Everything between the ENGINE markers is extracted as text at startup and
#  re-created inside every background runspace. A runspace inherits no state,
#  so this is how the GUI and the worker are guaranteed to run the same code.
#  Nothing in here may touch WinForms or any $script: variable owned by the UI.
# ===========================================================================

function Get-StVersion { '1.2.0' }

function Get-StJsonMarkers { @{ Begin = '<<<SMTPDIAG_JSON_BEGIN>>>'; End = '<<<SMTPDIAG_JSON_END>>>' } }

# ---------------------------------------------------------------------------
#  Result model
#
#  Ported from NetSpeed-Troubleshooter.Helpers.ps1. Two ideas carry across:
#
#  'basis' says HOW a result was arrived at, so a reader can tell a measurement
#  from an inference from a thing we declined to test. It ships in the JSON and
#  is a sortable column in the grid.
#
#  A section that did not run must record WHY. "unsupported" with no reason is
#  worse than no section at all: the reader cannot tell a limitation from a bug.
# ---------------------------------------------------------------------------

function New-StRow {
    param(
        [string]$Group,
        [string]$Check,
        [string]$Result,
        [ValidateSet('measured','server-reported','inference','registry','not measured','withheld')]
        [string]$Basis = 'measured',
        [ValidateSet('ok','warn','fail','info')][string]$Status = 'info',
        [string]$Detail = ''
    )
    [pscustomobject]@{
        group  = $Group
        check  = $Check
        result = $Result
        basis  = $Basis
        status = $Status
        detail = $Detail
    }
}

function New-StSection {
    param(
        [ValidateSet('measured','failed','skipped','unsupported')]
        [string]$State,
        [string]$Reason      = '',
        [string]$Method      = '',
        [string[]]$Caveats   = @(),
        $Rows                = @(),
        $Data                = $null,
        [datetime]$StartedUtc = [datetime]::MinValue,
        [double]$DurationSec  = 0
    )
    if ($State -ne 'measured' -and [string]::IsNullOrWhiteSpace($Reason)) {
        $Reason = 'No reason recorded - this is a bug in the tool, not a finding about the mail path.'
    }
    $started = $null
    if ($StartedUtc -ne [datetime]::MinValue) { $started = ConvertTo-StIso8601 $StartedUtc }
    $reasonVal = $null
    if (-not [string]::IsNullOrWhiteSpace($Reason)) { $reasonVal = $Reason }

    [pscustomobject]@{
        state       = $State
        reason      = $reasonVal
        startedUtc  = $started
        durationSec = [math]::Round($DurationSec, 2)
        method      = $Method
        caveats     = [object[]]$Caveats
        rows        = [object[]]$Rows
        data        = $Data
    }
}

function Set-StSectionTiming {
    param($Section, [datetime]$StartedUtc, [double]$DurationSec)
    if (-not $Section) { return $Section }
    try {
        $Section.startedUtc  = ConvertTo-StIso8601 $StartedUtc
        $Section.durationSec = [math]::Round($DurationSec, 2)
    } catch { }
    return $Section
}

function New-StSubscore {
    param(
        [string]$Name,
        [ValidateSet('green','amber','red','notmeasured')][string]$Level,
        [string]$Summary,
        [string]$Remedy = ''
    )
    [pscustomobject]@{ name = $Name; level = $Level; summary = $Summary; remedy = $Remedy }
}

function Get-StHealth {
    # Worst of the MEASURED subscores. Never an average, and no 0-100 score:
    # collapsing a TLS handshake, a registry value and an SMTP reply code into
    # "78" would be false precision.
    #
    # 'notmeasured' is a first-class fourth state. It is always listed, so a
    # healthy path can be told apart from an untested one, but it never drags
    # the overall verdict down.
    param($Subscores)
    $list = @($Subscores | Where-Object { $_ })
    $out = [pscustomobject]@{
        overall   = 'notmeasured'
        headline  = 'Nothing was measured.'
        subscores = [object[]]$list
        notes     = [object[]]@()
    }
    if ($list.Count -eq 0) { return $out }

    $measured = @($list | Where-Object { $_.level -ne 'notmeasured' })
    if ($measured.Count -eq 0) {
        $out.headline = 'No test produced a usable result. See each section''s reason.'
        return $out
    }

    $rank  = @{ 'green' = 0; 'amber' = 1; 'red' = 2 }
    $worst = $measured[0]
    foreach ($s in $measured) {
        if ($rank[$s.level] -gt $rank[$worst.level]) { $worst = $s }
    }
    $out.overall = $worst.level

    if ($worst.level -eq 'green') {
        $out.headline = 'Nothing measured here stops this machine sending mail through this server.'
    } else {
        $out.headline = ('{0}: {1}' -f $worst.name, $worst.summary)
        if ($worst.remedy) { $out.headline = ('{0} {1}' -f $out.headline, $worst.remedy) }
    }

    $notMeasured = @($list | Where-Object { $_.level -eq 'notmeasured' } | ForEach-Object { $_.name })
    if ($notMeasured.Count -gt 0) {
        $out.notes = [object[]]@(('Not measured, and therefore not counted either way: {0}.' -f ($notMeasured -join ', ')))
    }
    return $out
}

function Get-StKnownLimits {
    # Published false negatives. A clean report here is not a promise that mail
    # will flow - it is a statement about this machine, at this moment, to this
    # server.
    [object[]]@(
        'Every test runs from THIS machine as THIS user. An application running as a different account - a service, a scheduled task, another logged-on user - can have a different proxy, a different certificate store and different credentials, and is not covered.',
        'A pass here says the server accepted the envelope, not that the message was delivered. Content filtering, recipient-side rules and reputation all act after the point this tool stops.',
        'The .NET reproduction uses the same class the application uses, but not the application''s own code. If the application hardcodes a host or port, or sets its own timeout, this tool cannot see that from outside - the app config scan is the only visibility into it.',
        'Certificate interception is inferred from the issuer and the chain, not proven. A corporate CA that legitimately signs internal mail servers looks the same as an inspecting proxy.',
        'The port matrix reflects this network path right now. A firewall rule that applies per-process, or an outbound filter that only triggers on volume, will not show up in a single connect.',
        'IP authentication is checked by comparing this host''s public IP with the address the server echoes in its EHLO greeting. If the site egresses through more than one IP, a single sample can be misleading.'
    )
}

# ---------------------------------------------------------------------------
#  Small conversions and formatting
# ---------------------------------------------------------------------------

function ConvertTo-StIso8601 {
    # ConvertTo-Json renders [datetime] as /Date(1785985400391)/, which nothing
    # downstream wants. Every timestamp is pre-formatted as a string.
    param([datetime]$When)
    if ($When -eq [datetime]::MinValue) { return $null }
    return $When.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Get-StDomainFromAddress {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return '' }
    $at = $Address.LastIndexOf('@')
    if ($at -lt 0 -or $at -eq ($Address.Length - 1)) { return '' }
    return $Address.Substring($at + 1).Trim().Trim('>').Trim()
}

function Test-StEmailAddress {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    return ($Address -match '^[^@\s<>,;]+@[^@\s<>,;]+\.[^@\s<>,;]+$')
}

function ConvertTo-StBase64 {
    param([string]$Text)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text))
}

function Get-StRedactedLine {
    # Redaction happens where a line is LOGGED, not where it is rendered, so a
    # secret can never reach the transcript buffer in the first place. Anything
    # that reads the transcript later - JSON, clipboard, a report - is safe by
    # construction rather than by remembering to filter.
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $Line }
    $t = $Line
    # AUTH PLAIN <base64>  /  AUTH LOGIN <base64>
    $t = [regex]::Replace($t, '(?i)^(C:\s*)?(AUTH\s+(?:PLAIN|LOGIN|CRAM-MD5|XOAUTH2|SCRAM-SHA-\d+))\s+\S+.*$', '$1$2 <redacted>')

    # A bare base64 continuation line - the username and password turns of
    # AUTH LOGIN. Those turns are already redacted at the point they are sent
    # (Write-StCommand -Secret); this is the second line of defence, so it has
    # to be narrow enough not to eat protocol verbs. "STARTTLS" is eight
    # characters of base64 alphabet, which a naive rule happily destroys -
    # requiring padding, a digit, or a lower-case letter keeps the verbs intact.
    if ($t -match '^(C:\s*)?([A-Za-z0-9+/]{12,}={0,2})$') {
        $payload = $Matches[2]
        if ($payload -match '[a-z]' -or $payload -match '\d' -or $payload -match '[+/=]') {
            $t = ('{0}<redacted>' -f $Matches[1])
        }
    }
    return $t
}

function New-StTranscriptSink {
    # A synchronized list so the worker runspace can append while the UI thread
    # reads. Handed to every connection the worker opens.
    #
    # The leading comma is load-bearing HERE, and wrong elsewhere - the
    # difference is what the caller does with it.
    #
    # This function returns ONE object that happens to be a collection, and the
    # caller always assigns it directly. Without the comma an empty list
    # unrolls to nothing, comes back as $null, and every subsequent
    # Write-StTranscript silently does nothing.
    #
    # A function returning a LIST of things must not do this, because a caller
    # wrapping it in @() then gets a one-element array containing the array.
    # See ConvertTo-StPortList for the other side of that coin. Never wrap a
    # call to this one in @().
    $list = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    return ,$list
}

function Write-StTranscript {
    param($Sink, [string]$Direction, [string]$Text)
    if ($null -eq $Sink) { return }
    $line = ('{0} {1}' -f $Direction, $Text)
    [void]$Sink.Add((Get-StRedactedLine $line))
}

function Write-StNote {
    param($Sink, [string]$Text)
    if ($null -eq $Sink) { return }
    [void]$Sink.Add(('-- {0}' -f $Text))
}

# ---------------------------------------------------------------------------
#  Targets and presets
# ---------------------------------------------------------------------------

function Get-StPresets {
    # SMTP2GO first because that is what this tool was built against, but the
    # shape is generic - anything with a host and a port list works.
    #
    # 'kind' changes how results are READ, not how they are measured:
    #   generic         a hosted relay with SMTP AUTH (SMTP2GO, Google, ...)
    #   exchange-onprem an internal Exchange server - internal CA certificates
    #                   are normal, and failed logins count against an AD account
    #   exo-smtpauth    Exchange Online client submission (smtp.office365.com)
    #   exo-connector   Exchange Online MX on 25, authorised by an inbound
    #                   connector that matches this site's public IP
    # A blank server on exo-connector means "use the From domain's MX".
    [object[]]@(
        [pscustomobject]@{ name = 'SMTP2GO - Australia'; server = 'mail-au.smtp2go.com'; explicit = @(2525,587,8025,80,25); implicit = @(465,8465,443); preferred = 2525; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'SMTP2GO - Global (nearest)'; server = 'mail.smtp2go.com'; explicit = @(2525,587,8025,80,25); implicit = @(465,8465,443); preferred = 2525; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'SMTP2GO - Europe'; server = 'mail-eu.smtp2go.com'; explicit = @(2525,587,8025,80,25); implicit = @(465,8465,443); preferred = 2525; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'SMTP2GO - United States'; server = 'mail-us.smtp2go.com'; explicit = @(2525,587,8025,80,25); implicit = @(465,8465,443); preferred = 2525; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'Exchange Server (on-premises)'; server = ''; explicit = @(587,25); implicit = @(); preferred = 587; kind = 'exchange-onprem'; expectInternalCa = $true; serverHint = 'Enter the Exchange server name, e.g. mail.contoso.local' }
        [pscustomobject]@{ name = 'Microsoft 365 - SMTP AUTH (587)'; server = 'smtp.office365.com'; explicit = @(587); implicit = @(); preferred = 587; kind = 'exo-smtpauth'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'Microsoft 365 - inbound connector (25)'; server = ''; explicit = @(25); implicit = @(); preferred = 25; kind = 'exo-connector'; expectInternalCa = $false; serverHint = 'Leave blank to use the From domain''s MX (*.mail.protection.outlook.com)' }
        [pscustomobject]@{ name = 'Google Workspace'; server = 'smtp.gmail.com'; explicit = @(587); implicit = @(465); preferred = 587; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
        [pscustomobject]@{ name = 'Custom / other server'; server = ''; explicit = @(587,25,2525); implicit = @(465); preferred = 587; kind = 'generic'; expectInternalCa = $false; serverHint = '' }
    )
}

function Get-StReplyMeaning {
    # One place that turns an SMTP reply into a sentence a technician can act
    # on. The enhanced status code (5.7.57) carries far more than the basic one
    # (530), and Exchange in particular overloads a handful of basic codes, so
    # both are read. Meanings are the best-known ones; the server's own text is
    # always shown beside them by the caller.
    param([int]$Code, [string]$Text = '', [string]$Kind = 'generic', [bool]$HasCredentials = $false)

    $enh = ''
    if ($Text -match '\b([245]\.\d{1,3}\.\d{1,3})\b') { $enh = $Matches[1] }
    $t = $Text

    # --- Exchange Online
    if ($enh -eq '5.7.139') {
        if ($t -match '(?i)tenant') { return 'Exchange Online has SMTP AUTH (basic authentication) switched off for the whole tenant. An application that logs in with a username and password cannot send this way any more; use an inbound connector on port 25, authorised by this site''s public IP, instead.' }
        if ($t -match '(?i)mailbox|user') { return 'SMTP AUTH is switched off for this mailbox in Exchange Online. It can be re-enabled per mailbox only while the tenant still permits basic authentication; otherwise use an inbound connector on port 25.' }
        return 'Exchange Online refused basic authentication. SMTP AUTH with a username and password is disabled or retired for this tenant; an inbound connector on port 25 is the supported route for a device or application that cannot do OAuth.'
    }
    if ($enh -eq '5.7.64' -or $t -match '(?i)TenantAttribution') {
        return 'Exchange Online could not attribute this connection to a tenant: no inbound connector matched the address this site connects from, so relaying to an outside recipient was refused. Add the IP shown in the Egress section to the connector, or check that the From domain is an accepted domain in the tenant.'
    }
    if ($enh -eq '5.7.708' -or $enh -eq '5.7.606' -or $enh -eq '5.7.511') {
        return 'Exchange Online is refusing traffic from this IP address outright - it is on Microsoft''s block list, usually after spam was seen from it. This is not a connector setting; the address has to be delisted.'
    }
    if ($enh -eq '4.4.62') {
        return 'The message reached the wrong Exchange Online tenant or region. The connector, or the MX host being used, belongs to a different organisation from the From domain.'
    }

    # --- Exchange on-premises (and Exchange generally)
    if ($enh -eq '5.7.57' -or ($Code -eq 530 -and $t -match '(?i)not authenticated')) {
        if ($HasCredentials) { return 'The server says the client is not authenticated - yet credentials were supplied. That means the login was rejected. System.Net.Mail.SmtpClient carries on after a failed login, so a WRONG PASSWORD surfaces here, at MAIL FROM, as "requires a secure connection or the client was not authenticated". Check the password before touching TLS settings.' }
        return 'This connector requires authentication and none was offered. Either give the application a username and password, or have the Exchange administrator add this machine''s address to an anonymous relay connector.'
    }
    if ($enh -eq '5.7.54' -or ($Code -eq 550 -and $t -match '(?i)unable to relay')) {
        return 'Exchange accepted the connection but will not relay to an outside recipient for this address. No receive connector grants anonymous relay to the IP this machine connects from. Internal recipients would probably still be accepted, which is why "it works for internal mail" is a common symptom.'
    }
    if ($enh -eq '5.7.60' -or $t -match '(?i)SendAsDenied') {
        return 'The login worked, but this account is not allowed to send as the From address. Grant Send As on that mailbox, or make the From address the authenticated user''s own.'
    }
    if ($enh -eq '5.7.3' -or $enh -eq '5.7.8' -or $Code -eq 535) {
        if ($Kind -eq 'exchange-onprem') { return 'Wrong username or password. On Exchange that is an Active Directory account, and every failed attempt counts towards its lockout threshold - this tool stops trying after the first rejection for that reason.' }
        return 'Wrong username or password for this server. Check them against the provider portal - a mail user is usually not the same as the portal login.'
    }
    if ($enh -eq '5.7.1' -and $t -match '(?i)block|spamhaus|reputation') {
        return 'The server rejected this IP on reputation. That is a block list problem, not a configuration one.'
    }
    if ($Code -eq 530 -or $enh -eq '5.7.0') {
        return 'The server wants an encrypted channel (STARTTLS) before it will go further. On this port the application''s SSL / TLS option has to be on.'
    }
    if ($Code -eq 550 -and $t -match '(?i)relay') {
        if ($HasCredentials) { return 'Logged in, but this account is still not permitted to relay to that recipient. Check the sending permissions on the account.' }
        return 'The server will not relay for this recipient without authentication. No credentials were supplied, so this also settles the IP authentication question: the address this machine leaves from is NOT authorised. Either authorise it, or give the application a username and password.'
    }
    if ($Code -eq 550 -or $Code -eq 553) { return 'The address was refused. Check it is spelled correctly and that the domain exists before assuming a configuration fault.' }
    if ($Code -eq 421 -or $Code -eq 451 -or $Code -eq 452) { return 'A temporary refusal - throttling, a busy server or greylisting. Retry later before changing anything.' }
    return ''
}

function Get-StAppProfiles {
    # What the sending APPLICATION can do, and what has to be done to it after
    # the settings are changed.
    #
    # The second part matters more than it looks. Every item in these checklists
    # is a way for correct settings to appear to have no effect - a cached
    # configuration, the wrong button, a dropdown elsewhere on the form that
    # silently overrides the fields you just filled in. A technician who does
    # not know about them concludes the settings are being ignored and goes
    # back to blaming the network, which is where this fault started.
    [object[]]@(
        [pscustomobject]@{
            name          = 'Practice-management app (STARTTLS only, cached settings)'
            client        = 'System.Net.Mail.SmtpClient'
            implicitSsl   = $false
            timeoutMs     = 100000
            settingsPath  = 'The application''s email or practice settings screen (often under Setup or Options)'
            fields        = 'Server, Port, Use SSL, Username, Password, From address'
            sslMeaning    = 'Ticking "Use SSL" selects STARTTLS. Applications built this way cannot use implicit SSL / SMTPS (port 465) at all.'
            timeoutNote   = 'It gives up after 100 seconds with "The operation has timed out." Tick the 100-second option on this tab to reproduce the user''s wait exactly.'
            summary       = 'Three things commonly make CORRECT settings look ignored: a settings cache, a second Send Test button that belongs to a different block, and a provider dropdown that overrides the fields. Check them before concluding the settings did not take.'
            checklist     = [object[]]@(
                'Enter the settings in the block the application actually uses for outgoing mail: Server, Port, Use SSL, Username, Password, From address.'
                'Leave "Use SSL" TICKED on a STARTTLS port. In applications like this the checkbox means STARTTLS, not implicit SSL.'
                'After you Save, wait 15-30 seconds before testing. Some applications cache their settings and refresh them on a timer, so a test straight after saving can still use the OLD values - which looks exactly like the new settings being ignored.'
                'If the screen has more than one Send Test (or similar) button, use the one that belongs to the block you just filled in.'
                'If there is a provider or Email Host dropdown, set it to Custom (or equivalent). On any other selection many applications ignore the Server, Port and Use SSL you entered and connect to that provider directly.'
                'A port that this tool reports as working will still fail in the application if any of the points above were missed. Check them before escalating.'
            )
        }
        [pscustomobject]@{
            name          = 'Generic .NET application (System.Net.Mail)'
            client        = 'System.Net.Mail.SmtpClient'
            implicitSsl   = $false
            timeoutMs     = 100000
            settingsPath  = ''
            fields        = ''
            sslMeaning    = 'In almost every application built on this class, a "use SSL" or "use TLS" checkbox sets EnableSsl, which performs STARTTLS only. Microsoft state that establishing SSL before any protocol command - the port 465 model - is not supported by this class.'
            timeoutNote   = 'The class default is a 100-second timeout unless the application sets its own.'
            summary       = 'Correct settings that appear to do nothing usually mean a cached configuration, a second place in the UI where the same values live, or a .config file overriding them.'
            checklist     = [object[]]@(
                'Point the application at a STARTTLS port with its SSL or TLS option enabled.'
                'If the application appears to ignore the settings, look for a cached or delayed configuration reload, and for a second place in the UI where the same values can be set.'
                'Check the Environment tab for a mail configuration in a .config file - those override the settings screen silently.'
            )
        }
        [pscustomobject]@{
            name          = 'Application that supports implicit SSL'
            client        = 'MailKit, or another modern client'
            implicitSsl   = $true
            timeoutMs     = 0
            settingsPath  = ''
            fields        = ''
            sslMeaning    = 'A client of this kind can use either style, so both the STARTTLS ports and the implicit-SSL ports are genuine options.'
            timeoutNote   = ''
            summary       = ''
            checklist     = [object[]]@(
                'Both groups of ports are available. Prefer STARTTLS on the provider''s recommended port unless there is a reason not to.'
            )
        }
        [pscustomobject]@{
            name          = 'Unknown / not sure'
            client        = ''
            implicitSsl   = $false
            timeoutMs     = 0
            settingsPath  = ''
            fields        = ''
            sslMeaning    = 'Assumed to be STARTTLS-only, because that is the safe assumption: a STARTTLS port works for both kinds of client, an implicit-SSL port only for one.'
            timeoutNote   = ''
            summary       = ''
            checklist     = [object[]]@(
                'If the application''s error is a timeout rather than a refusal, and it is pointed at 465, 8465 or 443, it is almost certainly a STARTTLS-only client on an implicit-SSL port.'
            )
        }
    )
}

function Get-StAppProfile {
    param([string]$Name)
    $all = @(Get-StAppProfiles)
    foreach ($p in $all) { if ($p.name -eq $Name) { return $p } }
    return $all[1]   # the generic .NET profile
}

function Get-StAllPorts {
    param($Preset)
    if (-not $Preset) { return @(25,80,465,587,2525,8025,8465,443) }
    $all = @()
    $all += @($Preset.explicit)
    $all += @($Preset.implicit)
    return @($all | Sort-Object -Unique)
}

function Get-StPortSecurity {
    # What a port conventionally means, used when the security mode is Auto and
    # to label the matrix. It is a convention, not a guarantee, which is exactly
    # why the tool tests rather than assumes.
    param([int]$Port)
    switch ($Port) {
        465  { return 'implicit' }
        8465 { return 'implicit' }
        443  { return 'implicit' }
        default { return 'starttls' }
    }
}

function Resolve-StSecurityMode {
    param([string]$Mode, [int]$Port)
    if ($Mode -eq 'auto' -or [string]::IsNullOrWhiteSpace($Mode)) { return (Get-StPortSecurity -Port $Port) }
    return $Mode
}

# ---------------------------------------------------------------------------
#  TLS protocol helpers
# ---------------------------------------------------------------------------

function Get-StSslProtocol {
    # Tls13 is only defined on .NET Framework 4.8 with a recent OS. Asking for
    # a value the enum does not have throws, so every lookup is guarded and the
    # caller finds out what it actually got.
    param([string]$Name)
    try {
        $v = [System.Security.Authentication.SslProtocols]::$Name
        if ($null -eq $v) { return $null }
        return $v
    } catch { return $null }
}

function Get-StDefaultSslProtocols {
    $tls12 = Get-StSslProtocol 'Tls12'
    $tls13 = Get-StSslProtocol 'Tls13'
    if ($null -ne $tls13) {
        try { return ([System.Security.Authentication.SslProtocols]($tls12 -bor $tls13)) } catch { }
    }
    if ($null -ne $tls12) { return $tls12 }
    return [System.Security.Authentication.SslProtocols]::Default
}

function Set-StNetDefaults {
    # For the tool's own HTTPS calls (public IP lookup). Deliberately does not
    # touch anything that would change how OTHER processes behave.
    try {
        $cur = [System.Net.ServicePointManager]::SecurityProtocol
        $tls12 = Get-StSslProtocol 'Tls12'
        if ($null -ne $tls12) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]($cur -bor 3072)
        }
    } catch { }
    try { [System.Net.ServicePointManager]::Expect100Continue = $false } catch { }
}

# ---------------------------------------------------------------------------
#  SMTP protocol engine
#
#  Two things here are deliberately not done the obvious way:
#
#  1. Lines are read a BYTE at a time rather than through a StreamReader.
#     A StreamReader over a NetworkStream reads ahead into its own buffer, and
#     after STARTTLS those stolen bytes are the beginning of the TLS handshake.
#     The handshake then fails in a way that looks like a server fault.
#
#  2. A multiline reply ends at "250 " and continues at "250-", per RFC 5321.
#     Looping until '^250 ' - which every hand-rolled probe script does - hangs
#     forever the moment the server answers 4xx or 5xx, which is precisely when
#     you most want to see the reply.
# ---------------------------------------------------------------------------

function New-StConnection {
    param(
        [string]$Server,
        [int]$Port,
        [ValidateSet('implicit','starttls','none')][string]$Security = 'starttls',
        [int]$ConnectTimeoutMs = 10000,
        [int]$IoTimeoutMs = 15000,
        $Sink = $null,
        $SslProtocols = $null,
        [switch]$RequireValidCert
    )
    # Returns a bag rather than a class: PS 5.1 classes cannot be redefined in a
    # live runspace, and this whole region gets re-created inside workers.
    $conn = @{
        Server           = $Server
        Port             = $Port
        Security         = $Security
        Tcp              = $null
        Raw              = $null      # NetworkStream
        Ssl              = $null      # SslStream once secured
        Stream           = $null      # whichever of the two is live
        Sink             = $Sink
        Secure           = $false
        IoTimeoutMs      = $IoTimeoutMs
        ConnectTimeoutMs = $ConnectTimeoutMs
        SslProtocols     = $SslProtocols
        RequireValidCert = [bool]$RequireValidCert
        Tls              = $null      # populated by Enable-StTls
        Cert             = $null
        Caps             = @()
        AuthMechs        = @()
        Greeting         = ''
        RemoteIpSeen     = ''
        Closed           = $false
    }
    if ($null -eq $conn.SslProtocols) { $conn.SslProtocols = Get-StDefaultSslProtocols }
    return $conn
}

function Connect-StSmtp {
    # Opens the TCP connection and, on an implicit port, completes the TLS
    # handshake before anything else is read. Does NOT read the banner - the
    # caller does that, because "how long until a banner arrives" is itself a
    # measurement worth timing separately.
    param($Conn)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $tcp = New-Object System.Net.Sockets.TcpClient
    $Conn.Tcp = $tcp
    try {
        $iar = $tcp.BeginConnect($Conn.Server, $Conn.Port, $null, $null)
        $ok  = $iar.AsyncWaitHandle.WaitOne($Conn.ConnectTimeoutMs, $false)
        if (-not $ok -or -not $tcp.Connected) {
            try { $tcp.Close() } catch { }
            throw ('TCP connect to {0}:{1} did not complete within {2} ms.' -f $Conn.Server, $Conn.Port, $Conn.ConnectTimeoutMs)
        }
        $tcp.EndConnect($iar)
    } catch {
        $sw.Stop()
        throw
    }
    $sw.Stop()

    $tcp.ReceiveTimeout = $Conn.IoTimeoutMs
    $tcp.SendTimeout    = $Conn.IoTimeoutMs
    $Conn.Raw = $tcp.GetStream()
    $Conn.Raw.ReadTimeout  = $Conn.IoTimeoutMs
    $Conn.Raw.WriteTimeout = $Conn.IoTimeoutMs
    $Conn.Stream = $Conn.Raw
    $Conn.ConnectMs = $sw.ElapsedMilliseconds

    Write-StNote $Conn.Sink ('TCP connected to {0}:{1} in {2} ms' -f $Conn.Server, $Conn.Port, $sw.ElapsedMilliseconds)

    if ($Conn.Security -eq 'implicit') {
        Enable-StTls -Conn $Conn -Phase 'implicit'
    }
    return $Conn
}

function Enable-StTls {
    # Wraps the live stream in TLS and records everything worth reporting about
    # the handshake. The validation callback always returns true so that the
    # audit can describe a bad certificate instead of the connection simply
    # failing - the chain and the policy errors are captured, not discarded, and
    # RequireValidCert is enforced afterwards by the caller if it asked for it.
    param($Conn, [string]$Phase = 'starttls')

    # A hashtable, so the closure mutates the caller's copy rather than a clone.
    #
    # Everything wanted from the certificate and the chain is extracted INSIDE
    # the callback. Both objects are only valid for the duration of the call -
    # SslStream frees the underlying handle afterwards, and a reference kept
    # past that point reads back with an empty Subject and no dates, which looks
    # like a parsing bug rather than a lifetime one.
    $captured = @{ Raw = $null; ChainSubjects = @(); ChainStatus = @(); Errors = 'None' }
    $validator = {
        param($sender, $certificate, $chain, $sslPolicyErrors)
        try { $captured.Errors = $sslPolicyErrors.ToString() } catch { }
        try { if ($certificate) { $captured.Raw = $certificate.GetRawCertData() } } catch { }
        try {
            if ($chain) {
                $subs = @()
                foreach ($el in $chain.ChainElements) { $subs += $el.Certificate.Subject }
                $captured.ChainSubjects = $subs
                $st = @()
                foreach ($s in $chain.ChainStatus) { $st += ('{0}: {1}' -f $s.Status, $s.StatusInformation.Trim()) }
                $captured.ChainStatus = $st
            }
        } catch { }
        return $true
    }.GetNewClosure()
    $cb = [System.Net.Security.RemoteCertificateValidationCallback]$validator

    $ssl = New-Object System.Net.Security.SslStream($Conn.Stream, $false, $cb)
    $sw  = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $ssl.AuthenticateAsClient($Conn.Server, $null, $Conn.SslProtocols, $false)
    } catch {
        $sw.Stop()
        Write-StNote $Conn.Sink ('TLS handshake failed after {0} ms: {1}' -f $sw.ElapsedMilliseconds, $_.Exception.Message)
        throw
    }
    $sw.Stop()

    $ssl.ReadTimeout  = $Conn.IoTimeoutMs
    $ssl.WriteTimeout = $Conn.IoTimeoutMs
    $Conn.Ssl    = $ssl
    $Conn.Stream = $ssl
    $Conn.Secure = $true

    $Conn.Tls = [pscustomobject]@{
        phase          = $Phase
        protocol       = $ssl.SslProtocol.ToString()
        cipher         = $ssl.CipherAlgorithm.ToString()
        cipherBits     = [int]$ssl.CipherStrength
        hash           = $ssl.HashAlgorithm.ToString()
        keyExchange    = $ssl.KeyExchangeAlgorithm.ToString()
        keyExchangeBits= [int]$ssl.KeyExchangeStrength
        handshakeMs    = [int]$sw.ElapsedMilliseconds
        policyErrors   = $captured.Errors
    }
    $Conn.Cert = Get-StCertificateFacts -RawCertData $captured.Raw -ChainSubjects $captured.ChainSubjects `
                                        -ChainStatus $captured.ChainStatus -PolicyErrors $captured.Errors -ExpectedHost $Conn.Server

    Write-StNote $Conn.Sink ('TLS established: {0} / {1}-{2} in {3} ms' -f $Conn.Tls.protocol, $Conn.Tls.cipher, $Conn.Tls.cipherBits, $Conn.Tls.handshakeMs)
    return $Conn
}

function Read-StRawLine {
    # One byte at a time to the LF. Slow in theory, irrelevant in practice at
    # SMTP volumes, and the only way to guarantee nothing is buffered past the
    # end of the line - which is what makes STARTTLS work reliably.
    param($Conn)
    $bytes = New-Object System.Collections.Generic.List[byte]
    while ($true) {
        $b = -1
        try { $b = $Conn.Stream.ReadByte() }
        catch [System.IO.IOException] { throw ('Timed out waiting for a reply from {0}:{1} after {2} ms.' -f $Conn.Server, $Conn.Port, $Conn.IoTimeoutMs) }
        if ($b -lt 0) {
            if ($bytes.Count -eq 0) { throw ('{0}:{1} closed the connection without replying.' -f $Conn.Server, $Conn.Port) }
            break
        }
        if ($b -eq 10) { break }          # LF ends the line
        if ($b -ne 13) { [void]$bytes.Add([byte]$b) }   # drop the CR
    }
    if ($bytes.Count -eq 0) { return '' }
    return [Text.Encoding]::UTF8.GetString($bytes.ToArray())
}

function Read-StResponse {
    # RFC 5321 continuation rule: "nnn-" continues, "nnn " terminates.
    param($Conn)
    $lines = New-Object System.Collections.Generic.List[string]
    $code  = 0
    while ($true) {
        $line = Read-StRawLine -Conn $Conn
        Write-StTranscript $Conn.Sink 'S:' $line
        [void]$lines.Add($line)
        if ($line -match '^(\d{3})([ -])(.*)$') {
            $code = [int]$Matches[1]
            if ($Matches[2] -eq ' ') { break }
        } else {
            # Not a well-formed reply line. Stop rather than block forever.
            break
        }
    }
    $arr = $lines.ToArray()
    return [pscustomobject]@{
        code    = $code
        lines   = [object[]]$arr
        text    = ($arr -join "`n")
        success = ($code -ge 200 -and $code -lt 400)
    }
}

function Write-StCommand {
    # -Secret means the payload never reaches the transcript buffer at all,
    # which is what makes "credentials cannot leak into the report" a property
    # of the code rather than a promise about filtering later.
    param($Conn, [string]$Command, [switch]$Secret)
    if ($Secret) { Write-StTranscript $Conn.Sink 'C:' '<redacted>' }
    else         { Write-StTranscript $Conn.Sink 'C:' $Command }
    $bytes = [Text.Encoding]::UTF8.GetBytes($Command + "`r`n")
    $Conn.Stream.Write($bytes, 0, $bytes.Length)
    $Conn.Stream.Flush()
}

function Invoke-StCommand {
    param($Conn, [string]$Command)
    Write-StCommand -Conn $Conn -Command $Command
    return (Read-StResponse -Conn $Conn)
}

function Read-StGreeting {
    param($Conn)
    $r = Read-StResponse -Conn $Conn
    $Conn.Greeting = ($r.lines -join ' ')
    return $r
}

function Invoke-StEhlo {
    # Parses the capability list, the AUTH mechanisms, and the address the
    # server says it sees us coming from. That last one is the whole answer to
    # "did we add the right IP for IP authentication".
    param($Conn, [string]$HeloName)
    if ([string]::IsNullOrWhiteSpace($HeloName)) { $HeloName = $env:COMPUTERNAME }
    $r = Invoke-StCommand -Conn $Conn -Command ('EHLO {0}' -f $HeloName)

    $caps  = New-Object System.Collections.Generic.List[string]
    $mechs = New-Object System.Collections.Generic.List[string]
    $seenIp = ''
    foreach ($line in $r.lines) {
        if ($line -match '^\d{3}[ -](.*)$') {
            $body = $Matches[1].Trim()
            if ($body) { [void]$caps.Add($body) }
            if ($body -match '(?i)^AUTH\s+(.+)$') {
                foreach ($m in ($Matches[1] -split '\s+')) { if ($m) { [void]$mechs.Add($m.ToUpperInvariant()) } }
            }
            if ($body -match '\[((?:\d{1,3}\.){3}\d{1,3}|[0-9a-fA-F:]{3,})\]') { $seenIp = $Matches[1] }
        }
    }
    $Conn.Caps      = [object[]]$caps.ToArray()
    $Conn.AuthMechs = [object[]]$mechs.ToArray()
    if ($seenIp) { $Conn.RemoteIpSeen = $seenIp }

    return [pscustomobject]@{
        code         = $r.code
        success      = $r.success
        lines        = $r.lines
        text         = $r.text
        caps         = [object[]]$caps.ToArray()
        authMechs    = [object[]]$mechs.ToArray()
        startTls     = [bool]($caps -match '(?i)^STARTTLS')
        size         = [bool]($caps -match '(?i)^SIZE')
        remoteIpSeen = $seenIp
    }
}

function Invoke-StStartTls {
    param($Conn)
    $r = Invoke-StCommand -Conn $Conn -Command 'STARTTLS'
    if ($r.code -ne 220) {
        throw ('Server refused STARTTLS: {0}' -f ($r.lines -join ' '))
    }
    Enable-StTls -Conn $Conn -Phase 'starttls'
    return $r
}

function Invoke-StAuthLogin {
    # AUTH LOGIN is the mechanism System.Net.Mail.SmtpClient reaches for first
    # against most servers, so it is the one worth reproducing. Both turns are
    # redacted at the point of logging.
    param($Conn, [string]$Username, [string]$Password)

    $r = Invoke-StCommand -Conn $Conn -Command 'AUTH LOGIN'
    if ($r.code -ne 334) { return [pscustomobject]@{ mech='LOGIN'; success=$false; code=$r.code; text=($r.lines -join ' ') } }

    Write-StCommand -Conn $Conn -Command (ConvertTo-StBase64 $Username) -Secret
    $r = Read-StResponse -Conn $Conn
    if ($r.code -ne 334) { return [pscustomobject]@{ mech='LOGIN'; success=$false; code=$r.code; text=($r.lines -join ' ') } }

    Write-StCommand -Conn $Conn -Command (ConvertTo-StBase64 $Password) -Secret
    $r = Read-StResponse -Conn $Conn
    return [pscustomobject]@{ mech='LOGIN'; success=($r.code -eq 235); code=$r.code; text=($r.lines -join ' ') }
}

function Invoke-StAuthPlain {
    param($Conn, [string]$Username, [string]$Password)
    $payload = ConvertTo-StBase64 ("`0" + $Username + "`0" + $Password)
    $r = Invoke-StCommand -Conn $Conn -Command ('AUTH PLAIN {0}' -f $payload)
    return [pscustomobject]@{ mech='PLAIN'; success=($r.code -eq 235); code=$r.code; text=($r.lines -join ' ') }
}

function Invoke-StAuthenticate {
    param($Conn, [string]$Username, [string]$Password, [string]$Prefer = 'LOGIN')
    $mechs = @($Conn.AuthMechs)
    $order = @()
    if ($Prefer -eq 'PLAIN') { $order = @('PLAIN','LOGIN') } else { $order = @('LOGIN','PLAIN') }
    foreach ($m in $order) {
        if ($mechs.Count -gt 0 -and ($mechs -notcontains $m)) { continue }
        if ($m -eq 'LOGIN') { return (Invoke-StAuthLogin -Conn $Conn -Username $Username -Password $Password) }
        else                { return (Invoke-StAuthPlain -Conn $Conn -Username $Username -Password $Password) }
    }
    return [pscustomobject]@{ mech=''; success=$false; code=0; text=('Server advertised no mechanism this tool can use: {0}' -f (($mechs -join ' '))) }
}

function Invoke-StEnvelope {
    # MAIL FROM / RCPT TO and then RSET. Stops short of DATA, so it proves the
    # sender domain is accepted and the recipient is routable without putting a
    # message in anyone's mailbox. This is what every sweep uses.
    param($Conn, [string]$From, [string]$To)
    $out = [pscustomobject]@{
        mailFromCode = 0; mailFromText = ''
        rcptToCode   = 0; rcptToText   = ''
        success      = $false
    }
    $r = Invoke-StCommand -Conn $Conn -Command ('MAIL FROM:<{0}>' -f $From)
    $out.mailFromCode = $r.code
    $out.mailFromText = ($r.lines -join ' ')
    if ($r.code -ne 250) { return $out }

    $r = Invoke-StCommand -Conn $Conn -Command ('RCPT TO:<{0}>' -f $To)
    $out.rcptToCode = $r.code
    $out.rcptToText = ($r.lines -join ' ')
    if ($r.code -ne 250 -and $r.code -ne 251) { return $out }

    [void](Invoke-StCommand -Conn $Conn -Command 'RSET')
    $out.success = $true
    return $out
}

function Invoke-StSendMessage {
    # A full send through this tool's own engine. Unlike SmtpClient it handles
    # implicit TLS, so it can genuinely deliver on 465 - which is how the tool
    # proves the port is healthy and the application's library is the problem.
    param($Conn, [string]$From, [string]$To, [string]$Subject, [string]$Body)

    $r = Invoke-StCommand -Conn $Conn -Command ('MAIL FROM:<{0}>' -f $From)
    if ($r.code -ne 250) { return [pscustomobject]@{ success=$false; stage='MAIL FROM'; code=$r.code; text=($r.lines -join ' ') } }

    $r = Invoke-StCommand -Conn $Conn -Command ('RCPT TO:<{0}>' -f $To)
    if ($r.code -ne 250 -and $r.code -ne 251) { return [pscustomobject]@{ success=$false; stage='RCPT TO'; code=$r.code; text=($r.lines -join ' ') } }

    $r = Invoke-StCommand -Conn $Conn -Command 'DATA'
    if ($r.code -ne 354) { return [pscustomobject]@{ success=$false; stage='DATA'; code=$r.code; text=($r.lines -join ' ') } }

    $msgId = ('<{0}@{1}>' -f ([guid]::NewGuid().ToString('N')), (Get-StDomainFromAddress $From))
    $date  = (Get-Date).ToString('R')
    $headers = @(
        ('Date: {0}' -f $date)
        ('From: {0}' -f $From)
        ('To: {0}' -f $To)
        ('Subject: {0}' -f $Subject)
        ('Message-ID: {0}' -f $msgId)
        'MIME-Version: 1.0'
        'Content-Type: text/plain; charset=utf-8'
        'Content-Transfer-Encoding: 8bit'
        ('X-Mailer: SMTP Toolkit {0}' -f (Get-StVersion))
    )
    $sb = New-Object System.Text.StringBuilder
    foreach ($h in $headers) { [void]$sb.Append($h).Append("`r`n") }
    [void]$sb.Append("`r`n")
    foreach ($line in ($Body -split "`r?`n")) {
        # Dot-stuffing, per RFC 5321 4.5.2.
        if ($line.StartsWith('.')) { [void]$sb.Append('.') }
        [void]$sb.Append($line).Append("`r`n")
    }
    [void]$sb.Append(".`r`n")

    Write-StNote $Conn.Sink ('DATA: {0} header lines plus body, terminated with a lone dot' -f $headers.Count)
    $bytes = [Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $Conn.Stream.Write($bytes, 0, $bytes.Length)
    $Conn.Stream.Flush()

    $r = Read-StResponse -Conn $Conn
    return [pscustomobject]@{
        success   = ($r.code -eq 250)
        stage     = 'end of DATA'
        code      = $r.code
        text      = ($r.lines -join ' ')
        messageId = $msgId
    }
}

function Close-StConnection {
    param($Conn, [switch]$NoQuit)
    if ($null -eq $Conn -or $Conn.Closed) { return }
    if (-not $NoQuit -and $Conn.Tcp -and $Conn.Tcp.Connected) {
        try { [void](Invoke-StCommand -Conn $Conn -Command 'QUIT') } catch { }
    }
    try { if ($Conn.Ssl) { $Conn.Ssl.Dispose() } } catch { }
    try { if ($Conn.Raw) { $Conn.Raw.Dispose() } } catch { }
    try { if ($Conn.Tcp) { $Conn.Tcp.Close() } } catch { }
    $Conn.Closed = $true
}

# ---------------------------------------------------------------------------
#  Certificate facts
# ---------------------------------------------------------------------------

function Get-StCertificateFacts {
    # Takes the DER bytes rather than a live certificate object, because the
    # object handed to a validation callback does not survive the callback.
    param($RawCertData, $ChainSubjects = @(), $ChainStatus = @(), [string]$PolicyErrors, [string]$ExpectedHost)
    if ($null -eq $RawCertData -or $RawCertData.Length -eq 0) {
        return [pscustomobject]@{ present = $false }
    }
    $c2 = $null
    try { $c2 = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([byte[]]$RawCertData) } catch { }
    if ($null -eq $c2) {
        return [pscustomobject]@{ present = $false }
    }

    $subject = $c2.Subject
    $issuer  = $c2.Issuer
    $notBefore = $c2.NotBefore
    $notAfter  = $c2.NotAfter
    $thumb  = $c2.Thumbprint
    $serial = $c2.SerialNumber
    $sig    = ''
    $sans   = @()
    try { $sig = $c2.SignatureAlgorithm.FriendlyName } catch { }
    try {
        foreach ($ext in $c2.Extensions) {
            if ($ext.Oid.Value -eq '2.5.29.17') {
                $txt = $ext.Format($false)
                foreach ($p in ($txt -split ',')) {
                    $p = $p.Trim()
                    if ($p -match '(?i)DNS Name[=:]\s*(.+)$') { $sans += $Matches[1].Trim() }
                }
            }
        }
    } catch { }

    $daysLeft  = $null
    $notBeforeIso = $null
    $notAfterIso  = $null
    if ($notBefore) { $notBeforeIso = ConvertTo-StIso8601 ([datetime]$notBefore) }
    if ($notAfter) {
        $notAfterIso = ConvertTo-StIso8601 ([datetime]$notAfter)
        $daysLeft = [int][math]::Floor(($notAfter - (Get-Date)).TotalDays)
    }

    $subs        = @($ChainSubjects)
    $chainStatus = @($ChainStatus)
    $chainDepth  = $subs.Count
    $rootSubject = ''
    if ($chainDepth -gt 0) { $rootSubject = $subs[$chainDepth - 1] }

    # Interception heuristic. A public mail host presents a certificate from a
    # public CA; an inspecting proxy or AV re-signs with a local CA that is only
    # trusted because it was pushed into the machine store. Both validate
    # cleanly, so the issuer - not the validation result - is the tell.
    $issuerCn = ''
    if ($issuer -match '(?i)CN=([^,]+)') { $issuerCn = $Matches[1].Trim() }
    $publicCaHints = @('DigiCert','Sectigo','Comodo','Let''s Encrypt','GlobalSign','GoDaddy','Entrust','Amazon','Google Trust','ISRG','Thawte','GeoTrust','RapidSSL','Certum','Buypass','ZeroSSL','Microsoft Azure','Actalis','SSL.com','IdenTrust','Baltimore','USERTrust','QuoVadis','Starfield','E1','E5','E6','R10','R11','R3','WE1','WR1','GTS')
    $looksPublic = $false
    foreach ($h in $publicCaHints) { if ($issuerCn -like ('*{0}*' -f $h) -or $rootSubject -like ('*{0}*' -f $h)) { $looksPublic = $true; break } }
    $selfSigned = ($subject -eq $issuer -and $subject -ne '')

    $hostMatches = $true
    if ($PolicyErrors -match '(?i)RemoteCertificateNameMismatch') { $hostMatches = $false }

    return [pscustomobject]@{
        present          = $true
        subject          = $subject
        issuer           = $issuer
        issuerCn         = $issuerCn
        rootSubject      = $rootSubject
        sans             = [object[]]$sans
        notBefore        = $notBeforeIso
        notAfter         = $notAfterIso
        daysUntilExpiry  = $daysLeft
        thumbprint       = $thumb
        serialNumber     = $serial
        signatureAlgo    = $sig
        chainDepth       = $chainDepth
        chainStatus      = [object[]]$chainStatus
        policyErrors     = $PolicyErrors
        valid            = ($PolicyErrors -eq 'None')
        hostNameMatches  = $hostMatches
        issuerLooksPublic= $looksPublic
        selfSigned       = $selfSigned
        expectedHost     = $ExpectedHost
    }
}

# ---------------------------------------------------------------------------
#  Session helper
# ---------------------------------------------------------------------------

function Open-StSession {
    # Everything up to "ready to issue commands": connect, banner, EHLO, and on
    # an explicit port the STARTTLS upgrade and the second EHLO that must follow
    # it. Returns the connection plus both capability sets, because the
    # difference between them is itself diagnostic - some servers only advertise
    # AUTH once the channel is encrypted, which is exactly why an application
    # that authenticates too early fails.
    param(
        [string]$Server,
        [int]$Port,
        [string]$Security,
        [string]$HeloName,
        $Sink = $null,
        [int]$IoTimeoutMs = 15000,
        [int]$ConnectTimeoutMs = 10000,
        $SslProtocols = $null
    )
    $sec = Resolve-StSecurityMode -Mode $Security -Port $Port
    $conn = New-StConnection -Server $Server -Port $Port -Security $sec -Sink $Sink `
                             -IoTimeoutMs $IoTimeoutMs -ConnectTimeoutMs $ConnectTimeoutMs -SslProtocols $SslProtocols
    $greeting = $null
    $ehlo1 = $null
    $ehlo2 = $null
    try {
        [void](Connect-StSmtp -Conn $conn)

        $greeting = Read-StGreeting -Conn $conn
        $ehlo1 = Invoke-StEhlo -Conn $conn -HeloName $HeloName

        if ($sec -eq 'starttls') {
            if (-not $ehlo1.startTls) {
                throw ('{0}:{1} did not advertise STARTTLS. An application that requires an encrypted channel on this port will fail here.' -f $Server, $Port)
            }
            [void](Invoke-StStartTls -Conn $conn)
            $ehlo2 = Invoke-StEhlo -Conn $conn -HeloName $HeloName
        }
    } catch {
        # Every caller closes the session it is handed - but a throw here
        # means nothing is handed back, so the socket would otherwise stay
        # open until garbage collection. Exchange limits concurrent
        # connections per source IP, and a sweep fails a lot of probes.
        Close-StConnection -Conn $conn -NoQuit
        throw
    }

    return [pscustomobject]@{
        conn     = $conn
        security = $sec
        greeting = $greeting
        ehlo     = $ehlo1
        ehloTls  = $ehlo2
        caps     = $(if ($ehlo2) { $ehlo2 } else { $ehlo1 })
    }
}

# ---------------------------------------------------------------------------
#  Host context
# ---------------------------------------------------------------------------

function Test-StElevated {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Get-StDotNetRelease {
    # The Release DWORD is the only reliable way to name the installed 4.x
    # version. It matters here because TLS 1.2 is only the default from 4.7.
    try {
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full')
        if (-not $key) { return [pscustomobject]@{ release = 0; version = 'not found' } }
        $rel = [int]$key.GetValue('Release', 0)
        $key.Close()
        $name = switch ($rel) {
            { $_ -ge 533320 } { '4.8.1 or later'; break }
            { $_ -ge 528040 } { '4.8'; break }
            { $_ -ge 461808 } { '4.7.2'; break }
            { $_ -ge 461308 } { '4.7.1'; break }
            { $_ -ge 460798 } { '4.7'; break }
            { $_ -ge 394802 } { '4.6.2'; break }
            { $_ -ge 393295 } { '4.6.1'; break }
            { $_ -ge 379893 } { '4.5.2'; break }
            default           { ('4.x (release {0})' -f $rel) }
        }
        return [pscustomobject]@{ release = $rel; version = $name }
    } catch { return [pscustomobject]@{ release = 0; version = 'unreadable' } }
}

function Get-StHostContext {
    $ctx = [pscustomobject]@{
        computerName = $env:COMPUTERNAME
        domain       = $env:USERDNSDOMAIN
        osCaption    = ''
        osBuild      = ''
        psVersion    = $PSVersionTable.PSVersion.ToString()
        clrVersion   = [System.Environment]::Version.ToString()
        dotNet       = (Get-StDotNetRelease)
        runAs        = ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
        isSystem     = $false
        elevated     = (Test-StElevated)
        pid          = $PID
        is64Process  = [Environment]::Is64BitProcess
    }
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $ctx.osCaption = $os.Caption
        $ctx.osBuild   = $os.BuildNumber
    } catch {
        $ctx.osCaption = [System.Environment]::OSVersion.VersionString
    }
    try {
        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $ctx.isSystem = ($sid -eq 'S-1-5-18')
    } catch { }
    if ([string]::IsNullOrWhiteSpace($ctx.domain)) { $ctx.domain = 'WORKGROUP' }
    return $ctx
}

function Test-StContextSection {
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ctx = Get-StHostContext
    $rows = @()

    $who = $ctx.runAs
    if ($ctx.isSystem) { $who = ('{0}  [SYSTEM]' -f $who) }
    elseif ($ctx.elevated) { $who = ('{0}  [elevated]' -f $who) }
    $rows += New-StRow -Group 'Host' -Check 'Running as' -Result $who -Status 'info' `
        -Detail 'Mail settings, proxy configuration and certificate stores can all differ per account. If the application runs as a service, test as that account too.'
    $rows += New-StRow -Group 'Host' -Check 'Machine' -Result ('{0} ({1})' -f $ctx.computerName, $ctx.domain) -Status 'info'
    $rows += New-StRow -Group 'Host' -Check 'Operating system' -Result ('{0} build {1}' -f $ctx.osCaption, $ctx.osBuild) -Status 'info'
    $rows += New-StRow -Group 'Host' -Check 'PowerShell' -Result ('{0} ({1}-bit process)' -f $ctx.psVersion, $(if ($ctx.is64Process) { 64 } else { 32 })) -Status 'info'

    $netStatus = 'ok'
    $netDetail = 'TLS 1.2 is the default for applications built against this version.'
    if ($ctx.dotNet.release -lt 460798 -and $ctx.dotNet.release -gt 0) {
        $netStatus = 'warn'
        $netDetail = 'Below 4.7, .NET applications do not default to TLS 1.2. The registry values in the Environment tab decide what such an application negotiates.'
    }
    $rows += New-StRow -Group 'Host' -Check '.NET Framework' -Result $ctx.dotNet.version -Basis 'registry' -Status $netStatus -Detail $netDetail

    $sw.Stop()
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Local OS and registry inspection' -Rows $rows -Data $ctx) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  DNS
# ---------------------------------------------------------------------------

function Test-StDns {
    param([string]$Server)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = @()

    if ([string]::IsNullOrWhiteSpace($Server)) {
        $sw.Stop()
        return (New-StSection -State 'skipped' -Reason 'No server name was supplied.')
    }

    $addresses = @()
    $resolveMs = 0
    try {
        $rsw = [System.Diagnostics.Stopwatch]::StartNew()
        $addresses = @([System.Net.Dns]::GetHostAddresses($Server))
        $rsw.Stop()
        $resolveMs = $rsw.ElapsedMilliseconds
    } catch {
        $sw.Stop()
        $rows += New-StRow -Group 'DNS' -Check 'Forward lookup' -Result 'FAILED' -Status 'fail' -Detail $_.Exception.Message
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason ('{0} could not be resolved: {1}' -f $Server, $_.Exception.Message) -Method 'System.Net.Dns' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    }

    $v4 = @($addresses | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString })
    $v6 = @($addresses | Where-Object { $_.AddressFamily -eq 'InterNetworkV6' } | ForEach-Object { $_.IPAddressToString })

    $lat = 'ok'
    if ($resolveMs -gt 1500) { $lat = 'warn' }
    $rows += New-StRow -Group 'DNS' -Check 'Forward lookup' -Result ('{0} address(es) in {1} ms' -f $addresses.Count, $resolveMs) -Status $lat `
        -Detail $(if ($resolveMs -gt 1500) { 'Slow resolution eats into an application''s connect timeout before a single byte is sent.' } else { '' })
    if ($v4.Count -gt 0) { $rows += New-StRow -Group 'DNS' -Check 'IPv4' -Result ($v4 -join ', ') -Status 'ok' }
    if ($v6.Count -gt 0) {
        $rows += New-StRow -Group 'DNS' -Check 'IPv6' -Result ($v6 -join ', ') -Status 'info' `
            -Detail 'If IPv6 resolves but is not routable, some clients try it first and stall until they fall back.'
    }

    # Reverse lookup on the first v4, purely informational but it names the POP.
    $ptr = ''
    if ($v4.Count -gt 0) {
        try { $ptr = [System.Net.Dns]::GetHostEntry($v4[0]).HostName } catch { $ptr = '' }
        if ($ptr) { $rows += New-StRow -Group 'DNS' -Check 'Reverse lookup' -Result $ptr -Basis 'server-reported' -Status 'info' }
    }

    # A hosts-file override for the mail server is rare, deliberate, and worth
    # shouting about when present.
    $hostsHit = ''
    try {
        $hostsPath = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
        if (Test-Path -LiteralPath $hostsPath) {
            foreach ($line in (Get-Content -LiteralPath $hostsPath -ErrorAction Stop)) {
                if ($line -match '^\s*#') { continue }
                if ($line -match ('(?i)\s' + [regex]::Escape($Server) + '\s*$') -or $line -match ('(?i)\s' + [regex]::Escape($Server) + '\s')) {
                    $hostsHit = $line.Trim()
                    break
                }
            }
        }
    } catch { }
    if ($hostsHit) {
        $rows += New-StRow -Group 'DNS' -Check 'hosts file override' -Result $hostsHit -Status 'warn' `
            -Detail 'This machine is not using public DNS for the mail server. Whatever that entry points at is what every application here will talk to.'
    }

    $sw.Stop()
    $data = [pscustomobject]@{ server = $Server; ipv4 = [object[]]$v4; ipv6 = [object[]]$v6; resolveMs = $resolveMs; ptr = $ptr; hostsOverride = $hostsHit }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'System.Net.Dns plus a hosts file scan' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Invoke-StDnsQuery {
    # Public resolver first. From inside a customer network the AD-integrated
    # zone often answers authoritatively for the public domain while holding no
    # MX, SPF or DMARC records at all, so a local lookup reports them "missing"
    # when they are perfectly fine on the internet. Falls back to the local
    # resolver when outbound DNS to 1.1.1.1 is blocked.
    [CmdletBinding()]
    param([string]$Name, [string]$Type, [switch]$LocalOnly)
    if (-not $LocalOnly) {
        try {
            $pub = @(Resolve-DnsName -Name $Name -Type $Type -Server '1.1.1.1' -DnsOnly -QuickTimeout -ErrorAction Stop)
            if ($pub.Count -gt 0) { return $pub }
        } catch { }
    }
    return @(Resolve-DnsName -Name $Name -Type $Type -DnsOnly -ErrorAction Stop)
}

function Resolve-StMx {
    # The From domain's MX, publicly and locally, plus whether it is hosted in
    # Exchange Online. The public answer is the one that matters for sending.
    param([string]$Domain)
    $out = [pscustomobject]@{
        domain = $Domain; publicMx = [object[]]@(); localMx = [object[]]@()
        best = ''; isExchangeOnline = $false; splitHorizon = $false; error = ''
    }
    if ([string]::IsNullOrWhiteSpace($Domain)) { $out.error = 'No domain to look up.'; return $out }
    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { $out.error = 'Resolve-DnsName is not available on this machine.'; return $out }

    $pubList = @()
    try {
        $r = @(Resolve-DnsName -Name $Domain -Type MX -Server '1.1.1.1' -DnsOnly -QuickTimeout -ErrorAction Stop | Where-Object { $_.QueryType -eq 'MX' })
        $pubList = @($r | Sort-Object Preference | ForEach-Object { [string]$_.NameExchange })
    } catch { }
    $locList = @()
    try {
        $r = @(Resolve-DnsName -Name $Domain -Type MX -DnsOnly -ErrorAction Stop | Where-Object { $_.QueryType -eq 'MX' })
        $locList = @($r | Sort-Object Preference | ForEach-Object { [string]$_.NameExchange })
    } catch { }

    $out.publicMx = [object[]]$pubList
    $out.localMx  = [object[]]$locList
    if ($pubList.Count -gt 0) { $out.best = $pubList[0] } elseif ($locList.Count -gt 0) { $out.best = $locList[0] }
    if (-not $out.best) { $out.error = ('No MX record found for {0}.' -f $Domain) }
    $out.isExchangeOnline = ($out.best -match '(?i)\.mail\.protection\.outlook\.com\.?$')
    $out.splitHorizon = ($pubList.Count -gt 0 -and ((@($pubList) -join ',') -ne (@($locList) -join ',')))
    $out.best = $out.best.TrimEnd('.')
    return $out
}

function Test-StSenderDns {
    # SPF / DKIM / DMARC for the sending domain. Transport can be perfect and
    # mail still vanish downstream, so a vendor-ready report has to cover this.
    param([string]$FromAddress, [string]$DkimSelector = '')
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $domain = Get-StDomainFromAddress $FromAddress
    if (-not $domain) {
        return (New-StSection -State 'skipped' -Reason 'No usable sender domain - fill in the From address.')
    }
    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
        return (New-StSection -State 'unsupported' -Reason 'Resolve-DnsName is not available on this machine, so TXT records cannot be queried. Everything else in this report is unaffected.')
    }

    $rows = @()
    $spf = ''; $dmarc = ''; $dkim = ''; $mx = @()

    try {
        $txt = @(Invoke-StDnsQuery -Name $domain -Type TXT -ErrorAction Stop)
        foreach ($r in $txt) {
            $s = ''
            if ($r.PSObject.Properties['Strings']) { $s = ($r.Strings -join '') }
            if ($s -match '(?i)^v=spf1') { $spf = $s }
        }
    } catch { }

    if ($spf) {
        $rows += New-StRow -Group 'Sender DNS' -Check 'SPF' -Result $spf -Basis 'server-reported' -Status 'ok'
        if ($spf -notmatch '(?i)(include:|ip4:|ip6:|a[: ]|mx[: ])') {
            $rows += New-StRow -Group 'Sender DNS' -Check 'SPF content' -Result 'No senders authorised' -Status 'warn' -Detail 'The record exists but names no permitted sender.'
        }
    } else {
        $rows += New-StRow -Group 'Sender DNS' -Check 'SPF' -Result 'Not published' -Status 'warn' `
            -Detail ('{0} has no SPF record. Relayed mail is far more likely to be treated as spam.' -f $domain)
    }

    try {
        $d = @(Invoke-StDnsQuery -Name ('_dmarc.{0}' -f $domain) -Type TXT -ErrorAction Stop)
        foreach ($r in $d) {
            $s = ''
            if ($r.PSObject.Properties['Strings']) { $s = ($r.Strings -join '') }
            if ($s -match '(?i)^v=DMARC1') { $dmarc = $s }
        }
    } catch { }
    if ($dmarc) { $rows += New-StRow -Group 'Sender DNS' -Check 'DMARC' -Result $dmarc -Basis 'server-reported' -Status 'ok' }
    else        { $rows += New-StRow -Group 'Sender DNS' -Check 'DMARC' -Result 'Not published' -Status 'info' -Detail 'Optional, but its absence weakens SPF and DKIM.' }

    # DKIM needs the selector, which is provider-specific. SMTP2GO publishes
    # under a per-account selector, so a miss here is not a finding.
    $selectors = @()
    if ($DkimSelector) { $selectors += $DkimSelector }
    $selectors += @('s1','s2','smtp2go','default','google','selector1','selector2','k1')
    foreach ($sel in ($selectors | Select-Object -Unique)) {
        try {
            $k = @(Invoke-StDnsQuery -Name ('{0}._domainkey.{1}' -f $sel, $domain) -Type TXT -ErrorAction Stop)
            foreach ($r in $k) {
                $s = ''
                if ($r.PSObject.Properties['Strings']) { $s = ($r.Strings -join '') }
                if ($s -match '(?i)(v=DKIM1|p=)') { $dkim = ('{0}: {1}' -f $sel, $s); break }
            }
        } catch { }
        if ($dkim) { break }
    }
    if ($dkim) {
        $rows += New-StRow -Group 'Sender DNS' -Check 'DKIM' -Result $dkim -Basis 'server-reported' -Status 'ok'
    } else {
        $rows += New-StRow -Group 'Sender DNS' -Check 'DKIM' -Result 'No key found under the selectors tried' -Basis 'not measured' -Status 'info' `
            -Detail 'DKIM selectors are provider-specific and cannot be enumerated from DNS. This is not evidence that DKIM is missing.'
    }

    try {
        $mxr = @(Invoke-StDnsQuery -Name $domain -Type MX -ErrorAction Stop | Where-Object { $_.QueryType -eq 'MX' })
        $mx = @($mxr | Sort-Object Preference | ForEach-Object { ('{0} (pref {1})' -f $_.NameExchange, $_.Preference) })
        if ($mx.Count -gt 0) { $rows += New-StRow -Group 'Sender DNS' -Check 'MX' -Result ($mx -join ', ') -Basis 'server-reported' -Status 'info' -Detail 'Where replies to this domain are delivered. Not used for sending.' }
    } catch { }

    $sw.Stop()
    $data = [pscustomobject]@{ domain = $domain; spf = $spf; dmarc = $dmarc; dkim = $dkim; mx = [object[]]$mx }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Resolve-DnsName TXT and MX lookups against 1.1.1.1, falling back to the local resolver' -Rows $rows -Data $data `
            -Caveats @('DKIM cannot be proven absent - selectors are not enumerable.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  Reachability
# ---------------------------------------------------------------------------

function Test-StTcpPort {
    param([string]$Server, [int]$Port, [int]$TimeoutMs = 4000)
    $tcp = New-Object System.Net.Sockets.TcpClient
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $open = $false
    $err = ''
    try {
        $iar = $tcp.BeginConnect($Server, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok -and $tcp.Connected) { $tcp.EndConnect($iar); $open = $true }
        elseif (-not $ok) { $err = ('no answer within {0} ms' -f $TimeoutMs) }
        else { $err = 'refused' }
    } catch {
        $err = $_.Exception.Message
    } finally {
        $sw.Stop()
        try { $tcp.Close() } catch { }
    }
    return [pscustomobject]@{
        port      = $Port
        open      = $open
        latencyMs = $(if ($open) { [int]$sw.ElapsedMilliseconds } else { $null })
        security  = (Get-StPortSecurity -Port $Port)
        error     = $err
    }
}

function Test-StTcpMatrix {
    param([string]$Server, [int[]]$Ports, [int]$TimeoutMs = 4000, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = @()
    $results = @()

    foreach ($p in $Ports) {
        if ($Shared -and $Shared.Cancel) { break }
        if ($Shared) { $Shared.Status = ('Testing TCP {0}:{1}' -f $Server, $p) }
        $r = Test-StTcpPort -Server $Server -Port $p -TimeoutMs $TimeoutMs
        $results += $r
        $label = ('Port {0} ({1})' -f $p, $(if ($r.security -eq 'implicit') { 'implicit SSL' } else { 'plain / STARTTLS' }))
        if ($r.open) {
            # An open port is the single most misleading result in this whole
            # tool, and it is what sent the original investigation after
            # firewalls. Test-NetConnection reports exactly this and stops here.
            $openDetail = 'This completes the TCP handshake only. It says something answered, not that a client can hold a session.'
            if ($r.security -eq 'implicit') {
                $openDetail = 'Open, but this port expects TLS the instant the socket opens. A STARTTLS-only client connects here and then waits for a greeting that never comes. "Open" is not evidence the application can use it - this is the exact trap a port test sets.'
            }
            $rows += New-StRow -Group 'Reachability' -Check $label -Result ('OPEN, {0} ms' -f $r.latencyMs) -Status 'ok' -Detail $openDetail
        } else {
            $rows += New-StRow -Group 'Reachability' -Check $label -Result 'CLOSED or FILTERED' -Status 'fail' -Detail $r.error
        }
    }

    $openCount = @($results | Where-Object { $_.open }).Count
    $sw.Stop()

    $state = 'measured'
    $reason = ''
    if ($results.Count -eq 0) { $state = 'skipped'; $reason = 'Cancelled before any port was tested.' }

    $data = [pscustomobject]@{ server = $Server; ports = [object[]]$results; openCount = $openCount; testedCount = $results.Count }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method ('TcpClient connect with a {0} ms budget per port' -f $TimeoutMs) -Rows $rows -Data $data `
            -Caveats @(
                'A TCP connect proves reachability and nothing else. Test-NetConnection stops at this point too, which is why it reports success on an implicit-SSL port that the application can never actually use.',
                'A single connect cannot see a per-process firewall rule or an outbound filter that only reacts to volume.'
            )) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StPublicIp {
    # Two questions in one: what IP does the world see us as, and what IP does
    # the mail server say it sees. When a provider is configured for IP
    # authentication and mail still bounces, this is almost always the answer.
    param([string]$IpSeenByServer = '', [int]$TimeoutSec = 8)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Set-StNetDefaults

    $rows = @()
    $public = ''
    $source = ''
    foreach ($u in @('https://api.ipify.org','https://checkip.amazonaws.com','https://ifconfig.me/ip')) {
        try {
            $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop
            $txt = ($resp.Content | Out-String).Trim()
            if ($txt -match '^((?:\d{1,3}\.){3}\d{1,3})$') { $public = $Matches[1]; $source = $u; break }
        } catch { }
    }

    if ($public) {
        $rows += New-StRow -Group 'Egress' -Check 'Public IP of this site' -Result $public -Status 'info' -Detail ('Reported by {0}.' -f $source)
    } else {
        $rows += New-StRow -Group 'Egress' -Check 'Public IP of this site' -Result 'Could not be determined' -Basis 'not measured' -Status 'info' `
            -Detail 'No outbound HTTPS lookup succeeded. A proxy or an outbound web block would do this without affecting SMTP.'
    }

    if ($IpSeenByServer) {
        $rows += New-StRow -Group 'Egress' -Check 'IP the mail server sees' -Result $IpSeenByServer -Basis 'server-reported' -Status 'info' `
            -Detail 'Taken from the server''s own EHLO greeting. This is the address to authorise if the provider uses IP authentication.'
        if ($public -and $public -ne $IpSeenByServer) {
            $rows += New-StRow -Group 'Egress' -Check 'Egress consistency' -Result 'Different addresses' -Status 'warn' `
                -Detail ('Web traffic leaves as {0} but SMTP arrives from {1}. If IP authentication was set up from a "what is my IP" page, the wrong address may have been authorised.' -f $public, $IpSeenByServer)
        } elseif ($public) {
            $rows += New-StRow -Group 'Egress' -Check 'Egress consistency' -Result 'Same address for web and SMTP' -Status 'ok'
        }
    } else {
        $rows += New-StRow -Group 'Egress' -Check 'IP the mail server sees' -Result 'Not observed' -Basis 'not measured' -Status 'info' `
            -Detail 'The server did not echo an address in its greeting, or no session was opened.'
    }

    $sw.Stop()
    $data = [pscustomobject]@{ publicIp = $public; source = $source; ipSeenByServer = $IpSeenByServer; match = ($public -ne '' -and $public -eq $IpSeenByServer) }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'HTTPS echo service cross-checked against the SMTP greeting' -Rows $rows -Data $data `
            -Caveats @('A site with more than one egress path can legitimately show two addresses.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  TLS probes
# ---------------------------------------------------------------------------

function Test-StTlsPort {
    # One port, taken as far as a working encrypted SMTP channel, reporting each
    # step. Works for both implicit and STARTTLS ports.
    param(
        [string]$Server, [int]$Port, [string]$Security, [string]$HeloName,
        $Sink = $null, [int]$IoTimeoutMs = 15000, [int]$ConnectTimeoutMs = 10000
    )
    $sec = Resolve-StSecurityMode -Mode $Security -Port $Port
    $out = [pscustomobject]@{
        port = $Port; security = $sec; connected = $false; greeting = ''
        tls = $null; cert = $null; caps = [object[]]@(); authMechs = [object[]]@()
        startTlsAdvertised = $false; secured = $false; remoteIpSeen = ''
        error = ''; elapsedMs = 0
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $s = $null
    try {
        Write-StNote $Sink ('--- {0}:{1} ({2}) ---' -f $Server, $Port, $sec)
        $s = Open-StSession -Server $Server -Port $Port -Security $sec -HeloName $HeloName -Sink $Sink `
                            -IoTimeoutMs $IoTimeoutMs -ConnectTimeoutMs $ConnectTimeoutMs
        $out.connected          = $true
        $out.greeting           = ($s.greeting.lines -join ' ')
        $out.startTlsAdvertised = $s.ehlo.startTls
        $out.caps               = $s.caps.caps
        $out.authMechs          = $s.caps.authMechs
        $out.remoteIpSeen       = $s.caps.remoteIpSeen
        if (-not $out.remoteIpSeen) { $out.remoteIpSeen = $s.ehlo.remoteIpSeen }
        $out.tls                = $s.conn.Tls
        $out.cert               = $s.conn.Cert
        $out.secured            = $s.conn.Secure
    } catch {
        $out.error = $_.Exception.Message
    } finally {
        $sw.Stop()
        $out.elapsedMs = [int]$sw.ElapsedMilliseconds
        if ($s -and $s.conn) { Close-StConnection -Conn $s.conn }
    }
    return $out
}

function Test-StImplicitTls {
    param([string]$Server, [int[]]$Ports, [string]$HeloName, $Sink = $null, [int]$IoTimeoutMs = 15000, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $ports = @($Ports | Where-Object { (Get-StPortSecurity -Port $_) -eq 'implicit' })
    if ($ports.Count -eq 0) {
        return (New-StSection -State 'skipped' -Reason 'This server has no implicit-SSL ports in the list, so there was nothing to test.')
    }

    $rows = @()
    $results = @()
    foreach ($p in $ports) {
        if ($Shared -and $Shared.Cancel) { break }
        if ($Shared) { $Shared.Status = ('Implicit TLS on port {0}' -f $p) }
        $r = Test-StTlsPort -Server $Server -Port $p -Security 'implicit' -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs
        $results += $r
        if ($r.connected -and $r.secured) {
            $rows += New-StRow -Group 'Implicit SSL' -Check ('Port {0} handshake' -f $p) `
                -Result ('{0} / {1}-{2} in {3} ms' -f $r.tls.protocol, $r.tls.cipher, $r.tls.cipherBits, $r.tls.handshakeMs) -Status 'ok'
            $rows += New-StRow -Group 'Implicit SSL' -Check ('Port {0} SMTP over TLS' -f $p) -Result $r.greeting -Basis 'server-reported' -Status 'ok' `
                -Detail 'The banner arrived only after TLS was established. That is what implicit SSL means, and it is why a STARTTLS-only client hangs here.'
            if ($r.authMechs.Count -gt 0) {
                $rows += New-StRow -Group 'Implicit SSL' -Check ('Port {0} AUTH offered' -f $p) -Result ($r.authMechs -join ' ') -Basis 'server-reported' -Status 'ok'
            }
        } else {
            $rows += New-StRow -Group 'Implicit SSL' -Check ('Port {0}' -f $p) -Result 'FAILED' -Status 'fail' -Detail $r.error
        }
    }

    $sw.Stop()
    $anyOk = @($results | Where-Object { $_.secured }).Count -gt 0
    $state = 'measured'
    $reason = ''
    if ($results.Count -eq 0) { $state = 'skipped'; $reason = 'Cancelled before any implicit port was tested.' }
    elseif (-not $anyOk) { $state = 'failed'; $reason = 'No implicit-SSL port completed a TLS handshake.' }

    $data = [pscustomobject]@{ ports = [object[]]$results }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'TCP connect, immediate TLS handshake, then the SMTP banner and EHLO inside the tunnel' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StStartTls {
    param([string]$Server, [int[]]$Ports, [string]$HeloName, $Sink = $null, [int]$IoTimeoutMs = 15000, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $ports = @($Ports | Where-Object { (Get-StPortSecurity -Port $_) -eq 'starttls' })
    if ($ports.Count -eq 0) {
        return (New-StSection -State 'skipped' -Reason 'No explicit / STARTTLS ports were in the list.')
    }

    $rows = @()
    $results = @()
    foreach ($p in $ports) {
        if ($Shared -and $Shared.Cancel) { break }
        if ($Shared) { $Shared.Status = ('STARTTLS on port {0}' -f $p) }
        $r = Test-StTlsPort -Server $Server -Port $p -Security 'starttls' -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs
        $results += $r
        if ($r.connected -and $r.secured) {
            $rows += New-StRow -Group 'STARTTLS' -Check ('Port {0} banner' -f $p) -Result $r.greeting -Basis 'server-reported' -Status 'ok'
            $rows += New-StRow -Group 'STARTTLS' -Check ('Port {0} upgrade' -f $p) `
                -Result ('{0} / {1}-{2} in {3} ms' -f $r.tls.protocol, $r.tls.cipher, $r.tls.cipherBits, $r.tls.handshakeMs) -Status 'ok' `
                -Detail 'Plaintext banner, EHLO, STARTTLS, then the encrypted channel. This is the only sequence System.Net.Mail.SmtpClient can perform.'
            if ($r.authMechs.Count -gt 0) {
                $rows += New-StRow -Group 'STARTTLS' -Check ('Port {0} AUTH offered' -f $p) -Result ($r.authMechs -join ' ') -Basis 'server-reported' -Status 'ok'
            }
        } elseif ($r.connected) {
            $rows += New-StRow -Group 'STARTTLS' -Check ('Port {0}' -f $p) -Result 'Connected but not encrypted' -Status 'warn' -Detail $r.error
        } else {
            $rows += New-StRow -Group 'STARTTLS' -Check ('Port {0}' -f $p) -Result 'FAILED' -Status 'fail' -Detail $r.error
        }
    }

    $sw.Stop()
    $anyOk = @($results | Where-Object { $_.secured }).Count -gt 0
    $state = 'measured'
    $reason = ''
    if ($results.Count -eq 0) { $state = 'skipped'; $reason = 'Cancelled before any explicit port was tested.' }
    elseif (-not $anyOk) { $state = 'failed'; $reason = 'No explicit port completed a STARTTLS upgrade.' }

    $data = [pscustomobject]@{ ports = [object[]]$results }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'Plaintext banner, EHLO, STARTTLS, TLS handshake, second EHLO' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StTlsVersions {
    # Which protocol versions the server will actually accept on the port the
    # application uses. An app pinned to TLS 1.0 against a 1.2-only server fails
    # in exactly the way this tool exists to explain.
    param([string]$Server, [int]$Port, [string]$Security, [string]$HeloName, $Sink = $null, [int]$IoTimeoutMs = 12000, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $sec = Resolve-StSecurityMode -Mode $Security -Port $Port

    $rows = @()
    $results = @()
    foreach ($name in @('Tls','Tls11','Tls12','Tls13')) {
        if ($Shared -and $Shared.Cancel) { break }
        $label = switch ($name) { 'Tls' { 'TLS 1.0' } 'Tls11' { 'TLS 1.1' } 'Tls12' { 'TLS 1.2' } 'Tls13' { 'TLS 1.3' } }
        $proto = Get-StSslProtocol $name
        if ($null -eq $proto) {
            $rows += New-StRow -Group 'TLS versions' -Check $label -Result 'Not available on this .NET/OS' -Basis 'not measured' -Status 'info'
            $results += [pscustomobject]@{ protocol = $label; accepted = $null; error = 'not available in this runtime' }
            continue
        }
        if ($Shared) { $Shared.Status = ('Trying {0} on port {1}' -f $label, $Port) }

        $conn = $null
        $accepted = $false
        $err = ''
        try {
            $conn = New-StConnection -Server $Server -Port $Port -Security $sec -Sink $null -IoTimeoutMs $IoTimeoutMs -SslProtocols $proto
            [void](Connect-StSmtp -Conn $conn)
            if ($sec -eq 'starttls') {
                [void](Read-StGreeting -Conn $conn)
                $e = Invoke-StEhlo -Conn $conn -HeloName $HeloName
                if (-not $e.startTls) { throw 'STARTTLS not advertised on this port.' }
                [void](Invoke-StStartTls -Conn $conn)
            }
            $accepted = $conn.Secure
        } catch {
            $err = $_.Exception.Message
        } finally {
            if ($conn) { Close-StConnection -Conn $conn -NoQuit }
        }

        $results += [pscustomobject]@{ protocol = $label; accepted = $accepted; error = $err }
        if ($accepted) {
            $status = 'ok'
            $detail = ''
            if ($name -eq 'Tls' -or $name -eq 'Tls11') {
                $status = 'info'
                $detail = 'Still accepted. Useful for an old application, but it will be withdrawn eventually.'
            }
            $rows += New-StRow -Group 'TLS versions' -Check $label -Result 'Accepted' -Status $status -Detail $detail
        } else {
            $rows += New-StRow -Group 'TLS versions' -Check $label -Result 'Rejected' -Status 'info' -Detail $err
        }
    }

    $sw.Stop()
    $modern = @($results | Where-Object { $_.accepted -and ($_.protocol -eq 'TLS 1.2' -or $_.protocol -eq 'TLS 1.3') }).Count
    $legacyOnly = ($modern -eq 0 -and @($results | Where-Object { $_.accepted }).Count -gt 0)
    if ($modern -eq 0 -and @($results | Where-Object { $_.accepted }).Count -eq 0) {
        $rows += New-StRow -Group 'TLS versions' -Check 'Summary' -Result 'No version negotiated' -Status 'fail' `
            -Detail 'Nothing this machine offers is acceptable to the server on this port, or the port does not speak TLS at all.'
    } elseif ($legacyOnly) {
        $rows += New-StRow -Group 'TLS versions' -Check 'Summary' -Result 'Only legacy versions accepted' -Status 'warn'
    }

    $data = [pscustomobject]@{ port = $Port; security = $sec; versions = [object[]]$results }
    $state = 'measured'
    $reason = ''
    if ($results.Count -eq 0) { $state = 'skipped'; $reason = 'Cancelled before any version was tried.' }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method ('One connection per protocol version against port {0}' -f $Port) -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StCertificate {
    param([string]$Server, [int]$Port, [string]$Security, [string]$HeloName, $Sink = $null, [int]$IoTimeoutMs = 15000,
          # On-prem Exchange routinely presents an internal-CA or self-signed
          # certificate. That is normal there, so it must not read as TLS
          # interception - but an untrusted chain is still reported, because
          # SmtpClient will refuse it.
          [bool]$ExpectInternalCa = $false)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $r = Test-StTlsPort -Server $Server -Port $Port -Security $Security -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs
    if (-not $r.secured -or -not $r.cert -or -not $r.cert.present) {
        $sw.Stop()
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason ('No certificate was presented on port {0}: {1}' -f $Port, $(if ($r.error) { $r.error } else { 'the channel never became encrypted' })) -Method 'TLS handshake certificate capture') -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    }

    $c = $r.cert
    $rows = @()
    $rows += New-StRow -Group 'Certificate' -Check 'Subject' -Result $c.subject -Basis 'server-reported' -Status 'info'
    $rows += New-StRow -Group 'Certificate' -Check 'Issuer' -Result $c.issuer -Basis 'server-reported' -Status 'info'
    if ($c.sans.Count -gt 0) {
        $shown = $c.sans
        if ($shown.Count -gt 8) { $shown = @($shown[0..7]) + @(('... and {0} more' -f ($c.sans.Count - 8))) }
        $rows += New-StRow -Group 'Certificate' -Check 'Subject alternative names' -Result ($shown -join ', ') -Basis 'server-reported' -Status 'info'
    }

    $vStatus = 'ok'; $vDetail = ''
    if (-not $c.valid) {
        $vStatus = 'fail'
        $vDetail = ('Chain validation reported: {0}. {1}' -f $c.policyErrors, ($c.chainStatus -join '; '))
    }
    $rows += New-StRow -Group 'Certificate' -Check 'Chain validation' -Result $(if ($c.valid) { 'Valid' } else { 'INVALID' }) -Status $vStatus -Detail $vDetail

    $nameStatus = if ($c.hostNameMatches) { 'ok' } else { 'fail' }
    $rows += New-StRow -Group 'Certificate' -Check 'Host name match' -Result $(if ($c.hostNameMatches) { ('Matches {0}' -f $Server) } else { ('Does NOT match {0}' -f $Server) }) -Status $nameStatus

    $expStatus = 'ok'
    if ($null -ne $c.daysUntilExpiry) {
        if ($c.daysUntilExpiry -lt 0)      { $expStatus = 'fail' }
        elseif ($c.daysUntilExpiry -lt 14) { $expStatus = 'warn' }
        $rows += New-StRow -Group 'Certificate' -Check 'Expiry' -Result ('{0} ({1} days)' -f $c.notAfter, $c.daysUntilExpiry) -Status $expStatus
    }

    # The interception tell. Both a legitimate public CA and an inspecting proxy
    # produce a chain that validates; only the issuer distinguishes them.
    $intercept = $false
    if ($ExpectInternalCa -and ($c.selfSigned -or -not $c.issuerLooksPublic)) {
        $kindWord = if ($c.selfSigned) { 'Self-signed' } else { ('Internal CA "{0}"' -f $c.issuerCn) }
        $trustDetail = 'Normal for an on-premises Exchange server. Every client that sends through it has to trust that CA - and this machine does.'
        $st = 'info'
        if (-not $c.valid) {
            $st = 'fail'
            $trustDetail = 'Normal for an on-premises Exchange server, but THIS machine does not trust it. System.Net.Mail.SmtpClient will refuse the connection after STARTTLS with "The remote certificate is invalid according to the validation procedure". Install the issuing CA certificate in the machine''s Trusted Root store, or put a publicly issued certificate on the Exchange receive connector.'
        }
        $rows += New-StRow -Group 'Certificate' -Check 'Issuer' -Result ('{0} - expected for on-prem Exchange' -f $kindWord) -Basis 'inference' -Status $st -Detail $trustDetail
    } elseif ($c.selfSigned) {
        $intercept = $true
        $rows += New-StRow -Group 'Certificate' -Check 'TLS interception' -Result 'Self-signed certificate presented' -Basis 'inference' -Status 'warn' `
            -Detail 'A public mail relay does not present a self-signed certificate. Something between this machine and the server is terminating TLS.'
    } elseif (-not $c.issuerLooksPublic) {
        $intercept = $true
        $rows += New-StRow -Group 'Certificate' -Check 'TLS interception' -Result ('Issued by "{0}", which is not a recognised public CA' -f $c.issuerCn) -Basis 'inference' -Status 'warn' `
            -Detail 'Antivirus or firewall mail scanning re-signs the connection with a local CA. That breaks clients which pin certificates, and adds a delay that can look exactly like a timeout. Check the security product on this machine before blaming the mail provider.'
    } else {
        $rows += New-StRow -Group 'Certificate' -Check 'TLS interception' -Result ('None detected - issued by {0}' -f $c.issuerCn) -Basis 'inference' -Status 'ok'
    }

    $sw.Stop()
    $data = [pscustomobject]@{ port = $Port; certificate = $c; tls = $r.tls; interceptionSuspected = $intercept }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Certificate and chain captured during the live handshake' -Rows $rows -Data $data `
            -Caveats @('Interception is inferred from the issuer, not proven. A corporate CA that legitimately signs internal servers looks identical.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  Authentication and envelope
# ---------------------------------------------------------------------------

function Test-StAuth {
    param(
        [string]$Server, [int]$Port, [string]$Security, [string]$HeloName,
        [string]$Username, [string]$Password, $Sink = $null, [int]$IoTimeoutMs = 15000,
        [string]$Kind = 'generic'
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if ([string]::IsNullOrEmpty($Username) -and [string]::IsNullOrEmpty($Password)) {
        return (New-StSection -State 'skipped' -Reason 'No credentials were entered. If the provider uses IP authentication that is correct - the Egress section is the check that matters instead.')
    }

    $rows = @()
    $s = $null
    $result = $null
    try {
        Write-StNote $Sink ('--- AUTH test on {0}:{1} ---' -f $Server, $Port)
        $s = Open-StSession -Server $Server -Port $Port -Security $Security -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs
        $mechs = @($s.caps.authMechs)

        if ($mechs.Count -eq 0 -and $Kind -eq 'exo-connector') {
            # Expected, not a fault: an Exchange Online MX host never offers a
            # login. The connector authorises by IP. Measured against the real
            # class: SmtpClient with credentials set simply skips AUTH when none
            # is offered, so an application that insists on a username and
            # password still works through a connector.
            $rows += New-StRow -Group 'Authentication' -Check 'AUTH offered' -Result 'None - as expected for an Exchange Online MX' -Basis 'server-reported' -Status 'info' `
                -Detail 'Authorisation on this route is the inbound connector matching this site''s public IP. Any username and password the application holds are ignored: System.Net.Mail.SmtpClient skips AUTH when the server offers none. See the Relay tab.'
            $sw.Stop()
            return (Set-StSectionTiming -Section (New-StSection -State 'skipped' -Reason 'The Exchange Online MX offers no login; this route is authorised by connector IP, so there is nothing to authenticate.' -Method 'EHLO capability inspection' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
        }
        if ($mechs.Count -eq 0) {
            $rows += New-StRow -Group 'Authentication' -Check 'AUTH offered' -Result 'None advertised' -Basis 'server-reported' -Status 'fail' `
                -Detail 'The server offers no authentication on this port in this state. An application that insists on supplying a username and password has nowhere to send them.'
            $sw.Stop()
            return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason 'The server advertised no AUTH mechanisms.' -Method 'EHLO capability inspection' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
        }

        $rows += New-StRow -Group 'Authentication' -Check 'Mechanisms offered' -Result ($mechs -join ' ') -Basis 'server-reported' -Status 'ok'
        $winMech = @($mechs | Where-Object { $_ -eq 'GSSAPI' -or $_ -eq 'NTLM' })
        if ($winMech.Count -gt 0) {
            $rows += New-StRow -Group 'Authentication' -Check 'What the application will use' -Result ('{0} first, then LOGIN' -f $winMech[0]) -Basis 'measured' -Status 'warn' `
                -Detail ('Measured against the real class: System.Net.Mail.SmtpClient tries {0} before LOGIN whenever it is offered, and falls back to LOGIN if that fails. This test uses LOGIN or PLAIN, so its answer can differ from the application''s - confirm with Reproduce app send. It also means a wrong password costs TWO failed logins per send, which matters against an Active Directory lockout policy.' -f $winMech[0])
        }
        $rows += New-StRow -Group 'Authentication' -Check 'Channel' -Result $(if ($s.conn.Secure) { ('Encrypted ({0})' -f $s.conn.Tls.protocol) } else { 'PLAINTEXT' }) `
            -Status $(if ($s.conn.Secure) { 'ok' } else { 'warn' }) `
            -Detail $(if ($s.conn.Secure) { '' } else { 'Credentials would cross the network unprotected. Most servers refuse AUTH here, which an application usually reports as a login failure rather than a security one.' })

        $result = Invoke-StAuthenticate -Conn $s.conn -Username $Username -Password $Password
        if ($result.success) {
            $rows += New-StRow -Group 'Authentication' -Check ('AUTH {0}' -f $result.mech) -Result ('Accepted ({0})' -f $result.code) -Status 'ok' `
                -Detail 'These credentials are valid for this server. A send failure from here is not a password problem.'
        } else {
            $detail = Get-StReplyMeaning -Code $result.code -Text $result.text -Kind $Kind -HasCredentials $true
            if (-not $detail) { $detail = 'The server rejected these credentials.' }
            if ($result.code -eq 538) { $detail = 'The server requires encryption for this mechanism.' }
            $rows += New-StRow -Group 'Authentication' -Check ('AUTH {0}' -f $result.mech) -Result ('REJECTED ({0})' -f $result.code) -Status 'fail' -Detail ('{0} Server said: {1}' -f $detail, $result.text)
        }
    } catch {
        $rows += New-StRow -Group 'Authentication' -Check 'Session' -Result 'FAILED' -Status 'fail' -Detail $_.Exception.Message
        $sw.Stop()
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason $_.Exception.Message -Method 'AUTH LOGIN or AUTH PLAIN over the encrypted channel' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    } finally {
        if ($s -and $s.conn) { Close-StConnection -Conn $s.conn }
    }

    $sw.Stop()
    $data = [pscustomobject]@{
        port = $Port; mechanismsOffered = [object[]]@($s.caps.authMechs)
        mechanismUsed = $result.mech; accepted = $result.success; code = $result.code
        text = $result.text; encrypted = $s.conn.Secure
    }
    $state = if ($result.success) { 'measured' } else { 'failed' }
    $reason = if ($result.success) { '' } else { ('The server rejected the credentials with {0}.' -f $result.code) }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method ('AUTH {0} over the live channel' -f $result.mech) -Rows $rows -Data $data `
            -Caveats @('The password is never written to the transcript, the JSON or any file.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StEnvelope {
    # The most useful single test in the tool: it proves the server will accept
    # this sender and this recipient, without putting a message anywhere.
    param(
        [string]$Server, [int]$Port, [string]$Security, [string]$HeloName,
        [string]$Username, [string]$Password, [string]$From, [string]$To,
        $Sink = $null, [int]$IoTimeoutMs = 15000, [string]$Kind = 'generic'
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $hasCreds = (-not [string]::IsNullOrEmpty($Username)) -or (-not [string]::IsNullOrEmpty($Password))

    if (-not (Test-StEmailAddress $From)) { return (New-StSection -State 'skipped' -Reason 'A valid From address is needed before the envelope can be tested.') }
    if (-not (Test-StEmailAddress $To))   { return (New-StSection -State 'skipped' -Reason 'A valid To address is needed before the envelope can be tested.') }

    $rows = @()
    $s = $null
    $envRes = $null
    $authOk = $null
    try {
        Write-StNote $Sink ('--- Envelope test on {0}:{1} (no message will be delivered) ---' -f $Server, $Port)
        $s = Open-StSession -Server $Server -Port $Port -Security $Security -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs

        if ($hasCreds -and @($s.caps.authMechs).Count -eq 0) {
            # Same as SmtpClient: credentials held but no login offered means
            # carry on unauthenticated rather than fail.
            $rows += New-StRow -Group 'Envelope' -Check 'Authentication' -Result 'Skipped - the server offers no login' -Basis 'measured' -Status 'info' `
                -Detail 'System.Net.Mail.SmtpClient does exactly this: with credentials set and no AUTH offered, it sends without logging in. Authorisation then has to come from the connection itself - an IP-scoped connector.'
        } elseif ($hasCreds) {
            $a = Invoke-StAuthenticate -Conn $s.conn -Username $Username -Password $Password
            $authOk = $a.success
            if (-not $a.success) {
                $am = Get-StReplyMeaning -Code $a.code -Text $a.text -Kind $Kind -HasCredentials $true
                $rows += New-StRow -Group 'Envelope' -Check 'Authentication' -Result ('REJECTED ({0})' -f $a.code) -Status 'fail' -Detail (('{0}  {1}' -f $a.text, $am).Trim())
                $sw.Stop()
                return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason ('Could not authenticate, so the envelope was never offered: {0}' -f $a.text) -Method 'MAIL FROM / RCPT TO with no DATA' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
            }
            $rows += New-StRow -Group 'Envelope' -Check 'Authentication' -Result ('Accepted via {0}' -f $a.mech) -Status 'ok'
        } else {
            $rows += New-StRow -Group 'Envelope' -Check 'Authentication' -Result 'Not attempted (no credentials)' -Basis 'not measured' -Status 'info' `
                -Detail 'Testing the path a provider''s IP authentication would take.'
        }

        $envRes = Invoke-StEnvelope -Conn $s.conn -From $From -To $To

        if ($envRes.mailFromCode -eq 250) {
            $rows += New-StRow -Group 'Envelope' -Check ('MAIL FROM <{0}>' -f $From) -Result 'Accepted (250)' -Status 'ok' `
                -Detail 'The server is willing to relay for this sender. That covers sender-domain verification and, where it applies, IP authentication.'
        } else {
            $d = $envRes.mailFromText
            $mm = Get-StReplyMeaning -Code $envRes.mailFromCode -Text $envRes.mailFromText -Kind $Kind -HasCredentials $hasCreds
            if ($mm) { $d = ('{0}  {1}' -f $d, $mm) }
            elseif ($envRes.mailFromCode -eq 550 -or $envRes.mailFromCode -eq 553) { $d = ('{0}  The sending domain is most likely not verified with the provider, or this account is not permitted to send as it.' -f $d) }
            $rows += New-StRow -Group 'Envelope' -Check ('MAIL FROM <{0}>' -f $From) -Result ('REJECTED ({0})' -f $envRes.mailFromCode) -Status 'fail' -Detail $d
        }

        if ($envRes.rcptToCode -ne 0) {
            if ($envRes.rcptToCode -eq 250 -or $envRes.rcptToCode -eq 251) {
                $rows += New-StRow -Group 'Envelope' -Check ('RCPT TO <{0}>' -f $To) -Result ('Accepted ({0})' -f $envRes.rcptToCode) -Status 'ok'
            } else {
                # "relay access denied" at RCPT TO is the provider saying it
                # does not know who we are. When no credentials were supplied
                # that is also the definitive answer on IP authentication:
                # this egress address is not on the authorised list.
                $rd = $envRes.rcptToText
                $rm = Get-StReplyMeaning -Code $envRes.rcptToCode -Text $envRes.rcptToText -Kind $Kind -HasCredentials ($hasCreds -and [bool]$authOk)
                if ($rm) { $rd = ('{0}  {1}' -f $rd, $rm) }
                $rows += New-StRow -Group 'Envelope' -Check ('RCPT TO <{0}>' -f $To) -Result ('REJECTED ({0})' -f $envRes.rcptToCode) -Status 'fail' -Detail $rd
            }
        }

        if ($envRes.success) {
            $rows += New-StRow -Group 'Envelope' -Check 'Outcome' -Result 'Server would have accepted the message' -Basis 'inference' -Status 'ok' `
                -Detail 'The session was reset before DATA, so nothing was delivered. Everything up to the point of handing over content works.'
        }
    } catch {
        $rows += New-StRow -Group 'Envelope' -Check 'Session' -Result 'FAILED' -Status 'fail' -Detail $_.Exception.Message
        $sw.Stop()
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason $_.Exception.Message -Method 'MAIL FROM / RCPT TO with no DATA' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    } finally {
        if ($s -and $s.conn) { Close-StConnection -Conn $s.conn }
    }

    $sw.Stop()
    $data = [pscustomobject]@{
        port = $Port; from = $From; to = $To; authenticated = $authOk
        mailFromCode = $envRes.mailFromCode; rcptToCode = $envRes.rcptToCode; accepted = $envRes.success
        mailFromText = $envRes.mailFromText; rcptToText = $envRes.rcptToText
    }
    $state = if ($envRes.success) { 'measured' } else { 'failed' }
    $reason = if ($envRes.success) { '' } else { 'The server refused the envelope. The row detail carries its exact reply.' }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'MAIL FROM / RCPT TO / RSET - no DATA, nothing delivered' -Rows $rows -Data $data `
            -Caveats @('Acceptance of an envelope is not delivery. Content filtering and recipient rules act after this point.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StRelay {
    # Anonymous relay: no credentials, an EXTERNAL recipient, stop before DATA.
    #
    # This is how scanners, copiers and line-of-business applications usually
    # send through on-prem Exchange (an IP-scoped relay connector) and through
    # Exchange Online (an inbound connector matching the site's public IP, on
    # port 25 to the tenant's MX host).
    #
    # The recipient has to be outside the organisation. Exchange Online accepts
    # mail for its own users from anyone ("direct send"), and on-prem Exchange
    # accepts internal recipients anonymously by default, so an internal
    # recipient succeeds whether or not any connector exists - it proves
    # nothing about the thing being tested.
    param(
        [string]$Server, [int]$Port = 25, [string]$HeloName,
        [string]$From, [string]$ExternalTo,
        [string[]]$ExpectedIps = @(), [string]$PublicIp = '',
        [string]$Kind = 'generic', [bool]$PortOpen = $true,
        $Sink = $null, [int]$IoTimeoutMs = 15000
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if (-not (Test-StEmailAddress $From))       { return (New-StSection -State 'skipped' -Reason 'A valid From address is needed for the relay test.') }
    if (-not (Test-StEmailAddress $ExternalTo)) { return (New-StSection -State 'skipped' -Reason 'Enter an external recipient on the Relay tab - an address outside this organisation, such as your own.') }

    $rows = @()
    $routeWord = if ($Kind -eq 'exo-connector') { 'inbound connector' } else { 'relay connector' }
    $fromDom = (Get-StDomainFromAddress $From).ToLowerInvariant()
    $toDom   = (Get-StDomainFromAddress $ExternalTo).ToLowerInvariant()
    $internal = ($fromDom -and $fromDom -eq $toDom)
    if ($internal) {
        $rows += New-StRow -Group 'Relay' -Check 'Recipient' -Result ('{0} is in the sender''s own domain' -f $ExternalTo) -Status 'warn' `
            -Detail 'An internal recipient is accepted with or without a connector, so a pass here would prove nothing. Use an address outside the organisation.'
    }

    if (-not $PortOpen) {
        $why = ('Port {0} did not accept a connection.' -f $Port)
        if ($Port -eq 25) { $why = 'Outbound port 25 is blocked from this site. Most often that is the ISP - residential and many business services block it to stop spam - or the site firewall. The connector cannot work until 25 is open to the mail host.' }
        $rows += New-StRow -Group 'Relay' -Check ('Port {0}' -f $Port) -Result 'CLOSED or FILTERED' -Status 'fail' -Detail $why
        $sw.Stop()
        $data = [pscustomobject]@{ port = $Port; portOpen = $false; accepted = $false; seenIp = ''; expectedIps = [object[]]$ExpectedIps; internalRecipient = $internal }
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason $why -Method 'Anonymous MAIL FROM / RCPT TO to an external recipient' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    }

    $conn = $null
    $tlsUsed = $false
    $seenIp = ''
    $envRes = $null
    $meaning = ''
    try {
        Write-StNote $Sink ('--- Relay test on {0}:{1}, no credentials, external recipient (nothing is delivered) ---' -f $Server, $Port)
        # Plaintext first, upgrading only if STARTTLS is offered. A relay
        # connector without TLS is legal - and worth knowing about - so its
        # absence is reported, not treated as a connection failure.
        $conn = New-StConnection -Server $Server -Port $Port -Security 'none' -Sink $Sink -IoTimeoutMs $IoTimeoutMs
        [void](Connect-StSmtp -Conn $conn)
        $g = Read-StGreeting -Conn $conn
        if ($g.code -ne 220) { throw ('The server greeted with {0}: {1}' -f $g.code, ($g.lines -join ' ')) }
        $e = Invoke-StEhlo -Conn $conn -HeloName $HeloName
        if ($e.startTls) {
            [void](Invoke-StStartTls -Conn $conn)
            $tlsUsed = $true
            $e = Invoke-StEhlo -Conn $conn -HeloName $HeloName
        }
        $seenIp = $conn.RemoteIpSeen
        $rows += New-StRow -Group 'Relay' -Check 'Encryption' -Result $(if ($tlsUsed) { ('STARTTLS, {0}' -f $conn.Tls.protocol) } else { 'None offered' }) `
            -Status $(if ($tlsUsed) { 'ok' } else { 'warn' }) `
            -Detail $(if ($tlsUsed) { '' } else { 'The server did not offer STARTTLS on this port, so mail would cross the network unencrypted. Untick SSL in the application for this route, or ask for TLS to be enabled on the connector.' })

        $envRes = Invoke-StEnvelope -Conn $conn -From $From -To $ExternalTo
    } catch {
        $rows += New-StRow -Group 'Relay' -Check 'Session' -Result 'FAILED' -Status 'fail' -Detail $_.Exception.Message
        $sw.Stop()
        $data = [pscustomobject]@{ port = $Port; portOpen = $true; accepted = $false; seenIp = $seenIp; expectedIps = [object[]]$ExpectedIps; internalRecipient = $internal; error = $_.Exception.Message }
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason $_.Exception.Message -Method 'Anonymous MAIL FROM / RCPT TO to an external recipient' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    } finally {
        if ($conn) { Close-StConnection -Conn $conn }
    }

    # Which address the server attributes us to. That - not what a "what is my
    # IP" page says - is what a connector is matched against.
    $expected = @($ExpectedIps | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $ipMatch = $null
    if ($seenIp) {
        $ipDetail = 'Taken from the server''s EHLO reply. This is the address the connector is matched against.'
        $ipStatus = 'info'
        if ($expected.Count -gt 0) {
            $ipMatch = ($expected -contains $seenIp)
            if ($ipMatch) {
                $ipStatus = 'ok'
                $ipDetail = ('Matches the connector address you entered ({0}).' -f ($expected -join ', '))
            } else {
                $ipStatus = 'fail'
                $ipDetail = ('The connector is set for {0}, but this site is arriving as {1}. Typical causes are a second internet connection, a failover link, or an ISP change. Add {1} to the connector.' -f ($expected -join ', '), $seenIp)
            }
        }
        $rows += New-StRow -Group 'Relay' -Check 'Address the server sees' -Result $seenIp -Basis 'server-reported' -Status $ipStatus -Detail $ipDetail
        if ($PublicIp -and $PublicIp -ne $seenIp) {
            $rows += New-StRow -Group 'Relay' -Check 'Web egress address' -Result $PublicIp -Status 'warn' `
                -Detail ('Web traffic leaves as {0} but SMTP arrives as {1}. If the connector was set up from a "what is my IP" page, it will have the wrong address.' -f $PublicIp, $seenIp)
        }
    } else {
        $rows += New-StRow -Group 'Relay' -Check 'Address the server sees' -Result 'Not reported by the server' -Basis 'not measured' -Status 'info' `
            -Detail $(if ($PublicIp) { ('The server did not echo an address. This site''s web egress is {0}, which is the best available guess.' -f $PublicIp) } else { 'The server did not echo an address in its EHLO reply.' })
    }

    if ($envRes.mailFromCode -eq 250) {
        $rows += New-StRow -Group 'Relay' -Check ('MAIL FROM <{0}>' -f $From) -Result 'Accepted (250)' -Status 'ok'
    } else {
        $meaning = Get-StReplyMeaning -Code $envRes.mailFromCode -Text $envRes.mailFromText -Kind $Kind -HasCredentials $false
        $rows += New-StRow -Group 'Relay' -Check ('MAIL FROM <{0}>' -f $From) -Result ('REJECTED ({0})' -f $envRes.mailFromCode) -Status 'fail' -Detail (('{0}  {1}' -f $envRes.mailFromText, $meaning).Trim())
    }
    if ($envRes.rcptToCode -ne 0) {
        if ($envRes.rcptToCode -eq 250 -or $envRes.rcptToCode -eq 251) {
            $rows += New-StRow -Group 'Relay' -Check ('RCPT TO <{0}> (external)' -f $ExternalTo) -Result ('Accepted ({0})' -f $envRes.rcptToCode) -Status 'ok'
        } else {
            $meaning = Get-StReplyMeaning -Code $envRes.rcptToCode -Text $envRes.rcptToText -Kind $Kind -HasCredentials $false
            $rows += New-StRow -Group 'Relay' -Check ('RCPT TO <{0}> (external)' -f $ExternalTo) -Result ('REJECTED ({0})' -f $envRes.rcptToCode) -Status 'fail' -Detail (('{0}  {1}' -f $envRes.rcptToText, $meaning).Trim())
        }
    }

    $accepted = [bool]$envRes.success
    $who = if ($seenIp) { $seenIp } elseif ($PublicIp) { $PublicIp } else { 'this site' }
    if ($accepted -and -not $internal) {
        $rows += New-StRow -Group 'Relay' -Check 'Outcome' -Result ('The {0} accepts {1}' -f $routeWord, $who) -Basis 'inference' -Status 'ok' `
            -Detail ('An application on this machine can send to outside recipients through {0}:{1} with no username or password{2}. The session was reset before DATA, so nothing was delivered.' -f $Server, $Port, $(if ($tlsUsed) { ', SSL ticked' } else { '' }))
    } elseif ($accepted) {
        $rows += New-StRow -Group 'Relay' -Check 'Outcome' -Result 'Accepted, but only for an internal recipient' -Basis 'inference' -Status 'warn' `
            -Detail 'This would have succeeded without any connector. Re-run with an external recipient to test the connector itself.'
    } else {
        $rows += New-StRow -Group 'Relay' -Check 'Outcome' -Result ('The {0} does not accept {1}' -f $routeWord, $who) -Basis 'inference' -Status 'fail' `
            -Detail $(if ($meaning) { $meaning } else { 'The server refused to relay. See the reply above.' })
    }

    $sw.Stop()
    $data = [pscustomobject]@{
        port = $Port; portOpen = $true; tlsUsed = $tlsUsed
        seenIp = $seenIp; publicIp = $PublicIp; expectedIps = [object[]]$expected; ipMatchesExpected = $ipMatch
        mailFromCode = $envRes.mailFromCode; mailFromText = $envRes.mailFromText
        rcptToCode = $envRes.rcptToCode; rcptToText = $envRes.rcptToText
        accepted = ($accepted -and -not $internal); internalRecipient = $internal; meaning = $meaning
        route = $routeWord; server = $Server
    }
    $state = if ($accepted -and -not $internal) { 'measured' } else { 'failed' }
    $reason = if ($state -eq 'measured') { '' } elseif ($internal) { 'Only an internal recipient was tried, which proves nothing about a connector.' } else { ('Relay refused: {0}' -f $(if ($meaning) { $meaning } else { ($envRes.mailFromText + ' ' + $envRes.rcptToText).Trim() })) }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'Anonymous MAIL FROM / RCPT TO to an external recipient, reset before DATA' -Rows $rows -Data $data `
            -Caveats @('Acceptance of the envelope is not delivery: Exchange Online can still quarantine or reject the message afterwards on content or sender reputation.', 'The connector configuration itself cannot be read from here; this tests whether it accepts THIS site, right now.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Invoke-StHttpProbe {
    # One unauthenticated GET, headers only, redirects NOT followed. Built on
    # the same TcpClient + SslStream path as the SMTP engine rather than
    # HttpWebRequest, so the certificate can be captured and reported instead
    # of the request simply failing - and so no validation callback ever has
    # to run on a thread without a runspace.
    param([string]$HostName, [string]$Path = '/autodiscover/autodiscover.xml', [bool]$Https = $true,
          [int]$Port = 0, [int]$TimeoutMs = 8000)
    if ($Port -le 0) { $Port = $(if ($Https) { 443 } else { 80 }) }
    $out = [pscustomobject]@{
        url = ('{0}://{1}{2}{3}' -f $(if ($Https) { 'https' } else { 'http' }), $HostName, $(if (($Https -and $Port -ne 443) -or (-not $Https -and $Port -ne 80)) { ':' + $Port } else { '' }), $Path)
        connected = $false; status = 0; statusText = ''; location = ''; server = ''
        authMethods = [object[]]@(); cert = $null; error = ''
    }
    $tcp = $null
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false) -or -not $tcp.Connected) { throw ('No connection to {0}:{1} within {2} ms.' -f $HostName, $Port, $TimeoutMs) }
        $tcp.EndConnect($iar)
        $out.connected = $true
        $stream = $tcp.GetStream()
        $stream.ReadTimeout = $TimeoutMs; $stream.WriteTimeout = $TimeoutMs

        if ($Https) {
            $captured = @{ Raw = $null; Errors = 'None' }
            $validator = {
                param($sender, $certificate, $chain, $sslPolicyErrors)
                try { $captured.Errors = $sslPolicyErrors.ToString() } catch { }
                try { if ($certificate) { $captured.Raw = $certificate.GetRawCertData() } } catch { }
                return $true
            }.GetNewClosure()
            $ssl = New-Object System.Net.Security.SslStream($stream, $false, ([System.Net.Security.RemoteCertificateValidationCallback]$validator))
            $ssl.ReadTimeout = $TimeoutMs; $ssl.WriteTimeout = $TimeoutMs
            $ssl.AuthenticateAsClient($HostName, $null, (Get-StDefaultSslProtocols), $false)
            $out.cert = Get-StCertificateFacts -RawCertData $captured.Raw -PolicyErrors $captured.Errors -ExpectedHost $HostName
            $stream = $ssl
        }

        $req = ('GET {0} HTTP/1.1{1}Host: {2}{1}User-Agent: SMTP-Toolkit/{3}{1}Accept: */*{1}Connection: close{1}{1}' -f $Path, "`r`n", $HostName, (Get-StVersion))
        $bytes = [Text.Encoding]::ASCII.GetBytes($req)
        $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()

        # Headers only: read to the blank line and stop.
        $buf = [System.Collections.Generic.List[byte]]::new()
        while ($buf.Count -lt 65536) {
            $b = $stream.ReadByte()
            if ($b -lt 0) { break }
            $buf.Add([byte]$b)
            $n = $buf.Count
            if ($n -ge 4 -and $buf[$n-4] -eq 13 -and $buf[$n-3] -eq 10 -and $buf[$n-2] -eq 13 -and $buf[$n-1] -eq 10) { break }
        }
        $head = [Text.Encoding]::ASCII.GetString($buf.ToArray())
        $lines = @($head -split "`r`n")
        if ($lines[0] -match '^HTTP/\d\.\d\s+(\d{3})\s*(.*)$') { $out.status = [int]$Matches[1]; $out.statusText = $Matches[2].Trim() }
        $auth = @()
        foreach ($l in $lines) {
            if ($l -match '^(?i)Location:\s*(.+)$')         { $out.location = $Matches[1].Trim() }
            elseif ($l -match '^(?i)Server:\s*(.+)$')       { $out.server = $Matches[1].Trim() }
            elseif ($l -match '^(?i)WWW-Authenticate:\s*(\S+)') { $auth += $Matches[1].Trim() }
        }
        $out.authMethods = [object[]]@($auth | Select-Object -Unique)
    } catch {
        $out.error = $_.Exception.Message
        if ($_.Exception.InnerException) { $out.error = ('{0} {1}' -f $out.error, $_.Exception.InnerException.Message) }
    } finally {
        if ($tcp) { try { $tcp.Close() } catch { } }
    }
    return $out
}

function Get-StAutodiscoverAnswer {
    # Does this HTTP result count as "Autodiscover answered"? A 401 is the
    # normal unauthenticated answer from Exchange; 200 and 405 also mean the
    # service is there; a redirect counts only if it points at another
    # Autodiscover URL. 404, 403 and a website's home page do not.
    #
    # A redirect is only an answer in two cases: the plain-HTTP step, which
    # exists precisely to redirect (that is how Microsoft 365 vanity domains
    # work), or a redirect straight into Microsoft's own service. A website
    # bouncing https://domain to https://autodiscover.domain is just a pointer
    # to a step that is tested separately - counting it would report success
    # for a domain where nothing actually answers.
    param($Probe, [bool]$Https = $true)
    if (-not $Probe -or $Probe.status -eq 0) { return $false }
    if ($Probe.status -eq 401 -or $Probe.status -eq 200 -or $Probe.status -eq 405) { return $true }
    $isRedirect = ($Probe.status -eq 301 -or $Probe.status -eq 302 -or $Probe.status -eq 307 -or $Probe.status -eq 308)
    if (-not $isRedirect -or $Probe.location -notmatch '(?i)autodiscover') { return $false }
    $target = ''
    try { $target = ([uri]$Probe.location).Host } catch { }
    if ($target -match '(?i)\.outlook\.com$|\.office365\.com$') { return $true }
    if (-not $Https -and $Probe.location -match '^(?i)https://') { return $true }
    return $false
}

function Test-StAutodiscover {
    # Quick answer to one question: does anything answer Autodiscover for this
    # address, and where? Unauthenticated throughout - no password is sent, so
    # there is nothing to lock out. It is the same set of requests Outlook
    # makes when it starts, in roughly the same order.
    param([string]$Email, [int]$TimeoutMs = 8000, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if (-not (Test-StEmailAddress $Email)) { return (New-StSection -State 'skipped' -Reason 'Enter a From address - Autodiscover is looked up for that mailbox.') }
    $domain = (Get-StDomainFromAddress $Email).ToLowerInvariant()
    $adHost = ('autodiscover.{0}' -f $domain)
    $rows = @()
    $answeredBy = ''
    $where = ''

    # 1. Microsoft 365 knows where a mailbox lives, and says so without a
    #    password. This is the one Outlook for Microsoft 365 trusts first.
    if ($Shared) { $Shared.Status = 'Asking Microsoft 365 where the mailbox lives' }
    $v2Url = ''
    $v2Status = ''
    try {
        Set-StNetDefaults
        $u = ('https://autodiscover-s.outlook.com/autodiscover/autodiscover.json/v1.0/{0}?Protocol=AutodiscoverV1' -f [uri]::EscapeDataString($Email))
        $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec ([math]::Max(5, [int]($TimeoutMs / 1000))) -ErrorAction Stop
        $j = $resp.Content | ConvertFrom-Json
        $v2Url = [string]$j.Url
        $v2Status = [string]$resp.StatusCode
    } catch {
        $v2Status = 'no answer'
        try { if ($_.Exception.Response) { $v2Status = [string][int]$_.Exception.Response.StatusCode } } catch { }
    }
    if ($v2Url) {
        $h = ''
        try { $h = ([uri]$v2Url).Host } catch { }
        if ($h -match '(?i)\.outlook\.com$|\.office365\.com$') {
            $where = 'Exchange Online'
            $rows += New-StRow -Group 'Autodiscover' -Check 'Microsoft 365 lookup' -Result 'Mailbox is in Exchange Online' -Basis 'server-reported' -Status 'ok' -Detail ('Microsoft returned {0}' -f $v2Url)
        } else {
            $where = ('on-premises ({0})' -f $h)
            $rows += New-StRow -Group 'Autodiscover' -Check 'Microsoft 365 lookup' -Result ('Points to {0}' -f $h) -Basis 'server-reported' -Status 'info' `
                -Detail ('Microsoft returned {0}. That is usually a hybrid organisation whose mailbox is still on-premises.' -f $v2Url)
        }
        $answeredBy = 'Microsoft 365 Autodiscover'
    }
    # The Microsoft 365 row is written after DNS, because what a miss means
    # depends on where the domain's autodiscover record points.

    # 2. DNS for the autodiscover host, through the public resolver.
    if ($Shared) { $Shared.Status = ('Looking up {0}' -f $adHost) }
    $cname = ''
    $adIps = @()
    try {
        $rr = @(Invoke-StDnsQuery -Name $adHost -Type A -ErrorAction Stop)
        $c = @($rr | Where-Object { $_.QueryType -eq 'CNAME' } | Select-Object -First 1)
        if ($c.Count) { $cname = [string]$c[0].NameHost }
        $adIps = @($rr | Where-Object { $_.QueryType -eq 'A' } | ForEach-Object { [string]$_.IPAddress })
    } catch { }
    $srvTarget = ''
    try {
        $s = @(Invoke-StDnsQuery -Name ('_autodiscover._tcp.{0}' -f $domain) -Type SRV -ErrorAction Stop | Where-Object { $_.QueryType -eq 'SRV' } | Select-Object -First 1)
        if ($s.Count) { $srvTarget = ('{0}:{1}' -f $s[0].NameTarget, $s[0].Port) }
    } catch { }
    $cnameToMs = ($cname -match '(?i)outlook\.com\.?$')
    if ($cname -or $adIps.Count) {
        $rows += New-StRow -Group 'Autodiscover' -Check ('DNS {0}' -f $adHost) -Result $(if ($cname) { ('CNAME {0}' -f $cname.TrimEnd('.')) } else { ($adIps -join ', ') }) -Basis 'server-reported' -Status 'ok' `
            -Detail $(if ($cnameToMs) { 'Points at Microsoft 365, as it should for mailboxes in Exchange Online.' } else { '' })
    } else {
        $rows += New-StRow -Group 'Autodiscover' -Check ('DNS {0}' -f $adHost) -Result 'No record' -Basis 'server-reported' -Status 'info' `
            -Detail 'Outlook falls back to the other methods. For Exchange Online, Microsoft recommends a CNAME to autodiscover.outlook.com.'
    }
    if ($srvTarget) { $rows += New-StRow -Group 'Autodiscover' -Check 'DNS SRV _autodiscover._tcp' -Result $srvTarget -Basis 'server-reported' -Status 'info' }
    if (-not $v2Url) {
        if ($cnameToMs) {
            $rows = @((New-StRow -Group 'Autodiscover' -Check 'Microsoft 365 lookup' -Result ('Microsoft 365 does not know this address ({0})' -f $v2Status) -Basis 'server-reported' -Status 'warn' `
                -Detail ('The domain''s autodiscover record points at Microsoft 365, so the domain is there - but this particular mailbox is not. Check the spelling of {0}, or whether it is an alias rather than a mailbox.' -f $Email))) + $rows
        } else {
            $rows = @((New-StRow -Group 'Autodiscover' -Check 'Microsoft 365 lookup' -Result ('Not known to Microsoft 365 ({0})' -f $v2Status) -Basis 'server-reported' -Status 'info' `
                -Detail 'Normal for an on-premises-only Exchange organisation. The checks below look for the on-premises service.')) + $rows
        }
    }

    # 3. The HTTP endpoints, in the order Outlook tries them. The bare domain
    #    usually lands on the company website: a certificate there that does
    #    not match is the classic cause of an Outlook security prompt at every
    #    start-up, so it is worth flagging even when a later step answers.
    $probes = @(
        @{ label = ('https://{0}' -f $domain); host = $domain; https = $true }
        @{ label = ('https://{0}' -f $adHost); host = $adHost; https = $true }
        @{ label = ('http://{0} (redirect)' -f $adHost); host = $adHost; https = $false }
    )
    $promptRisk = $false
    foreach ($p in $probes) {
        if ($Shared) { $Shared.Status = ('Trying {0}' -f $p.label) }
        $r = Invoke-StHttpProbe -HostName $p.host -Https $p.https -TimeoutMs $TimeoutMs
        $ok = Get-StAutodiscoverAnswer -Probe $r -Https $p.https
        if (-not $r.connected) {
            $nd = $r.error
            if ($p.https -and $p.host -eq $adHost -and $cnameToMs) { $nd = ('{0}  Expected with a CNAME to Microsoft 365: Microsoft does not serve HTTPS for customer host names here, so Outlook uses the redirect or the Microsoft 365 lookup.' -f $r.error) }
            $rows += New-StRow -Group 'Autodiscover' -Check $p.label -Result 'No connection' -Status 'info' -Detail $nd
            continue
        }
        $res = if ($r.status) { ('HTTP {0} {1}' -f $r.status, $r.statusText).Trim() } else { 'No HTTP reply' }
        $detail = @()
        if ($r.location) { $detail += ('Redirects to {0}' -f $r.location) }
        if (@($r.authMethods).Count) { $detail += ('Login methods offered: {0}' -f (@($r.authMethods) -join ', ')) }
        if ($r.error) { $detail += $r.error }
        $status = if ($ok) { 'ok' } else { 'info' }

        if ($p.https -and $r.cert -and $r.cert.present -and -not $r.cert.hostNameMatches) {
            if ($p.host -eq $adHost -and $cnameToMs) {
                $detail += 'Certificate is Microsoft''s, not this domain''s - expected with a CNAME to Microsoft 365. Outlook uses the redirect or the Microsoft 365 lookup instead.'
            } else {
                $promptRisk = $true
                $status = 'warn'
                $detail += ('The certificate here does not cover {0} ({1}). Outlook tries this address and shows a certificate warning when it does - the classic "security alert at every start-up".' -f $p.host, $r.cert.subject)
            }
        }
        if (@($r.authMethods) -contains 'Basic' -and $ok) {
            $status = 'warn'
            $detail += 'Basic authentication is still offered on an internet-facing endpoint.'
        }
        $rows += New-StRow -Group 'Autodiscover' -Check $p.label -Result $(if ($ok) { ('ANSWERS - {0}' -f $res) } else { $res }) -Basis 'measured' -Status $status -Detail ($detail -join '  ')
        if ($ok -and -not $answeredBy) {
            $answeredBy = $p.label
            if (-not $where) {
                $where = $p.host
                if ($r.location) {
                    $lh = ''
                    try { $lh = ([uri]$r.location).Host } catch { }
                    if ($lh -match '(?i)\.outlook\.com$|\.office365\.com$') { $where = 'Exchange Online' } elseif ($lh) { $where = $lh }
                }
            }
        }
    }

    if ($answeredBy) {
        $rows += New-StRow -Group 'Autodiscover' -Check 'Outcome' -Result ('Answers via {0}' -f $answeredBy) -Basis 'inference' -Status $(if ($promptRisk) { 'warn' } else { 'ok' }) `
            -Detail $(if ($where -and $v2Url) { ('Mailbox location: {0}.' -f $where) } elseif ($where) { ('Autodiscover service: {0}. That says where the domain is served, not that this mailbox exists.' -f $where) } else { '' })
    } else {
        $rows += New-StRow -Group 'Autodiscover' -Check 'Outcome' -Result 'Nothing answers Autodiscover for this address' -Basis 'inference' -Status 'fail' `
            -Detail 'Outlook will not be able to set this mailbox up automatically. Add the autodiscover DNS record for the domain, or check the address is right.'
    }

    $sw.Stop()
    $data = [pscustomobject]@{
        email = $Email; domain = $domain; answered = [bool]$answeredBy; answeredBy = $answeredBy
        mailboxLocation = $where; microsoftUrl = $v2Url; cname = $cname; srv = $srvTarget; certificatePromptRisk = $promptRisk
    }
    $state = if ($answeredBy) { 'measured' } else { 'failed' }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $(if ($answeredBy) { '' } else { 'Nothing answered Autodiscover for this address.' }) `
            -Method 'Unauthenticated requests only, in the order Outlook makes them. The address is sent to Microsoft''s public Autodiscover service, as Outlook does.' -Rows $rows -Data $data) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StRawSend {
    # A real delivery through this tool's own engine. On an implicit port this
    # is the proof that matters: if the tool can send on 465 and the
    # application cannot, the port is not the fault - the application's client
    # library is.
    param(
        [string]$Server, [int]$Port, [string]$Security, [string]$HeloName,
        [string]$Username, [string]$Password, [string]$From, [string]$To,
        [string]$Subject = '', [string]$Body = '',
        $Sink = $null, [int]$IoTimeoutMs = 30000, [string]$Kind = 'generic'
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if (-not (Test-StEmailAddress $From)) { return (New-StSection -State 'skipped' -Reason 'A valid From address is required to send.') }
    if (-not (Test-StEmailAddress $To))   { return (New-StSection -State 'skipped' -Reason 'A valid To address is required to send.') }

    $sec = Resolve-StSecurityMode -Mode $Security -Port $Port
    if (-not $Subject) { $Subject = ('SMTP Toolkit test - {0}:{1} ({2})' -f $Server, $Port, $sec) }
    if (-not $Body) {
        $Body = @(
            'This message was sent by SMTP Toolkit as a delivery test.'
            ''
            ('Host        : {0}' -f $env:COMPUTERNAME)
            ('Sent at     : {0}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
            ('Server      : {0}:{1}' -f $Server, $Port)
            ('Security    : {0}' -f $sec)
            ('Engine      : SMTP Toolkit {0} native client (not System.Net.Mail)' -f (Get-StVersion))
            ''
            'If you received this but the application still cannot send, the transport is'
            'proven and the fault is in the application or its settings.'
        ) -join "`r`n"
    }

    $rows = @()
    $s = $null
    $res = $null
    try {
        Write-StNote $Sink ('--- Live send via the native engine on {0}:{1} ---' -f $Server, $Port)
        $s = Open-StSession -Server $Server -Port $Port -Security $sec -HeloName $HeloName -Sink $Sink -IoTimeoutMs $IoTimeoutMs

        if (-not [string]::IsNullOrEmpty($Username) -or -not [string]::IsNullOrEmpty($Password)) {
            $a = Invoke-StAuthenticate -Conn $s.conn -Username $Username -Password $Password
            if (-not $a.success) {
                $rows += New-StRow -Group 'Live send' -Check 'Authentication' -Result ('REJECTED ({0})' -f $a.code) -Status 'fail' -Detail $a.text
                $sw.Stop()
                return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason ('Authentication failed, nothing was sent: {0}' -f $a.text) -Method 'Native SMTP client' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
            }
            $rows += New-StRow -Group 'Live send' -Check 'Authentication' -Result ('Accepted via {0}' -f $a.mech) -Status 'ok'
        }

        $res = Invoke-StSendMessage -Conn $s.conn -From $From -To $To -Subject $Subject -Body $Body
        if ($res.success) {
            $rows += New-StRow -Group 'Live send' -Check ('Delivery on port {0} ({1})' -f $Port, $sec) -Result ('ACCEPTED ({0})' -f $res.code) -Status 'ok' -Detail $res.text
            $rows += New-StRow -Group 'Live send' -Check 'Message-ID' -Result $res.messageId -Status 'info' `
                -Detail 'Quote this when asking the provider to trace the message through their logs.'
        } else {
            $sendMeaning = Get-StReplyMeaning -Code $res.code -Text $res.text -Kind $Kind -HasCredentials ((-not [string]::IsNullOrEmpty($Username)) -or (-not [string]::IsNullOrEmpty($Password)))
            $rows += New-StRow -Group 'Live send' -Check ('Delivery on port {0} ({1})' -f $Port, $sec) -Result ('FAILED at {0} ({1})' -f $res.stage, $res.code) -Status 'fail' -Detail (('{0}  {1}' -f $res.text, $sendMeaning).Trim())
        }
    } catch {
        $rows += New-StRow -Group 'Live send' -Check 'Session' -Result 'FAILED' -Status 'fail' -Detail $_.Exception.Message
        $sw.Stop()
        return (Set-StSectionTiming -Section (New-StSection -State 'failed' -Reason $_.Exception.Message -Method 'Native SMTP client' -Rows $rows) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
    } finally {
        if ($s -and $s.conn) { Close-StConnection -Conn $s.conn }
    }

    $sw.Stop()
    $data = [pscustomobject]@{ port = $Port; security = $sec; accepted = $res.success; code = $res.code; messageId = $res.messageId; subject = $Subject }
    $state = if ($res.success) { 'measured' } else { 'failed' }
    $reason = if ($res.success) { '' } else { ('Rejected at {0}: {1}' -f $res.stage, $res.text) }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'This tool''s own SMTP client, which supports implicit SSL as well as STARTTLS' -Rows $rows -Data $data `
            -Caveats @('Acceptance by the relay is not proof of final delivery to the recipient''s mailbox.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  Application emulation
#
#  This is the part of the tool that answers the question people actually ask:
#  "the network is fine, so why can't the program send mail?"
#
#  System.Net.Mail.SmtpClient - the class inside a very large number of
#  line-of-business Windows applications - only implements STARTTLS. Microsoft
#  documents this: an SSL session established before any protocol commands are
#  sent, which is what port 465 requires, is not supported. Point such an
#  application at 465 with its "use SSL" box ticked and it connects, then waits
#  for a plaintext greeting that the server will never send unencrypted. Both
#  ends wait. The application reports a timeout, which sends everyone looking
#  at firewalls.
# ---------------------------------------------------------------------------

function Invoke-StSmtpClientProbe {
    # Replays what SmtpClient does at the protocol level, stopping before DATA
    # so a sweep can cover every port and SSL combination without delivering a
    # single message. The 465 deadlock reproduces faithfully here, because the
    # deadlock IS waiting for a plaintext banner on an implicit-TLS port.
    param(
        [string]$Server, [int]$Port, [bool]$EnableSsl,
        [string]$Username, [string]$Password, [string]$From, [string]$To,
        [string]$HeloName, [int]$TimeoutMs = 20000, $Sink = $null,
        [string]$Kind = 'generic',
        # Set by the matrix once a login has been rejected: the transport is
        # still measured, the password is not tried again.
        [bool]$SkipAuth = $false
    )
    $hasCreds = (-not [string]::IsNullOrEmpty($Username)) -or (-not [string]::IsNullOrEmpty($Password))
    $out = [pscustomobject]@{
        port = $Port; enableSsl = $EnableSsl; ok = $false
        stage = 'tcp'; failure = ''; elapsedMs = 0; encrypted = $false
        wouldSend = $false; diagnosis = ''
        authTested = $false          # were credentials actually exercised here
        credentialsInClear = $false  # ... and did they cross unencrypted
        authRejected = $false        # the server said no to the password
        authSkipped = ''             # why the login was deliberately not tried
        credentialsIgnored = $false  # no AUTH offered, so SmtpClient sends without
        clientWouldUse = ''          # GSSAPI / NTLM, which the real class prefers
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $conn = $null
    try {
        # 'none' - never 'implicit'. This is the whole point: SmtpClient always
        # opens in plaintext, whatever the port conventionally means.
        $conn = New-StConnection -Server $Server -Port $Port -Security 'none' -Sink $Sink `
                                 -IoTimeoutMs $TimeoutMs -ConnectTimeoutMs ([math]::Min($TimeoutMs, 10000))
        [void](Connect-StSmtp -Conn $conn)

        $out.stage = 'banner'
        $greet = Read-StGreeting -Conn $conn

        $out.stage = 'ehlo'
        $ehlo = Invoke-StEhlo -Conn $conn -HeloName $HeloName

        if ($EnableSsl) {
            $out.stage = 'starttls'
            if (-not $ehlo.startTls) {
                $out.failure = 'The server did not advertise STARTTLS on this port.'
                $out.diagnosis = 'SmtpClient raises "The SMTP server does not support secure connections" here. Untick SSL, or move to a port that offers STARTTLS.'
                return $out
            }
            [void](Invoke-StStartTls -Conn $conn)
            $out.encrypted = $true
            $out.stage = 'ehlo2'
            $ehlo = Invoke-StEhlo -Conn $conn -HeloName $HeloName
        }

        $offered = @($ehlo.authMechs)
        $winMech = @($offered | Where-Object { $_ -eq 'GSSAPI' -or $_ -eq 'NTLM' })
        if ($winMech.Count -gt 0) { $out.clientWouldUse = $winMech[0] }

        if ($hasCreds -and $offered.Count -eq 0) {
            # Measured against the real class: with credentials set and no AUTH
            # offered, SmtpClient skips the login and sends anyway. Replaying
            # that is what lets this probe answer "will an app that insists on
            # a password work through an IP connector" - it will.
            $out.credentialsIgnored = $true
        } elseif ($hasCreds -and $SkipAuth) {
            $out.stage = 'auth'
            $out.authSkipped = 'protecting the account after a rejected password'
            $out.failure = 'Login not attempted - an earlier combination already had this password rejected.'
            $out.diagnosis = 'The transport on this port and SSL setting reached the login step. The password was not tried again, so the account is not pushed towards lockout.'
            return $out
        } elseif ($hasCreds) {
            $out.stage = 'auth'
            $out.authTested = $true
            # SMTP2GO - and Exim generally - advertises AUTH on the plaintext
            # channel and will process a login there. So an unencrypted
            # combination does not fail safely: it succeeds, with the password
            # in clear text across the internet. That has to be recorded as a
            # finding, not quietly counted as a pass.
            if (-not $conn.Secure) { $out.credentialsInClear = $true }
            $a = Invoke-StAuthenticate -Conn $conn -Username $Username -Password $Password
            if (-not $a.success) {
                $out.failure = ('AUTH rejected ({0}): {1}' -f $a.code, $a.text)
                $out.authRejected = ($a.code -eq 535 -or $a.text -match '\b5\.7\.(3|8|139)\b')
                $out.diagnosis = Get-StReplyMeaning -Code $a.code -Text $a.text -Kind $Kind -HasCredentials $true
                if ($a.code -eq 530) { $out.diagnosis = 'The server refuses authentication on an unencrypted channel. The application''s SSL setting has to be on for this port.' }
                return $out
            }
        }

        if ((Test-StEmailAddress $From) -and (Test-StEmailAddress $To)) {
            $out.stage = 'envelope'
            $envRes = Invoke-StEnvelope -Conn $conn -From $From -To $To
            if (-not $envRes.success) {
                $out.failure = ('Envelope refused - MAIL FROM {0}, RCPT TO {1}: {2}' -f $envRes.mailFromCode, $envRes.rcptToCode, ($envRes.mailFromText + ' ' + $envRes.rcptToText).Trim())
                if ($envRes.mailFromCode -ne 250) { $out.diagnosis = Get-StReplyMeaning -Code $envRes.mailFromCode -Text $envRes.mailFromText -Kind $Kind -HasCredentials ($hasCreds -and -not $out.credentialsIgnored) }
                else { $out.diagnosis = Get-StReplyMeaning -Code $envRes.rcptToCode -Text $envRes.rcptToText -Kind $Kind -HasCredentials ($hasCreds -and -not $out.credentialsIgnored) }
                return $out
            }
            $out.wouldSend = $true
        }

        $out.stage = 'complete'
        $out.ok = $true
        $out.diagnosis = 'This combination works. An application using System.Net.Mail.SmtpClient can send with these settings.'
        if ($out.credentialsIgnored) {
            $out.diagnosis = 'This combination works. The server offers no login, so the application''s username and password are simply not used - System.Net.Mail.SmtpClient skips AUTH in that case. The connection itself is what is authorised.'
        }
    } catch {
        $msg = $_.Exception.Message
        $out.failure = $msg
        if ($out.stage -eq 'banner' -and $msg -match '(?i)timed out') {
            # The signature fault this tool was built to name.
            $out.diagnosis = ('Connected to port {0}, then waited for a plaintext greeting that never came. That is implicit SSL: the server expects the TLS handshake first. System.Net.Mail.SmtpClient cannot do that on any port, so the application will sit here until its own timeout expires and then report a timeout. This is the classic false "firewall" fault.' -f $Port)
        } elseif ($out.stage -eq 'tcp') {
            $out.diagnosis = 'The port did not accept a connection at all. This one really is a network or firewall question.'
        }
    } finally {
        $sw.Stop()
        $out.elapsedMs = [int]$sw.ElapsedMilliseconds
        if ($conn) { Close-StConnection -Conn $conn -NoQuit:(-not $out.ok) }
    }
    return $out
}

function Test-StDotNetSmtpClient {
    # The real class, really sending. Nothing else carries the same weight with
    # a vendor: it is their library, their settings, their machine.
    param(
        [string]$Server, [int]$Port, [bool]$EnableSsl,
        [string]$Username, [string]$Password, [string]$From, [string]$To,
        [int]$TimeoutMs = 20000, $Sink = $null, [string]$Kind = 'generic'
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $hasCreds = (-not [string]::IsNullOrEmpty($Username)) -or (-not [string]::IsNullOrEmpty($Password))
    $authRejectedLikely = $false

    if (-not (Test-StEmailAddress $From)) { return (New-StSection -State 'skipped' -Reason 'A valid From address is required - SmtpClient will not send without one.') }
    if (-not (Test-StEmailAddress $To))   { return (New-StSection -State 'skipped' -Reason 'A valid To address is required.') }

    $rows = @()
    $ok = $false
    $errType = ''
    $errMsg = ''
    $statusCode = ''

    Write-StNote $Sink ('--- System.Net.Mail.SmtpClient: {0}:{1}, EnableSsl={2}, timeout {3} ms ---' -f $Server, $Port, $EnableSsl, $TimeoutMs)

    $smtp = $null
    $mail = $null
    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($Server, $Port)
        $smtp.EnableSsl             = $EnableSsl
        $smtp.Timeout               = $TimeoutMs
        $smtp.DeliveryMethod        = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.UseDefaultCredentials = $false
        if (-not [string]::IsNullOrEmpty($Username) -or -not [string]::IsNullOrEmpty($Password)) {
            $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
        } else {
            # Blank credentials rather than none, mirroring how applications that
            # insist on populated fields behave when a site uses IP authentication.
            $smtp.Credentials = New-Object System.Net.NetworkCredential('', '')
        }

        $mail = New-Object System.Net.Mail.MailMessage
        $mail.From = $From
        $mail.To.Add($To)
        $mail.Subject = ('SMTP Toolkit - System.Net.Mail test on port {0}' -f $Port)
        $mail.Body = @(
            'Sent by SMTP Toolkit using System.Net.Mail.SmtpClient - the same class used by many Windows line-of-business applications.'
            ''
            ('Host      : {0}' -f $env:COMPUTERNAME)
            ('Server    : {0}:{1}' -f $Server, $Port)
            ('EnableSsl : {0}' -f $EnableSsl)
            ('Timeout   : {0} ms' -f $TimeoutMs)
            ('Sent at   : {0}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
        ) -join "`r`n"

        $smtp.Send($mail)
        $ok = $true
    } catch [System.Net.Mail.SmtpException] {
        $errType = 'SmtpException'
        $errMsg  = $_.Exception.Message
        try { $statusCode = $_.Exception.StatusCode.ToString() } catch { }
        if ($_.Exception.InnerException) { $errMsg = ('{0} -> {1}' -f $errMsg, $_.Exception.InnerException.Message) }
    } catch {
        $errType = $_.Exception.GetType().Name
        $errMsg  = $_.Exception.Message
        if ($_.Exception.InnerException) { $errMsg = ('{0} -> {1}' -f $errMsg, $_.Exception.InnerException.Message) }
    } finally {
        if ($mail) { try { $mail.Dispose() } catch { } }
        if ($smtp) { try { $smtp.Dispose() } catch { } }
        $sw.Stop()
    }

    $sslLabel = if ($EnableSsl) { 'ticked' } else { 'unticked' }
    if ($ok) {
        $rows += New-StRow -Group 'Application emulation' -Check ('Port {0}, SSL {1}' -f $Port, $sslLabel) -Result ('SENT in {0:N1} s' -f $sw.Elapsed.TotalSeconds) -Status 'ok' `
            -Detail 'System.Net.Mail.SmtpClient delivered successfully with these exact settings. An application using this class should work when configured the same way.'
    } else {
        $detail = $errMsg
        if ($errMsg -match '(?i)operation has timed out' -and (Get-StPortSecurity -Port $Port) -eq 'implicit') {
            $detail = ('{0}{1}{1}This is the documented incompatibility, not a network fault. Port {2} requires implicit SSL - the TLS handshake happens before any SMTP command. SmtpClient only implements STARTTLS, so it connects and waits for a plaintext greeting that never arrives. Microsoft state that the implicit-SSL connection method is not supported by this class.' -f $errMsg, [Environment]::NewLine, $Port)
        } elseif ($errMsg -match '(?i)does not support secure') {
            $detail = ('{0}{1}{1}The server did not offer STARTTLS on this port. Either untick SSL for this port or move to a port that offers it.' -f $errMsg, [Environment]::NewLine)
        } elseif ($hasCreds -and ($errMsg -match '(?i)\b5\.7\.57\b|not authenticated|5\.7\.3\b|5\.7\.8\b|5\.7\.139')) {
            # The misleading one. Measured: after a rejected login SmtpClient
            # carries on to MAIL FROM, the server refuses the unauthenticated
            # sender, and the exception reads "requires a secure connection or
            # the client was not authenticated" with status MustIssueStartTlsFirst.
            # It is a password fault dressed as a TLS fault.
            $authRejectedLikely = $true
            $code = 530
            if ($errMsg -match '\b5\.7\.139\b') { $code = 535 }
            $detail = ('{0}{1}{1}{2}' -f $errMsg, [Environment]::NewLine, (Get-StReplyMeaning -Code $code -Text $errMsg -Kind $Kind -HasCredentials $true))
        } else {
            $m = [regex]::Match($errMsg, '\b([245])(\d\d)\b')
            $basic = 0
            if ($m.Success) { $basic = [int]($m.Groups[1].Value + $m.Groups[2].Value) }
            $mean = Get-StReplyMeaning -Code $basic -Text $errMsg -Kind $Kind -HasCredentials $hasCreds
            if ($mean) { $detail = ('{0}{1}{1}{2}' -f $errMsg, [Environment]::NewLine, $mean) }
        }
        $rows += New-StRow -Group 'Application emulation' -Check ('Port {0}, SSL {1}' -f $Port, $sslLabel) -Result ('FAILED after {0:N1} s' -f $sw.Elapsed.TotalSeconds) -Status 'fail' -Detail $detail
        if ($statusCode) { $rows += New-StRow -Group 'Application emulation' -Check 'SMTP status code' -Result $statusCode -Basis 'server-reported' -Status 'info' }
    }

    $data = [pscustomobject]@{
        port = $Port; enableSsl = $EnableSsl; timeoutMs = $TimeoutMs
        sent = $ok; errorType = $errType; error = $errMsg; statusCode = $statusCode
        elapsedSec = [math]::Round($sw.Elapsed.TotalSeconds, 2)
        authRejectedLikely = $authRejectedLikely
    }
    $state = if ($ok) { 'measured' } else { 'failed' }
    $reason = if ($ok) { '' } else { ('SmtpClient could not send: {0}' -f $errMsg) }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method 'System.Net.Mail.SmtpClient, the class named in most vendor stack traces' -Rows $rows -Data $data `
            -Caveats @('This uses the same class as the application, not the application''s own code. A hardcoded host, port or timeout inside the application is invisible from here - the app config scan is the only view of that.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Invoke-StAppMatrix {
    # Every port against both states of the application's "use SSL" checkbox,
    # so the answer stops being an argument and becomes a grid.
    param(
        [string]$Server, [int[]]$Ports, [string]$Username, [string]$Password,
        [string]$From, [string]$To, [string]$HeloName,
        [int]$TimeoutMs = 15000, [bool]$LiveSend = $false,
        $Sink = $null, $Shared = $null, [string]$Kind = 'generic'
    )
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $rows = @()
    $cells = @()
    $total = ($Ports.Count * 2)
    $done = 0
    # Lockout guard. Against Active Directory every rejected login counts, and
    # the real SmtpClient makes TWO attempts per send when NTLM or GSSAPI is
    # offered. A mistyped password tried in every cell of a four-port matrix
    # could lock the user out in the middle of the diagnosis. So the password
    # is tried until it is rejected once, and never again in this run.
    $authRejected = $false
    $guardCell = $null

    foreach ($p in $Ports) {
        foreach ($ssl in @($true, $false)) {
            if ($Shared -and $Shared.Cancel) { break }
            $done++
            if ($Shared) {
                $Shared.Status = ('Application emulation {0}/{1}: port {2}, SSL {3}' -f $done, $total, $p, $(if ($ssl) { 'on' } else { 'off' }))
                $Shared.Progress = [int](($done / [math]::Max($total,1)) * 100)
            }

            if ($LiveSend -and $authRejected) {
                # The real class cannot be told "transport only", so once the
                # password has been refused the remaining live sends are skipped
                # outright rather than spending two more failed logins each.
                $cell = [pscustomobject]@{
                    port = $p; enableSsl = $ssl; ok = $false; mode = 'live'
                    stage = 'auth'; failure = 'Not sent - an earlier combination already had this password rejected.'
                    elapsedMs = 0; diagnosis = 'Skipped to protect the account from lockout.'
                    authRejected = $false; authSkipped = 'protecting the account after a rejected password'
                    credentialsIgnored = $false; clientWouldUse = ''; authTested = $false; credentialsInClear = $false
                }
            } elseif ($LiveSend) {
                $sec = Test-StDotNetSmtpClient -Server $Server -Port $p -EnableSsl $ssl -Username $Username -Password $Password `
                                               -From $From -To $To -TimeoutMs $TimeoutMs -Sink $Sink -Kind $Kind
                $cellOk = ($sec.state -eq 'measured')
                $rej = [bool]($sec.data -and $sec.data.authRejectedLikely)
                $cell = [pscustomobject]@{
                    port = $p; enableSsl = $ssl; ok = $cellOk; mode = 'live'
                    stage = $(if ($cellOk) { 'sent' } elseif ($rej) { 'auth' } else { 'failed' })
                    failure = $(if ($cellOk) { '' } else { $sec.reason })
                    elapsedMs = [int]($sec.durationSec * 1000)
                    diagnosis = $(if ($rej) { 'The password was rejected. SmtpClient reports this as "requires a secure connection or the client was not authenticated".' } else { '' })
                    authRejected = $rej; authSkipped = ''; credentialsIgnored = $false; clientWouldUse = ''
                    authTested = $true; credentialsInClear = (-not $ssl)
                }
            } else {
                $cell = Invoke-StSmtpClientProbe -Server $Server -Port $p -EnableSsl $ssl -Username $Username -Password $Password `
                                                 -From $From -To $To -HeloName $HeloName -TimeoutMs $TimeoutMs -Sink $Sink `
                                                 -Kind $Kind -SkipAuth $authRejected
                $cell | Add-Member -NotePropertyName 'mode' -NotePropertyValue 'probe' -Force
            }
            $cells += $cell
            if ($cell.authRejected -and -not $authRejected) {
                $authRejected = $true
                $guardCell = $cell
            }

            $sslLabel = if ($ssl) { 'ticked' } else { 'unticked' }
            if ($cell.ok) {
                # A combination that "works" without encryption is not a result
                # to hand anyone as a recommendation, and when no credentials
                # were supplied it did not even exercise the login - which is
                # the step most likely to be what fails in the field.
                $status = 'ok'
                $result = ('WORKS ({0:N1} s)' -f ($cell.elapsedMs / 1000))
                $detail = $cell.diagnosis
                if (-not $ssl) {
                    $status = 'warn'
                    if ($cell.PSObject.Properties['credentialsInClear'] -and $cell.credentialsInClear) {
                        $result = ('works, BUT UNENCRYPTED ({0:N1} s)' -f ($cell.elapsedMs / 1000))
                        $detail = 'The server accepted the login over a plaintext connection, so the username and password crossed the internet in clear text. This combination must not be used, however well it appears to work.'
                    } elseif ($cell.PSObject.Properties['authTested'] -and -not $cell.authTested) {
                        $result = ('reaches the server unencrypted ({0:N1} s)' -f ($cell.elapsedMs / 1000))
                        $detail = 'No credentials were supplied, so the login step was never tested here - this is not evidence that an application configured this way will work. It is also unencrypted. Enter the credentials and run this again.'
                    } else {
                        $detail = 'Works, but with no encryption. Use the SSL-ticked row on the same port instead.'
                    }
                }
                if ($cell.PSObject.Properties['clientWouldUse'] -and $cell.clientWouldUse) {
                    $detail = ('{0}  The server offers {1}: the real application will try {1} before LOGIN, which this probe does not replay - confirm with Reproduce app send.' -f $detail, $cell.clientWouldUse).Trim()
                    if ($status -eq 'ok') { $status = 'warn' }
                }
                $rows += New-StRow -Group 'Settings matrix' -Check ('Port {0}, SSL {1}' -f $p, $sslLabel) `
                    -Result $result -Status $status -Detail $detail
            } elseif ($cell.PSObject.Properties['authSkipped'] -and $cell.authSkipped) {
                $rows += New-StRow -Group 'Settings matrix' -Check ('Port {0}, SSL {1}' -f $p, $sslLabel) `
                    -Result 'reached the login - password not retried' -Basis 'not measured' -Status 'info' -Detail $cell.diagnosis
            } else {
                $failDetail = $(if ($cell.diagnosis) { ('{0}  [{1}]' -f $cell.diagnosis, $cell.failure) } else { $cell.failure })
                # A rejected login still travelled. Someone testing with the
                # real password needs to know it went across unencrypted, even
                # though this particular attempt failed.
                if ($cell.PSObject.Properties['credentialsInClear'] -and $cell.credentialsInClear) {
                    $failDetail = ('{0}  Note: this attempt was made over a plaintext connection, so whatever password was entered crossed the internet in clear text. Change it if it was the live one.' -f $failDetail)
                }
                $rows += New-StRow -Group 'Settings matrix' -Check ('Port {0}, SSL {1}' -f $p, $sslLabel) `
                    -Result ('fails at {0} ({1:N1} s)' -f $cell.stage, ($cell.elapsedMs / 1000)) -Status 'fail' `
                    -Detail $failDetail
            }
        }
        if ($Shared -and $Shared.Cancel) { break }
    }

    $sw.Stop()
    if ($guardCell) {
        $rows += New-StRow -Group 'Settings matrix' -Check 'Lockout guard' -Result ('Password tried once, rejected on port {0}' -f $guardCell.port) -Basis 'measured' -Status 'warn' `
            -Detail 'After the first rejection the password was not tried again in this run, so an Active Directory account is not pushed towards its lockout threshold. Later combinations still measured everything up to the login. Fix the password, then run the matrix again.'
    }
    $working = @($cells | Where-Object { $_.ok })
    $secure  = @($working | Where-Object { $_.enableSsl })
    $state = 'measured'
    $reason = ''
    if ($cells.Count -eq 0) { $state = 'skipped'; $reason = 'Cancelled before any combination was tried.' }

    $data = [pscustomobject]@{
        mode = $(if ($LiveSend) { 'live' } else { 'probe' })
        timeoutMs = $TimeoutMs
        cells = [object[]]$cells
        workingCount = $working.Count
        secureWorkingCount = $secure.Count
        credentialsTested = [bool]((-not [string]::IsNullOrEmpty($Username)) -or (-not [string]::IsNullOrEmpty($Password)))
        lockoutGuardTripped = [bool]$guardCell
        loginAttempts = @($cells | Where-Object { $_.authTested }).Count
    }
    $method = if ($LiveSend) {
        'Real System.Net.Mail.SmtpClient sends, one per combination'
    } else {
        'Protocol-level replay of System.Net.Mail.SmtpClient, stopped before DATA so nothing is delivered'
    }
    return (Set-StSectionTiming -Section (New-StSection -State $state -Reason $reason -Method $method -Rows $rows -Data $data `
            -Caveats @($(if ($LiveSend) { 'Every passing combination delivered a real message.' } else { 'Probe mode reproduces the protocol exchange, not the .NET library itself. Use the single live reproduction button for evidence a vendor cannot wave away.' }))) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  Environment
# ---------------------------------------------------------------------------

function Get-StRegValue {
    # The .NET registry API rather than the PowerShell provider: it is an order
    # of magnitude quicker, and it behaves under SYSTEM in a session 0 shell
    # where the provider can be sluggish.
    param([string]$Path, [string]$Name)
    try {
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($Path)
        if (-not $key) { return $null }
        $v = $key.GetValue($Name, $null)
        $key.Close()
        return $v
    } catch { return $null }
}

function Test-StDotNetCrypto {
    # Reports only. Which TLS version a .NET application negotiates is decided
    # by these values, and getting them wrong produces failures that look
    # exactly like a server problem. Changing them affects every .NET
    # application on the machine, so that stays a decision for the technician.
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = @()

    $dotNet = Get-StDotNetRelease
    $hives = @(
        [pscustomobject]@{ label = '64-bit .NET 4.x'; path = 'SOFTWARE\Microsoft\.NETFramework\v4.0.30319' }
        [pscustomobject]@{ label = '32-bit .NET 4.x'; path = 'SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319' }
    )
    $hiveData = @()
    foreach ($h in $hives) {
        $strong = Get-StRegValue -Path $h.path -Name 'SchUseStrongCrypto'
        $sysTls = Get-StRegValue -Path $h.path -Name 'SystemDefaultTlsVersions'
        $hiveData += [pscustomobject]@{ hive = $h.label; schUseStrongCrypto = $strong; systemDefaultTlsVersions = $sysTls }

        # Below 4.7 these values are what stands between an application and
        # TLS 1.0. From 4.7 the framework defaults to the OS choice anyway, so
        # their absence stops being a finding.
        $status = 'ok'
        $detail = ''
        if ($strong -ne 1) {
            if ($dotNet.release -ge 460798) {
                $status = 'info'
                $detail = 'Not set, but this .NET version already defaults to the operating system''s protocol choice, so it is not a problem here.'
            } else {
                $status = 'warn'
                $detail = 'Not set on a .NET below 4.7. Applications built for this framework may still offer TLS 1.0 first and be refused by a modern mail server.'
            }
        }
        $rows += New-StRow -Group '.NET crypto' -Check ('{0} - SchUseStrongCrypto' -f $h.label) `
            -Result $(if ($null -eq $strong) { 'not set' } else { [string]$strong }) -Basis 'registry' -Status $status -Detail $detail

        $status2 = if ($sysTls -eq 1) { 'ok' } else { 'info' }
        $rows += New-StRow -Group '.NET crypto' -Check ('{0} - SystemDefaultTlsVersions' -f $h.label) `
            -Result $(if ($null -eq $sysTls) { 'not set' } else { [string]$sysTls }) -Basis 'registry' -Status $status2 `
            -Detail $(if ($sysTls -eq 1) { 'The application inherits the operating system''s protocol list, which is what you want.' } else { 'Not set. The framework picks its own protocol list rather than following the OS.' })
    }

    # SCHANNEL client switches. A machine hardened with a script that disabled
    # TLS 1.2 by mistake fails everything, everywhere, and this is where it shows.
    $protoData = @()
    foreach ($proto in @('SSL 3.0','TLS 1.0','TLS 1.1','TLS 1.2','TLS 1.3')) {
        $path = ('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\{0}\Client' -f $proto)
        $enabled  = Get-StRegValue -Path $path -Name 'Enabled'
        $disabled = Get-StRegValue -Path $path -Name 'DisabledByDefault'
        if ($null -eq $enabled -and $null -eq $disabled) { continue }

        $offState = ($enabled -eq 0)
        $protoData += [pscustomobject]@{ protocol = $proto; enabled = $enabled; disabledByDefault = $disabled; off = $offState }

        $status = 'info'
        $detail = ''
        if ($proto -eq 'TLS 1.2' -and $offState) {
            $status = 'fail'
            $detail = 'TLS 1.2 is switched off for client connections on this machine. Almost nothing will connect to a modern mail server until this is put back.'
        } elseif (($proto -eq 'SSL 3.0' -or $proto -eq 'TLS 1.0') -and $offState) {
            $status = 'ok'
            $detail = 'Disabled, which is correct - but an old application pinned to this version will fail here rather than at the server.'
        }
        $rows += New-StRow -Group 'SCHANNEL' -Check ('{0} (client)' -f $proto) `
            -Result ('Enabled={0}, DisabledByDefault={1}' -f $(if ($null -eq $enabled) { '-' } else { $enabled }), $(if ($null -eq $disabled) { '-' } else { $disabled })) `
            -Basis 'registry' -Status $status -Detail $detail
    }
    if ($protoData.Count -eq 0) {
        $rows += New-StRow -Group 'SCHANNEL' -Check 'Protocol overrides' -Result 'None configured' -Basis 'registry' -Status 'ok' `
            -Detail 'No explicit overrides, so Windows negotiates its own defaults. This is the normal state.'
    }

    $spm = ''
    try { $spm = [System.Net.ServicePointManager]::SecurityProtocol.ToString() } catch { }
    $rows += New-StRow -Group '.NET crypto' -Check 'This process negotiates' -Result $spm -Status 'info' `
        -Detail 'What this PowerShell session offers. An application sets its own, so treat this as context rather than a finding.'

    $sw.Stop()
    $data = [pscustomobject]@{ dotNet = $dotNet; hives = [object[]]$hiveData; schannel = [object[]]$protoData; processSecurityProtocol = $spm }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Registry read only - nothing is changed' -Rows $rows -Data $data `
            -Caveats @('These values are machine-wide. Changing them affects every .NET application, which is why this tool reports rather than fixes.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Test-StLocalInterference {
    # Security software, firewall state and proxy configuration - the three
    # local things that most often sit between an application and a mail server
    # while every network test still passes.
    param([bool]$CertInterceptionSuspected = $false, [string]$InterceptingIssuer = '')
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rows = @()
    $av = @()

    try {
        $products = @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName 'AntiVirusProduct' -ErrorAction Stop)
        foreach ($p in $products) {
            # productState is a bit field; bit 0x1000 in the middle byte means
            # real-time protection is on.
            $state = [int]$p.productState
            $rtOn = (($state -band 0x1000) -ne 0)
            $av += [pscustomobject]@{ name = $p.displayName; realtime = $rtOn; state = ('0x{0:X}' -f $state) }
            $rows += New-StRow -Group 'Local software' -Check 'Security product' -Result ('{0} (real-time {1})' -f $p.displayName, $(if ($rtOn) { 'on' } else { 'off' })) -Status 'info' `
                -Detail 'If this product offers mail scanning or SSL inspection, it can proxy SMTP and add latency or substitute the certificate. Check the certificate section alongside this.'
        }
        if ($products.Count -eq 0) {
            $rows += New-StRow -Group 'Local software' -Check 'Security product' -Result 'None registered' -Status 'info'
        }
    } catch {
        $rows += New-StRow -Group 'Local software' -Check 'Security product' -Result 'Could not be enumerated' -Basis 'not measured' -Status 'info' `
            -Detail 'The Security Center namespace is not present. That is normal on Windows Server, where antivirus does not register there.'
    }

    if ($CertInterceptionSuspected) {
        $rows += New-StRow -Group 'Local software' -Check 'TLS inspection' -Result ('Likely - certificate issued by "{0}"' -f $InterceptingIssuer) -Basis 'inference' -Status 'warn' `
            -Detail 'Something is terminating and re-signing TLS between this machine and the mail server. Exclude the mail server, or the application, from that product''s SSL scanning and re-test.'
    }

    $fwData = @()
    # Not $profile - that is an automatic variable holding the profile script path.
    foreach ($fwProfile in @('DomainProfile','StandardProfile','PublicProfile')) {
        $path = ('SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\{0}' -f $fwProfile)
        $on = Get-StRegValue -Path $path -Name 'EnableFirewall'
        if ($null -ne $on) { $fwData += [pscustomobject]@{ profile = $fwProfile; enabled = ($on -eq 1) } }
    }
    if ($fwData.Count -gt 0) {
        $onList = @($fwData | Where-Object { $_.enabled } | ForEach-Object { $_.profile })
        $rows += New-StRow -Group 'Local software' -Check 'Windows Firewall' -Result $(if ($onList.Count -gt 0) { ('On for: {0}' -f ($onList -join ', ')) } else { 'Off for all profiles' }) -Basis 'registry' -Status 'info' `
            -Detail 'Outbound connections are allowed by default. A blocking rule would have shown as a closed port in the Reachability section.'
    }

    $proxy = ''
    try {
        $pe = Get-StRegValue -Path 'SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings' -Name 'ProxyEnable'
        $ps = Get-StRegValue -Path 'SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings' -Name 'ProxyServer'
        if ($pe -eq 1 -and $ps) { $proxy = $ps }
    } catch { }
    if ($proxy) {
        $rows += New-StRow -Group 'Local software' -Check 'System proxy' -Result $proxy -Basis 'registry' -Status 'info' `
            -Detail 'SMTP does not use an HTTP proxy, so this does not affect sending. It does tell you this network filters traffic, which is worth knowing.'
    } else {
        $rows += New-StRow -Group 'Local software' -Check 'System proxy' -Result 'None configured' -Basis 'registry' -Status 'ok'
    }

    $sw.Stop()
    $data = [pscustomobject]@{ antivirus = [object[]]$av; firewall = [object[]]$fwData; proxy = $proxy; tlsInterceptionSuspected = $CertInterceptionSuspected }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Security Center inventory plus registry inspection' -Rows $rows -Data $data `
            -Caveats @('The presence of a security product is not evidence that it is interfering. The certificate issuer is the only direct evidence this tool has.')) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

function Get-StConfigCandidates {
    # A depth-limited, time-budgeted walk. Recursing all of Program Files with
    # a wildcard takes minutes and trips over reparse points; this returns what
    # it found in the time allowed and says so rather than hanging the sweep.
    param([int]$MaxDepth = 5, [int]$BudgetSec = 25)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $found = New-Object System.Collections.Generic.List[string]

    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData)) {
        if ($r -and (Test-Path -LiteralPath $r)) { [void]$roots.Add($r) }
    }
    # Line-of-business applications of this vintage very often sit in a folder
    # straight off the system drive, so those are worth a look too.
    try {
        $sysDrive = ($env:SystemDrive + '\')
        foreach ($d in [System.IO.Directory]::GetDirectories($sysDrive)) {
            $leaf = Split-Path $d -Leaf
            if ($leaf -match '(?i)^(Windows|Users|\$Recycle\.Bin|System Volume Information|PerfLogs|Program Files|Program Files \(x86\)|ProgramData|Recovery|MSOCache|Intel|AMD|NVIDIA)$') { continue }
            [void]$roots.Add($d)
        }
    } catch { }

    $stack = New-Object System.Collections.Generic.Stack[object]
    foreach ($r in $roots) { $stack.Push([pscustomobject]@{ path = $r; depth = 0 }) }

    $timedOut = $false
    while ($stack.Count -gt 0) {
        if ($sw.Elapsed.TotalSeconds -gt $BudgetSec) { $timedOut = $true; break }
        $item = $stack.Pop()
        try {
            foreach ($f in [System.IO.Directory]::GetFiles($item.path, '*.config')) {
                $name = [System.IO.Path]::GetFileName($f)
                if ($name -match '(?i)\.(exe|dll)\.config$' -or $name -match '(?i)^app\.config$') { [void]$found.Add($f) }
            }
        } catch { }
        if ($item.depth -lt $MaxDepth) {
            try {
                foreach ($d in [System.IO.Directory]::GetDirectories($item.path)) {
                    $di = New-Object System.IO.DirectoryInfo($d)
                    if (($di.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    $stack.Push([pscustomobject]@{ path = $d; depth = ($item.depth + 1) })
                }
            } catch { }
        }
    }
    $sw.Stop()
    return [pscustomobject]@{ files = [object[]]$found.ToArray(); timedOut = $timedOut; elapsedSec = [math]::Round($sw.Elapsed.TotalSeconds, 1); rootCount = $roots.Count }
}

function Test-StAppConfigScan {
    # The one piece of visibility this tool has into what an application does
    # internally. A <mailSettings> block or a hardcoded host in a .config
    # silently overrides whatever the application's own settings screen shows,
    # and it is invisible from the network side.
    param([string]$Server = '', [int]$BudgetSec = 25, $Shared = $null)
    $t0 = (Get-Date).ToUniversalTime()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if ($Shared) { $Shared.Status = 'Scanning application configuration files' }
    $scan = Get-StConfigCandidates -BudgetSec $BudgetSec
    $rows = @()
    $hits = @()

    foreach ($f in $scan.files) {
        if ($Shared -and $Shared.Cancel) { break }
        $text = ''
        try {
            $fi = New-Object System.IO.FileInfo($f)
            if ($fi.Length -gt 512000) { continue }
            $text = [System.IO.File]::ReadAllText($f)
        } catch { continue }
        if (-not $text) { continue }

        $markers = @()
        if ($text -match '(?i)<mailSettings')  { $markers += 'mailSettings' }
        if ($text -match '(?i)smtp2go')        { $markers += 'smtp2go' }
        if ($Server -and $text -match ('(?i)' + [regex]::Escape($Server))) { $markers += 'target server name' }
        if ($text -match '(?i)<smtp\s')        { $markers += 'smtp element' }
        if ($text -match '(?i)(smtpserver|smtphost|mailserver|smtpport)\s*"\s*value') { $markers += 'smtp appSetting' }
        if ($markers.Count -eq 0) { continue }

        $host_ = ''; $port_ = ''; $ssl_ = ''
        if ($text -match '(?is)<smtp\b[^>]*>.*?<network\b([^>]*)>') {
            $net = $Matches[1]
            if ($net -match '(?i)host\s*=\s*"([^"]+)"')            { $host_ = $Matches[1] }
            if ($net -match '(?i)port\s*=\s*"([^"]+)"')            { $port_ = $Matches[1] }
            if ($net -match '(?i)enableSsl\s*=\s*"([^"]+)"')       { $ssl_  = $Matches[1] }
        }

        $hits += [pscustomobject]@{ path = $f; markers = [object[]]$markers; host = $host_; port = $port_; enableSsl = $ssl_ }

        $result = ($markers -join ', ')
        $detail = $f
        $status = 'info'
        if ($host_ -or $port_) {
            $result = ('host={0} port={1} enableSsl={2}' -f $(if ($host_) { $host_ } else { '-' }), $(if ($port_) { $port_ } else { '-' }), $(if ($ssl_) { $ssl_ } else { '-' }))
            $status = 'warn'
            $detail = ('{0}{1}{1}These values are applied by the .NET framework itself and override whatever the application''s settings screen shows. If they disagree with the settings the user can see, this file wins.' -f $f, [Environment]::NewLine)
            if ($port_ -eq '465') {
                $status = 'fail'
                $detail = ('{0}{1}{1}Port 465 hardcoded in a .NET mail configuration is the exact fault this tool exists to find. System.Net.Mail.SmtpClient cannot speak implicit SSL, so this configuration times out no matter what the user selects in the application.' -f $f, [Environment]::NewLine)
            }
        }
        $rows += New-StRow -Group 'App configuration' -Check ([System.IO.Path]::GetFileName($f)) -Result $result -Status $status -Detail $detail
    }

    if ($hits.Count -eq 0) {
        $rows += New-StRow -Group 'App configuration' -Check 'Mail configuration in .config files' -Result 'None found' -Status 'ok' `
            -Detail ('{0} configuration files were read. None carried a mail server override.' -f $scan.files.Count)
    }

    $caveats = @('Only .exe.config, .dll.config and app.config files are read, to a limited depth. An application storing its settings in a database or a proprietary file is out of reach.')
    if ($scan.timedOut) {
        $caveats += ('The scan hit its {0} second budget and stopped early, so coverage is partial. Increase the budget if a specific application needs checking.' -f $BudgetSec)
        $rows += New-StRow -Group 'App configuration' -Check 'Scan coverage' -Result 'Partial - time budget reached' -Basis 'not measured' -Status 'info'
    }

    $sw.Stop()
    $data = [pscustomobject]@{ filesScanned = $scan.files.Count; hits = [object[]]$hits; timedOut = $scan.timedOut; scanSec = $scan.elapsedSec }
    return (Set-StSectionTiming -Section (New-StSection -State 'measured' -Method 'Depth-limited walk of Program Files, ProgramData and system-drive application folders' -Rows $rows -Data $data -Caveats $caveats) -StartedUtc $t0 -DurationSec $sw.Elapsed.TotalSeconds)
}

# ---------------------------------------------------------------------------
#  Verdict
# ---------------------------------------------------------------------------

function Get-StRecommendation {
    # Turns what passed into the settings to type into the application. The
    # order of preference is deliberate: a STARTTLS port that a .NET client can
    # actually drive beats an implicit-SSL port that only this tool can.
    param($Sections, [string]$Server, [string]$AppProfileName = '')

    $app = Get-StAppProfile -Name $AppProfileName
    $rec = [pscustomobject]@{
        available   = $false
        server      = $Server
        port        = $null
        useSsl      = $null
        security    = ''
        confidence  = ''
        statement   = ''
        alternatives= [object[]]@()
        avoid       = [object[]]@()
        application = [pscustomobject]@{
            name         = $app.name
            client       = $app.client
            sslMeaning   = $app.sslMeaning
            settingsPath = $app.settingsPath
            fields       = $app.fields
            timeoutNote  = $app.timeoutNote
            summary      = $app.summary
            checklist    = [object[]]@($app.checklist)
        }
    }

    $matrix = $Sections['appMatrix']
    $working = @()
    if ($matrix -and $matrix.data -and $matrix.data.cells) {
        $working = @($matrix.data.cells | Where-Object { $_.ok })
    }

    if ($working.Count -gt 0) {
        # Prefer an encrypted combination, then the provider's usual port, then
        # the lowest-numbered thing that works.
        $ranked = @($working | Sort-Object `
            @{ Expression = { if ($_.enableSsl) { 0 } else { 1 } } }, `
            @{ Expression = { switch ($_.port) { 2525 { 0 } 587 { 1 } 8025 { 2 } 25 { 3 } 80 { 4 } default { 5 } } } }, `
            @{ Expression = { $_.port } })

        $best = $ranked[0]
        $rec.available  = $true
        $rec.port       = $best.port
        $rec.useSsl     = [bool]$best.enableSsl
        $rec.security   = if ($best.enableSsl) { 'STARTTLS' } else { 'none' }
        $rec.confidence = if ($matrix.data.mode -eq 'live') { 'proven by a real delivery' } else { 'proven to the point of handing over the message body' }
        $rec.statement  = ('Set the application to {0}, port {1}, with its SSL / TLS option {2}.' -f $Server, $best.port, $(if ($best.enableSsl) { 'TICKED' } else { 'UNTICKED' }))
        if (-not $best.enableSsl) {
            $rec.statement = ('{0} Only the unencrypted combination worked, which is worth raising with the provider - credentials should not cross the internet in the clear.' -f $rec.statement)
        }

        # Alternatives are encrypted options only. A working unencrypted
        # combination is not a fallback - it is a way to publish the mail
        # password, so it belongs in the list of things not to use.
        $rec.alternatives = [object[]]@($ranked | Select-Object -Skip 1 | Where-Object { $_.enableSsl } | ForEach-Object {
            ('port {0} with SSL ticked' -f $_.port)
        })

        $avoid = @()
        foreach ($c in @($matrix.data.cells | Where-Object { -not $_.ok })) {
            $avoid += ('port {0} with SSL {1} - fails at {2}' -f $c.port, $(if ($c.enableSsl) { 'ticked' } else { 'unticked' }), $c.stage)
        }
        if ($best.enableSsl) {
            foreach ($c in @($working | Where-Object { -not $_.enableSsl })) {
                $avoid += ('port {0} with SSL unticked - it works, but with no encryption' -f $c.port)
            }
        }
        $rec.avoid = [object[]]$avoid
        return $rec
    }

    # No matrix run. Fall back to what the transport tests showed, and say so.
    $st = $Sections['startTls']
    if ($st -and $st.state -eq 'measured' -and $st.data -and $st.data.ports) {
        $ok = @($st.data.ports | Where-Object { $_.secured })
        if ($ok.Count -gt 0) {
            $ranked = @($ok | Sort-Object @{ Expression = { switch ($_.port) { 2525 { 0 } 587 { 1 } 8025 { 2 } default { 5 } } } }, @{ Expression = { $_.port } })
            $rec.available  = $true
            $rec.port       = $ranked[0].port
            $rec.useSsl     = $true
            $rec.security   = 'STARTTLS'
            $rec.confidence = 'inferred from the transport tests - run the settings matrix to confirm end to end'
            $why = 'This port completed a STARTTLS upgrade.'
            if (-not $app.implicitSsl) { $why = 'This port completed a STARTTLS upgrade, which is the only encrypted sequence this application can perform.' }
            $rec.statement  = ('Set the application to {0}, port {1}, with its SSL / TLS option TICKED. {2}' -f $Server, $ranked[0].port, $why)
            $rec.alternatives = [object[]]@($ranked | Select-Object -Skip 1 | ForEach-Object { ('port {0} with SSL ticked' -f $_.port) })
        }
    }

    # A healthy implicit-SSL port is a warning for a STARTTLS-only client and a
    # genuine option for anything else, so which list it lands in depends on
    # what is actually doing the sending.
    $imp = $Sections['implicitTls']
    if ($imp -and $imp.data -and $imp.data.ports) {
        $impOk = @($imp.data.ports | Where-Object { $_.secured })
        if ($impOk.Count -gt 0) {
            if ($app.implicitSsl) {
                $rec.alternatives = [object[]]@(@($rec.alternatives) + @($impOk | ForEach-Object { ('port {0} using implicit SSL' -f $_.port) }))
            } else {
                $rec.avoid = [object[]]@($impOk | ForEach-Object {
                    ('port {0} - healthy, but it is implicit SSL and {1} cannot use it' -f $_.port, $(if ($app.client) { $app.client } else { 'this application' }))
                })
            }
        }
    }
    return $rec
}

function Get-StVerdict {
    param($Sections, [string]$Server, [string]$AppProfileName = '', [string]$Kind = 'generic')

    $subs = @()

    # --- Reachability
    $tcp = $Sections['tcp']
    if ($tcp -and $tcp.state -eq 'measured' -and $tcp.data) {
        if ($tcp.data.openCount -eq 0) {
            $subs += New-StSubscore -Name 'Reachability' -Level 'red' -Summary 'no port on this server accepted a connection' -Remedy 'This is a network, firewall or DNS question - nothing further can be tested until it is fixed.'
        } elseif ($tcp.data.openCount -lt $tcp.data.testedCount) {
            $closed = @($tcp.data.ports | Where-Object { -not $_.open } | ForEach-Object { $_.port })
            $subs += New-StSubscore -Name 'Reachability' -Level 'amber' -Summary ('{0} of {1} ports are open; blocked: {2}' -f $tcp.data.openCount, $tcp.data.testedCount, ($closed -join ', ')) -Remedy 'Configure the application on a port that is open.'
        } else {
            $subs += New-StSubscore -Name 'Reachability' -Level 'green' -Summary 'every port tested is open'
        }
    } else {
        $subs += New-StSubscore -Name 'Reachability' -Level 'notmeasured' -Summary 'the port matrix did not run'
    }

    # --- Transport security
    $st  = $Sections['startTls']
    $imp = $Sections['implicitTls']
    $stOk  = ($st  -and $st.data  -and @($st.data.ports  | Where-Object { $_.secured }).Count -gt 0)
    $impOk = ($imp -and $imp.data -and @($imp.data.ports | Where-Object { $_.secured }).Count -gt 0)
    if ($stOk -or $impOk) {
        $bits = @()
        if ($stOk)  { $bits += 'STARTTLS' }
        if ($impOk) { $bits += 'implicit SSL' }
        $lvl = 'green'
        $sum = ('encryption negotiates cleanly ({0})' -f ($bits -join ' and '))
        $rem = ''
        if (-not $stOk -and $impOk) {
            $lvl = 'amber'
            $sum = 'only implicit SSL works - no port offered a STARTTLS upgrade'
            $rem = 'An application built on System.Net.Mail.SmtpClient cannot use implicit SSL at all, so it has no working encrypted option against this server.'
        }
        $subs += New-StSubscore -Name 'Transport security' -Level $lvl -Summary $sum -Remedy $rem
    } elseif (($st -and $st.state -ne 'skipped') -or ($imp -and $imp.state -ne 'skipped')) {
        $subs += New-StSubscore -Name 'Transport security' -Level 'red' -Summary 'no port completed a TLS handshake' -Remedy 'Check the TLS versions section and the SCHANNEL settings on this machine.'
    } else {
        $subs += New-StSubscore -Name 'Transport security' -Level 'notmeasured' -Summary 'the TLS probes did not run'
    }

    # --- Certificate
    $cert = $Sections['certificate']
    if ($cert -and $cert.state -eq 'measured' -and $cert.data) {
        $c = $cert.data.certificate
        if (-not $c.valid) {
            $subs += New-StSubscore -Name 'Certificate' -Level 'red' -Summary ('the server certificate does not validate ({0})' -f $c.policyErrors) -Remedy 'A client that checks certificates properly will refuse this connection.'
        } elseif ($cert.data.interceptionSuspected) {
            $subs += New-StSubscore -Name 'Certificate' -Level 'amber' -Summary ('the certificate is issued by "{0}", not a public CA' -f $c.issuerCn) -Remedy 'Something local is inspecting TLS. Exclude the mail server from that product and re-test.'
        } elseif ($null -ne $c.daysUntilExpiry -and $c.daysUntilExpiry -lt 14) {
            $subs += New-StSubscore -Name 'Certificate' -Level 'amber' -Summary ('the server certificate expires in {0} days' -f $c.daysUntilExpiry)
        } elseif ($c.selfSigned -or -not $c.issuerLooksPublic) {
            # Only reachable when interception was NOT suspected, i.e. the
            # server kind expects an internal CA and this machine trusts it.
            $subs += New-StSubscore -Name 'Certificate' -Level 'green' -Summary ('valid and trusted here - issued by internal CA "{0}", normal for on-prem Exchange' -f $c.issuerCn) `
                -Remedy 'Any other machine that sends through this server must trust the same CA.'
        } else {
            $subs += New-StSubscore -Name 'Certificate' -Level 'green' -Summary 'valid, matching, publicly issued'
        }
    } else {
        $subs += New-StSubscore -Name 'Certificate' -Level 'notmeasured' -Summary 'the certificate was not inspected'
    }

    # --- Authentication
    $auth = $Sections['auth']
    if ($auth -and $auth.state -eq 'measured' -and $auth.data -and $auth.data.accepted) {
        $subs += New-StSubscore -Name 'Authentication' -Level 'green' -Summary ('the credentials are accepted (AUTH {0})' -f $auth.data.mechanismUsed)
    } elseif ($auth -and $auth.state -eq 'failed') {
        $authRemedy = $auth.reason
        if ($auth.data -and $auth.data.code) {
            $am = Get-StReplyMeaning -Code ([int]$auth.data.code) -Text ([string]$auth.data.text) -Kind $Kind -HasCredentials $true
            if ($am) { $authRemedy = $am }
        }
        $authSum = 'the server rejected the credentials'
        if ($auth.data -and [string]$auth.data.text -match '\b5\.7\.139\b') { $authSum = 'Exchange Online refused basic authentication (5.7.139)' }
        $subs += New-StSubscore -Name 'Authentication' -Level 'red' -Summary $authSum -Remedy $authRemedy
    } else {
        $subs += New-StSubscore -Name 'Authentication' -Level 'notmeasured' -Summary 'no credentials were tested'
    }

    # --- Sender acceptance
    $envSec = $Sections['envelope']
    if ($envSec -and $envSec.state -eq 'measured' -and $envSec.data -and $envSec.data.accepted) {
        $subs += New-StSubscore -Name 'Sender acceptance' -Level 'green' -Summary 'the server accepts this sender and recipient'
    } elseif ($envSec -and $envSec.state -eq 'failed') {
        $d = $envSec.data
        $relayDenied = ($d -and $d.rcptToCode -eq 550 -and $d.mailFromCode -eq 250)
        $noCreds     = ($d -and $null -eq $d.authenticated)
        if ($relayDenied -and $noCreds) {
            # Expected, not broken: we deliberately offered no credentials. It
            # is still a finding, because it rules out IP authentication.
            $subs += New-StSubscore -Name 'Sender acceptance' -Level 'amber' `
                -Summary 'the server accepted the sender but would not relay without authentication' `
                -Remedy 'No credentials were supplied, so this also rules out IP authentication for this site''s public address. Enter the provider''s SMTP username and password and run the sweep again.'
        } elseif ($d -and $d.mailFromCode -ne 250) {
            $subs += New-StSubscore -Name 'Sender acceptance' -Level 'red' `
                -Summary 'the server refused the sending address' `
                -Remedy 'The sending domain is most likely not verified with the provider, or this account may not send as it.'
        } else {
            $subs += New-StSubscore -Name 'Sender acceptance' -Level 'red' -Summary 'the server refused the envelope' -Remedy $envSec.reason
        }
    } else {
        $subs += New-StSubscore -Name 'Sender acceptance' -Level 'notmeasured' -Summary 'the envelope was not offered'
    }

    # --- Application compatibility. The subscore this tool exists for.
    $matrix = $Sections['appMatrix']
    $dotNet = $Sections['dotNetSmtpClient']
    if ($matrix -and $matrix.data -and $matrix.data.cells -and @($matrix.data.cells).Count -gt 0) {
        $cells = @($matrix.data.cells)
        $work  = @($cells | Where-Object { $_.ok })
        if ($work.Count -eq 0) {
            # "Nothing worked" covers two completely different situations, and
            # conflating them sends a technician back to changing ports when
            # the ports were never the problem. A cell that got as far as AUTH
            # or the envelope proved the transport on that port; only the
            # credentials or the relay permission stopped it.
            $reached = @($cells | Where-Object { $_.stage -eq 'auth' -or $_.stage -eq 'envelope' })
            if ($reached.Count -gt 0) {
                $reachedPorts = @($reached | ForEach-Object { $_.port } | Sort-Object -Unique)
                $stageWord = 'the login'
                if (@($reached | Where-Object { $_.stage -eq 'envelope' }).Count -gt 0) { $stageWord = 'the login or relay permission' }
                $subs += New-StSubscore -Name 'Application compatibility' -Level 'amber' `
                    -Summary ('the transport works on port(s) {0} - what failed was {1}, not the port or SSL setting' -f ($reachedPorts -join ', '), $stageWord) `
                    -Remedy 'Supply the provider''s SMTP username and password and run this again. Changing ports will not help; those are already proven.'
            } else {
                $subs += New-StSubscore -Name 'Application compatibility' -Level 'red' -Summary 'no port and SSL combination reached the server' -Remedy 'Nothing got as far as a usable session. This is a transport problem - see the Reachability and STARTTLS sections.'
            }
        } else {
            $implicitBroken = @($cells | Where-Object { -not $_.ok -and (Get-StPortSecurity -Port $_.port) -eq 'implicit' })
            $secureWork = @($work | Where-Object { $_.enableSsl })
            $credsTested = $true
            if ($matrix.data.PSObject.Properties['credentialsTested']) { $credsTested = [bool]$matrix.data.credentialsTested }

            # Only the encrypted combinations count towards a pass. An
            # unencrypted one that happens to work is a finding, not a result.
            $sum = ('{0} of {1} combinations work with encryption' -f $secureWork.Count, $cells.Count)
            $rem = ''
            if ($implicitBroken.Count -gt 0) {
                $ports = @($implicitBroken | ForEach-Object { $_.port } | Sort-Object -Unique)
                $rem = ('Implicit-SSL ports ({0}) fail for this class of application by design, not because of anything on this network.' -f ($ports -join ', '))
            }
            $lvl = 'green'
            if ($secureWork.Count -eq 0) {
                $lvl = 'amber'
                $sum = ('only unencrypted combinations work - {0} of {1}, none of them with TLS' -f $work.Count, $cells.Count)
                $rem = ('Do not configure the application this way. The credentials would cross the internet in clear text. {0}' -f $rem).Trim()
            } elseif (-not $credsTested) {
                $lvl = 'amber'
                $sum = ('{0} of {1} combinations reach the server with encryption, but no credentials were supplied' -f $secureWork.Count, $cells.Count)
                $rem = ('The login step was never exercised, which is the part most likely to fail in the field. Enter the credentials and run the matrix again. {0}' -f $rem).Trim()
            }
            $subs += New-StSubscore -Name 'Application compatibility' -Level $lvl -Summary $sum -Remedy $rem
        }
    } elseif ($dotNet -and $dotNet.state -eq 'measured') {
        $subs += New-StSubscore -Name 'Application compatibility' -Level 'green' -Summary 'System.Net.Mail.SmtpClient sent successfully with the current settings'
    } elseif ($dotNet -and $dotNet.state -eq 'failed') {
        $subs += New-StSubscore -Name 'Application compatibility' -Level 'red' -Summary 'System.Net.Mail.SmtpClient could not send with the current settings' -Remedy 'Run the settings matrix to find a combination that does work.'
    } else {
        $subs += New-StSubscore -Name 'Application compatibility' -Level 'notmeasured' -Summary 'the application emulation did not run'
    }

    # --- Relay / connector: anonymous relay to an outside recipient. The
    # deciding test for the Exchange Online connector route, and a useful one
    # for on-prem anonymous relay; for a hosted relay with SMTP AUTH it only
    # appears when someone asked for it.
    $rel = $Sections['relay']
    $relayRelevant = ($Kind -eq 'exo-connector' -or $Kind -eq 'exchange-onprem' -or ($rel -and $rel.state -ne 'skipped'))
    if ($relayRelevant) {
        $routeName = if ($Kind -eq 'exo-connector') { 'Connector' } else { 'Anonymous relay' }
        if ($rel -and $rel.state -eq 'measured' -and $rel.data -and $rel.data.accepted) {
            $who = if ($rel.data.seenIp) { $rel.data.seenIp } else { 'this site' }
            $subs += New-StSubscore -Name $routeName -Level 'green' -Summary ('accepts {0} for outside recipients, no credentials needed' -f $who)
        } elseif ($rel -and $rel.state -eq 'failed' -and $rel.data) {
            $d = $rel.data
            if ($d.PSObject.Properties['portOpen'] -and -not $d.portOpen) {
                $subs += New-StSubscore -Name $routeName -Level 'red' -Summary ('outbound port {0} is blocked from this site' -f $d.port) -Remedy 'Usually the ISP. Ask for port 25 to be opened to the mail host, or use an authenticated route on 587 instead.'
            } elseif ($d.internalRecipient) {
                $subs += New-StSubscore -Name $routeName -Level 'notmeasured' -Summary 'only an internal recipient was tried, which proves nothing about a connector'
            } elseif ($d.PSObject.Properties['ipMatchesExpected'] -and $d.ipMatchesExpected -eq $false) {
                $subs += New-StSubscore -Name $routeName -Level 'amber' `
                    -Summary ('the connector is set for {0}, but this site arrives as {1}' -f (@($d.expectedIps) -join ', '), $d.seenIp) `
                    -Remedy ('Add {0} to the connector.' -f $d.seenIp)
            } elseif ([string]$d.meaning -match '(?i)block list') {
                $subs += New-StSubscore -Name $routeName -Level 'red' -Summary 'this site''s IP address is on a block list' -Remedy $d.meaning
            } else {
                $lvl = if ($Kind -eq 'exo-connector') { 'red' } else { 'amber' }
                $who = if ($d.seenIp) { $d.seenIp } else { 'this site' }
                $rem = if ($d.meaning) { $d.meaning } else { $rel.reason }
                $subs += New-StSubscore -Name $routeName -Level $lvl -Summary ('refuses to relay for {0}' -f $who) -Remedy $rem
            }
        } else {
            $subs += New-StSubscore -Name $routeName -Level 'notmeasured' -Summary 'the relay test did not run - enter an external recipient on the Relay tab'
        }
    }

    # --- Autodiscover: only when someone pressed the button.
    $ad = $Sections['autodiscover']
    if ($ad -and $ad.state -ne 'skipped' -and $ad.data) {
        if ($ad.data.answered -and $ad.data.certificatePromptRisk) {
            $subs += New-StSubscore -Name 'Autodiscover' -Level 'amber' -Summary ('answers ({0}), but a certificate mismatch will prompt in Outlook' -f $ad.data.answeredBy) -Remedy 'See the Autodiscover rows: the bare domain presents a certificate that does not cover it.'
        } elseif ($ad.data.answered) {
            $subs += New-StSubscore -Name 'Autodiscover' -Level 'green' -Summary ('answers via {0}' -f $ad.data.answeredBy)
        } else {
            $subs += New-StSubscore -Name 'Autodiscover' -Level 'red' -Summary ('nothing answers Autodiscover for {0}' -f $ad.data.email) -Remedy 'Add the autodiscover DNS record for the domain.'
        }
    }

    # --- Environment
    $crypto = $Sections['crypto']
    if ($crypto -and $crypto.state -eq 'measured') {
        $bad = @($crypto.rows | Where-Object { $_.status -eq 'fail' })
        $warn = @($crypto.rows | Where-Object { $_.status -eq 'warn' })
        if ($bad.Count -gt 0) {
            $subs += New-StSubscore -Name 'Environment' -Level 'red' -Summary $bad[0].check -Remedy $bad[0].detail
        } elseif ($warn.Count -gt 0) {
            $subs += New-StSubscore -Name 'Environment' -Level 'amber' -Summary 'the .NET crypto registry settings could hold an old application on TLS 1.0' -Remedy 'See the Environment tab.'
        } else {
            $subs += New-StSubscore -Name 'Environment' -Level 'green' -Summary 'nothing on this machine forces an obsolete protocol'
        }
    } else {
        $subs += New-StSubscore -Name 'Environment' -Level 'notmeasured' -Summary 'the environment checks did not run'
    }

    $health = Get-StHealth -Subscores $subs
    $rec = Get-StRecommendation -Sections $Sections -Server $Server -AppProfileName $AppProfileName

    # The connector route has its own recommendation: no credentials, port 25,
    # and SSL only if the server offered STARTTLS. It wins whenever it was
    # proven, because it is the thing that was being tested.
    if ($Kind -eq 'exo-connector' -and $rel -and $rel.state -eq 'measured' -and $rel.data -and $rel.data.accepted) {
        $d = $rel.data
        $rec.available  = $true
        $rec.server     = $d.server
        $rec.port       = $d.port
        $rec.useSsl     = [bool]$d.tlsUsed
        $rec.security   = $(if ($d.tlsUsed) { 'STARTTLS' } else { 'none' })
        $rec.confidence = 'proven to the point of handing over the message body, with no credentials'
        $rec.statement  = ('Set the application to {0}, port {1}, with its SSL / TLS option {2}, and NO username or password. The inbound connector authorises this site by its IP address ({3}).' -f $d.server, $d.port, $(if ($d.tlsUsed) { 'TICKED' } else { 'UNTICKED' }), $(if ($d.seenIp) { $d.seenIp } else { 'as seen by the server' }))
        $rec.statement  = ('{0} If the application refuses to save without credentials, any value will do: the server offers no login, and System.Net.Mail.SmtpClient skips AUTH when none is offered (measured).' -f $rec.statement)
        $rec.alternatives = [object[]]@()
    }

    # A working recommendation is the most useful sentence in the report, so it
    # replaces the generic headline whenever the overall verdict is not red.
    if ($rec.available -and $health.overall -ne 'red') {
        $health.headline = $rec.statement
    } elseif ($rec.available -and $health.overall -eq 'red') {
        $health.headline = ('{0}  A working combination does exist: {1}' -f $health.headline, $rec.statement)
    }

    return [pscustomobject]@{
        overall        = $health.overall
        headline       = $health.headline
        subscores      = $health.subscores
        notes          = $health.notes
        recommendation = $rec
    }
}

# ---------------------------------------------------------------------------
#  JSON report
# ---------------------------------------------------------------------------

function New-StReport {
    param(
        $Context, $Target, $Sections, $Verdict, $Transcript,
        [string]$Mode = 'gui',
        [datetime]$StartedUtc = [datetime]::MinValue,
        [double]$DurationSec = 0
    )
    # The target block records whether credentials were supplied, never what
    # they were. The username survives because it is diagnostic - a wrong
    # username is a real finding - the password never touches this object.
    return [pscustomobject]@{
        schema        = 'smtp-diagnostic'
        schemaVersion = 1
        tool          = [pscustomobject]@{ name = 'SMTP Toolkit'; version = (Get-StVersion); mode = $Mode }
        collectedUtc  = (ConvertTo-StIso8601 ((Get-Date).ToUniversalTime()))
        startedUtc    = (ConvertTo-StIso8601 $StartedUtc)
        durationSec   = [math]::Round($DurationSec, 2)
        host          = $Context
        target        = $Target
        sections      = $Sections
        verdict       = $Verdict
        transcript    = [object[]]$Transcript
        knownLimits   = (Get-StKnownLimits)
    }
}

function New-StTarget {
    param(
        [string]$Server, [int]$Port, [string]$Security, [string]$From, [string]$To,
        [string]$Username, [string]$HeloName, [int[]]$Ports, [int]$TimeoutMs,
        [string]$Kind = 'generic', [string]$ExternalTo = '', [string[]]$ConnectorIps = @()
    )
    [pscustomobject]@{
        kind            = $Kind
        externalTo      = $ExternalTo
        connectorIps    = [object[]]$ConnectorIps
        server          = $Server
        port            = $Port
        security        = $Security
        portsTested     = [object[]]$Ports
        from            = $From
        to              = $To
        username        = $Username
        hasCredentials  = (-not [string]::IsNullOrEmpty($Username))
        authMode        = $(if ([string]::IsNullOrEmpty($Username)) { 'none / IP authentication' } else { 'username and password' })
        heloName        = $HeloName
        timeoutMs       = $TimeoutMs
    }
}

function ConvertTo-StJson {
    param($Report, [switch]$Compress)
    if ($Compress) { return ($Report | ConvertTo-Json -Depth 14 -Compress) }
    return ($Report | ConvertTo-Json -Depth 14)
}

function Get-StMarkedJson {
    # Wrapped in markers so a report can be recovered from a pasted RMM log or a
    # console transcript, not only from a file.
    param($Report, [switch]$Compress)
    $m = Get-StJsonMarkers
    $json = ConvertTo-StJson -Report $Report -Compress:$Compress
    return ('{0}{1}{2}{1}{3}' -f $m.Begin, [Environment]::NewLine, $json, $m.End)
}

#endregion ENGINE

if ($EngineOnly) { return }

# ===========================================================================
#  From here down is the user interface. Nothing below this line may be
#  referenced by the engine, because the engine also runs inside background
#  runspaces where none of it exists.
# ===========================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ---------------------------------------------------------------------------
#  Console window
# ---------------------------------------------------------------------------
if (-not $NoHideConsole) {
    try {
        if (-not ('StNative.Win' -as [type])) {
            Add-Type -Namespace 'StNative' -Name 'Win' -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("user32.dll")]   public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
        }
        $h = [StNative.Win]::GetConsoleWindow()
        if ($h -ne [System.IntPtr]::Zero) { [void][StNative.Win]::ShowWindow($h, 0) }
    } catch { }
}

# ---------------------------------------------------------------------------
#  Engine source, extracted from this very file
#
#  A background runspace inherits nothing, so the tests have to be rebuilt
#  inside it. Taking the text from the running script - rather than from a
#  second file - is what keeps this tool a single artefact that can be dropped
#  onto a machine through a remote session.
# ---------------------------------------------------------------------------
$SelfSource = ''
try { $SelfSource = $MyInvocation.MyCommand.ScriptBlock.ToString() } catch { }
if ([string]::IsNullOrWhiteSpace($SelfSource)) {
    $selfPath = $PSCommandPath
    if ([string]::IsNullOrWhiteSpace($selfPath)) { $selfPath = $MyInvocation.MyCommand.Path }
    if ($selfPath -and (Test-Path -LiteralPath $selfPath)) { $SelfSource = [System.IO.File]::ReadAllText($selfPath) }
}
$EngineSource = ''
$m = [regex]::Match($SelfSource, '(?s)#region ENGINE(.*?)#endregion ENGINE')
if ($m.Success) { $EngineSource = $m.Groups[1].Value }
if ([string]::IsNullOrWhiteSpace($EngineSource)) {
    [void][System.Windows.Forms.MessageBox]::Show(
        'The engine section of this script could not be read, so background tests cannot run.' + [Environment]::NewLine + [Environment]::NewLine +
        'This happens if the file has been edited and the "#region ENGINE" or "#endregion ENGINE" marker was removed.',
        'SMTP Toolkit', 'OK', 'Error')
    return
}

# ---------------------------------------------------------------------------
#  UI state
# ---------------------------------------------------------------------------
$script:Shared     = $null
$script:ps         = $null
$script:rs         = $null
$script:async      = $null
$script:timer      = $null
$script:busy       = $false
$script:logIndex   = 0
$script:transIndex = 0
$script:Sections   = @{}
$script:Report     = $null
$script:Verdict    = $null
$script:startedUtc = [datetime]::MinValue
$script:startedSw  = $null
$script:Context    = Get-StHostContext
$script:Presets    = @(Get-StPresets)

# ---------------------------------------------------------------------------
#  Control factories
# ---------------------------------------------------------------------------
function New-StButton {
    param([string]$Text, [int]$W = 150, [int]$H = 30, [string]$Tip = '')
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Width = $W; $b.Height = $H
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    if ($Tip) { $script:tips.SetToolTip($b, $Tip) }
    return $b
}

function New-StLabel {
    param([string]$Text, [bool]$Bold = $false, [int]$Size = 9)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true          # so a longer string at a higher DPI cannot clip
    $style = [System.Drawing.FontStyle]::Regular
    if ($Bold) { $style = [System.Drawing.FontStyle]::Bold }
    $l.Font = New-Object System.Drawing.Font('Segoe UI', $Size, $style)
    return $l
}

function New-StTextBox {
    param([int]$W = 220, [string]$Text = '', [bool]$Password = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $W
    $t.Text = $Text
    $t.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    if ($Password) { $t.UseSystemPasswordChar = $true }
    return $t
}

function New-StResultGrid {
    # One grid shape everywhere. Basis is a visible column on purpose: a
    # technician can see at a glance which lines were measured, which were
    # inferred and which were only read out of the registry.
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $false
    $g.AutoSizeColumnsMode = 'Fill'
    $g.BorderStyle = 'None'
    $g.BackgroundColor = [System.Drawing.Color]::White
    $g.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    # A DataGridView re-selects its first row whenever the tab it sits on is
    # activated, and clearing the selection does not survive that. Rather than
    # fight it, the selection is made pale enough that it cannot be mistaken
    # for a status colour or hide the one underneath it.
    $g.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(214, 228, 245)
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

    foreach ($c in @(
        @{ n = 'group';  h = 'Area';   w = 14 },
        @{ n = 'check';  h = 'Check';  w = 22 },
        @{ n = 'result'; h = 'Result'; w = 26 },
        @{ n = 'basis';  h = 'Basis';  w = 10 },
        @{ n = 'detail'; h = 'What it means'; w = 28 }
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.n
        $col.HeaderText = $c.h
        $col.FillWeight = $c.w
        $g.Columns.Add($col) | Out-Null
    }
    return $g
}

function Add-StGridRow {
    param($Grid, $Row)
    $i = $Grid.Rows.Add(@($Row.group, $Row.check, $Row.result, $Row.basis, $Row.detail))
    $r = $Grid.Rows[$i]
    switch ($Row.status) {
        'ok'   { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 247, 234) }
        'warn' { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 246, 224) }
        'fail' { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232) }
        default { }
    }
    return $i
}

function New-StMonoBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true
    $t.ReadOnly = $true
    $t.Dock = 'Fill'
    $t.ScrollBars = 'Both'
    $t.WordWrap = $false
    $t.BackColor = [System.Drawing.Color]::White
    $t.Font = New-Object System.Drawing.Font('Consolas', 9)
    return $t
}

# ---------------------------------------------------------------------------
#  Background worker
#
#  Runs in its own runspace with the engine rebuilt from source. Results are
#  handed back through a synchronized hashtable that a WinForms timer polls -
#  a single .NET send against an implicit-SSL port blocks for its full timeout,
#  and a window that stops repainting for a hundred seconds looks broken.
# ---------------------------------------------------------------------------
$WorkerScript = {
    param($Shared, $EngineSource, $Job, $Opt)

    . ([scriptblock]::Create($EngineSource))

    function Log { param([string]$m) [void]$Shared.Log.Add($m) }
    function Publish {
        param([string]$Name, $Section)
        $Shared.Sections[$Name] = $Section
        $Shared.Version = $Shared.Version + 1
    }
    function Stopping { return [bool]$Shared.Cancel }

    $sink = $Shared.Transcript
    $server   = [string]$Opt.Server
    $port     = [int]$Opt.Port
    $security = [string]$Opt.Security
    $helo     = [string]$Opt.HeloName
    $user     = [string]$Opt.Username
    $pass     = [string]$Opt.Password
    $from     = [string]$Opt.From
    $to       = [string]$Opt.To
    $ioMs     = [int]$Opt.TimeoutMs
    $ports    = @($Opt.Ports)
    $kind     = [string]$Opt.ServerKind
    if (-not $kind) { $kind = 'generic' }
    $expectCa = [bool]$Opt.ExpectInternalCa
    $extTo    = [string]$Opt.ExternalTo
    $connIps  = @($Opt.ConnectorIps)

    # Where the relay test points. The Exchange Online connector route is port
    # 25 by definition; on-prem anonymous relay is normally 25 too.
    $relayPort = $port
    if ($kind -ne 'exo-connector' -and ($ports -contains 25)) { $relayPort = 25 }

    function Invoke-Relay {
        param([bool]$PortOpen = $true)
        $pubIp = ''
        $pubSec = $Shared.Sections['publicIp']
        if ($pubSec -and $pubSec.data) { $pubIp = [string]$pubSec.data.publicIp }
        $Shared.Status = ('Relay test on port {0} (no credentials, nothing is delivered)' -f $relayPort)
        $rel = Test-StRelay -Server $server -Port $relayPort -HeloName $helo -From $from -ExternalTo $extTo `
                            -ExpectedIps $connIps -PublicIp $pubIp -Kind $kind -PortOpen $PortOpen -Sink $sink -IoTimeoutMs $ioMs
        Publish 'relay' $rel
        if ($rel.data -and $rel.data.accepted) { Log ('Relay accepted: {0} lets this site send to outside recipients without credentials.' -f $server) }
        elseif ($rel.state -ne 'skipped') { Log ('Relay refused: {0}' -f $rel.reason) }
    }

    try {
        Set-StNetDefaults

        # Exchange Online connector route: the host is the tenant MX, taken
        # from the From domain using a public resolver so an AD-integrated
        # zone inside the network cannot hand back an empty answer.
        if ($kind -eq 'exo-connector') {
            $Shared.Status = 'Looking up the From domain''s MX'
            $dom = Get-StDomainFromAddress $from
            $mxInfo = Resolve-StMx -Domain $dom
            $mxRows = @()
            if ($mxInfo.best) {
                $mxRows += New-StRow -Group 'MX' -Check ('MX for {0}' -f $dom) -Result ($(if (@($mxInfo.publicMx).Count) { @($mxInfo.publicMx) } else { @($mxInfo.localMx) }) -join ', ') -Basis 'server-reported' -Status $(if ($mxInfo.isExchangeOnline) { 'ok' } else { 'warn' }) `
                    -Detail $(if ($mxInfo.isExchangeOnline) { 'Hosted in Exchange Online. The connector route sends to this host on port 25.' } else { 'This MX is not an Exchange Online host (*.mail.protection.outlook.com). Either mail for this domain is filtered by a third party first, or it is not hosted in Exchange Online - in which case an Exchange Online connector is not the route to use.' })
                if ($mxInfo.splitHorizon) {
                    $mxRows += New-StRow -Group 'MX' -Check 'Inside this network' -Result ($(if (@($mxInfo.localMx).Count) { @($mxInfo.localMx) -join ', ' } else { 'no MX at all' })) -Basis 'server-reported' -Status 'info' `
                        -Detail 'The local DNS server answers differently from the internet - usually an Active Directory zone with the same name as the public domain. The public answer is used.'
                }
            } else {
                $mxRows += New-StRow -Group 'MX' -Check ('MX for {0}' -f $dom) -Result 'Not found' -Status 'fail' -Detail $mxInfo.error
            }
            $mxState = if ($mxInfo.best) { 'measured' } else { 'failed' }
            Publish 'mx' (New-StSection -State $mxState -Reason $(if ($mxInfo.best) { '' } else { $mxInfo.error }) -Method 'MX lookup against 1.1.1.1 and the local resolver' -Rows $mxRows -Data $mxInfo)
            if ([string]::IsNullOrWhiteSpace($server)) {
                if (-not $mxInfo.best) {
                    Log ('No server was entered and no MX could be found for "{0}". Enter the From address, or the tenant host name, and try again.' -f $dom)
                    $Shared.Done = $true
                    return
                }
                $server = $mxInfo.best
                $Shared.ResolvedServer = $server
                Log ('Using the MX for {0}: {1}' -f $dom, $server)
            }
        }

        switch ($Job) {

            'sweep' {
                $steps = 10
                $n = 0

                $Shared.Status = 'Reading this machine'
                $ctxSec = Test-StContextSection
                Publish 'context' $ctxSec
                $n++; $Shared.Progress = [int](($n / $steps) * 100)
                if ($ctxSec.data) {
                    $who = $ctxSec.data.runAs
                    if ($ctxSec.data.isSystem) { $who = ('{0} (SYSTEM)' -f $who) }
                    Log ('Testing from {0} as {1}.' -f $ctxSec.data.computerName, $who)
                }

                if (Stopping) { break }
                $Shared.Status = ('Resolving {0}' -f $server)
                $dns = Test-StDns -Server $server
                Publish 'dns' $dns
                $n++; $Shared.Progress = [int](($n / $steps) * 100)
                if ($dns.state -eq 'measured' -and $dns.data) {
                    Log ('{0} resolves to {1}.' -f $server, (@($dns.data.ipv4) -join ', '))
                } else {
                    Log ('{0} could not be resolved. Nothing further will work until that is fixed.' -f $server)
                    $Shared.Done = $true
                    return
                }

                if (Stopping) { break }
                $Shared.Status = 'Testing every port'
                $tcp = Test-StTcpMatrix -Server $server -Ports $ports -Shared $Shared
                Publish 'tcp' $tcp
                $n++; $Shared.Progress = [int](($n / $steps) * 100)
                if ($tcp.data) { Log ('{0} of {1} ports are open.' -f $tcp.data.openCount, $tcp.data.testedCount) }

                $openPorts = @()
                if ($tcp.data) { $openPorts = @($tcp.data.ports | Where-Object { $_.open } | ForEach-Object { $_.port }) }
                if ($openPorts.Count -eq 0) {
                    Log 'No port accepted a connection, so the encrypted tests were skipped.'
                    Publish 'implicitTls' (New-StSection -State 'skipped' -Reason 'No port was reachable.')
                    Publish 'startTls'    (New-StSection -State 'skipped' -Reason 'No port was reachable.')
                } else {
                    if (Stopping) { break }
                    $Shared.Status = 'Testing STARTTLS'
                    $st = Test-StStartTls -Server $server -Ports $openPorts -HeloName $helo -Sink $sink -IoTimeoutMs $ioMs -Shared $Shared
                    Publish 'startTls' $st
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)

                    if (Stopping) { break }
                    $Shared.Status = 'Testing implicit SSL'
                    $imp = Test-StImplicitTls -Server $server -Ports $openPorts -HeloName $helo -Sink $sink -IoTimeoutMs $ioMs -Shared $Shared
                    Publish 'implicitTls' $imp
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)

                    # The address the server echoes back is the one that matters
                    # for IP authentication, so take it from whichever probe saw it.
                    $seenIp = ''
                    foreach ($sec in @($st, $imp)) {
                        if ($sec -and $sec.data -and $sec.data.ports) {
                            foreach ($pr in $sec.data.ports) { if ($pr.remoteIpSeen) { $seenIp = $pr.remoteIpSeen; break } }
                        }
                        if ($seenIp) { break }
                    }
                    $Shared.SeenIp = $seenIp

                    if (Stopping) { break }
                    $Shared.Status = 'Checking TLS versions'
                    Publish 'tlsVersions' (Test-StTlsVersions -Server $server -Port $port -Security $security -HeloName $helo -Shared $Shared)
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)

                    if (Stopping) { break }
                    $Shared.Status = 'Auditing the certificate'
                    $cert = Test-StCertificate -Server $server -Port $port -Security $security -HeloName $helo -Sink $sink -IoTimeoutMs $ioMs -ExpectInternalCa $expectCa
                    Publish 'certificate' $cert
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)
                    if ($cert.data -and $cert.data.interceptionSuspected) {
                        Log ('The certificate is issued by "{0}", not a public CA. Something local is inspecting TLS.' -f $cert.data.certificate.issuerCn)
                    }

                    $Shared.Status = 'Checking egress addresses'
                    Publish 'publicIp' (Test-StPublicIp -IpSeenByServer $seenIp)

                    if (Stopping) { break }
                    $passwordRejected = $false
                    if ($user -or $pass) {
                        $Shared.Status = 'Testing the credentials'
                        $auth = Test-StAuth -Server $server -Port $port -Security $security -HeloName $helo -Username $user -Password $pass -Sink $sink -IoTimeoutMs $ioMs -Kind $kind
                        Publish 'auth' $auth
                        if ($auth.data -and $auth.data.accepted) { Log 'The credentials were accepted.' }
                        elseif ($auth.state -eq 'failed') {
                            Log ('Authentication failed: {0}' -f $auth.reason)
                            if ($auth.data -and ($auth.data.code -eq 535 -or [string]$auth.data.text -match '\b5\.7\.(3|8|139)\b')) { $passwordRejected = $true }
                        }
                    } else {
                        Publish 'auth' (New-StSection -State 'skipped' -Reason 'No credentials were entered. If the provider uses IP authentication that is correct - the Egress section is the check that matters instead.')
                    }
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)

                    if (Stopping) { break }
                    if ($passwordRejected) {
                        # Lockout guard: one rejected password per run.
                        Publish 'envelope' (New-StSection -State 'skipped' -Reason 'Not attempted: the password was just rejected, and it is not tried a second time so an Active Directory account is not pushed towards lockout. Fix the password and run the sweep again.')
                        Log 'Envelope skipped - the password is not retried after a rejection.'
                    } else {
                        $Shared.Status = 'Offering the envelope (nothing is delivered)'
                        $envSec = Test-StEnvelope -Server $server -Port $port -Security $security -HeloName $helo `
                                                  -Username $user -Password $pass -From $from -To $to -Sink $sink -IoTimeoutMs $ioMs -Kind $kind
                        Publish 'envelope' $envSec
                        if ($envSec.data -and $envSec.data.accepted) { Log 'The server accepted the sender and the recipient. No message was delivered.' }
                    }
                    $n++; $Shared.Progress = [int](($n / $steps) * 100)

                    if (Stopping) { break }
                    if ($extTo) {
                        $tcpSec = $Shared.Sections['tcp']
                        $open = $true
                        if ($tcpSec -and $tcpSec.data -and $tcpSec.data.ports) {
                            $hit = @($tcpSec.data.ports | Where-Object { $_.port -eq $relayPort })
                            if ($hit.Count -gt 0) { $open = [bool]$hit[0].open }
                        }
                        Invoke-Relay -PortOpen $open
                    } elseif ($kind -eq 'exo-connector' -or $kind -eq 'exchange-onprem') {
                        Publish 'relay' (New-StSection -State 'skipped' -Reason 'No external recipient was entered on the Relay tab, so the relay connector was not tested.')
                    }
                }

                if (Stopping) { break }
                $Shared.Status = 'Reading .NET and TLS settings'
                $crypto = Test-StDotNetCrypto
                Publish 'crypto' $crypto

                $intercepted = $false
                $issuer = ''
                $certSec = $Shared.Sections['certificate']
                if ($certSec -and $certSec.data) {
                    $intercepted = [bool]$certSec.data.interceptionSuspected
                    if ($certSec.data.certificate) { $issuer = [string]$certSec.data.certificate.issuerCn }
                }
                Publish 'interference' (Test-StLocalInterference -CertInterceptionSuspected $intercepted -InterceptingIssuer $issuer)

                if (Stopping) { break }
                Publish 'appConfig' (Test-StAppConfigScan -Server $server -BudgetSec ([int]$Opt.ConfigBudgetSec) -Shared $Shared)
                Publish 'senderDns' (Test-StSenderDns -FromAddress $from)
                $n++; $Shared.Progress = 100
            }

            'matrix' {
                $Shared.Status = 'Running the settings matrix'
                $mx = Invoke-StAppMatrix -Server $server -Ports $ports -Username $user -Password $pass `
                                         -From $from -To $to -HeloName $helo `
                                         -TimeoutMs ([int]$Opt.MatrixTimeoutMs) -LiveSend ([bool]$Opt.MatrixLive) `
                                         -Sink $sink -Shared $Shared -Kind $kind
                Publish 'appMatrix' $mx
                if ($mx.data) {
                    Log ('{0} of {1} port and SSL combinations work, {2} of them encrypted.' -f $mx.data.workingCount, @($mx.data.cells).Count, $mx.data.secureWorkingCount)
                    if ($mx.data.lockoutGuardTripped) { Log 'The password was rejected once and not tried again in this run.' }
                }
                $Shared.Progress = 100
            }

            'reproduce' {
                $Shared.Status = 'Reproducing the application''s send'
                $sec = Test-StDotNetSmtpClient -Server $server -Port $port -EnableSsl ([bool]$Opt.EnableSsl) `
                                               -Username $user -Password $pass -From $from -To $to `
                                               -TimeoutMs ([int]$Opt.ReproTimeoutMs) -Sink $sink -Kind $kind
                Publish 'dotNetSmtpClient' $sec
                if ($sec.state -eq 'measured') { Log 'System.Net.Mail.SmtpClient sent successfully with these settings.' }
                else { Log ('System.Net.Mail.SmtpClient failed: {0}' -f $sec.reason) }
                $Shared.Progress = 100
            }

            'send' {
                $Shared.Status = 'Sending a real message'
                $sec = Test-StRawSend -Server $server -Port $port -Security $security -HeloName $helo `
                                      -Username $user -Password $pass -From $from -To $to -Sink $sink -IoTimeoutMs ([math]::Max($ioMs, 30000)) -Kind $kind
                Publish 'send' $sec
                if ($sec.state -eq 'measured') { Log ('Message accepted by the server. Message-ID {0}' -f $sec.data.messageId) }
                else { Log ('Send failed: {0}' -f $sec.reason) }
                $Shared.Progress = 100
            }

            'relay' {
                $Shared.Status = ('Checking port {0}' -f $relayPort)
                $tp = Test-StTcpPort -Server $server -Port $relayPort -TimeoutMs 4000
                if (-not $Shared.Sections['publicIp']) {
                    $Shared.Status = 'Checking egress addresses'
                    Publish 'publicIp' (Test-StPublicIp)
                }
                Invoke-Relay -PortOpen ([bool]$tp.open)
                $Shared.Progress = 100
            }

            'autodiscover' {
                $ad = Test-StAutodiscover -Email $from -TimeoutMs ([math]::Min($ioMs, 10000)) -Shared $Shared
                Publish 'autodiscover' $ad
                if ($ad.data -and $ad.data.answered) { Log ('Autodiscover answers for {0} via {1}{2}.' -f $from, $ad.data.answeredBy, $(if ($ad.data.mailboxLocation) { ' - mailbox ' + $ad.data.mailboxLocation } else { '' })) }
                else { Log ('Nothing answers Autodiscover for {0}.' -f $from) }
                $Shared.Progress = 100
            }

            'environment' {
                $Shared.Status = 'Reading .NET and TLS settings'
                Publish 'crypto' (Test-StDotNetCrypto)
                Publish 'interference' (Test-StLocalInterference)
                $Shared.Status = 'Scanning application configuration files'
                Publish 'appConfig' (Test-StAppConfigScan -Server $server -BudgetSec ([int]$Opt.ConfigBudgetSec) -Shared $Shared)
                $Shared.Progress = 100
            }
        }

        # The verdict is recomputed from everything gathered so far, so a
        # matrix run after a sweep sharpens the recommendation rather than
        # replacing it with a narrower one.
        $Shared.Status = 'Working out the verdict'
        $snapshot = @{}
        foreach ($k in @($Shared.Sections.Keys)) { $snapshot[$k] = $Shared.Sections[$k] }
        $Shared.Verdict = Get-StVerdict -Sections $snapshot -Server $server -AppProfileName ([string]$Opt.AppProfile) -Kind $kind
        $Shared.Version = $Shared.Version + 1

    } catch {
        $Shared.Error = ('{0}{1}{2}' -f $_.Exception.Message, [Environment]::NewLine, $_.ScriptStackTrace)
        [void]$Shared.Log.Add(('The test run stopped with an unexpected error: {0}' -f $_.Exception.Message))
    } finally {
        $Shared.Status = 'Finished'
        $Shared.Done = $true
    }
}

# ---------------------------------------------------------------------------
#  Form
# ---------------------------------------------------------------------------
$script:tips = New-Object System.Windows.Forms.ToolTip
$script:tips.AutoPopDelay = 20000
$script:tips.InitialDelay = 400
$script:tips.ReshowDelay  = 200

$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = ('SMTP Toolkit  v{0}   -   mail sending diagnostics' -f (Get-StVersion))
$script:form.Size = New-Object System.Drawing.Size(1240, 900)
$script:form.MinimumSize = New-Object System.Drawing.Size(1040, 740)
$script:form.StartPosition = 'CenterScreen'
$script:form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'; $root.ColumnCount = 1; $root.RowCount = 3
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 222)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$script:form.Controls.Add($root)

# ---- connection panel -----------------------------------------------------
$conn = New-Object System.Windows.Forms.GroupBox
$conn.Text = ' Connection '
$conn.Dock = 'Fill'
$conn.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$root.Controls.Add($conn, 0, 0)

$cp = New-Object System.Windows.Forms.Panel
$cp.Dock = 'Fill'
$cp.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
$conn.Controls.Add($cp)

$colA = 12;  $colAv = 92
$colB = 360; $colBv = 430
$colC = 700; $colCv = 790
$rowH = 27
$y = 8

# Preset
$lblPreset = New-StLabel -Text 'Preset'
$lblPreset.Location = New-Object System.Drawing.Point($colA, ($y + 4))
$cp.Controls.Add($lblPreset)
$script:cboPreset = New-Object System.Windows.Forms.ComboBox
$script:cboPreset.Location = New-Object System.Drawing.Point($colAv, $y)
$script:cboPreset.Width = 236
$script:cboPreset.DropDownStyle = 'DropDownList'
foreach ($p in $script:Presets) { [void]$script:cboPreset.Items.Add($p.name) }
$script:cboPreset.SelectedIndex = 0
$cp.Controls.Add($script:cboPreset)

# Server
$lblServer = New-StLabel -Text 'Server'
$lblServer.Location = New-Object System.Drawing.Point($colB, ($y + 4))
$cp.Controls.Add($lblServer)
$script:txtServer = New-StTextBox -W 236 -Text 'mail-au.smtp2go.com'
$script:txtServer.Location = New-Object System.Drawing.Point($colBv, $y)
$cp.Controls.Add($script:txtServer)

# Port
$lblPort = New-StLabel -Text 'Port'
$lblPort.Location = New-Object System.Drawing.Point($colC, ($y + 4))
$cp.Controls.Add($lblPort)
$script:cboPort = New-Object System.Windows.Forms.ComboBox
$script:cboPort.Location = New-Object System.Drawing.Point($colCv, $y)
$script:cboPort.Width = 90
$script:cboPort.DropDownStyle = 'DropDown'
$cp.Controls.Add($script:cboPort)
$script:tips.SetToolTip($script:cboPort, 'The port the application is configured to use. Type any port if it is not in the list.')

$lblSec = New-StLabel -Text 'Security'
$lblSec.Location = New-Object System.Drawing.Point(($colC + 190), ($y + 4))
$cp.Controls.Add($lblSec)
$script:cboSecurity = New-Object System.Windows.Forms.ComboBox
$script:cboSecurity.Location = New-Object System.Drawing.Point(($colC + 250), $y)
$script:cboSecurity.Width = 130
$script:cboSecurity.DropDownStyle = 'DropDownList'
foreach ($s in @('Auto (by port)','STARTTLS','Implicit SSL','None')) { [void]$script:cboSecurity.Items.Add($s) }
$script:cboSecurity.SelectedIndex = 0
$cp.Controls.Add($script:cboSecurity)
$script:tips.SetToolTip($script:cboSecurity, 'Auto treats 465, 8465 and 443 as implicit SSL and everything else as STARTTLS.')

$y += $rowH

# From / To
$lblFrom = New-StLabel -Text 'From'
$lblFrom.Location = New-Object System.Drawing.Point($colA, ($y + 4))
$cp.Controls.Add($lblFrom)
$script:txtFrom = New-StTextBox -W 236
$script:txtFrom.Location = New-Object System.Drawing.Point($colAv, $y)
$cp.Controls.Add($script:txtFrom)
$script:tips.SetToolTip($script:txtFrom, 'The sending address the application uses. Its domain must be verified with the provider.')

$lblTo = New-StLabel -Text 'To'
$lblTo.Location = New-Object System.Drawing.Point($colB, ($y + 4))
$cp.Controls.Add($lblTo)
$script:txtTo = New-StTextBox -W 236
$script:txtTo.Location = New-Object System.Drawing.Point($colBv, $y)
$cp.Controls.Add($script:txtTo)
$script:tips.SetToolTip($script:txtTo, 'Where a test message would go. Sweeps stop before delivery, so nothing arrives unless you press Send test message.')

$lblHelo = New-StLabel -Text 'EHLO name'
$lblHelo.Location = New-Object System.Drawing.Point($colC, ($y + 4))
$cp.Controls.Add($lblHelo)
$script:txtHelo = New-StTextBox -W 150 -Text $env:COMPUTERNAME
$script:txtHelo.Location = New-Object System.Drawing.Point($colCv, $y)
$cp.Controls.Add($script:txtHelo)

$y += $rowH

# Credentials
$lblUser = New-StLabel -Text 'Username'
$lblUser.Location = New-Object System.Drawing.Point($colA, ($y + 4))
$cp.Controls.Add($lblUser)
$script:txtUser = New-StTextBox -W 236
$script:txtUser.Location = New-Object System.Drawing.Point($colAv, $y)
$cp.Controls.Add($script:txtUser)

$lblPass = New-StLabel -Text 'Password'
$lblPass.Location = New-Object System.Drawing.Point($colB, ($y + 4))
$cp.Controls.Add($lblPass)
$script:txtPass = New-StTextBox -W 236 -Password $true
$script:txtPass.Location = New-Object System.Drawing.Point($colBv, $y)
$cp.Controls.Add($script:txtPass)

$script:chkShowPass = New-Object System.Windows.Forms.CheckBox
$script:chkShowPass.Text = 'Show'
$script:chkShowPass.AutoSize = $true
$script:chkShowPass.Location = New-Object System.Drawing.Point(($colBv + 244), ($y + 3))
$cp.Controls.Add($script:chkShowPass)

# Clear of the Show box, which is AutoSize and grows with the DPI. Sitting this
# at $colC put the two checkboxes on top of each other at 100% scaling.
$script:chkNoCreds = New-Object System.Windows.Forms.CheckBox
$script:chkNoCreds.Text = 'No credentials (IP authentication)'
$script:chkNoCreds.AutoSize = $true
$script:chkNoCreds.Location = New-Object System.Drawing.Point(($colCv - 30), ($y + 3))
$cp.Controls.Add($script:chkNoCreds)
$script:tips.SetToolTip($script:chkNoCreds, 'Test the path a provider''s IP authentication takes. The Egress section then tells you which address to authorise.')

$y += $rowH + 4

# Timeouts
$lblTo2 = New-StLabel -Text 'Read timeout'
$lblTo2.Location = New-Object System.Drawing.Point($colA, ($y + 4))
$cp.Controls.Add($lblTo2)
$script:numTimeout = New-Object System.Windows.Forms.NumericUpDown
$script:numTimeout.Location = New-Object System.Drawing.Point($colAv, $y)
$script:numTimeout.Width = 70
$script:numTimeout.Minimum = 3
$script:numTimeout.Maximum = 300
$script:numTimeout.Value = 15
$cp.Controls.Add($script:numTimeout)
$lblTo2b = New-StLabel -Text 'seconds'
$lblTo2b.Location = New-Object System.Drawing.Point(($colAv + 76), ($y + 4))
$cp.Controls.Add($lblTo2b)

$lblCfg = New-StLabel -Text 'Config scan budget'
$lblCfg.Location = New-Object System.Drawing.Point($colB, ($y + 4))
$cp.Controls.Add($lblCfg)
$script:numConfig = New-Object System.Windows.Forms.NumericUpDown
$script:numConfig.Location = New-Object System.Drawing.Point(($colB + 116), $y)
$script:numConfig.Width = 70
$script:numConfig.Minimum = 5
$script:numConfig.Maximum = 600
$script:numConfig.Value = 25
$cp.Controls.Add($script:numConfig)
$lblCfgb = New-StLabel -Text 'seconds'
$lblCfgb.Location = New-Object System.Drawing.Point(($colB + 192), ($y + 4))
$cp.Controls.Add($lblCfgb)
$script:tips.SetToolTip($script:numConfig, 'How long the scan of application .config files may run before it stops and says its coverage was partial.')

$lblPorts = New-StLabel -Text 'Ports'
$lblPorts.Location = New-Object System.Drawing.Point($colC, ($y + 4))
$cp.Controls.Add($lblPorts)
$script:txtPorts = New-StTextBox -W 260
$script:txtPorts.Location = New-Object System.Drawing.Point($colCv, $y)
$cp.Controls.Add($script:txtPorts)
$script:tips.SetToolTip($script:txtPorts, 'The ports the sweep and the settings matrix cover. Comma separated.')

$y += $rowH + 6

# Actions
$script:btnSweep = New-StButton -Text 'Run full sweep' -W 150 -Tip 'DNS, every port, TLS, certificate, credentials and the envelope. Nothing is delivered.'
$script:btnSweep.Location = New-Object System.Drawing.Point($colA, $y)
$script:btnSweep.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$cp.Controls.Add($script:btnSweep)

$script:btnMatrix = New-StButton -Text 'Settings matrix' -W 140 -Tip 'Every port against both states of the application''s SSL checkbox. This is the evidence a vendor cannot argue with.'
$script:btnMatrix.Location = New-Object System.Drawing.Point(($colA + 158), $y)
$cp.Controls.Add($script:btnMatrix)

$script:btnRepro = New-StButton -Text 'Reproduce app send' -W 150 -Tip 'One real System.Net.Mail.SmtpClient send with exactly the settings above. This DELIVERS a message when it succeeds.'
$script:btnRepro.Location = New-Object System.Drawing.Point(($colA + 306), $y)
$cp.Controls.Add($script:btnRepro)

$script:btnSend = New-StButton -Text 'Send test message' -W 140 -Tip 'A real send through this tool''s own client, which can use implicit SSL. DELIVERS a message.'
$script:btnSend.Location = New-Object System.Drawing.Point(($colA + 464), $y)
$cp.Controls.Add($script:btnSend)

$script:btnEnv = New-StButton -Text 'Environment only' -W 130 -Tip 'Registry crypto settings, security software and application config files. No network traffic.'
$script:btnEnv.Location = New-Object System.Drawing.Point(($colA + 612), $y)
$cp.Controls.Add($script:btnEnv)

$script:btnCancel = New-StButton -Text 'Cancel' -W 90
$script:btnCancel.Location = New-Object System.Drawing.Point(($colA + 750), $y)
$script:btnCancel.Enabled = $false
$cp.Controls.Add($script:btnCancel)

$script:btnSaveJson = New-StButton -Text 'Save report' -W 110 -Tip 'Writes the JSON report. Credentials are never included.'
$script:btnSaveJson.Location = New-Object System.Drawing.Point(($colA + 848), $y)
$cp.Controls.Add($script:btnSaveJson)

$y += $rowH + 8
$script:lblStatus = New-StLabel -Text 'Ready.' -Bold $true
$script:lblStatus.Location = New-Object System.Drawing.Point($colA, ($y + 2))
$cp.Controls.Add($script:lblStatus)

$script:progress = New-Object System.Windows.Forms.ProgressBar
$script:progress.Location = New-Object System.Drawing.Point(($colC + 90), $y)
$script:progress.Width = 290
$script:progress.Height = 18
$cp.Controls.Add($script:progress)

# ---- tabs -----------------------------------------------------------------
$script:tabs = New-Object System.Windows.Forms.TabControl
$script:tabs.Dock = 'Fill'
$script:tabs.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$root.Controls.Add($script:tabs, 0, 1)

# Verdict tab
$tabVerdict = New-Object System.Windows.Forms.TabPage
$tabVerdict.Text = 'Verdict'
$tabVerdict.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabVerdict)

$vLayout = New-Object System.Windows.Forms.TableLayoutPanel
$vLayout.Dock = 'Fill'; $vLayout.ColumnCount = 1; $vLayout.RowCount = 3
[void]$vLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 96)))
[void]$vLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$vLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 236)))
$tabVerdict.Controls.Add($vLayout)

$script:panelHeadline = New-Object System.Windows.Forms.Panel
$script:panelHeadline.Dock = 'Fill'
$script:panelHeadline.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$vLayout.Controls.Add($script:panelHeadline, 0, 0)

$script:lblOverall = New-StLabel -Text 'Not tested yet' -Bold $true -Size 14
$script:lblOverall.Location = New-Object System.Drawing.Point(14, 12)
$script:panelHeadline.Controls.Add($script:lblOverall)

$script:lblHeadline = New-Object System.Windows.Forms.Label
$script:lblHeadline.Location = New-Object System.Drawing.Point(14, 44)
$script:lblHeadline.Size = New-Object System.Drawing.Size(1150, 44)
$script:lblHeadline.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$script:lblHeadline.Text = 'Fill in the connection details above and press Run full sweep.'
$script:panelHeadline.Controls.Add($script:lblHeadline)

$script:gridVerdict = New-Object System.Windows.Forms.DataGridView
$script:gridVerdict.Dock = 'Fill'
$script:gridVerdict.ReadOnly = $true
$script:gridVerdict.AllowUserToAddRows = $false
$script:gridVerdict.RowHeadersVisible = $false
$script:gridVerdict.SelectionMode = 'FullRowSelect'
$script:gridVerdict.AutoSizeColumnsMode = 'Fill'
$script:gridVerdict.BorderStyle = 'None'
$script:gridVerdict.BackgroundColor = [System.Drawing.Color]::White
$script:gridVerdict.EnableHeadersVisualStyles = $false
$script:gridVerdict.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$script:gridVerdict.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:gridVerdict.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(214, 228, 245)
$script:gridVerdict.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
foreach ($c in @(@{n='area';h='Area';w=16},@{n='state';h='State';w=10},@{n='summary';h='What was found';w=38},@{n='remedy';h='What to do';w=36})) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.n; $col.HeaderText = $c.h; $col.FillWeight = $c.w
    [void]$script:gridVerdict.Columns.Add($col)
}
$vLayout.Controls.Add($script:gridVerdict, 0, 1)

$script:txtRecommend = New-Object System.Windows.Forms.TextBox
$script:txtRecommend.Multiline = $true
$script:txtRecommend.ReadOnly = $true
$script:txtRecommend.Dock = 'Fill'
$script:txtRecommend.ScrollBars = 'Vertical'
$script:txtRecommend.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$script:txtRecommend.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 240)
$vLayout.Controls.Add($script:txtRecommend, 0, 2)

# Results tab
$tabResults = New-Object System.Windows.Forms.TabPage
$tabResults.Text = 'Results'
$tabResults.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabResults)
$script:gridResults = New-StResultGrid
$tabResults.Controls.Add($script:gridResults)

# Transcript tab
$tabTrans = New-Object System.Windows.Forms.TabPage
$tabTrans.Text = 'SMTP transcript'
$tabTrans.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabTrans)
$script:txtTranscript = New-StMonoBox
$tabTrans.Controls.Add($script:txtTranscript)

# App simulation tab
$tabApp = New-Object System.Windows.Forms.TabPage
$tabApp.Text = 'Application simulation'
$tabApp.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabApp)

$appLayout = New-Object System.Windows.Forms.TableLayoutPanel
$appLayout.Dock = 'Fill'; $appLayout.ColumnCount = 1; $appLayout.RowCount = 2
[void]$appLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 240)))
[void]$appLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabApp.Controls.Add($appLayout)

$appTop = New-Object System.Windows.Forms.Panel
$appTop.Dock = 'Fill'
$appLayout.Controls.Add($appTop, 0, 0)

$lblAppIntro = New-Object System.Windows.Forms.Label
$lblAppIntro.Location = New-Object System.Drawing.Point(12, 10)
$lblAppIntro.Size = New-Object System.Drawing.Size(1160, 60)
$lblAppIntro.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblAppIntro.Text = @'
These are the fields a typical Windows practice-management or accounting application exposes, and they map straight onto System.Net.Mail.SmtpClient: the SSL box sets EnableSsl, which only ever means STARTTLS. The matrix tries every port against both states of that box.
Probe mode replays the same protocol exchange but stops before the message body, so a full sweep of every combination delivers nothing. Tick Live send only when you want real messages in the mailbox.
'@
$appTop.Controls.Add($lblAppIntro)

$ay = 74
$lblAppProfile = New-StLabel -Text 'Target application'
$lblAppProfile.Location = New-Object System.Drawing.Point(12, ($ay + 4))
$appTop.Controls.Add($lblAppProfile)
$script:cboAppProfile = New-Object System.Windows.Forms.ComboBox
$script:cboAppProfile.Location = New-Object System.Drawing.Point(170, $ay)
$script:cboAppProfile.Width = 300
$script:cboAppProfile.DropDownStyle = 'DropDownList'
foreach ($ap in (Get-StAppProfiles)) { [void]$script:cboAppProfile.Items.Add($ap.name) }
$script:cboAppProfile.SelectedIndex = 0
$appTop.Controls.Add($script:cboAppProfile)
$script:tips.SetToolTip($script:cboAppProfile, 'Decides which ports are recommended, and which post-configuration steps appear on the Verdict tab.')

$lblAppProfileNote = New-StLabel -Text 'adds that application''s own gotchas to the verdict'
$lblAppProfileNote.Location = New-Object System.Drawing.Point(480, ($ay + 4))
$lblAppProfileNote.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$appTop.Controls.Add($lblAppProfileNote)

$ay += 30
$lblMxTo = New-StLabel -Text 'Per-combination timeout'
$lblMxTo.Location = New-Object System.Drawing.Point(12, ($ay + 4))
$appTop.Controls.Add($lblMxTo)
$script:numMatrixTimeout = New-Object System.Windows.Forms.NumericUpDown
$script:numMatrixTimeout.Location = New-Object System.Drawing.Point(170, $ay)
$script:numMatrixTimeout.Width = 70
$script:numMatrixTimeout.Minimum = 3
$script:numMatrixTimeout.Maximum = 300
$script:numMatrixTimeout.Value = 15
$appTop.Controls.Add($script:numMatrixTimeout)
$lblMxTo2 = New-StLabel -Text 'seconds  (an implicit-SSL port burns this whole budget before failing)'
$lblMxTo2.Location = New-Object System.Drawing.Point(246, ($ay + 4))
$appTop.Controls.Add($lblMxTo2)

$ay += 30
$script:chkMatrixLive = New-Object System.Windows.Forms.CheckBox
$script:chkMatrixLive.Text = 'Live send - use the real System.Net.Mail.SmtpClient and DELIVER a message for every combination that works'
$script:chkMatrixLive.AutoSize = $true
$script:chkMatrixLive.Location = New-Object System.Drawing.Point(12, $ay)
$appTop.Controls.Add($script:chkMatrixLive)

$ay += 28
$script:chkReproSsl = New-Object System.Windows.Forms.CheckBox
$script:chkReproSsl.Text = 'Use SSL  (the application''s checkbox)'
$script:chkReproSsl.AutoSize = $true
$script:chkReproSsl.Checked = $true
$script:chkReproSsl.Location = New-Object System.Drawing.Point(12, $ay)
$appTop.Controls.Add($script:chkReproSsl)

$lblRepTo = New-StLabel -Text 'Reproduction timeout'
$lblRepTo.Location = New-Object System.Drawing.Point(300, ($ay + 1))
$appTop.Controls.Add($lblRepTo)
$script:numReproTimeout = New-Object System.Windows.Forms.NumericUpDown
$script:numReproTimeout.Location = New-Object System.Drawing.Point(440, ($ay - 2))
$script:numReproTimeout.Width = 70
$script:numReproTimeout.Minimum = 3
$script:numReproTimeout.Maximum = 600
$script:numReproTimeout.Value = 20
$appTop.Controls.Add($script:numReproTimeout)
$script:chkReproDefault = New-Object System.Windows.Forms.CheckBox
$script:chkReproDefault.Text = 'Use the .NET default of 100 seconds (reproduces the user''s wait exactly)'
$script:chkReproDefault.AutoSize = $true
$script:chkReproDefault.Location = New-Object System.Drawing.Point(520, $ay)
$appTop.Controls.Add($script:chkReproDefault)

$ay += 30
$lblAppNote = New-Object System.Windows.Forms.Label
$lblAppNote.Location = New-Object System.Drawing.Point(12, $ay)
$lblAppNote.Size = New-Object System.Drawing.Size(1160, 34)
$lblAppNote.ForeColor = [System.Drawing.Color]::FromArgb(120, 80, 0)
$lblAppNote.Text = 'Reproduce app send and Send test message both deliver real mail when they succeed. Everything else stops before the message body.'
$appTop.Controls.Add($lblAppNote)

$script:gridMatrix = New-StResultGrid
$appLayout.Controls.Add($script:gridMatrix, 0, 1)

# Relay & connectors tab
$tabRelay = New-Object System.Windows.Forms.TabPage
$tabRelay.Text = 'Relay & Autodiscover'
$tabRelay.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabRelay)

$relLayout = New-Object System.Windows.Forms.TableLayoutPanel
$relLayout.Dock = 'Fill'; $relLayout.ColumnCount = 1; $relLayout.RowCount = 2
[void]$relLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
[void]$relLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabRelay.Controls.Add($relLayout)

$relTop = New-Object System.Windows.Forms.Panel
$relTop.Dock = 'Fill'
$relLayout.Controls.Add($relTop, 0, 0)

$lblRelIntro = New-Object System.Windows.Forms.Label
$lblRelIntro.Location = New-Object System.Drawing.Point(12, 10)
$lblRelIntro.Size = New-Object System.Drawing.Size(1160, 48)
$lblRelIntro.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblRelIntro.Text = @'
Tests sending with NO credentials to an outside recipient - the route a scanner or application takes through an Exchange Online inbound connector (port 25 to the tenant MX, authorised by this site's public IP) or an on-prem Exchange anonymous relay connector. The recipient must be external: internal recipients are accepted with or without a connector. The session is reset before DATA, so nothing is delivered.
'@
$relTop.Controls.Add($lblRelIntro)

$lblExt = New-StLabel -Text 'External recipient'
$lblExt.Location = New-Object System.Drawing.Point(12, 70)
$relTop.Controls.Add($lblExt)
$script:txtExternalTo = New-StTextBox -W 260
$script:txtExternalTo.Location = New-Object System.Drawing.Point(150, 66)
$relTop.Controls.Add($script:txtExternalTo)
$script:tips.SetToolTip($script:txtExternalTo, 'An address outside the organisation, for example your own. Nothing is delivered.')

$lblConnIp = New-StLabel -Text 'Connector IP(s)'
$lblConnIp.Location = New-Object System.Drawing.Point(430, 70)
$relTop.Controls.Add($lblConnIp)
$script:txtConnectorIps = New-StTextBox -W 240
$script:txtConnectorIps.Location = New-Object System.Drawing.Point(540, 66)
$relTop.Controls.Add($script:txtConnectorIps)
$script:tips.SetToolTip($script:txtConnectorIps, 'Optional. The IP address(es) configured on the connector, comma separated. The test compares them with the address the server actually sees.')

$script:btnRelay = New-StButton -Text 'Relay test' -W 120 -Tip 'Anonymous MAIL FROM / RCPT TO to the external recipient. Nothing is delivered.'
$script:btnRelay.Location = New-Object System.Drawing.Point(800, 63)
$relTop.Controls.Add($script:btnRelay)

$script:btnAutodiscover = New-StButton -Text 'Autodiscover check' -W 150 -Tip 'Does anything answer Autodiscover for the From address, and where is the mailbox? Unauthenticated - no password is sent.'
$script:btnAutodiscover.Location = New-Object System.Drawing.Point(930, 63)
$relTop.Controls.Add($script:btnAutodiscover)

$lblRelNote = New-Object System.Windows.Forms.Label
$lblRelNote.Location = New-Object System.Drawing.Point(12, 104)
$lblRelNote.Size = New-Object System.Drawing.Size(1160, 36)
$lblRelNote.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$lblRelNote.Text = 'For the Exchange Online connector, pick "Microsoft 365 - inbound connector (25)" and leave Server blank: the From domain''s MX is looked up through a public resolver. Run full sweep also runs this test whenever an external recipient is filled in.'
$relTop.Controls.Add($lblRelNote)

$script:gridRelay = New-StResultGrid
$relLayout.Controls.Add($script:gridRelay, 0, 1)

# Environment tab
$tabEnvT = New-Object System.Windows.Forms.TabPage
$tabEnvT.Text = 'Environment'
$tabEnvT.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabEnvT)
$script:gridEnv = New-StResultGrid
$tabEnvT.Controls.Add($script:gridEnv)

# JSON tab
$tabJson = New-Object System.Windows.Forms.TabPage
$tabJson.Text = 'JSON report'
$tabJson.BackColor = [System.Drawing.Color]::White
$script:tabs.TabPages.Add($tabJson)

$jLayout = New-Object System.Windows.Forms.TableLayoutPanel
$jLayout.Dock = 'Fill'; $jLayout.ColumnCount = 1; $jLayout.RowCount = 2
[void]$jLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
[void]$jLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabJson.Controls.Add($jLayout)

$jTop = New-Object System.Windows.Forms.Panel
$jTop.Dock = 'Fill'
$jLayout.Controls.Add($jTop, 0, 0)

$script:btnJsonRefresh = New-StButton -Text 'Rebuild' -W 90
$script:btnJsonRefresh.Location = New-Object System.Drawing.Point(10, 7)
$jTop.Controls.Add($script:btnJsonRefresh)
$script:btnJsonCopy = New-StButton -Text 'Copy to clipboard' -W 140
$script:btnJsonCopy.Location = New-Object System.Drawing.Point(108, 7)
$jTop.Controls.Add($script:btnJsonCopy)
$script:chkJsonCompress = New-Object System.Windows.Forms.CheckBox
$script:chkJsonCompress.Text = 'Compact'
$script:chkJsonCompress.AutoSize = $true
$script:chkJsonCompress.Location = New-Object System.Drawing.Point(260, 13)
$jTop.Controls.Add($script:chkJsonCompress)
$lblJsonNote = New-StLabel -Text 'Schema smtp-diagnostic v1. Passwords are never included. Markers around the JSON let a report be recovered from a pasted log.'
$lblJsonNote.Location = New-Object System.Drawing.Point(340, 13)
$jTop.Controls.Add($lblJsonNote)

$script:txtJson = New-StMonoBox
$jLayout.Controls.Add($script:txtJson, 0, 1)

# ---- activity log ---------------------------------------------------------
$logBox = New-Object System.Windows.Forms.GroupBox
$logBox.Text = ' Activity '
$logBox.Dock = 'Fill'
$logBox.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$root.Controls.Add($logBox, 0, 2)

$script:txtLog = New-Object System.Windows.Forms.TextBox
$script:txtLog.Multiline = $true
$script:txtLog.ReadOnly = $true
$script:txtLog.Dock = 'Fill'
$script:txtLog.ScrollBars = 'Vertical'
$script:txtLog.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
$script:txtLog.BackColor = [System.Drawing.Color]::White
$logBox.Controls.Add($script:txtLog)

# ---------------------------------------------------------------------------
#  UI behaviour
# ---------------------------------------------------------------------------
function Write-StLog {
    param([string]$Message)
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $script:txtLog.AppendText(('[{0}] {1}{2}' -f $stamp, $Message, [Environment]::NewLine))
}

function Set-StBusy {
    param([bool]$On)
    $script:busy = $On
    foreach ($b in @($script:btnSweep, $script:btnMatrix, $script:btnRepro, $script:btnSend, $script:btnEnv, $script:btnRelay, $script:btnAutodiscover)) {
        $b.Enabled = (-not $On)
    }
    $script:btnCancel.Enabled = $On
    if ($On) { $script:form.Cursor = [System.Windows.Forms.Cursors]::AppStarting }
    else     { $script:form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Get-StSelectedPreset {
    $i = $script:cboPreset.SelectedIndex
    if ($i -lt 0 -or $i -ge $script:Presets.Count) { return $script:Presets[0] }
    return $script:Presets[$i]
}

function Set-StCueBanner {
    # Grey placeholder text inside an empty TextBox (EM_SETCUEBANNER). Purely
    # cosmetic, so any failure is ignored.
    param($TextBox, [string]$Text)
    try {
        if (-not ('StNative.Cue' -as [type])) {
            Add-Type -Namespace 'StNative' -Name 'Cue' -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern System.IntPtr SendMessage(System.IntPtr hWnd, int msg, System.IntPtr wParam, string lParam);
'@
        }
        [void][StNative.Cue]::SendMessage($TextBox.Handle, 0x1501, [System.IntPtr]1, $Text)
    } catch { }
}

function Update-StPreset {
    $p = Get-StSelectedPreset
    if ($p.server) { $script:txtServer.Text = $p.server }
    elseif ($p.kind -ne 'generic') {
        # A preset with no fixed host (on-prem Exchange, the connector route)
        # must not leave the previous preset's host sitting in the box.
        $script:txtServer.Text = ''
    }
    Set-StCueBanner -TextBox $script:txtServer -Text $p.serverHint
    $script:tips.SetToolTip($script:txtServer, $(if ($p.serverHint) { $p.serverHint } else { 'The mail server host name.' }))
    $ports = @(Get-StAllPorts -Preset $p)
    $script:txtPorts.Text = ($ports -join ', ')
    $current = $script:cboPort.Text
    $script:cboPort.Items.Clear()
    foreach ($pt in $ports) { [void]$script:cboPort.Items.Add([string]$pt) }
    if ($current -and ($ports -contains ([int]($current -as [int])))) { $script:cboPort.Text = $current }
    else { $script:cboPort.Text = [string]$p.preferred }
}

function ConvertTo-StPortList {
    param([string]$Text)
    $out = New-Object System.Collections.Generic.List[int]
    foreach ($piece in ($Text -split '[,;\s]+')) {
        if (-not $piece) { continue }
        $n = 0
        if ([int]::TryParse($piece.Trim(), [ref]$n) -and $n -gt 0 -and $n -le 65535) {
            if (-not $out.Contains($n)) { [void]$out.Add($n) }
        }
    }
    # NO leading comma. Comma-wrapping stops the array unrolling on return,
    # which is right when the caller assigns it directly - but every caller
    # here wraps in @(), and @() over a comma-wrapped return produces a
    # one-element array CONTAINING the array. Passing that to an [int[]]
    # parameter fails with "Cannot convert System.Object[] to System.Int32".
    # Returning it plain and letting the callers wrap is the shape that works
    # both ways.
    return @($out.ToArray())
}

function Get-StSecurityMode {
    switch ($script:cboSecurity.SelectedIndex) {
        1 { return 'starttls' }
        2 { return 'implicit' }
        3 { return 'none' }
        default { return 'auto' }
    }
}

function Get-StUiOptions {
    param([string]$Job)

    $preset = Get-StSelectedPreset
    $server = $script:txtServer.Text.Trim()
    if (-not $server -and $preset.kind -eq 'exo-connector') {
        # Resolved from the From domain's MX by the worker.
        if (-not (Test-StEmailAddress $script:txtFrom.Text.Trim())) {
            [void][System.Windows.Forms.MessageBox]::Show('Enter the From address so its domain''s MX can be looked up, or type the tenant host name (for example contoso-com.mail.protection.outlook.com) into Server.', 'SMTP Toolkit', 'OK', 'Warning')
            return $null
        }
    } elseif (-not $server) {
        [void][System.Windows.Forms.MessageBox]::Show('Enter the mail server name.', 'SMTP Toolkit', 'OK', 'Warning')
        return $null
    }
    $portText = $script:cboPort.Text.Trim()
    $port = 0
    if (-not [int]::TryParse($portText, [ref]$port) -or $port -le 0 -or $port -gt 65535) {
        [void][System.Windows.Forms.MessageBox]::Show('The port is not a number between 1 and 65535.', 'SMTP Toolkit', 'OK', 'Warning')
        return $null
    }
    $ports = @(ConvertTo-StPortList -Text $script:txtPorts.Text)
    if ($ports.Count -eq 0) { $ports = @($port) }

    $user = $script:txtUser.Text
    $pass = $script:txtPass.Text
    if ($script:chkNoCreds.Checked) { $user = ''; $pass = '' }

    $from = $script:txtFrom.Text.Trim()
    $to   = $script:txtTo.Text.Trim()

    # Only the jobs that actually need addresses insist on them, so a
    # connectivity check on a machine where nobody knows the sending address
    # still runs.
    $extTo = $script:txtExternalTo.Text.Trim()
    $connIps = @($script:txtConnectorIps.Text -split '[,;\s]+' | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' })
    if ($Job -eq 'autodiscover' -and -not (Test-StEmailAddress $from)) {
        [void][System.Windows.Forms.MessageBox]::Show('Enter the From address - Autodiscover is looked up for that mailbox.', 'SMTP Toolkit', 'OK', 'Warning')
        return $null
    }
    if ($Job -eq 'relay') {
        if (-not (Test-StEmailAddress $from) -or -not (Test-StEmailAddress $extTo)) {
            [void][System.Windows.Forms.MessageBox]::Show('The relay test needs a From address and an external recipient - an address outside this organisation, such as your own. Nothing is delivered.', 'SMTP Toolkit', 'OK', 'Warning')
            return $null
        }
    }

    $needsAddresses = @('reproduce','send')
    if ($needsAddresses -contains $Job) {
        if (-not (Test-StEmailAddress $from) -or -not (Test-StEmailAddress $to)) {
            [void][System.Windows.Forms.MessageBox]::Show('This test sends a real message, so it needs a valid From and To address.', 'SMTP Toolkit', 'OK', 'Warning')
            return $null
        }
    }

    $reproMs = ([int]$script:numReproTimeout.Value) * 1000
    if ($script:chkReproDefault.Checked) { $reproMs = 100000 }

    return @{
        Server          = $server
        Port            = $port
        Ports           = $ports
        Security        = (Get-StSecurityMode)
        From            = $from
        To              = $to
        Username        = $user
        Password        = $pass
        HeloName        = $(if ($script:txtHelo.Text.Trim()) { $script:txtHelo.Text.Trim() } else { $env:COMPUTERNAME })
        TimeoutMs       = ([int]$script:numTimeout.Value) * 1000
        MatrixTimeoutMs = ([int]$script:numMatrixTimeout.Value) * 1000
        MatrixLive      = [bool]$script:chkMatrixLive.Checked
        EnableSsl       = [bool]$script:chkReproSsl.Checked
        ReproTimeoutMs  = $reproMs
        ConfigBudgetSec = [int]$script:numConfig.Value
        AppProfile      = [string]$script:cboAppProfile.SelectedItem
        ServerKind      = [string]$preset.kind
        ExpectInternalCa= [bool]$preset.expectInternalCa
        ExternalTo      = $extTo
        ConnectorIps    = $connIps
    }
}

function Start-StJob {
    param([string]$Job)
    if ($script:busy) { return }
    $opt = Get-StUiOptions -Job $Job
    if ($null -eq $opt) { return }

    # Anything that puts a message in a mailbox asks first. Sweeps and probe
    # runs never do, so they never interrupt.
    if ($Job -eq 'reproduce' -or $Job -eq 'send' -or ($Job -eq 'matrix' -and $opt.MatrixLive)) {
        $what = switch ($Job) {
            'reproduce' { ('one message via System.Net.Mail.SmtpClient to {0}' -f $opt.To) }
            'send'      { ('one message via this tool''s own client to {0}' -f $opt.To) }
            default     { ('up to {0} messages to {1} - one for every combination that works' -f ($opt.Ports.Count * 2), $opt.To) }
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ('This will deliver {0}.' + [Environment]::NewLine + [Environment]::NewLine + 'Continue?') -f $what,
            'SMTP Toolkit - real mail will be sent', 'YesNo', 'Question')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $sh = [hashtable]::Synchronized(@{})
    $sh.Log        = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    $sh.Transcript = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    $sh.Sections   = [hashtable]::Synchronized(@{})
    # Carry forward what earlier runs established, so the verdict keeps getting
    # sharper instead of resetting every time a button is pressed.
    foreach ($k in @($script:Sections.Keys)) { $sh.Sections[$k] = $script:Sections[$k] }
    $sh.Status = 'Starting'
    $sh.Progress = 0
    $sh.Done = $false
    $sh.Cancel = $false
    $sh.Error = ''
    $sh.Version = 0
    $sh.Verdict = $null
    $sh.SeenIp = ''
    $script:Shared = $sh
    $script:logIndex = 0
    $script:transIndex = 0

    if ($script:startedUtc -eq [datetime]::MinValue) {
        $script:startedUtc = (Get-Date).ToUniversalTime()
        $script:startedSw  = [System.Diagnostics.Stopwatch]::StartNew()
    }

    try {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $script:rs = [runspacefactory]::CreateRunspace($iss)
        $script:rs.ApartmentState = 'STA'
        $script:rs.Open()
        $script:ps = [powershell]::Create()
        $script:ps.Runspace = $script:rs
        [void]$script:ps.AddScript($WorkerScript.ToString())
        [void]$script:ps.AddParameter('Shared', $sh)
        [void]$script:ps.AddParameter('EngineSource', $EngineSource)
        [void]$script:ps.AddParameter('Job', $Job)
        [void]$script:ps.AddParameter('Opt', $opt)
        $script:async = $script:ps.BeginInvoke()
        Set-StBusy -On $true
        $script:timer.Start()
        $startMsg = switch ($Job) {
            'sweep'       { ('Full sweep of {0} started.' -f $opt.Server) }
            'matrix'      { ('Settings matrix started against {0} on ports {1}.' -f $opt.Server, ($opt.Ports -join ', ')) }
            'reproduce'   { ('Reproducing the application send: port {0}, SSL {1}, timeout {2:N0} s.' -f $opt.Port, $(if ($opt.EnableSsl) { 'ticked' } else { 'unticked' }), ($opt.ReproTimeoutMs / 1000)) }
            'send'        { ('Sending a test message via port {0}.' -f $opt.Port) }
            'environment' { 'Reading local settings only - no network traffic.' }
            'relay'       { ('Relay test: no credentials, external recipient {0}. Nothing is delivered.' -f $opt.ExternalTo) }
            'autodiscover'{ ('Autodiscover check for {0} - unauthenticated, no password is sent.' -f $opt.From) }
            default       { 'Started.' }
        }
        Write-StLog $startMsg
        if ($opt.ServerKind -eq 'exchange-onprem' -and ($opt.Username -or $opt.Password) -and ($Job -eq 'sweep' -or $Job -eq 'matrix')) {
            Write-StLog 'Exchange on-premises: the password is tried at most once per run and never again after a rejection, so the Active Directory account is not pushed towards lockout. Reproduce app send uses the real .NET class, which makes up to two attempts (Kerberos/NTLM, then LOGIN).'
        }
    } catch {
        Write-StLog ('The background worker would not start: {0}' -f $_.Exception.Message)
        Set-StBusy -On $false
    }
}

function Complete-StWorker {
    try { if ($script:ps -and $script:async) { [void]$script:ps.EndInvoke($script:async) } } catch { }
    try { if ($script:ps) { $script:ps.Dispose() } } catch { }
    try { if ($script:rs) { $script:rs.Close(); $script:rs.Dispose() } } catch { }
    $script:ps = $null; $script:rs = $null; $script:async = $null
}

function Get-StSectionOrder {
    # Plain return for the same reason as ConvertTo-StPortList: this is only
    # ever consumed by foreach, which unrolls either shape, but keeping one
    # convention across the file removes the chance of the next caller
    # wrapping it in @() and quietly getting a nested array.
    return @(
        @{ key = 'context';          title = 'This machine';            grid = 'results' }
        @{ key = 'mx';               title = 'MX lookup';               grid = 'results' }
        @{ key = 'relay';            title = 'Relay / connector';       grid = 'relay'   }
        @{ key = 'autodiscover';     title = 'Autodiscover';            grid = 'relay'   }
        @{ key = 'dns';              title = 'DNS';                     grid = 'results' }
        @{ key = 'tcp';              title = 'Reachability';            grid = 'results' }
        @{ key = 'startTls';         title = 'STARTTLS';                grid = 'results' }
        @{ key = 'implicitTls';      title = 'Implicit SSL';            grid = 'results' }
        @{ key = 'tlsVersions';      title = 'TLS versions';            grid = 'results' }
        @{ key = 'certificate';      title = 'Certificate';             grid = 'results' }
        @{ key = 'publicIp';         title = 'Egress addresses';        grid = 'results' }
        @{ key = 'auth';             title = 'Authentication';          grid = 'results' }
        @{ key = 'envelope';         title = 'Envelope';                grid = 'results' }
        @{ key = 'send';             title = 'Live send';               grid = 'results' }
        @{ key = 'senderDns';        title = 'Sender DNS';              grid = 'results' }
        @{ key = 'dotNetSmtpClient'; title = 'Application reproduction';grid = 'matrix'  }
        @{ key = 'appMatrix';        title = 'Settings matrix';         grid = 'matrix'  }
        @{ key = 'crypto';           title = '.NET and TLS settings';   grid = 'env'     }
        @{ key = 'interference';     title = 'Local software';          grid = 'env'     }
        @{ key = 'appConfig';        title = 'Application config files';grid = 'env'     }
    )
}

function Update-StGrids {
    $grids = @{ results = $script:gridResults; matrix = $script:gridMatrix; env = $script:gridEnv; relay = $script:gridRelay }
    foreach ($g in $grids.Values) { $g.SuspendLayout(); $g.Rows.Clear() }

    foreach ($entry in (Get-StSectionOrder)) {
        $sec = $script:Sections[$entry.key]
        if (-not $sec) { continue }
        $grid = $grids[$entry.grid]

        if ($sec.state -ne 'measured') {
            # A section that did not run is shown, with its reason. Hiding it
            # would let an untested path read as a healthy one.
            [void](Add-StGridRow -Grid $grid -Row (New-StRow -Group $entry.title -Check ('({0})' -f $sec.state) `
                -Result $(if ($sec.state -eq 'failed') { 'Did not pass' } else { 'Not run' }) `
                -Basis 'not measured' -Status $(if ($sec.state -eq 'failed') { 'fail' } else { 'info' }) -Detail $sec.reason))
        }
        foreach ($row in @($sec.rows)) {
            if ($row) { [void](Add-StGridRow -Grid $grid -Row $row) }
        }
    }
    # Without this the first row sits selected in blue, which hides the status
    # colour on the one row a reader looks at first. Clearing the current cell
    # as well as the selection stops the grid re-selecting it when it is shown.
    foreach ($g in $grids.Values) {
        $g.ClearSelection()
        try { $g.CurrentCell = $null } catch { }
        $g.ResumeLayout()
    }
}

function Update-StVerdictUi {
    $v = $script:Verdict
    if (-not $v) { return }

    $colour = switch ($v.overall) {
        'green'       { [System.Drawing.Color]::FromArgb(20, 120, 40) }
        'amber'       { [System.Drawing.Color]::FromArgb(170, 110, 0) }
        'red'         { [System.Drawing.Color]::FromArgb(180, 30, 30) }
        default       { [System.Drawing.Color]::FromArgb(90, 90, 90) }
    }
    $label = switch ($v.overall) {
        'green'       { 'GO' }
        'amber'       { 'GO, WITH A CAVEAT' }
        'red'         { 'NO GO' }
        default       { 'NOT TESTED' }
    }
    $script:lblOverall.Text = $label
    $script:lblOverall.ForeColor = $colour
    $script:lblHeadline.Text = $v.headline

    $script:gridVerdict.Rows.Clear()
    foreach ($s in @($v.subscores)) {
        $stateText = switch ($s.level) {
            'green'       { 'pass' }
            'amber'       { 'caution' }
            'red'         { 'fail' }
            default       { 'not measured' }
        }
        $i = $script:gridVerdict.Rows.Add(@($s.name, $stateText, $s.summary, $s.remedy))
        $r = $script:gridVerdict.Rows[$i]
        switch ($s.level) {
            'green' { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(233, 247, 234) }
            'amber' { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 246, 224) }
            'red'   { $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 232, 232) }
            default { $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120) }
        }
    }
    $script:gridVerdict.ClearSelection()
    try { $script:gridVerdict.CurrentCell = $null } catch { }

    $rec = $v.recommendation
    $sb = New-Object System.Text.StringBuilder
    if ($rec -and $rec.available) {
        [void]$sb.AppendLine('RECOMMENDED SETTINGS')
        [void]$sb.AppendLine($rec.statement)
        [void]$sb.AppendLine(('Confidence: {0}.' -f $rec.confidence))
        if (@($rec.alternatives).Count -gt 0) {
            [void]$sb.AppendLine(('Also works: {0}.' -f ((@($rec.alternatives) | Select-Object -First 4) -join '; ')))
        }
        if (@($rec.avoid).Count -gt 0) {
            [void]$sb.AppendLine(('Do not use: {0}.' -f ((@($rec.avoid) | Select-Object -First 4) -join '; ')))
        }
    } else {
        [void]$sb.AppendLine('No working combination has been proven yet.')
        [void]$sb.AppendLine('Run the settings matrix on the Application simulation tab to find one.')
    }

    # The checklist belongs here rather than on a tab of its own. Correct
    # settings that appear to do nothing are what send a technician back to
    # blaming the network, and this is the panel they are reading when they go
    # to type the settings in.
    $app = $null
    if ($rec) { $app = $rec.application }
    if ($app) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(('AFTER CHANGING THE SETTINGS - {0}' -f $app.name))
        # The summary goes above the numbered list so the traps are visible
        # without scrolling, whatever the window has been resized to.
        if ($app.summary) { [void]$sb.AppendLine($app.summary) }
        if ($app.settingsPath) { [void]$sb.AppendLine(('Settings live at: {0} ({1}).' -f $app.settingsPath, $app.fields)) }
        $n = 1
        foreach ($item in @($app.checklist)) {
            [void]$sb.AppendLine(('{0}. {1}' -f $n, $item))
            $n++
        }
        if ($app.timeoutNote) { [void]$sb.AppendLine($app.timeoutNote) }
    }

    if (@($v.notes).Count -gt 0) {
        [void]$sb.AppendLine('')
        foreach ($n2 in @($v.notes)) { [void]$sb.AppendLine($n2) }
    }
    $script:txtRecommend.Text = ($sb.ToString() -replace "`r?`n", [Environment]::NewLine)
}

function New-StUiReport {
    $target = New-StTarget -Server $script:txtServer.Text.Trim() `
                           -Port ([int]($script:cboPort.Text -as [int])) `
                           -Security (Get-StSecurityMode) `
                           -From $script:txtFrom.Text.Trim() `
                           -To $script:txtTo.Text.Trim() `
                           -Username $(if ($script:chkNoCreds.Checked) { '' } else { $script:txtUser.Text }) `
                           -HeloName $script:txtHelo.Text.Trim() `
                           -Ports (ConvertTo-StPortList -Text $script:txtPorts.Text) `
                           -TimeoutMs (([int]$script:numTimeout.Value) * 1000) `
                           -Kind ([string](Get-StSelectedPreset).kind) `
                           -ExternalTo $script:txtExternalTo.Text.Trim() `
                           -ConnectorIps @($script:txtConnectorIps.Text -split '[,;\s]+' | Where-Object { $_ })

    $dur = 0
    if ($script:startedSw) { $dur = $script:startedSw.Elapsed.TotalSeconds }
    $transcript = @()
    if ($script:Shared -and $script:Shared.Transcript) { $transcript = @($script:Shared.Transcript.ToArray()) }

    return (New-StReport -Context $script:Context -Target $target -Sections ([pscustomobject]$script:Sections) `
                         -Verdict $script:Verdict -Transcript $transcript -Mode 'gui' `
                         -StartedUtc $script:startedUtc -DurationSec $dur)
}

function Update-StJsonTab {
    try {
        $script:Report = New-StUiReport
        $script:txtJson.Text = ((Get-StMarkedJson -Report $script:Report -Compress:$script:chkJsonCompress.Checked) -replace "`r?`n", [Environment]::NewLine)
    } catch {
        $script:txtJson.Text = ('The report could not be built: {0}' -f $_.Exception.Message)
    }
}

# ---------------------------------------------------------------------------
#  Timer - the only place the UI reads worker state
# ---------------------------------------------------------------------------
$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Interval = 300
$script:timer.Add_Tick({
    $sh = $script:Shared
    if (-not $sh) { $script:timer.Stop(); return }

    try {
        while ($script:logIndex -lt $sh.Log.Count) {
            Write-StLog ([string]$sh.Log[$script:logIndex])
            $script:logIndex++
        }
        if ($sh.Transcript.Count -gt $script:transIndex) {
            $new = New-Object System.Text.StringBuilder
            while ($script:transIndex -lt $sh.Transcript.Count) {
                [void]$new.Append([string]$sh.Transcript[$script:transIndex]).Append([Environment]::NewLine)
                $script:transIndex++
            }
            $script:txtTranscript.AppendText($new.ToString())
        }
    } catch { }

    $script:lblStatus.Text = [string]$sh.Status
    $p = [int]$sh.Progress
    if ($p -lt 0) { $p = 0 }
    if ($p -gt 100) { $p = 100 }
    $script:progress.Value = $p

    if ($sh.Version -ne $script:lastVersion) {
        $script:lastVersion = $sh.Version
        foreach ($k in @($sh.Sections.Keys)) { $script:Sections[$k] = $sh.Sections[$k] }
        if ($sh.Verdict) { $script:Verdict = $sh.Verdict }
        Update-StGrids
        Update-StVerdictUi
    }

    if ($sh.Done) {
        $script:timer.Stop()
        Complete-StWorker
        Set-StBusy -On $false
        # The connector route resolves its host from the MX; show it, so the
        # next run (and the report) uses the same host.
        if ($sh.ResolvedServer -and -not $script:txtServer.Text.Trim()) { $script:txtServer.Text = [string]$sh.ResolvedServer }
        foreach ($k in @($sh.Sections.Keys)) { $script:Sections[$k] = $sh.Sections[$k] }
        if ($sh.Verdict) { $script:Verdict = $sh.Verdict }
        Update-StGrids
        Update-StVerdictUi
        Update-StJsonTab
        if ($sh.Error) {
            Write-StLog ('Error: {0}' -f ($sh.Error -split "`n")[0])
        }
        if ($sh.Cancel) { Write-StLog 'Stopped at the end of the step that was running.' }
        else { Write-StLog 'Finished.' }
        $script:lblStatus.Text = 'Ready.'
    }
})
$script:lastVersion = -1

# ---------------------------------------------------------------------------
#  Events
# ---------------------------------------------------------------------------
$script:cboPreset.Add_SelectedIndexChanged({ Update-StPreset })

$script:chkShowPass.Add_CheckedChanged({
    $script:txtPass.UseSystemPasswordChar = (-not $script:chkShowPass.Checked)
})

$script:chkNoCreds.Add_CheckedChanged({
    $on = (-not $script:chkNoCreds.Checked)
    $script:txtUser.Enabled = $on
    $script:txtPass.Enabled = $on
})

$script:chkReproDefault.Add_CheckedChanged({
    $script:numReproTimeout.Enabled = (-not $script:chkReproDefault.Checked)
})

$script:btnSweep.Add_Click({ Start-StJob -Job 'sweep' })
$script:btnMatrix.Add_Click({ Start-StJob -Job 'matrix' })
$script:btnRepro.Add_Click({ Start-StJob -Job 'reproduce' })
$script:btnSend.Add_Click({ Start-StJob -Job 'send' })
$script:btnEnv.Add_Click({ Start-StJob -Job 'environment' })
$script:btnRelay.Add_Click({ Start-StJob -Job 'relay' })
$script:btnAutodiscover.Add_Click({ Start-StJob -Job 'autodiscover' })

$script:btnCancel.Add_Click({
    if ($script:Shared) {
        $script:Shared.Cancel = $true
        # Honest about what cancel can and cannot do: a socket already waiting
        # on a read cannot be interrupted, so the current step runs to its
        # timeout. Saying so beats a window that looks hung.
        Write-StLog 'Cancelling. The step already running will finish or time out first - a blocked socket read cannot be interrupted.'
        $script:btnCancel.Enabled = $false
    }
})

$script:btnJsonRefresh.Add_Click({ Update-StJsonTab })
$script:chkJsonCompress.Add_CheckedChanged({ Update-StJsonTab })

$script:btnJsonCopy.Add_Click({
    try {
        if (-not $script:txtJson.Text) { Update-StJsonTab }
        [System.Windows.Forms.Clipboard]::SetText($script:txtJson.Text)
        Write-StLog 'The JSON report is on the clipboard.'
    } catch {
        Write-StLog ('Could not reach the clipboard: {0}' -f $_.Exception.Message)
    }
})

$script:btnSaveJson.Add_Click({
    try {
        $script:Report = New-StUiReport
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = 'JSON report (*.json)|*.json|All files (*.*)|*.*'
        $dlg.FileName = ('SmtpDiag_{0}_{1}.json' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $dlg.InitialDirectory = $(if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::GetFolderPath('Desktop') })
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $json = Get-StMarkedJson -Report $script:Report
            [System.IO.File]::WriteAllText($dlg.FileName, $json, (New-Object System.Text.UTF8Encoding $false))
            Write-StLog ('Report written to {0}' -f $dlg.FileName)
            Update-StJsonTab
        }
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(('The report could not be saved: {0}' -f $_.Exception.Message), 'SMTP Toolkit', 'OK', 'Error')
    }
})

$script:form.Add_FormClosing({
    if ($script:Shared) { $script:Shared.Cancel = $true }
    try { if ($script:timer) { $script:timer.Stop() } } catch { }
    # The worker cannot be killed mid-socket, so give it a moment and then let
    # the process exit take the runspace with it.
    if ($script:ps -and $script:async -and -not $script:async.IsCompleted) {
        [void]$script:async.AsyncWaitHandle.WaitOne(1500, $false)
    }
    Complete-StWorker
})

# ---------------------------------------------------------------------------
#  Start
# ---------------------------------------------------------------------------
Update-StPreset

$sysTag = ''
if ($script:Context.isSystem)     { $sysTag = '  [SYSTEM]' }
elseif ($script:Context.elevated) { $sysTag = '  [elevated]' }
Write-StLog ('SMTP Toolkit {0} on {1}, running as {2}{3}, PowerShell {4}.' -f (Get-StVersion), $script:Context.computerName, $script:Context.runAs, $sysTag, $script:Context.psVersion)
Write-StLog 'Fill in the server, addresses and any credentials, then press Run full sweep. Nothing is delivered unless you press one of the two send buttons.'

[void]$script:form.ShowDialog()
$script:form.Dispose()

