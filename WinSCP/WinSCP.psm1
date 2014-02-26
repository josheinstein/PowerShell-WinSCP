[WinSCP.Session]$WinSCP_Session = $null

#.SYNOPSIS
# Opens an active WinSCP session so that other commands can be
# invoked without specifying connection details to each command.
function Open-WinSCPSession {

    [CmdletBinding()]
    param(

        # The host name or ip address of the remote host to connect to
        [Alias('h')]
        [Parameter(Mandatory=$true, Position=1)]
        [String]$HostName,

        # Specifies a non-standard port number to use to connect.
        # A value of zero will use the default port number for the
        # selected protocol.
        [Parameter()]
        [Int32]$Port = 0,

        # The transfer protocol to use. The default is FTP.
        [Parameter()]
        [WinSCP.Protocol]$Protocol = 'Ftp',

        # The type of FTP security to use. The default value is None which
        # transmits commands and data in plain text.
        [Parameter()]
        [WinSCP.FtpSecure]$Security = 'None',

        # The user name to use to authenticate with the remote host.
        [Alias('u')]
        [Parameter()]
        [String]$UserName,

        # The password to use to authenticate with the remote host.
        [Alias('p')]
        [Parameter()]
        [String]$Password,

        # The SSH host key fingerprint to use with the remote host when
        # using the SFTP protocol.
        [Parameter()]
        [String]$HostKey,

        # When this switch is used, FTP will work in active mode which does
        # requires that the client be able to accept incoming data connections
        # back from the remote host. The default mode is passive.
        [Parameter()]
        [Switch]$Active,

        # When this switch is used in conjunction with secure protocols, the
        # server's identity is not verified.
        [Parameter()]
        [Switch]$IgnoreHostSecurity,

        # The timeout when communicating with the remote server.
        # The default value is 1 minute.
        [Parameter()]
        [TimeSpan]$Timeout = [TimeSpan]::FromMinutes(1),

        # Returns the WinSCP Session object from the command.
        [Parameter()]
        [Switch]$PassThru

    )

    process {

        if ($WinSCP_Session) {
            Write-Error "Session is already open. Use Close-WinSCPSession to close the active session."
            Return
        }

        $SessionOptions = New-Object WinSCP.SessionOptions

        $SessionOptions.HostName = $HostName
        $SessionOptions.Protocol = $Protocol
        $SessionOptions.FtpSecure = $Security
        $SessionOptions.Timeout = $Timeout

        if ($Active.IsPresent) { $SessionOptions.FtpMode = 'Active' }
        else { $SessionOptions.FtpMode = 'Passive' }

        if ($Port) { $SessionOptions.PortNumber = $Port }
        if ($UserName) { $SessionOptions.UserName = $UserName }
        if ($Password) { $SessionOptions.Password = $Password }
        if ($HostKey) { $SessionOptions.SshHostKeyFingerprint = $HostKey }

        if ($IgnoreHostSecurity) {
            $SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
            $SessionOptions.GiveUpSecurityAndAcceptAnySslHostCertificate = $true
            $SessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
        }

        $Script:WinSCP_Session = New-Object WinSCP.Session
        $Script:WinSCP_Session.Open($SessionOptions)

        if ($PassThru) {
            Write-Output $Script:WinSCP_Session
        }

    }

}

#.SYNOPSIS
# Closes the currently active WinSCP session.
function Close-WinSCPSession {

    [CmdletBinding()]
    param(

        # The session to close. If not specified, the currently active
        # default session is closed.
        [Parameter()]
        [WinSCP.Session]$Session = $Null

    )

    process {

        if ($Session) {
            $Session.Dispose()
        }
        elseif ($Script:WinSCP_Session) {
            $Script:WinSCP_Session.Dispose()
            $Script:WinSCP_Session = $Null
        }
        else {
            Write-Warning "There is no currently active session. Use Open-WinSCPSession to open an active session."
        }

    }

}

#.SYNOPSIS
# Gets a directory listing from the remote host in the currently active
# WinSCP session.
function Get-WinSCPDirectory {

    [CmdletBinding()]
    param(

        # The path on the remote host to get a directory listing for.
        [Parameter(Position=1)]
        [String]$Path = '.',

        # The remote session to use.
        [Parameter()]
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        if ($Result = $Session.ListDirectory($Path)) {
            Write-Output $Result.Files
        }

    }

}

#.SYNOPSIS
# Downloads one or more files from the remote host and saves them to
# the current working directory or a specified local path.
function Get-WinSCPFiles {

    [CmdletBinding()]
    param(

        # The path of a file on the remote host (wildcards allowed) or the
        # path to a directory on the remote host (all files in the directory)
        # to download.
        [Parameter(Position=1, Mandatory=$true)]
        [String]$RemotePath,

        # A path on the local system to download the file to. If not specified,
        # the current working directory will be used.
        [Parameter(Position=2)]
        [String]$LocalPath = '.',

        # When specified, the files on the remote server will be deleted once
        # they have been downloaded.
        [Parameter()]
        [Switch]$Remove,

        # The remote session to use.
        [Parameter()]
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        $LocalPathInfo = Resolve-Path -LiteralPath:$LocalPath
        if ($LocalPathInfo.Provider.Name -ne 'FileSystem') {
            Write-Error "Local path must be a valid FileSystem path."
            Return
        }
        else {
            # Stuff it back into the LocalPath variable.
            # If it's a directory, though, we need to tack on the \* or else
            # WinSCP will complain.
            $LocalPath = $LocalPathInfo.ProviderPath
            if (Test-Path -LiteralPath:$LocalPath -PathType Container) {
                $LocalPath += '\*'
            }
        }

        $TransferOptions = New-Object WinSCP.TransferOptions
        $TransferOptions.TransferMode = 'Binary'

        $Result = $Session.GetFiles($RemotePath, $LocalPath, $Remove.IsPresent, $TransferOptions)
        
        # Write the FileInfo about the successfully downloaded files to the pipeline
        foreach ($Transfer in $Result.Transfers) {
            if ($Transfer.Error) {
                Write-Warning "$($Transfer.FileName) - $($Transfer.Error.Message)"
            }
            else {
                Get-Item -LiteralPath:$Transfer.Destination -ErrorAction SilentlyContinue
            }
        }

    }

}

Export-ModuleMember -Function *-*
