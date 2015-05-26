[WinSCP.Session]$WinSCP_Session = $null
[Int32]$WinSCP_ProgressID = Get-Random

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
            
            if ($Protocol -in ([WinScp.Protocol]::Sftp, [WinScp.Protocol]::Scp)) {
                $SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
            }
            
            if ($Security -eq [WinScp.FtpSecure]::Implicit) {
                $SessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
                $SessionOptions.GiveUpSecurityAndAcceptAnySslHostCertificate = $true
            }
            elseif ($Security -eq [WinScp.FtpSecure]::ExplicitSsl) {
                $SessionOptions.GiveUpSecurityAndAcceptAnySslHostCertificate = $true
            }
            elseif ($Security -eq [WinScp.FtpSecure]::ExplicitTls) {
                $SessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
            }

        }

        try {

            $Script:WinSCP_Session = New-Object WinSCP.Session

            $Script:WinSCP_Session.add_FileTransferProgress({
                param($sender,[WinSCP.FileTransferProgressEventArgs]$e)
                $CPS = [Int64]$e.CPS
                $ProgressArgs = @{
                    Id              = $Script:WinSCP_ProgressID
                    Activity        = "$(if ($e.Side -eq 'Local') {'Sending'} else {'Receiving'}) ($CPS/s)"
                    Status          = $e.FileName
                    PercentComplete = $e.OverallProgress * 100
                    Completed       = $(if ($e.OverallProgress -ge 1) {$true} else {$false})
                }
                Write-Progress @ProgressArgs
            })

            $Script:WinSCP_Session.Open($SessionOptions)

            if ($PassThru) {
                Write-Output $Script:WinSCP_Session
            }

        }
        catch {

            if ($Script:WinSCP_Session) {
                $Script:WinSCP_Session.Dispose();
                $Script:WinSCP_Session = $null
            }

            Write-Error $_

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
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        if ($Session) {
            
            $Session.Dispose()

            if ($Session -eq $Script:WinSCP_Session) {
                $Script:WinSCP_Session = $Null
            }

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
    [OutputType([WinSCP.RemoteFileInfo])]
    param(

        # The path on the remote host to get a directory listing for.
        [Parameter(Position=1)]
        [String]$Path = '.',

        # When specified without -Directory, gets only files.
        [Parameter()]
        [Switch]$File,

        # When specified without -File, gets only directories.
        [Parameter()]
        [Switch]$Directory,

        # When specified, recursively lists directories that are returned
        [Parameter()]
        [Switch]$Recurse,

        # When specified, limits the output to only files or directories whose names
        # match the given pattern(s).
        [Parameter()]
        [String[]]$Include,

        # When specified, limits the output to only files or directories whose names
        # do not match the given pattern(s).
        [Parameter()]
        [String[]]$Exclude,

        # The remote session to use.
        [Parameter()]
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        if ($Result = $Session.ListDirectory($Path)) {

            if (!$Path.EndsWith('/')) { $Path += '/' }

            foreach ($RemoteFileInfo in $Result.Files) {

                # Skip the useless . and .. entries
                if ($RemoteFileInfo.Name -match '^\.\.?$') { continue; }

                # Adds the full path of the file to the output object
                $RemoteFileInfo = $RemoteFileInfo |
                    Add-Member -PassThru NoteProperty Path "${Path}$($RemoteFileInfo.Name)" |
                    Add-Member -PassThru NoteProperty BaseName ([IO.Path]::GetFileNameWithoutExtension($RemoteFileInfo.Name))

                # Determine inclusion
                $IsIncluded = $true
                if ($Include.Length) {
                    $IsIncluded = $false
                    foreach ($Pattern in $Include) {
                        if ($RemoteFileInfo.Name -like $Pattern) {
                            Write-Verbose "$($RemoteFileInfo.Name) included by $Pattern"
                            $IsIncluded = $true
                            break
                        }
                    }
                }

                # Determine exclusion
                $IsExcluded = $False
                if ($Exclude.Length) {
                    $IsExcluded = $False
                    foreach ($Pattern in $Exclude) {
                        if ($RemoteFileInfo.Name -like $Pattern) {
                            Write-Verbose "$($RemoteFileInfo.Name) excluded by $Pattern"
                            $IsExcluded = $True
                            break
                        }
                    }
                }

                # If neither or both of these switches is specified, the
                # behavior is to include both directories and files.
                # If one or the other is specified, it skips output accordingly.
                if ($File.IsPresent -ne $Directory.IsPresent) {
                    if ($File.IsPresent -eq $RemoteFileInfo.IsDirectory) { $IsIncluded = $False }
                    if ($Directory.IsPresent -ne $RemoteFileInfo.IsDirectory) { $IsIncluded = $False }
                }

                # Send the file/directory result to the pipeline
                if ($IsIncluded -and !$IsExcluded) { Write-Output $RemoteFileInfo }

                # Attempt to recursively call the function
                # Skip directories beginning with a dot
                if ($Recurse -and $RemoteFileInfo.IsDirectory -and $RemoteFileInfo.Name -notlike '.*') {
                    Get-WinSCPDirectory -Path "$($RemoteFileInfo.Path)/" -File:$File -Directory:$Directory -Recurse -Include:$Include -Exclude:$Exclude -Session $Session
                }

            }
        }

    }

}

#.SYNOPSIS
# Uploads one or more files to the remote host from
# the current working directory or a specified local path.
function Send-WinSCPFiles {

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(

        # A path on the local system to upload the file from.
        [Alias('Path')]
        [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
        [String]$LocalPath,

        # The path of a file on the remote host or the path to a directory
        # on the remote host to upload to.
        [Parameter(Position=2)]
        [String]$RemotePath = '.',

        # Optionally specifies one or more wildcard masks that are matched against
        # the local path to limit the uploaded files to those that match at
        # least one of the masks specified.
        [Parameter()]
        [String[]]$Include,

        # Optionally specifies one or more wildcard masks that are matched against
        # the local path to limit the uploaded files to
        # those that do not match any of the masks specified.
        [Parameter()]
        [String[]]$Exclude,

        # The remote session to use.
        [Parameter()]
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        # Expand wildcards and apply include/exclude logic
        $LocalFiles = @(Get-Item -Path:$LocalPath -Include:$Include -Exclude:$Exclude)

        foreach ($LocalFile in $LocalFiles) {
        
            if ($PSCmdlet.ShouldProcess($LocalFile, "Upload $LocalFile")) {

                $TransferOptions = New-Object WinSCP.TransferOptions
                $TransferOptions.TransferMode = 'Binary'

                $RemoteFile = $RemotePath
                if ($RemoteFile.EndsWith('/')) { $RemoteFile += (Split-Path -Leaf $LocalFile) }

                Write-Verbose "PutFiles: Local=$LocalFile, Remote=$RemoteFile"

                $Result = $Session.PutFiles($LocalFile, $RemoteFile, $False, $TransferOptions)
        
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

    }

}

#.SYNOPSIS
# Downloads one or more files from the remote host and saves them to
# the current working directory or a specified local path.
function Receive-WinSCPFiles {

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(

        # The path of a file on the remote host (wildcards allowed) or the
        # path to a directory on the remote host (all files in the directory)
        # to download.
        [Alias('Path')]
        [Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [String]$RemotePath,

        # A path on the local system to download the file to. If not specified,
        # the current working directory will be used.
        [Parameter(Position=2)]
        [String]$LocalPath = '.',

        # When specified, the files on the remote server will be deleted once
        # they have been downloaded.
        [Parameter()]
        [Switch]$Remove,

        # Optionally specifies one or more wildcard masks that are matched against
        # the remote path to limit the downloaded files to those that match at
        # least one of the masks specified.
        # Note that this parameter will not behave as expected when -RemotePath
        # specifies a wildcard or directory name, because it is matched against
        # the remote path *before* wildcard expansion or directory listing.
        [Parameter()]
        [String[]]$Include,

        # Optionally specifies one or more wildcard masks that are matched against
        # the remote path to limit the downloaded files to
        # those that do not match any of the masks specified.
        # Note that this parameter will not behave as expected when -RemotePath
        # specifies a wildcard or directory name, because it is matched against
        # the remote path *before* wildcard expansion or directory listing.
        [Parameter()]
        [String[]]$Exclude,

        # The remote session to use.
        [Parameter()]
        [WinSCP.Session]$Session = $Script:WinSCP_Session

    )

    process {

        # Determine inclusion
        $IsIncluded = $true
        if ($Include.Length) {
            $IsIncluded = $false
            foreach ($Pattern in $Include) {
                if ((Split-Path -Leaf $RemotePath) -like $Pattern) {
                    Write-Verbose "$RemotePath included by $Pattern"
                    $IsIncluded = $true
                    break
                }
            }
        }

        # Determine exclusion
        $IsExcluded = $False
        if ($Exclude.Length) {
            $IsExcluded = $False
            foreach ($Pattern in $Exclude) {
                if ((Split-Path -Leaf $RemotePath) -like $Pattern) {
                    Write-Verbose "$RemotePath excluded by $Pattern"
                    $IsExcluded = $True
                    break
                }
            }
        }

        if ($IsIncluded -and !$IsExcluded) {

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
                    if (!$LocalPath.EndsWith('\')) {
                        $LocalPath += '\'
                    }
                    #$LocalPath += '*'
                }
            }

            if ($PSCmdlet.ShouldProcess($RemotePath, "Download to $LocalPath")) {

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

    }

}

Export-ModuleMember -Function *-*

Set-Alias Get-WinSCPFiles Receive-WinSCPFiles 
Export-ModuleMember -Alias Get-WinSCPFiles
