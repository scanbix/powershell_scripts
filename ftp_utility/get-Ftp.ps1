# Load WinSCP .NET assembly
# Add-Type -Path "WinSCPnet.dll"
Add-Type -Path (Join-Path $PSScriptRoot "WinSCPnet.dll")

# Call example
# (Get_PreviousWeekDate 'Sunday').GetType()
# (Get_PreviousWeekDate 'Sunday').ToString('yyyyMMdd')
function Get_PreviousWeekDate {
    param (
        [Parameter()]
        [ValidateSet("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")]
        [string]
        $WeekDay
    )

    $date = Get-Date

    for($i=1; $i -le 7; $i++)
    {        
        if($date.AddDays(-$i).DayOfWeek -eq $WeekDay)
        {
            $date.AddDays(-$i)
            break
        }
    }

    # Uncommenting below breaks the function.
    # return Get-Date
    
}

function Open-Ftp {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Alias,
        [Parameter()]
        [string]
        $Protocol,
        [Parameter()]
        [string]
        $HostName,
        [Parameter()]
        [string]
        $UserName,
        [Parameter()]
        [string]
        $Password,
        [Parameter()]
        [string]
        $RemotePath,
        [Parameter()]
        [string]
        $LocalPath = '.\ftp\'
    )


    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Ftp
        HostName = $HostName
        UserName = $UserName
        Password = $Password
    }

    $session = New-Object WinSCP.Session

    try
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName -ForegroundColor DarkGreen
        
        # Connect
        $session.Open($sessionOptions)

        # Download files
        $session.GetFiles($RemotePath, $LocalPath).Check()


        # Exception, and ErrorDetails
    }
    catch
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName ($_) -ForegroundColor DarkRed
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
}

function Open-Sftp {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Alias,
        [Parameter()]
        [string]
        $Protocol,
        [Parameter()]
        [string]
        $HostName,
        [Parameter()]
        [string]
        $UserName,
        [Parameter()]
        [string]
        $SshHostKeyFingerprint,
        [Parameter()]
        [string]
        $Password,
        [Parameter()]
        [string]
        $RemotePath,
        [Parameter()]
        [string]
        $LocalPath = '.\ftp\'
    )


    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $HostName
        UserName = $UserName
        SshHostKeyFingerprint = $SshHostKeyFingerprint
        Password = $Password
    }

    $session = New-Object WinSCP.Session

    try
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName -ForegroundColor DarkGreen
        
        # Connect
        $session.Open($sessionOptions)

        # Download files
        $session.GetFiles($RemotePath, $LocalPath).Check()
    }
    catch
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName ($_) -ForegroundColor DarkRed
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
}

function Open-SftpUsingKey {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Alias,
        [Parameter()]
        [string]
        $Protocol,
        [Parameter()]
        [string]
        $HostName,
        [Parameter()]
        [string]
        $UserName,
        [Parameter()]
        [string]
        $SshHostKeyFingerprint,
        [Parameter()]
        [string]
        $PrivateKeyPath,
        [Parameter()]
        [string]
        $Passphrase,
        [Parameter()]
        [string]
        $RemotePath,
        [Parameter()]
        [string]
        $LocalPath = '.\ftp\'
    )
    
    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $HostName
        UserName = $UserName
        SshHostKeyFingerprint = $SshHostKeyFingerprint
        PrivateKeyPassphrase = $Passphrase
        SshPrivateKeyPath = $PrivateKeyPath
    }

    $session = New-Object WinSCP.Session

    try
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName -ForegroundColor DarkGreen
        
        # Connect
        $session.Open($sessionOptions)

        # Download files
        $session.GetFiles($RemotePath, $LocalPath).Check()
    }
    catch
    {
        Write-Host $Alias - $Protocol : $HostName : $UserName ($_) -ForegroundColor DarkRed
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
}

# The list below follows Excel sheet connection order for PowerCenter

# Open-Ftp -Alias 'FTP_NBP' -Protocol 'FTP' -HostName 'xyz.com' -UserName 'aaa' -Password 'bb' -RemotePath '/Download/zyx/xyz.txt'
# Open-Sftp -Alias 'FTP_NBP' -Protocol 'SFTP' -HostName 'xyz.com' -UserName 'aaa' -SshHostKeyFingerprint 'ssh-rsa 1024 jxyz' -Password 'aa' -RemotePath '/Download/xyz/xyz.txt'
# Open-SftpUsingKey -Alias 'SAM_TEST_FTP_test' -Protocol 'SFTP' -HostName 'xyz.com' -UserName 'aaa' -SshHostKeyFingerprint 'ssh-rsa 1024 jxyz' -PrivateKeyPath 'C:\some_key.pk' -Passphrase 'aaa' -RemotePath '/Download/xyz/xyz.txt'

