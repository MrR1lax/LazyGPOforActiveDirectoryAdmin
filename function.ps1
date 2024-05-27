Function RSAT {
    $RSAT = Get-module -ListAvailable -Name ActiveDirectory
    If ($RSAT) {
        Return $true
    } Else {
        Return $false
    }
}
Function AllGPO {
    $all = Get-GPO -All
    $report = @()
    foreach ($gpo in $all) {
        $obj = New-Object Psobject
        $obj | Add-Member -Name "DisplayName" -membertype Noteproperty -Value $gpo.DisplayName
        $obj | Add-Member -Name "Id" -membertype Noteproperty -Value $gpo.Id
        $obj | Add-Member -Name "DomainName" -membertype Noteproperty -Value $gpo.DomainName
        $obj | Add-Member -Name "CreationTime" -membertype Noteproperty -Value $gpo.CreationTime
        $obj | Add-Member -Name "ModificationTime" -membertype Noteproperty -Value $gpo.ModificationTime
        $obj | Add-Member -Name "GpoStatus" -membertype Noteproperty -Value $gpo.GpoStatus
        $report += $obj
    }
    return $report
}
Function AllGroup {
    $all = Get-ADGroup -Filter * -Properties *
    $report = @()
    foreach ($group in $all) {
        $obj = New-Object Psobject
        $obj | Add-Member -Name "Name" -membertype Noteproperty -Value $group.Name
        $obj | Add-Member -Name "DistinguishedName" -membertype Noteproperty -Value $group.DistinguishedName
        $obj | Add-Member -Name "Description" -membertype Noteproperty -Value $group.Description
        $obj | Add-Member -Name "whenCreated" -membertype Noteproperty -Value $group.whenCreated
        $obj | Add-Member -Name "Created" -membertype Noteproperty -Value $group.Created
        $obj | Add-Member -Name "whenChanged" -membertype Noteproperty -Value $group.whenChanged
        $obj | Add-Member -Name "ObjectGUID" -membertype Noteproperty -Value $group.ObjectGUID
        $obj | Add-Member -Name "SID" -membertype Noteproperty -Value $group.SID
        $obj | Add-Member -Name "member" -membertype Noteproperty -Value $group.member
        $obj | Add-Member -Name "Members" -membertype Noteproperty -Value $group.Members
        $obj | Add-Member -Name "MemberOf" -membertype Noteproperty -Value $group.MemberOf
        $report += $obj
    }
    return $report
}
Function GroupUnitaire {
    param(
        [Parameter(Mandatory=$True,Position=0)]$Group
    )
    $result = Get-ADGroup -Identity $Group -Properties *
    Return $result
}
Function UserAD {
    param(
        [Parameter(Mandatory=$True,Position=0)]$User
    )
    $result = Get-ADUser -Identity $User -Properties *
    Return $result
}
Function PrenomNomAD {
    param(
        [Parameter(Mandatory=$True,Position=0)]$Prenom,
        [Parameter(Mandatory=$True,Position=1)]$Nom
    )
    $result = Get-ADUser -Filter "GivenName -eq '$Prenom' -and SurName -eq '$nom'" -Properties *
    Return $result
}
Function RegValue {
    param(
        [Parameter(Mandatory=$True,Position=0)]$Hive
    )
    If (Test-Path -Path $Hive) {
        $result = Get-ItemProperty -Path $Hive -ErrorAction 0
    } Else {
        $result = $null
    }
    Return $result
}
Function ComputerAD {
    param(
        [Parameter(Mandatory=$True,Position=0)]$Computer
    )
    $result = Get-ADComputer -Identity $Computer -Properties *
    Return $result
}
Function Update-Log {
    param(
        [string]$Message,
        [string]$Color,
        [switch]$NoNewLine
    )
    $LogTextBox.SelectionColor = $Color
    $LogTextBox.AppendText("$Message")
    if (-not $NoNewLine) { $LogTextBox.AppendText("`n") }
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}
Function Update-Log-Obj {
    param(
        [Parameter(Mandatory=$True,Position=0)]$Object
    )
    foreach($object_properties in $Object.PsObject.Properties) {
        If ($object_properties.Value.count -eq "0") {
            Update-Log $object_properties.Name -NoNewLine -Color "LightBlue"
            Update-Log " : " -NoNewLine -Color "LightBlue"
            Update-Log "Sa valeur est vide" -Color "Yellow"
        } Else {
            Update-Log $object_properties.Name -NoNewLine -Color "LightBlue"
            Update-Log " : " -NoNewLine -Color "LightBlue"
            Update-Log $object_properties.Value -Color "Yellow"
        }
    }
}
Function Set-LogoAD {
    Update-Log "                        _____  ___________________           .__" -Color "LightBlue"
    Update-Log "                       /  _  \ \_  __  \__    ___/___   ____ |  |   ______" -Color "LightBlue"
    Update-Log "                      /  /_\  \ |  | \  \|    | /  _ \ /  _ \|  |  /  ___/" -Color "LightBlue"
    Update-Log "                     /    |    \|  |_|   \    |(  <_> |  <_> )  |__\___ \" -Color "LightBlue"
    Update-Log "                     \____|__  /_______  /____| \____/ \____/|____/____  >" -Color "LightBlue"
    Update-Log "                             \/        \/                              \/" -Color "LightBlue"
    Update-Log "                                      by M@x" -Color "Gold"
    Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
}
Function Set-LogoProxy {
    Update-Log "                  .___       .__  __ __________                             " -Color "LightBlue"
    Update-Log "                  |   | ____ |__|/  |\______   \_______  _______  ______.__." -Color "LightBlue"
    Update-Log "                  |   |/    \|  \   __\     ___/\_  __ \/  _ \  \/  <   |  |" -Color "LightBlue"
    Update-Log "                  |   |   |  \  ||  | |    |     |  | \(  <_> >    < \___  |" -Color "LightBlue"
    Update-Log "                  |___|___|  /__||__| |____|     |__|   \____/__/\_ \/ ____|" -Color "LightBlue"
    Update-Log "                           \/                                      \/\/     " -Color "LightBlue"
    Update-Log "                                             by M@x" -Color "Gold"
    Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
}
Function DC {
    #param(
    #    [Parameter(Mandatory=$True,Position=0)]$Object
    #)
    $all = Get-ADDomainController -filter *
    #$all.Count
    $report = @()
    Foreach ($dc in $all) {
        $obj = New-Object Psobject
        $obj | Add-Member -Name "Name" -membertype Noteproperty -Value $dc.Name
        $obj | Add-Member -Name "Domain" -membertype Noteproperty -Value $dc.Domain
        $obj | Add-Member -Name "OperatingSystem" -membertype Noteproperty -Value $dc.OperatingSystem
        $obj | Add-Member -Name "Site" -membertype Noteproperty -Value $dc.Site
        $obj | Add-Member -Name "IPv4" -membertype Noteproperty -Value $dc.IPv4Address
        $obj | Add-Member -Name "IPv6" -membertype Noteproperty -Value $dc.IPv6Address
        If (Test-Path -Path "\\$($dc.Name)\SYSVOL\$($dc.Domain)\Policies\PolicyDefinitions") {
            $obj | Add-Member -Name "SysVolStatus" -membertype Noteproperty -Value "OK"
            $GroupPolicyDate = (Get-Item -Path "\\$($dc.Name)\SYSVOL\$($dc.Domain)\Policies\PolicyDefinitions\GroupPolicy.admx").LastWriteTime
            $obj | Add-Member -Name "GroupPolicyDate" -membertype Noteproperty -Value $GroupPolicyDate
        } Else {
            $obj | Add-Member -Name "SysVolStatus" -membertype Noteproperty -Value "KO"
            $obj | Add-Member -Name "GroupPolicyDate" -membertype Noteproperty -Value "KO"
        }
        $report += $obj
    }
    return $report
}