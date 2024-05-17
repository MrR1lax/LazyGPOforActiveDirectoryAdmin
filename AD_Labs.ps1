$ProxyRegistre = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
$ConnectionsRegistre = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"

$Path = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
."$Path\function.ps1"
$Date = Get-Date -format "dd-MM-yyyy_HH-mm-ss"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Outillage : Active Directory"
$Form.Size = New-Object System.Drawing.Size(1200, 550)
$Form.SizeGripStyle = "Hide"
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox = $false
$Form.StartPosition = "CenterScreen"
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Form.Icon = $Icon

$FormKO = New-Object System.Windows.Forms.Form
$FormKO.Text = "Outillage : Active Directory"
$FormKO.Size = New-Object System.Drawing.Size(300, 200)
$FormKO.SizeGripStyle = "Hide"
$FormKO.FormBorderStyle = "FixedSingle"
$FormKO.MaximizeBox = $false
$FormKO.StartPosition = "CenterScreen"
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$FormKO.Icon = $Icon

# Label 1 (Erreur !) dans FormKO
$LabelKO = New-Object System.Windows.Forms.Label
$LabelKO.Location = New-Object System.Drawing.Size(50, 50)
$LabelKO.Size = New-Object System.Drawing.Size(200, 30)
$LabelKO.TextAlign = "MiddleCenter"
$LabelKO.Text = "Erreur, les RSAT ne sont pas presents. Abandon."
$FormKO.Controls.Add($LabelKO)

# Création de la fenêtre de log
$LogTextBox = New-Object System.Windows.Forms.RichTextBox
$LogTextBox.Location = New-Object System.Drawing.Size(500, 30)
$LogTextBox.Size = New-Object System.Drawing.Size(675, 472)
$LogTextBox.ReadOnly = 'True'
$LogTextBox.BackColor = 'Black'
$LogTextBox.ForeColor = 'White'
$LogTextBox.Font = 'Consolas, 10'
$LogTextBox.DetectUrls = $false
Set-LogoAD
$Form.Controls.Add($LogTextBox)

# Creation de la table de controle
$TabControl = New-object System.Windows.Forms.TabControl
$TabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$TabControl.Location = New-Object System.Drawing.Size(10, 10)
$TabControl.Size = New-Object System.Drawing.Size(490, 490)
$Form.Controls.Add($TabControl)

### Tab 1 GPO ###
# Creation de l'onglet 1 GPO
$Tab1Page = New-Object System.Windows.Forms.TabPage
$Tab1Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab1Page.UseVisualStyleBackColor = $true
$Tab1Page.Text = "GPO"
$Tab1Page.AutoScroll = $true
$TabControl.Controls.Add($Tab1Page)
# Boutton 1 Analyse Tab 1
$Button1Tab1 = New-Object System.Windows.Forms.Button
$Button1Tab1.location = New-Object System.Drawing.Size(215, 10)
$Button1Tab1.Size = New-Object System.Drawing.Size(60, 20)
$Button1Tab1.Text = "Analyse"
$Tab1Page.Controls.Add($Button1Tab1)
# ComboBox 1 List de GPO Tab 1
$ComboBox1Tab1 = New-Object System.Windows.Forms.ComboBox
$ComboBox1Tab1.location = New-Object System.Drawing.Size(10, 10)
$ComboBox1Tab1.Size = New-Object System.Drawing.Size(460, 20)
# Boutton 2 Reset Tab 1 GPO
$Button2Tab1 = New-Object System.Windows.Forms.Button
$Button2Tab1.location = New-Object System.Drawing.Size(215, 420)
$Button2Tab1.Size = New-Object System.Drawing.Size(60, 20)
$Button2Tab1.Text = "Reset"
# Boutton 3 Information Tab 1 GPO
$Button3Tab1 = New-Object System.Windows.Forms.Button
$Button3Tab1.location = New-Object System.Drawing.Size(205, 40)
$Button3Tab1.Size = New-Object System.Drawing.Size(80, 20)
$Button3Tab1.Text = "Information"
# Bouton 4 Extraction unitaire (Extract HTML) Tab 1 GPO
$Button4Tab1 = New-Object System.Windows.Forms.Button
$Button4Tab1.location = New-Object System.Drawing.Size(185, 180)
$Button4Tab1.Size = New-Object System.Drawing.Size(120, 20)
$Button4Tab1.Text = "Extraction HTML"
# Bouton 5 Extraction unitaire (Extract XML) Tab 1 GPO
$Button5Tab1 = New-Object System.Windows.Forms.Button
$Button5Tab1.location = New-Object System.Drawing.Size(185, 210)
$Button5Tab1.Size = New-Object System.Drawing.Size(120, 20)
$Button5Tab1.Text = "Extraction XML"
# Boutton 6 Export-All-Xml Tab 1 GPO
$Button6Tab1 = New-Object System.Windows.Forms.Button
$Button6Tab1.location = New-Object System.Drawing.Size(185, 360)
$Button6Tab1.Size = New-Object System.Drawing.Size(120, 20)
$Button6Tab1.Text = "Export-All-Xml"
# Boutton 7 Export-All-Html Tab 1 GPO
$Button7Tab1 = New-Object System.Windows.Forms.Button
$Button7Tab1.location = New-Object System.Drawing.Size(185, 320)
$Button7Tab1.Size = New-Object System.Drawing.Size(120, 20)
$Button7Tab1.Text = "Export-All-Html"
# Label 1 dans Tab 1 (Information : Calcul du nombre de GPO) Tab 1 GPO
$Label1Tab1 = New-Object System.Windows.Forms.Label
$Label1Tab1.Location = New-Object System.Drawing.Size(250, 70)
$Label1Tab1.Size = New-Object System.Drawing.Size(60, 20)
$Label1Tab1.TextAlign = "MiddleCenter"
# Label 2 (Total de GPO :) dans Tab 1 GPO
$Label2Tab1 = New-Object System.Windows.Forms.Label
$Label2Tab1.Location = New-Object System.Drawing.Size(180, 70)
$Label2Tab1.Size = New-Object System.Drawing.Size(60, 20)
$Label2Tab1.TextAlign = "MiddleCenter"
$Label2Tab1.Text = "Total :"
# Label 3 (Id) dans Tab 1 GPO
$Label3Tab1 = New-Object System.Windows.Forms.Label
$Label3Tab1.Location = New-Object System.Drawing.Size(15, 100)
$Label3Tab1.Size = New-Object System.Drawing.Size(80, 30)
$Label3Tab1.TextAlign = "MiddleCenter"
$Label3Tab1.Text = "Id :"
# Label 4 (Domaine) dans Tab 1 GPO
$Label4Tab1 = New-Object System.Windows.Forms.Label
$Label4Tab1.Location = New-Object System.Drawing.Size(110, 100)
$Label4Tab1.Size = New-Object System.Drawing.Size(80, 30)
$Label4Tab1.TextAlign = "MiddleCenter"
$Label4Tab1.Text = "Domaine :"
# Label 5 (Date de creation) dans Tab 1 GPO
$Label5Tab1 = New-Object System.Windows.Forms.Label
$Label5Tab1.Location = New-Object System.Drawing.Size(205, 100)
$Label5Tab1.Size = New-Object System.Drawing.Size(80, 30)
$Label5Tab1.TextAlign = "MiddleCenter"
$Label5Tab1.Text = "Date de creation :"
# Label 6 (Date de modification) dans Tab 1 GPO
$Label6Tab1 = New-Object System.Windows.Forms.Label
$Label6Tab1.Location = New-Object System.Drawing.Size(300, 100)
$Label6Tab1.Size = New-Object System.Drawing.Size(80, 30)
$Label6Tab1.TextAlign = "MiddleCenter"
$Label6Tab1.Text = "Date de modification :"
# Label 7 (Statut) dans Tab 1 GPO
$Label7Tab1 = New-Object System.Windows.Forms.Label
$Label7Tab1.Location = New-Object System.Drawing.Size(395, 100)
$Label7Tab1.Size = New-Object System.Drawing.Size(80, 30)
$Label7Tab1.TextAlign = "MiddleCenter"
$Label7Tab1.Text = "Statut :"
# TextBox 1 (InfoId) dans Tab 1 GPO
$TextBox1Tab1 = New-Object System.Windows.Forms.TextBox
$TextBox1Tab1.Location = New-Object System.Drawing.Size(15, 140)
$TextBox1Tab1.Size = New-Object System.Drawing.Size(80, 30)
# TexBox 2 (Info Domaine) dans Tab 1 GPO
$TextBox2Tab1 = New-Object System.Windows.Forms.TextBox
$TextBox2Tab1.Location = New-Object System.Drawing.Size(110, 140)
$TextBox2Tab1.Size = New-Object System.Drawing.Size(80, 30)
# TexBox 3 (Info Date de creation) dans Tab 1 GPO
$TextBox3Tab1 = New-Object System.Windows.Forms.TextBox
$TextBox3Tab1.Location = New-Object System.Drawing.Size(205, 140)
$TextBox3Tab1.Size = New-Object System.Drawing.Size(80, 30)
# TextBox 4 (Info Date de modification) dans Tab 1 GPO
$TextBox4Tab1 = New-Object System.Windows.Forms.TextBox
$TextBox4Tab1.Location = New-Object System.Drawing.Size(300, 140)
$TextBox4Tab1.Size = New-Object System.Drawing.Size(80, 30)
# TextBox 5 (Info Statut) dans Tab 1 GPO
$TextBox5Tab1 = New-Object System.Windows.Forms.TextBox
$TextBox5Tab1.Location = New-Object System.Drawing.Size(395, 140)
$TextBox5Tab1.Size = New-Object System.Drawing.Size(80, 30)
### Action Tab1 GPO ###
# Bouton 1 Analyse Tab 1 GPO
$Button1Tab1.Add_Click(
    {
        $Tab1Page.Controls.Remove($Button1Tab1)
        Update-Log "Lancement de l'analyse de toute la structure de GPO de l'AD : " -Color "LightBlue" -NoNewLine
        Update-Log $env:USERDOMAIN -Color "Gold"
        Update-Log "Patienter..." -Color "LightBlue"
        $listGPO = AllGPO | sort-Object -Property DisplayName
        $ComboBox1Tab1.DropDownStyle = "DropDownList"
        $ComboBox1Tab1.Items.AddRange($ListGPO.DisplayName)
        #$ComboBox1Tab1.SelectedIndex = 0
        $Tab1Page.Controls.Add($ComboBox1Tab1)
        $Label1Tab1.Text = $ListGPO.Count
        $Tab1Page.Controls.Add($Label1Tab1)
        $Tab1Page.Controls.Add($Label2Tab1)
        $Tab1Page.Controls.Remove($Button1Tab1)
        $Tab1Page.Controls.Add($Button2Tab1)
        $Tab1Page.Controls.Add($Button3Tab1)
        $Tab1Page.Controls.Add($Button6Tab1)
        $Tab1Page.Controls.Add($Button7Tab1)
        Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
    }    
)
# Boutton 2 Reset Tab 1 GPO
$Button2Tab1.Add_Click(
    {
        $Tab1Page.Controls.Remove($ComboBox1Tab1)
        $Tab1Page.Controls.Remove($Label1Tab1)
        $Tab1Page.Controls.Remove($Label2Tab1)
        $Tab1Page.Controls.Remove($Label3Tab1)
        $Tab1Page.Controls.Remove($Label4Tab1)
        $Tab1Page.Controls.Remove($Label5Tab1)
        $Tab1Page.Controls.Remove($Label6Tab1)
        $Tab1Page.Controls.Remove($Label7Tab1)
        $Tab1Page.Controls.Remove($TextBox1Tab1)
        $Tab1Page.Controls.Remove($TextBox2Tab1)
        $Tab1Page.Controls.Remove($TextBox3Tab1)
        $Tab1Page.Controls.Remove($TextBox4Tab1)
        $Tab1Page.Controls.Remove($TextBox5Tab1)
        $Tab1Page.Controls.Remove($Button2Tab1)
        $Tab1Page.Controls.Remove($Button3Tab1)
        $Tab1Page.Controls.Remove($Button4Tab1)
        $Tab1Page.Controls.Remove($Button5Tab1)
        $Tab1Page.Controls.Add($Button1Tab1)
    }
)
# Boutton 3 Information Tab 1 GPO
$Button3Tab1.Add_Click(
    {
        If ($ComboBox1Tab1.SelectedItem -eq $null) {
            Update-Log "Il faut selectionner une GPO." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            $Tab1Page.Controls.Add($Button4Tab1)
            $Tab1Page.Controls.Add($Button5Tab1)
            $Tab1Page.Controls.Add($Label3Tab1)
            $Tab1Page.Controls.Add($Label4Tab1)
            $Tab1Page.Controls.Add($Label5Tab1)
            $Tab1Page.Controls.Add($Label6Tab1)
            $Tab1Page.Controls.Add($Label7Tab1)
            Update-Log "Selection de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $ComboBox1Tab1.Text -Color "Gold"
            $infoGPO = Get-GPO -Name $ComboBox1Tab1.Text
            $TextBox1Tab1.Text = $infoGPO.Id
            $TextBox1Tab1.TextAlign = "Center"
            $Tab1Page.Controls.Add($TextBox1Tab1)
            $TextBox2Tab1.Text = $infoGPO.DomainName
            $TextBox2Tab1.TextAlign = "Center"
            $Tab1Page.Controls.Add($TextBox2Tab1)
            $TextBox3Tab1.Text = $infoGPO.CreationTime
            $TextBox3Tab1.TextAlign = "Center"
            $Tab1Page.Controls.Add($TextBox3Tab1)
            $TextBox4Tab1.Text = $infoGPO.ModificationTime
            $TextBox4Tab1.TextAlign = "Center"
            $Tab1Page.Controls.Add($TextBox4Tab1)
            $TextBox5Tab1.Text = $infoGPO.GpoStatus -replace "Settings"," "
            $TextBox5Tab1.TextAlign = "Center"
            $Tab1Page.Controls.Add($TextBox5Tab1)
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }
)
# Boutton 4 Extraction Unitaire Tab 1 GPO
$Button4Tab1.Add_Click(
    {
        If ($ComboBox1Tab1.SelectedItem -eq $null) {
            Update-Log "Selectionner une GPO." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            if (!(Test-Path -Path $Path\temp)) {
                New-Item -Path "$Path\temp" -ItemType Directory -Force
            }
            Update-Log "Lancement de l'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $ComboBox1Tab1.Text -Color "Gold"
            Update-Log "Patienter..." -Color "LightBlue"
            $Date = Get-Date -format "dd-MM-yyyy_HH-mm-ss"
            Get-GPOReport -Name $ComboBox1Tab1.Text -ReportType Html -Path "$Path\temp\$($ComboBox1Tab1.Text)_$Date.html"
            $ie = New-Object -ComObject InternetExplorer.Application
            $ie.Navigate("$Path\temp\$($ComboBox1Tab1.Text)_$Date.html")
            Update-Log "L'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $ComboBox1Tab1.Text -Color "Gold" -NoNewLine
            Update-Log " est realisee" -Color "LightBlue"
            Update-Log "Elle se trouve a cet emplacement : " -Color "LightBlue" -NoNewLine
            Update-Log "$Path\temp\$($ComboBox1Tab1.Text)_$Date.html" -Color "Gold"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }
)
# Boutton 5 Extraction Unitaire Tab 1 GPO
$Button5Tab1.Add_Click(
    {
        If ($ComboBox1Tab1.SelectedItem -eq $null) {
            Update-Log "Selectionner une GPO." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            if (!(Test-Path -Path $Path\temp)) {
                New-Item -Path "$Path\temp" -ItemType Directory -Force
            }
            Update-Log "Lancement de l'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $ComboBox1Tab1.Text -Color "Gold"
            Update-Log "Patienter..." -Color "LightBlue"
            $Date = Get-Date -format "dd-MM-yyyy_HH-mm-ss"
            Get-GPOReport -Name $ComboBox1Tab1.Text -ReportType Xml -Path "$Path\temp\$($ComboBox1Tab1.Text)_$Date.xml"
            $ie = New-Object -ComObject InternetExplorer.Application
            $ie.Navigate("$Path\temp\$($ComboBox1Tab1.Text)_$Date.xml")
            Update-Log "L'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $ComboBox1Tab1.Text -Color "Gold" -NoNewLine
            Update-Log " est realisee" -Color "LightBlue"
            Update-Log "Elle se trouve a cet emplacement : " -Color "LightBlue" -NoNewLine
            Update-Log "$Path\temp\$($ComboBox1Tab1.Text)_$Date.xml" -Color "Gold"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }
)
# Boutton 6 Extraction All GPO
$Button6Tab1.Add_Click(
    {
        Update-Log "Lancement de l'extraction Xml total des GPOs, il faut patienter... " -Color "LightBlue"
        $listGPO = AllGPO
        $Domain = (Get-ADDomainController).Domain
        if (!(Test-Path -Path $Path\$Domain\Xml)) {
            New-Item -Path "$Path\$Domain\Xml" -ItemType Directory -Force
        }
        Foreach ($gpo in $listGPO) {
            Update-Log "Lancement de l'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $gpo.DisplayName -Color "Gold"
            $Date = Get-Date -format "dd-MM-yyyy_HH-mm-ss"
            Get-GPOReport -Name $gpo.DisplayName -ReportType Xml -Path "$Path\$Domain\Xml\$($gpo.DisplayName)_$Date.xml"
            Update-Log "L'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $gpo.DisplayName -Color "Gold" -NoNewLine
            Update-Log " est realisee" -Color "LightBlue"
            Update-Log "Elle se trouve a cet emplacement : " -Color "LightBlue" -NoNewLine
            Update-Log "$Path\$Domain\Xml\$($gpo.DisplayName)_$Date.xml" -Color "Gold"
        }
        Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
    }
)
# Boutton 7 Extraction All GPO
$Button7Tab1.Add_Click(
    {
        Update-Log "Lancement de l'extraction Html total des GPOs, il faut patienter... " -Color "LightBlue"
        $listGPO = AllGPO
        $Domain = (Get-ADDomainController).Domain
        if (!(Test-Path -Path $Path\$Domain\Html)) {
            New-Item -Path "$Path\$Domain\Html" -ItemType Directory -Force
        }
        Foreach ($gpo in $listGPO) {
            Update-Log "Lancement de l'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $gpo.DisplayName -Color "Gold"
            $Date = Get-Date -format "dd-MM-yyyy_HH-mm-ss"
            Get-GPOReport -Name $gpo.DisplayName -ReportType Html -Path "$Path\$Domain\Html\$($gpo.DisplayName)_$Date.html"
            Update-Log "L'extraction unitaire de la GPO : " -Color "LightBlue" -NoNewLine
            Update-Log $gpo.DisplayName -Color "Gold" -NoNewLine
            Update-Log " est realisee" -Color "LightBlue"
            Update-Log "Elle se trouve a cet emplacement : " -Color "LightBlue" -NoNewLine
            Update-Log "$Path\$Domain\Html\$($gpo.DisplayName)_$Date.html" -Color "Gold"
        }
        Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
    }
)

### Tab 2 Groupe ###
# Creation de l'onglet 2 : Groupe
$Tab2Page = New-Object System.Windows.Forms.TabPage
$Tab2Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab2Page.UseVisualStyleBackColor = $true
$Tab2Page.Text = "Groupe"
$Tab2Page.AutoScroll = $true
$TabControl.Controls.Add($Tab2Page)
# Boutton 1 Analyse Tab 2 Groupe
$Button1Tab2 = New-Object System.Windows.Forms.Button
$Button1Tab2.location = New-Object System.Drawing.Size(215, 10)
$Button1Tab2.Size = New-Object System.Drawing.Size(60, 20)
$Button1Tab2.Text = "Analyse"
$Tab2Page.Controls.Add($Button1Tab2)
# Boutton 2 Reset Tab 2 Groupe
$Button2Tab2 = New-Object System.Windows.Forms.Button
$Button2Tab2.location = New-Object System.Drawing.Size(215, 420)
$Button2Tab2.Size = New-Object System.Drawing.Size(60, 20)
$Button2Tab2.Text = "Reset"
# Boutton 3 Analyse Tab 2 Groupe
$Button3Tab2 = New-Object System.Windows.Forms.Button
$Button3Tab2.location = New-Object System.Drawing.Size(215, 120)
$Button3Tab2.Size = New-Object System.Drawing.Size(60, 20)
$Button3Tab2.Text = "Analyse"
$Tab2Page.Controls.Add($Button3Tab2)
# ComboBox 1 Liste de Groupe Tab 2 Groupe
$ComboBox1Tab2 = New-Object System.Windows.Forms.ComboBox
$ComboBox1Tab2.location = New-Object System.Drawing.Size(10, 10)
$ComboBox1Tab2.Size = New-Object System.Drawing.Size(460, 20)
# Label 1 (Information : Calcul du nombre de Groupe) dans Tab 2 Groupe
$Label1Tab2 = New-Object System.Windows.Forms.Label
$Label1Tab2.Location = New-Object System.Drawing.Size(250, 70)
$Label1Tab2.Size = New-Object System.Drawing.Size(60, 20)
$Label1Tab2.TextAlign = "MiddleCenter"
# TextBox 1 (Groupe) dans Tab 2 Groupe
$TextBox1Tab2 = New-Object System.Windows.Forms.TextBox
$TextBox1Tab2.Location = New-Object System.Drawing.Size(145, 90)
$TextBox1Tab2.Size = New-Object System.Drawing.Size(200, 20)
$Tab2Page.Controls.Add($TextBox1Tab2)
### Action Tab 2 Groupe ###
# Bouton 1 Analyse Tab 2 Groupe
$Button1Tab2.Add_Click(
    {
        $Tab2Page.Controls.Remove($Button1Tab2)
        Update-Log "Lancement de l'analyse de toute la structure des groupes de l'AD : " -Color "LightBlue" -NoNewLine
        Update-Log $env:USERDOMAIN -Color "Gold"
        Update-Log "Patienter..." -Color "LightBlue"
        $listGroup = AllGroup
        $listGroup = $listGroup | sort-Object -Property Name
        $ComboBox1Tab2.DropDownStyle = "DropDownList"
        $ComboBox1Tab2.Items.AddRange($ListGroup.Name)
        $Tab2Page.Controls.Add($ComboBox1Tab2)
        $Label1Tab2.Text = $ListGroup.Count
        $Tab2Page.Controls.Add($Label1Tab2)
        $Tab2Page.Controls.Add($Button2Tab2)
        Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
    }    
)
# Bouton 3 Analyse Tab 2 Groupe
$Button3Tab2.Add_Click(
    {
        If ($TextBox1Tab2.Text) {
            Try {
                Update-Log "Lancement de l'analyse du groupe : " -Color "LightBlue" -NoNewLine
                Update-Log $TextBox1Tab2.Text -Color "Gold"
                Update-Log "Patienter..." -Color "LightBlue"
                $group = GroupUnitaire -Group $TextBox1Tab2.Text
            } Catch {
                Update-Log $TextBox1Tab2.Text -NoNewLine -Color "Gold"
                Update-Log " n'existe pas dans l'AD " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color Gold
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
            If ($group) {
                Update-Log $group.Name -NoNewLine -Color "Gold"
                Update-Log " a ete trouve dans l'AD : " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color "Gold"
                Update-Log-Obj $group
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
        } Else {
            Update-Log "Il faut renseigner un groupe." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }    
)

### Tab 3 Utilisateur ###
# Creation de l'onglet 3 : Utilisateur
$Tab3Page = New-Object System.Windows.Forms.TabPage
$Tab3Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab3Page.UseVisualStyleBackColor = $true
$Tab3Page.Text = "Utilisateur"
$Tab3Page.AutoScroll = $true
$TabControl.Controls.Add($Tab3Page)
# Boutton 1 Analyse (idRH) Tab 3 Utilisateur
$Button1Tab3 = New-Object System.Windows.Forms.Button
$Button1Tab3.location = New-Object System.Drawing.Size(420, 10)
$Button1Tab3.Size = New-Object System.Drawing.Size(60, 20)
$Button1Tab3.Text = "Analyse"
$Tab3Page.Controls.Add($Button1Tab3)
# Boutton 2 Analyse (Nom Prenom) Tab 3 Utilisateur
$Button2Tab3 = New-Object System.Windows.Forms.Button
$Button2Tab3.location = New-Object System.Drawing.Size(420, 40)
$Button2Tab3.Size = New-Object System.Drawing.Size(60, 20)
$Button2Tab3.Text = "Analyse"
$Tab3Page.Controls.Add($Button2Tab3)
# Label 1 (idRH) dans Tab 3 Utilisateur
$Label1Tab3 = New-Object System.Windows.Forms.Label
$Label1Tab3.Location = New-Object System.Drawing.Size(10, 10)
$Label1Tab3.Size = New-Object System.Drawing.Size(80, 20)
$Label1Tab3.TextAlign = "MiddleCenter"
$Label1Tab3.Text = "idRH :"
$Tab3Page.Controls.Add($Label1Tab3)
# Label 1 (Nom Prenom) dans Tab 3 Utilisateur
$Label2Tab3 = New-Object System.Windows.Forms.Label
$Label2Tab3.Location = New-Object System.Drawing.Size(10, 40)
$Label2Tab3.Size = New-Object System.Drawing.Size(80, 20)
$Label2Tab3.TextAlign = "MiddleCenter"
$Label2Tab3.Text = "Prenom Nom :"
$Tab3Page.Controls.Add($Label2Tab3)
# TextBox 1 (idRH) dans Tab 3 Utilisateur
$TextBox1Tab3 = New-Object System.Windows.Forms.TextBox
$TextBox1Tab3.Location = New-Object System.Drawing.Size(215, 10)
$TextBox1Tab3.Size = New-Object System.Drawing.Size(80, 20)
$Tab3Page.Controls.Add($TextBox1Tab3)
# TextBox 2 (Prenom) dans Tab 3 Utilisateur
$TextBox2Tab3 = New-Object System.Windows.Forms.TextBox
$TextBox2Tab3.Location = New-Object System.Drawing.Size(150, 40)
$TextBox2Tab3.Size = New-Object System.Drawing.Size(80, 20)
$Tab3Page.Controls.Add($TextBox2Tab3)
# TextBox 2 (Nom) dans Tab 3 Utilisateur
$TextBox3Tab3 = New-Object System.Windows.Forms.TextBox
$TextBox3Tab3.Location = New-Object System.Drawing.Size(280, 40)
$TextBox3Tab3.Size = New-Object System.Drawing.Size(80, 20)
$Tab3Page.Controls.Add($TextBox3Tab3)
### Action Tab 3 Utilisateur ###
# Bouton 1 Analyse Tab 3 Utilisateur
$Button1Tab3.Add_Click(
    {        
        If ($TextBox1Tab3.Text) {
            Try {
                Update-Log "Lancement de l'analyse de l'utilisateur : " -Color "LightBlue" -NoNewLine
                Update-Log $TextBox1Tab3.Text -Color "Gold"
                Update-Log "Patienter..." -Color "LightBlue"
                $user = UserAD -User $TextBox1Tab3.Text
            } Catch {
                Update-Log $TextBox1Tab3.Text -NoNewLine -Color "Gold"
                Update-Log " n'existe pas dans l'AD " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color Gold
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
            If ($user) {
                Update-Log $user.Name -NoNewLine -Color "Gold"
                Update-Log " - " -NoNewLine -Color "LightBlue"
                Update-Log $user.UserPrincipalName -NoNewLine -Color "Gold"
                Update-Log " a ete trouve(e) dans l'AD : " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color "Gold"
                Update-Log-Obj $user
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
        } Else {
            Update-Log "Il faut renseigner un utilisateur." -Color "Red"
            Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }    
)
# Bouton 2 Analyse Tab 3 Utilisateur
$Button2Tab3.Add_Click(
    {
        If ($TextBox2Tab3.Text -and $TextBox3Tab3.Text) {
            Update-Log "Vous avez renseigne le prenom : " -Color "LightBlue" -NoNewLine
            Update-Log $TextBox2Tab3.Text -Color "Gold"
            Update-Log "Vous avez renseigne le nom : " -Color "LightBlue" -NoNewLine
            Update-Log $TextBox3Tab3.Text -Color "Gold"
            Update-Log "Patienter..." -Color "LightBlue"
            Try {
                $user = PrenomNomAD -Prenom $TextBox2Tab3.Text -Nom $TextBox3Tab3.Text
            } Catch {
                Update-Log " Le prenom : " -NoNewLine -Color "LightBlue"
                Update-Log $TextBox2Tab3.Text -NoNewLine -Color "Gold"
                Update-Log " et le nom : " -NoNewLine -Color "LightBlue"
                Update-Log $TextBox3Tab3.Text -NoNewLine -Color "Gold"
                Update-Log " n'existe pas dans l'AD : " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color Gold
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
            If ($user) {
                Update-Log "Nombre de compte(s) trouve(s) : " -NoNewLine -Color "LightBlue"
                Update-Log $user.Count -Color "Gold"
                Foreach ($obj in $user) {
                    Update-Log $obj.Name -NoNewLine -Color "Gold"
                    Update-Log " - " -NoNewLine -Color "LightBlue"
                    Update-Log $obj.UserPrincipalName -NoNewLine -Color "Gold"
                    Update-Log " a ete trouve(e) dans l'AD : " -NoNewLine -Color "LightBlue"
                    Update-Log $env:USERDOMAIN -Color "Gold"
                    Update-Log-Obj $obj
                    Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
                }
            } Else {
                Update-Log " Le prenom : " -NoNewLine -Color "LightBlue"
                Update-Log $TextBox2Tab3.Text -NoNewLine -Color "Gold"
                Update-Log " et le nom : " -NoNewLine -Color "LightBlue"
                Update-Log $TextBox3Tab3.Text -NoNewLine -Color "Gold"
                Update-Log " n'existe pas dans l'AD : " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color Gold
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
        } Else {
            If ($TextBox2Tab3.Text) {
                Update-Log "Vous avez renseigne le prenom : " -Color "LightBlue" -NoNewLine
                Update-Log $TextBox2Tab3.Text -Color "Gold"
            } Else {
                Update-Log "Il faut renseigner un prenom." -Color "Red"
                Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
            If ($TextBox3Tab3.Text) {
                Update-Log "Vous avez renseigne le nom : " -Color "LightBlue" -NoNewLine
                Update-Log $TextBox3Tab3.Text -Color "Gold"
            } Else {
                Update-Log "Il faut renseigner un nom." -Color "Red"
                Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
        }
    }
)

### Tab 4 Ordinateur ###
# Creation de l'onglet 4 Ordinateur
$Tab4Page = New-Object System.Windows.Forms.TabPage
$Tab4Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab4Page.UseVisualStyleBackColor = $true
$Tab4Page.Text = "Ordinateur"
$Tab4Page.AutoScroll = $true
$TabControl.Controls.Add($Tab4Page)
# Boutton 1 Analyse Tab 4
$Button1Tab4 = New-Object System.Windows.Forms.Button
$Button1Tab4.location = New-Object System.Drawing.Size(250, 10)
$Button1Tab4.Size = New-Object System.Drawing.Size(60, 20)
$Button1Tab4.Text = "Analyse"
$Tab4Page.Controls.Add($Button1Tab4)
# TextBox 1 (Ordinateur) dans Tab 4 Ordinateur
$TextBox1Tab4 = New-Object System.Windows.Forms.TextBox
$TextBox1Tab4.Location = New-Object System.Drawing.Size(180, 10)
$TextBox1Tab4.Size = New-Object System.Drawing.Size(60, 20)
$Tab4Page.Controls.Add($TextBox1Tab4)
### Action Tab 4 Ordinateur ###
# Bouton 1 Analyse Tab 4 Ordinateur
$Button1Tab4.Add_Click(
    {        
        If ($TextBox1Tab4.Text) {
            Try {
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
                Update-Log "Lancement de l'analyse de l'ordinateur : " -NoNewLine -Color "LightBlue"
                Update-Log $TextBox1Tab4.Text -Color "Gold"
                Update-Log "Patienter..." -Color "LightBlue"
                $Computer = ComputerAD -Computer $TextBox1Tab4.Text
            } Catch {
                Update-Log $TextBox1Tab4.Text -NoNewLine -Color "Gold"
                Update-Log " n'existe pas dans l'AD " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color Gold
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
            If ($Computer) {
                Update-Log $Computer.Name -NoNewLine -Color "Gold"
                Update-Log " a ete trouve dans l'AD : " -NoNewLine -Color "LightBlue"
                Update-Log $env:USERDOMAIN -Color "Gold"
                Update-Log-Obj $Computer
                Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
            }
        } Else {
            Update-Log "Il faut renseigner un nom d'ordinateur." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }        
    }    
)

### Tab 5 InitProxy ###
# Creation de l'onglet 5 InitProxy
$Tab5Page = New-Object System.Windows.Forms.TabPage
$Tab5Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab5Page.UseVisualStyleBackColor = $true
$Tab5Page.Text = "InitProxy"
$Tab5Page.AutoScroll = $true
$TabControl.Controls.Add($Tab5Page)
# Boutton 1 Analyse Tab 5 InitProxy
$Button1Tab5 = New-Object System.Windows.Forms.Button
$Button1Tab5.location = New-Object System.Drawing.Size(215, 10)
$Button1Tab5.Size = New-Object System.Drawing.Size(60, 20)
$Button1Tab5.Text = "Analyse"
$Tab5Page.Controls.Add($Button1Tab5)
# Boutton 2 Analyse Tab 5 InitProxy
$Button2Tab5 = New-Object System.Windows.Forms.Button
$Button2Tab5.location = New-Object System.Drawing.Size(195, 170)
$Button2Tab5.Size = New-Object System.Drawing.Size(100, 20)
$Button2Tab5.Text = "Set Registry"
# Boutton 3 Analyse Tab 5 InitProxy
$Button3Tab5 = New-Object System.Windows.Forms.Button
$Button3Tab5.location = New-Object System.Drawing.Size(195, 200)
$Button3Tab5.Size = New-Object System.Drawing.Size(100, 20)
$Button3Tab5.Text = "InitProxy ?"
# Label 1 (AutoconfigUrl) dans Tab 5 InitProxy
$Label1Tab5 = New-Object System.Windows.Forms.Label
$Label1Tab5.Location = New-Object System.Drawing.Size(10, 40)
$Label1Tab5.Size = New-Object System.Drawing.Size(90, 20)
$Label1Tab5.TextAlign = "MiddleCenter"
$Label1Tab5.Text = "AutoconfigUrl :"
# Label 2 (AutoDetect) dans Tab 5 InitProxy
$Label2Tab5 = New-Object System.Windows.Forms.Label
$Label2Tab5.Location = New-Object System.Drawing.Size(10, 70)
$Label2Tab5.Size = New-Object System.Drawing.Size(90, 20)
$Label2Tab5.TextAlign = "MiddleCenter"
$Label2Tab5.Text = "AutoDetect :"
# Label 3 (ProxyEnable) dans Tab 5 InitProxy
$Label3Tab5 = New-Object System.Windows.Forms.Label
$Label3Tab5.Location = New-Object System.Drawing.Size(140, 70)
$Label3Tab5.Size = New-Object System.Drawing.Size(90, 20)
$Label3Tab5.TextAlign = "MiddleCenter"
$Label3Tab5.Text = "ProxyEnable :"
# Label 4 (ProxyOverride) dans Tab 5 InitProxy
$Label4Tab5 = New-Object System.Windows.Forms.Label
$Label4Tab5.Location = New-Object System.Drawing.Size(10, 100)
$Label4Tab5.Size = New-Object System.Drawing.Size(90, 20)
$Label4Tab5.TextAlign = "MiddleCenter"
$Label4Tab5.Text = "ProxyOverride :"
# Label 5 (ProxyServer) dans Tab 5 InitProxy
$Label5Tab5 = New-Object System.Windows.Forms.Label
$Label5Tab5.Location = New-Object System.Drawing.Size(10, 130)
$Label5Tab5.Size = New-Object System.Drawing.Size(90, 20)
$Label5Tab5.TextAlign = "MiddleCenter"
$Label5Tab5.Text = "ProxyServer :"
# TextBox 1 (AutoConfigUrl) dans Tab 5 InitProxy
$TextBox1Tab5 = New-Object System.Windows.Forms.TextBox
$TextBox1Tab5.Location = New-Object System.Drawing.Size(100, 40)
$TextBox1Tab5.Size = New-Object System.Drawing.Size(370, 20)
# TextBox 2 (AutoDetect) dans Tab 5 InitProxy
$TextBox2Tab5 = New-Object System.Windows.Forms.TextBox
$TextBox2Tab5.Location = New-Object System.Drawing.Size(100, 70)
$TextBox2Tab5.Size = New-Object System.Drawing.Size(30, 20)
# TextBox 3 (ProxyEnable) dans Tab 5 InitProxy
$TextBox3Tab5 = New-Object System.Windows.Forms.TextBox
$TextBox3Tab5.Location = New-Object System.Drawing.Size(240, 70)
$TextBox3Tab5.Size = New-Object System.Drawing.Size(30, 20)
# TextBox 4 (ProxyOverride) dans Tab 5 InitProxy
$TextBox4Tab5 = New-Object System.Windows.Forms.TextBox
$TextBox4Tab5.Location = New-Object System.Drawing.Size(100, 100)
$TextBox4Tab5.Size = New-Object System.Drawing.Size(370, 20)
# TextBox 5 (ProxyServer) dans Tab 5 InitProxy
$TextBox5Tab5 = New-Object System.Windows.Forms.TextBox
$TextBox5Tab5.Location = New-Object System.Drawing.Size(100, 130)
$TextBox5Tab5.Size = New-Object System.Drawing.Size(370, 20)
### Action Tab 5 InitProxy ###
# Boutton 1 Analyse Tab 5 (InitProxy)
$Button1Tab5.Add_Click(
    {
        Update-Log "Lancement de l'analyse de la registre Proxy : " -Color "LightBlue"
        Update-Log $ProxyRegistre -Color "Gold"
        Update-Log "Patienter..." -Color "LightBlue"
        $ProxyList = RegValue -Hive $ProxyRegistre
        If ($null -ne $ProxyList) {
            Update-Log-Obj -Object $ProxyList
            Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            Update-Log "Cle de registre absente" -Color "LightBlue"
            Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
        $Tab5Page.Controls.Add($Label1Tab5)
        $Tab5Page.Controls.Add($Label2Tab5)
        $Tab5Page.Controls.Add($Label3Tab5)
        $Tab5Page.Controls.Add($Label4Tab5)
        $Tab5Page.Controls.Add($Label5Tab5)
        $Tab5Page.Controls.Add($TextBox1Tab5)
        $Tab5Page.Controls.Add($TextBox2Tab5)
        $Tab5Page.Controls.Add($TextBox3Tab5)
        $Tab5Page.Controls.Add($TextBox4Tab5)
        $Tab5Page.Controls.Add($TextBox5Tab5)
        $Tab5Page.Controls.Add($Button2Tab5)        
    }
)
# Boutton 2 Set Registry Tab 5 InitProxy
$Button2Tab5.Add_Click(
    {
        If ($TextBox1Tab5.Text) {

        }  Else {
            Update-Log "Aucune valeur renseigne pour la registre : " -Color "LightBlue" -NoNewLine
            Update-Log "AutoconfigUrl" -Color "Gold" -NoNewLine
            Update-Log ", aucune modification ne sera effectuee" -Color "LightBlue"
        }
        If ($TextBox2Tab5.Text) {

        } Else {
            Update-Log "Aucune valeur renseigne pour la registre : " -Color "LightBlue" -NoNewLine
            Update-Log "AutoDetect" -Color "Gold" -NoNewLine
            Update-Log ", aucune modification ne sera effectuee" -Color "LightBlue"
        }
        If ($TextBox3Tab5.Text) {

        } Else {
            Update-Log "Aucune valeur renseigne pour la registre : " -Color "LightBlue" -NoNewLine
            Update-Log "ProxyEnable" -Color "Gold" -NoNewLine
            Update-Log ", aucune modification ne sera effectuee" -Color "LightBlue"
        }
        If ($TextBox4Tab5.Text) {

        } Else {
            Update-Log "Aucune valeur renseigne pour : " -Color "LightBlue" -NoNewLine
            Update-Log "ProxyOverride" -Color "Gold" -NoNewLine
            Update-Log ", aucune modification ne sera effectuee" -Color "LightBlue"
        }
        If ($TextBox5Tab5.Text) {

        } Else {
            Update-Log "Aucune valeur renseigne pour : " -Color "LightBlue" -NoNewLine
            Update-Log "ProxyServer" -Color "Gold" -NoNewLine
            Update-Log ", aucune modification ne sera effectuee" -Color "LightBlue"
        }
        Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        $Tab5Page.Controls.Add($Button3Tab5)
    }
)
# Boutton 3 InitProxy Tab 5 InitProxy
$Button3Tab5.Add_Click(
    {
        Set-LogoProxy
        Update-Log "Lancement de l'analyse de la registre Proxy : " -Color "LightBlue"
        Update-Log $ConnectionsRegistre -Color "Gold"
        Update-Log "Patienter..." -Color "LightBlue"
        $ConnectionsList = RegValue -Hive $ConnectionsRegistre
        If ($null -ne $ConnectionsList) {
            Update-Log-Obj -Object $ConnectionsList
            Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            Update-Log "Cle de registre absente" -Color "LightBlue"
            Update-Log " ------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
        If (Test-Path -path "$path\InitProxy.exe") {
            Update-Log "Lancement de l'outil InitProxy.exe." -Color "LightBlue"
            #Start-Process "$path\InitProxy.exe" -Wait
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        } Else {
            Update-Log "L'outil InitProxy.exe n'est pas disponible." -Color "Red"
            Update-Log " -------------------------------------------------------------------------------------------- " -Color "LightBlue"
        }
    }
)

If (RSAT) {
    $Form.Add_Shown({$Form.Activate()})
    $Form.ShowDialog() | Out-Null
} Else {
    $FormKO.Add_Shown({$FormKO.Activate()})
    $FormKO.ShowDialog() | Out-Null
}