<# 
.SYNOPSIS
  Menú completo de administración de Active Directory (CRUD para OUs, Usuarios y Grupos)

.DESCRIPTION
  Interfaz interactiva para gestionar Active Directory con manejo robusto de errores.
  Solo se cierra cuando el usuario pulsa 0.

.NOTES
  Versión: 2.0 - Menú completo con autodetección y gestión integral
#>

Import-Module ActiveDirectory -ErrorAction Stop

# =======================================================
# ===== VARIABLES GLOBALES =====
# =======================================================

$script:Server = $null
$script:BaseDN = $null
$script:DomainUPN = $null

# =======================================================
# ===== FUNCIONES DE AYUDA (HELPERS) =====
# =======================================================

function Pause-Script { 
    Write-Host ""
    Read-Host "Presiona ENTER para continuar..." | Out-Null 
}

function Ask-Input {
    param([string]$Prompt)
    do { 
        $value = Read-Host $Prompt 
    } while ([string]::IsNullOrWhiteSpace($value))
    return $value
}

function Ask-YesNo {
    param([string]$Prompt)
    $response = Read-Host "$Prompt [s/n]"
    return ($response -match '^(s|si|sí|S|SI|SÍ)$')
}

function Show-Header {
    Clear-Host
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "====== ADMINISTRADOR DE ACTIVE DIRECTORY ====" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "Usuario: $(whoami)" -ForegroundColor DarkGray
    Write-Host "Host: $(hostname)" -ForegroundColor DarkGray
    Write-Host "Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor DarkGray
    Write-Host "=============================================" -ForegroundColor Cyan
    
    if ($script:Server) {
        Write-Host ""
        Write-Host "DOMINIO ACTIVO: $($script:Server)" -ForegroundColor Yellow
        Write-Host "Base DN: $($script:BaseDN)" -ForegroundColor DarkCyan
        Write-Host "---------------------------------------------" -ForegroundColor DarkCyan
    }
    Write-Host ""
}

function Normalize-SamAccountName {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    
    $formD = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $formD.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    $normalized = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
    return ($normalized -replace '[^a-zA-Z0-9\.\-_]', '' -replace '\s+', '.' ).ToLower().Trim('.')
}

function Show-ErrorMessage {
    param([string]$Message)
    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    Write-Host ""
}

function Show-SuccessMessage {
    param([string]$Message)
    Write-Host ""
    Write-Host "ÉXITO: $Message" -ForegroundColor Green
    Write-Host ""
}

function Show-WarningMessage {
    param([string]$Message)
    Write-Host ""
    Write-Host "ADVERTENCIA: $Message" -ForegroundColor Yellow
    Write-Host ""
}

# =======================================================
# ===== SELECCIÓN Y CONEXIÓN DE DOMINIO =====
# =======================================================

function Select-Domain {
    while ($true) {
        try {
            Show-Header
            Write-Host "--- SELECCIÓN DE DOMINIO ---" -ForegroundColor Yellow
            Write-Host ""
            
            # Autodetección
            Write-Host "Detectando dominios disponibles..." -ForegroundColor Cyan
            try {
                $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
                $domains = @($forest.Domains | ForEach-Object { $_.Name })
                
                if ($domains.Count -gt 0) {
                    Write-Host "Dominios encontrados:" -ForegroundColor Green
                    $i = 1
                    foreach ($dom in $domains) {
                        Write-Host "  $i) $dom" -ForegroundColor White
                        $i++
                    }
                    Write-Host ""
                    
                    $selection = Read-Host "Selecciona un dominio (número) o introduce uno manualmente (FQDN)"
                    
                    if ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $domains.Count) {
                            $selectedDomain = $domains[$idx]
                        } else {
                            Show-ErrorMessage "Número fuera de rango."
                            Pause-Script
                            continue
                        }
                    } else {
                        $selectedDomain = $selection.Trim()
                    }
                } else {
                    Show-WarningMessage "No se detectaron dominios automáticamente."
                    $selectedDomain = Ask-Input "Introduce el FQDN del dominio (ej: scs.local)"
                }
            } catch {
                Show-WarningMessage "No se pudo detectar el bosque: $($_.Exception.Message)"
                $selectedDomain = Ask-Input "Introduce el FQDN del dominio (ej: scs.local)"
            }
            
            # Intentar conectar
            Write-Host ""
            Write-Host "Conectando a $selectedDomain..." -ForegroundColor Cyan
            
            try {
                $adDomain = Get-ADDomain -Server $selectedDomain -ErrorAction Stop
                $script:Server = $adDomain.DNSRoot
                $script:BaseDN = $adDomain.DistinguishedName
                $script:DomainUPN = $adDomain.DNSRoot
                
                Show-SuccessMessage "Conectado exitosamente a: $($script:Server)"
                Pause-Script
                return
            } catch {
                Show-ErrorMessage "No se pudo conectar al dominio: $($_.Exception.Message)"
                Pause-Script
            }
            
        } catch {
            Show-ErrorMessage "Error inesperado: $($_.Exception.Message)"
            Pause-Script
        }
    }
}

# =======================================================
# ===== GESTIÓN DE OUs =====
# =======================================================

function Menu-OU {
    while ($true) {
        try {
            Show-Header
            Write-Host "--- GESTIÓN DE UNIDADES ORGANIZATIVAS (OU) ---" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "1. Buscar/Ver OUs" -ForegroundColor White
            Write-Host "2. Crear nueva OU" -ForegroundColor White
            Write-Host "3. Renombrar OU" -ForegroundColor White
            Write-Host "4. Borrar OU" -ForegroundColor White
            Write-Host "0. Volver al menú principal" -ForegroundColor White
            Write-Host ""
            
            $option = Read-Host "Selecciona una opción"
            
            switch ($option) {
                '1' { OU-Search }
                '2' { OU-Create }
                '3' { OU-Rename }
                '4' { OU-Delete }
                '0' { return }
                default { 
                    Show-ErrorMessage "Opción no válida. Por favor, selecciona una opción del menú."
                    Pause-Script
                }
            }
        } catch {
            Show-ErrorMessage "Error en menú OU: $($_.Exception.Message)"
            Pause-Script
        }
    }
}

function OU-Search {
    try {
        Show-Header
        Write-Host "--- BUSCAR/VER OUs ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar OU (deja vacío para ver todas)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        Write-Host ""
        Write-Host "Buscando OUs..." -ForegroundColor Cyan
        
        $ous = Get-ADOrganizationalUnit -Filter "Name -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -ErrorAction Stop | Sort-Object Name
        
        if ($ous) {
            Write-Host ""
            Write-Host "OUs encontradas:" -ForegroundColor Green
            Write-Host ""
            $ous | Format-Table Name, DistinguishedName -AutoSize
        } else {
            Show-WarningMessage "No se encontraron OUs con el término '$searchTerm'."
        }
        
    } catch {
        Show-ErrorMessage "No se pudieron listar las OUs: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function OU-Create {
    try {
        Show-Header
        Write-Host "--- CREAR NUEVA OU ---" -ForegroundColor Cyan
        Write-Host ""
        
        $ouName = Ask-Input "Nombre de la nueva OU"
        
        Write-Host ""
        Write-Host "¿Dónde crear la OU?" -ForegroundColor Yellow
        Write-Host "1) En la raíz del dominio ($($script:BaseDN))" -ForegroundColor White
        Write-Host "2) Dentro de otra OU existente" -ForegroundColor White
        
        $parentOption = Read-Host "Opción"
        
        if ($parentOption -eq '2') {
            $parentPath = Select-OU-Interactive "crear '$ouName' dentro de"
            if (-not $parentPath) {
                Show-WarningMessage "Operación cancelada."
                Pause-Script
                return
            }
        } else {
            $parentPath = $script:BaseDN
        }
        
        Write-Host ""
        Write-Host "Creando OU '$ouName' en $parentPath..." -ForegroundColor Cyan
        
        New-ADOrganizationalUnit -Name $ouName -Path $parentPath -Server $script:Server -ProtectedFromAccidentalDeletion $true -ErrorAction Stop
        
        Show-SuccessMessage "OU '$ouName' creada exitosamente en $parentPath"
        
    } catch {
        Show-ErrorMessage "No se pudo crear la OU: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function OU-Rename {
    try {
        $ouDN = Select-OU-Interactive "renombrar"
        if (-not $ouDN) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- RENOMBRAR OU ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "OU seleccionada: $ouDN" -ForegroundColor DarkGreen
        Write-Host ""
        
        $newName = Ask-Input "Nuevo nombre para la OU"
        
        if (Ask-YesNo "¿Confirmar renombrado de la OU a '$newName'?") {
            Write-Host ""
            Write-Host "Renombrando OU..." -ForegroundColor Cyan
            
            Rename-ADObject -Identity $ouDN -NewName $newName -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "OU renombrada exitosamente a '$newName'"
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo renombrar la OU: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function OU-Delete {
    try {
        $ouDN = Select-OU-Interactive "BORRAR"
        if (-not $ouDN) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- BORRAR OU ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "OU seleccionada: $ouDN" -ForegroundColor Red
        Write-Host ""
        Write-Host "ADVERTENCIA: Esta operación borrará la OU y TODO su contenido de forma RECURSIVA." -ForegroundColor Red
        Write-Host ""
        
        if (Ask-YesNo "¿Estás SEGURO de que quieres BORRAR esta OU y todo su contenido?") {
            Write-Host ""
            Write-Host "Desactivando protección y borrando OU..." -ForegroundColor Cyan
            
            Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -Server $script:Server -ErrorAction Stop
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "OU borrada exitosamente."
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo borrar la OU: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Select-OU-Interactive {
    param([string]$Action)
    
    try {
        Show-Header
        Write-Host "--- SELECCIÓN DE OU PARA $Action ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar OU (deja vacío para ver todas)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        $ous = Get-ADOrganizationalUnit -Filter "Name -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -ErrorAction Stop | Sort-Object Name
        
        if (-not $ous) {
            Show-WarningMessage "No se encontraron OUs."
            return $null
        }
        
        Write-Host ""
        Write-Host "OUs encontradas:" -ForegroundColor Green
        $i = 1
        foreach ($ou in $ous) {
            Write-Host "$i) $($ou.Name) - $($ou.DistinguishedName)" -ForegroundColor White
            $i++
        }
        
        Write-Host ""
        $selection = Read-Host "Selecciona el número de la OU (0 para cancelar)"
        
        if ($selection -eq '0') {
            return $null
        }
        
        $idx = [int]$selection - 1
        if ($idx -ge 0 -and $idx -lt $ous.Count) {
            return $ous[$idx].DistinguishedName
        } else {
            Show-ErrorMessage "Selección no válida."
            return $null
        }
        
    } catch {
        Show-ErrorMessage "Error al seleccionar OU: $($_.Exception.Message)"
        return $null
    }
}

# =======================================================
# ===== GESTIÓN DE GRUPOS =====
# =======================================================

function Menu-Groups {
    while ($true) {
        try {
            Show-Header
            Write-Host "--- GESTIÓN DE GRUPOS ---" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "1. Buscar/Ver Grupos" -ForegroundColor White
            Write-Host "2. Crear nuevo Grupo" -ForegroundColor White
            Write-Host "3. Editar Grupo (Nombre/Ámbito)" -ForegroundColor White
            Write-Host "4. Borrar Grupo" -ForegroundColor White
            Write-Host "5. Añadir miembros a Grupo" -ForegroundColor White
            Write-Host "6. Quitar miembros de Grupo" -ForegroundColor White
            Write-Host "7. Ver miembros de Grupo" -ForegroundColor White
            Write-Host "0. Volver al menú principal" -ForegroundColor White
            Write-Host ""
            
            $option = Read-Host "Selecciona una opción"
            
            switch ($option) {
                '1' { Group-Search }
                '2' { Group-Create }
                '3' { Group-Edit }
                '4' { Group-Delete }
                '5' { Group-AddMembers }
                '6' { Group-RemoveMembers }
                '7' { Group-ViewMembers }
                '0' { return }
                default { 
                    Show-ErrorMessage "Opción no válida. Por favor, selecciona una opción del menú."
                    Pause-Script
                }
            }
        } catch {
            Show-ErrorMessage "Error en menú Grupos: $($_.Exception.Message)"
            Pause-Script
        }
    }
}

function Group-Search {
    try {
        Show-Header
        Write-Host "--- BUSCAR/VER GRUPOS ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar grupo (deja vacío para ver todos)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        Write-Host ""
        Write-Host "Buscando grupos..." -ForegroundColor Cyan
        
        $groups = Get-ADGroup -Filter "Name -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -ErrorAction Stop | Sort-Object Name
        
        if ($groups) {
            Write-Host ""
            Write-Host "Grupos encontrados:" -ForegroundColor Green
            Write-Host ""
            $groups | Format-Table Name, GroupScope, GroupCategory, DistinguishedName -AutoSize
        } else {
            Show-WarningMessage "No se encontraron grupos con el término '$searchTerm'."
        }
        
    } catch {
        Show-ErrorMessage "No se pudieron listar los grupos: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Group-Create {
    try {
        Show-Header
        Write-Host "--- CREAR NUEVO GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        
        $groupName = Ask-Input "Nombre del grupo (ej: GG-Médico)"
        $samName = Normalize-SamAccountName $groupName
        
        Write-Host ""
        Write-Host "SamAccountName generado: $samName" -ForegroundColor DarkGray
        Write-Host ""
        
        Write-Host "Ámbito del grupo:" -ForegroundColor Yellow
        Write-Host "1) Global" -ForegroundColor White
        Write-Host "2) Universal" -ForegroundColor White
        Write-Host "3) DomainLocal" -ForegroundColor White
        
        $scopeOption = Read-Host "Opción"
        
        switch ($scopeOption) {
            '1' { $scope = 'Global' }
            '2' { $scope = 'Universal' }
            '3' { $scope = 'DomainLocal' }
            default { 
                Show-WarningMessage "Ámbito no válido, usando 'Global' por defecto."
                $scope = 'Global'
            }
        }
        
        Write-Host ""
        $description = Read-Host "Descripción del grupo (opcional)"
        
        Write-Host ""
        Write-Host "¿Dónde crear el grupo?" -ForegroundColor Yellow
        Write-Host "1) En la raíz del dominio" -ForegroundColor White
        Write-Host "2) Dentro de una OU existente" -ForegroundColor White
        
        $locationOption = Read-Host "Opción"
        
        if ($locationOption -eq '2') {
            $ouPath = Select-OU-Interactive "crear el grupo '$groupName' en"
            if (-not $ouPath) {
                Show-WarningMessage "Operación cancelada."
                Pause-Script
                return
            }
        } else {
            $ouPath = $script:BaseDN
        }
        
        Write-Host ""
        Write-Host "Creando grupo '$groupName' ($scope) en $ouPath..." -ForegroundColor Cyan
        
        $params = @{
            Name = $groupName
            SamAccountName = $samName
            GroupScope = $scope
            GroupCategory = 'Security'
            Path = $ouPath
            Server = $script:Server
            ErrorAction = 'Stop'
        }
        
        if (-not [string]::IsNullOrWhiteSpace($description)) {
            $params.Description = $description
        }
        
        New-ADGroup @params
        
        Show-SuccessMessage "Grupo '$groupName' creado exitosamente."
        
    } catch {
        Show-ErrorMessage "No se pudo crear el grupo: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Group-Edit {
    try {
        $group = Select-Group-Interactive "editar"
        if (-not $group) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- EDITAR GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Grupo seleccionado: $($group.Name)" -ForegroundColor DarkGreen
        Write-Host "Ámbito actual: $($group.GroupScope)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "1. Cambiar nombre" -ForegroundColor White
        Write-Host "2. Cambiar ámbito" -ForegroundColor White
        Write-Host "3. Cambiar descripción" -ForegroundColor White
        Write-Host "0. Cancelar" -ForegroundColor White
        Write-Host ""
        
        $editOption = Read-Host "Opción"
        
        switch ($editOption) {
            '1' {
                $newName = Ask-Input "Nuevo nombre"
                if (Ask-YesNo "¿Cambiar nombre de '$($group.Name)' a '$newName'?") {
                    Rename-ADObject -Identity $group.DistinguishedName -NewName $newName -Server $script:Server -ErrorAction Stop
                    Show-SuccessMessage "Nombre cambiado exitosamente."
                } else {
                    Show-WarningMessage "Operación cancelada."
                }
            }
            '2' {
                Write-Host ""
                Write-Host "Nuevo ámbito:" -ForegroundColor Yellow
                Write-Host "1) Global" -ForegroundColor White
                Write-Host "2) Universal" -ForegroundColor White
                Write-Host "3) DomainLocal" -ForegroundColor White
                
                $scopeOption = Read-Host "Opción"
                
                switch ($scopeOption) {
                    '1' { $newScope = 'Global' }
                    '2' { $newScope = 'Universal' }
                    '3' { $newScope = 'DomainLocal' }
                    default { 
                        Show-ErrorMessage "Ámbito no válido."
                        Pause-Script
                        return
                    }
                }
                
                if (Ask-YesNo "¿Cambiar ámbito de '$($group.Name)' a '$newScope'?") {
                    Set-ADGroup -Identity $group.DistinguishedName -GroupScope $newScope -Server $script:Server -ErrorAction Stop
                    Show-SuccessMessage "Ámbito cambiado exitosamente."
                } else {
                    Show-WarningMessage "Operación cancelada."
                }
            }
            '3' {
                $newDesc = Ask-Input "Nueva descripción"
                if (Ask-YesNo "¿Actualizar descripción de '$($group.Name)'?") {
                    Set-ADGroup -Identity $group.DistinguishedName -Description $newDesc -Server $script:Server -ErrorAction Stop
                    Show-SuccessMessage "Descripción actualizada exitosamente."
                } else {
                    Show-WarningMessage "Operación cancelada."
                }
            }
            '0' {
                Show-WarningMessage "Operación cancelada."
            }
            default {
                Show-ErrorMessage "Opción no válida."
            }
        }
        
    } catch {
        Show-ErrorMessage "No se pudo editar el grupo: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Group-Delete {
    try {
        $group = Select-Group-Interactive "BORRAR"
        if (-not $group) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- BORRAR GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Grupo seleccionado: $($group.Name)" -ForegroundColor Red
        Write-Host ""
        
        if (Ask-YesNo "¿Estás SEGURO de que quieres BORRAR el grupo '$($group.Name)'?") {
            Write-Host ""
            Write-Host "Borrando grupo..." -ForegroundColor Cyan
            
            Remove-ADGroup -Identity $group.DistinguishedName -Confirm:$false -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "Grupo borrado exitosamente."
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo borrar el grupo: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Group-AddMembers {
    try {
        $group = Select-Group-Interactive "añadir miembros a"
        if (-not $group) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- AÑADIR MIEMBROS A GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Grupo: $($group.Name)" -ForegroundColor DarkGreen
        Write-Host ""
        
        $memberSam = Ask-Input "SamAccountName del usuario o grupo a añadir"
        
        Write-Host ""
        Write-Host "Añadiendo miembro..." -ForegroundColor Cyan
        
        Add-ADGroupMember -Identity $group.DistinguishedName -Members $memberSam -Server $script:Server -ErrorAction Stop
        
        Show-SuccessMessage "Miembro '$memberSam' añadido exitosamente al grupo '$($group.Name)'."
        
    } catch {
        if ($_.Exception.Message -match 'already a member') {
            Show-WarningMessage "El miembro ya pertenece al grupo."
        } else {
            Show-ErrorMessage "No se pudo añadir el miembro: $($_.Exception.Message)"
        }
    }
    
    Pause-Script
}

function Group-RemoveMembers {
    try {
        $group = Select-Group-Interactive "quitar miembros de"
        if (-not $group) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- QUITAR MIEMBROS DE GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Grupo: $($group.Name)" -ForegroundColor DarkGreen
        Write-Host ""
        
        # Mostrar miembros actuales
        Write-Host "Miembros actuales:" -ForegroundColor Yellow
        try {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -Server $script:Server -ErrorAction Stop
            if ($members) {
                $members | Format-Table Name, SamAccountName, ObjectClass -AutoSize
            } else {
                Write-Host "El grupo no tiene miembros." -ForegroundColor Gray
            }
        } catch {
            Show-WarningMessage "No se pudieron listar los miembros."
        }
        
        Write-Host ""
        $memberSam = Ask-Input "SamAccountName del usuario o grupo a quitar"
        
        if (Ask-YesNo "¿Quitar '$memberSam' del grupo '$($group.Name)'?") {
            Write-Host ""
            Write-Host "Quitando miembro..." -ForegroundColor Cyan
            
            Remove-ADGroupMember -Identity $group.DistinguishedName -Members $memberSam -Confirm:$false -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "Miembro '$memberSam' quitado exitosamente del grupo."
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo quitar el miembro: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Group-ViewMembers {
    try {
        $group = Select-Group-Interactive "ver miembros de"
        if (-not $group) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- MIEMBROS DEL GRUPO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Grupo: $($group.Name)" -ForegroundColor DarkGreen
        Write-Host ""
        
        Write-Host "Obteniendo miembros..." -ForegroundColor Cyan
        $members = Get-ADGroupMember -Identity $group.DistinguishedName -Server $script:Server -ErrorAction Stop
        
        if ($members) {
            Write-Host ""
            Write-Host "Miembros encontrados:" -ForegroundColor Green
            Write-Host ""
            $members | Format-Table Name, SamAccountName, ObjectClass, DistinguishedName -AutoSize
        } else {
            Show-WarningMessage "El grupo no tiene miembros."
        }
        
    } catch {
        Show-ErrorMessage "No se pudieron listar los miembros: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Select-Group-Interactive {
    param([string]$Action)
    
    try {
        Show-Header
        Write-Host "--- SELECCIÓN DE GRUPO PARA $Action ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar grupo (deja vacío para ver todos)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        $groups = Get-ADGroup -Filter "Name -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -ErrorAction Stop | Sort-Object Name
        
        if (-not $groups) {
            Show-WarningMessage "No se encontraron grupos."
            return $null
        }
        
        Write-Host ""
        Write-Host "Grupos encontrados:" -ForegroundColor Green
        $i = 1
        foreach ($grp in $groups) {
            Write-Host "$i) $($grp.Name) ($($grp.GroupScope)) - $($grp.DistinguishedName)" -ForegroundColor White
            $i++
        }
        
        Write-Host ""
        $selection = Read-Host "Selecciona el número del grupo (0 para cancelar)"
        
        if ($selection -eq '0') {
            return $null
        }
        
        $idx = [int]$selection - 1
        if ($idx -ge 0 -and $idx -lt $groups.Count) {
            return $groups[$idx]
        } else {
            Show-ErrorMessage "Selección no válida."
            return $null
        }
        
    } catch {
        Show-ErrorMessage "Error al seleccionar grupo: $($_.Exception.Message)"
        return $null
    }
}

# =======================================================
# ===== GESTIÓN DE USUARIOS =====
# =======================================================

function Menu-Users {
    while ($true) {
        try {
            Show-Header
            Write-Host "--- GESTIÓN DE USUARIOS ---" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "1. Buscar/Ver Usuarios" -ForegroundColor White
            Write-Host "2. Crear nuevo Usuario" -ForegroundColor White
            Write-Host "3. Editar Usuario" -ForegroundColor White
            Write-Host "4. Borrar Usuario" -ForegroundColor White
            Write-Host "5. Habilitar/Deshabilitar Usuario" -ForegroundColor White
            Write-Host "6. Resetear password de Usuario" -ForegroundColor White
            Write-Host "7. Ver grupos de Usuario" -ForegroundColor White
            Write-Host "0. Volver al menú principal" -ForegroundColor White
            Write-Host ""
            
            $option = Read-Host "Selecciona una opción"
            
            switch ($option) {
                '1' { User-Search }
                '2' { User-Create }
                '3' { User-Edit }
                '4' { User-Delete }
                '5' { User-EnableDisable }
                '6' { User-ResetPassword }
                '7' { User-ViewGroups }
                '0' { return }
                default { 
                    Show-ErrorMessage "Opción no válida. Por favor, selecciona una opción del menú."
                    Pause-Script
                }
            }
        } catch {
            Show-ErrorMessage "Error en menú Usuarios: $($_.Exception.Message)"
            Pause-Script
        }
    }
}

function User-Search {
    try {
        Show-Header
        Write-Host "--- BUSCAR/VER USUARIOS ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar usuario (deja vacío para ver todos)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        Write-Host ""
        Write-Host "Buscando usuarios..." -ForegroundColor Cyan
        
        $users = Get-ADUser -Filter "Name -like '$filter' -or SamAccountName -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -Properties Enabled -ErrorAction Stop | Sort-Object SamAccountName
        
        if ($users) {
            Write-Host ""
            Write-Host "Usuarios encontrados:" -ForegroundColor Green
            Write-Host ""
            $users | Format-Table SamAccountName, Name, Enabled, DistinguishedName -AutoSize
        } else {
            Show-WarningMessage "No se encontraron usuarios con el término '$searchTerm'."
        }
        
    } catch {
        Show-ErrorMessage "No se pudieron listar los usuarios: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-Create {
    try {
        Show-Header
        Write-Host "--- CREAR NUEVO USUARIO ---" -ForegroundColor Cyan
        Write-Host ""
        
        $givenName = Ask-Input "Nombre"
        $surname = Ask-Input "Apellido"
        
        $samBase = Normalize-SamAccountName "$givenName.$surname"
        Write-Host ""
        Write-Host "SamAccountName generado: $samBase" -ForegroundColor DarkGray
        
        $upn = "$samBase@$($script:DomainUPN)"
        Write-Host "UserPrincipalName: $upn" -ForegroundColor DarkGray
        Write-Host ""
        
        $password = Read-Host "Password inicial (Enter para usar 'P@ssw0rd.2025')"
        if ([string]::IsNullOrWhiteSpace($password)) {
            $password = "P@ssw0rd.2025"
        }
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        
        Write-Host ""
        Write-Host "¿Dónde crear el usuario?" -ForegroundColor Yellow
        Write-Host "1) En la raíz del dominio" -ForegroundColor White
        Write-Host "2) En una OU existente" -ForegroundColor White
        
        $locationOption = Read-Host "Opción"
        
        if ($locationOption -eq '2') {
            $ouPath = Select-OU-Interactive "crear el usuario en"
            if (-not $ouPath) {
                Show-WarningMessage "Operación cancelada."
                Pause-Script
                return
            }
        } else {
            $ouPath = $script:BaseDN
        }
        
        Write-Host ""
        Write-Host "Creando usuario '$samBase'..." -ForegroundColor Cyan
        
        New-ADUser -Name "$givenName $surname" -GivenName $givenName -Surname $surname `
            -SamAccountName $samBase -UserPrincipalName $upn `
            -AccountPassword $securePassword -ChangePasswordAtLogon $true `
            -Enabled $true -Path $ouPath -Server $script:Server -ErrorAction Stop
        
        Show-SuccessMessage "Usuario '$samBase' creado exitosamente."
        
    } catch {
        Show-ErrorMessage "No se pudo crear el usuario: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-Edit {
    try {
        $user = Select-User-Interactive "editar"
        if (-not $user) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- EDITAR USUARIO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usuario seleccionado: $($user.SamAccountName)" -ForegroundColor DarkGreen
        Write-Host ""
        Write-Host "1. Cambiar nombre completo (DisplayName)" -ForegroundColor White
        Write-Host "2. Cambiar descripción" -ForegroundColor White
        Write-Host "0. Cancelar" -ForegroundColor White
        Write-Host ""
        
        $editOption = Read-Host "Opción"
        
        switch ($editOption) {
            '1' {
                $newDisplayName = Ask-Input "Nuevo nombre completo (DisplayName)"
                if (Ask-YesNo "¿Actualizar DisplayName de '$($user.SamAccountName)' a '$newDisplayName'?") {
                    Set-ADUser -Identity $user.DistinguishedName -DisplayName $newDisplayName -Server $script:Server -ErrorAction Stop
                    Show-SuccessMessage "DisplayName actualizado exitosamente."
                } else {
                    Show-WarningMessage "Operación cancelada."
                }
            }
            '2' {
                $newDesc = Ask-Input "Nueva descripción"
                if (Ask-YesNo "¿Actualizar descripción de '$($user.SamAccountName)'?") {
                    Set-ADUser -Identity $user.DistinguishedName -Description $newDesc -Server $script:Server -ErrorAction Stop
                    Show-SuccessMessage "Descripción actualizada exitosamente."
                } else {
                    Show-WarningMessage "Operación cancelada."
                }
            }
            '0' {
                Show-WarningMessage "Operación cancelada."
            }
            default {
                Show-ErrorMessage "Opción no válida."
            }
        }
        
    } catch {
        Show-ErrorMessage "No se pudo editar el usuario: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-Delete {
    try {
        $user = Select-User-Interactive "BORRAR"
        if (-not $user) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- BORRAR USUARIO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usuario seleccionado: $($user.SamAccountName)" -ForegroundColor Red
        Write-Host "Nombre: $($user.Name)" -ForegroundColor Red
        Write-Host ""
        
        if (Ask-YesNo "¿Estás SEGURO de que quieres BORRAR el usuario '$($user.SamAccountName)'?") {
            Write-Host ""
            Write-Host "Borrando usuario..." -ForegroundColor Cyan
            
            Remove-ADUser -Identity $user.DistinguishedName -Confirm:$false -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "Usuario borrado exitosamente."
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo borrar el usuario: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-EnableDisable {
    try {
        $user = Select-User-Interactive "habilitar/deshabilitar"
        if (-not $user) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        $userDetail = Get-ADUser -Identity $user.DistinguishedName -Properties Enabled -Server $script:Server -ErrorAction Stop
        
        Show-Header
        Write-Host "--- HABILITAR/DESHABILITAR USUARIO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usuario: $($userDetail.SamAccountName)" -ForegroundColor DarkGreen
        Write-Host "Estado actual: $(if($userDetail.Enabled){'HABILITADO'}else{'DESHABILITADO'})" -ForegroundColor $(if($userDetail.Enabled){'Green'}else{'Red'})
        Write-Host ""
        
        if ($userDetail.Enabled) {
            if (Ask-YesNo "¿Deshabilitar el usuario '$($userDetail.SamAccountName)'?") {
                Disable-ADAccount -Identity $userDetail.DistinguishedName -Server $script:Server -ErrorAction Stop
                Show-SuccessMessage "Usuario deshabilitado exitosamente."
            } else {
                Show-WarningMessage "Operación cancelada."
            }
        } else {
            if (Ask-YesNo "¿Habilitar el usuario '$($userDetail.SamAccountName)'?") {
                Enable-ADAccount -Identity $userDetail.DistinguishedName -Server $script:Server -ErrorAction Stop
                Show-SuccessMessage "Usuario habilitado exitosamente."
            } else {
                Show-WarningMessage "Operación cancelada."
            }
        }
        
    } catch {
        Show-ErrorMessage "No se pudo cambiar el estado del usuario: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-ResetPassword {
    try {
        $user = Select-User-Interactive "resetear password de"
        if (-not $user) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        Show-Header
        Write-Host "--- RESETEAR PASSWORD ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usuario: $($user.SamAccountName)" -ForegroundColor DarkGreen
        Write-Host ""
        
        $newPassword = Ask-Input "Nueva password"
        $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
        
        if (Ask-YesNo "¿Resetear password de '$($user.SamAccountName)'?") {
            Write-Host ""
            Write-Host "Reseteando password..." -ForegroundColor Cyan
            
            Set-ADAccountPassword -Identity $user.DistinguishedName -Reset -NewPassword $securePassword -Server $script:Server -ErrorAction Stop
            Set-ADUser -Identity $user.DistinguishedName -ChangePasswordAtLogon $true -Server $script:Server -ErrorAction Stop
            
            Show-SuccessMessage "Password reseteada exitosamente. El usuario deberá cambiarla en el próximo inicio de sesión."
        } else {
            Show-WarningMessage "Operación cancelada."
        }
        
    } catch {
        Show-ErrorMessage "No se pudo resetear la password: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function User-ViewGroups {
    try {
        $user = Select-User-Interactive "ver grupos de"
        if (-not $user) {
            Show-WarningMessage "Operación cancelada."
            Pause-Script
            return
        }
        
        $userDetail = Get-ADUser -Identity $user.DistinguishedName -Properties MemberOf -Server $script:Server -ErrorAction Stop
        
        Show-Header
        Write-Host "--- GRUPOS DEL USUARIO ---" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usuario: $($userDetail.SamAccountName)" -ForegroundColor DarkGreen
        Write-Host ""
        
        if ($userDetail.MemberOf) {
            Write-Host "Grupos (MemberOf):" -ForegroundColor Green
            Write-Host ""
            foreach ($groupDN in $userDetail.MemberOf) {
                try {
                    $groupObj = Get-ADGroup -Identity $groupDN -Server $script:Server -ErrorAction Stop
                    Write-Host "  - $($groupObj.Name) ($($groupObj.GroupScope))" -ForegroundColor White
                } catch {
                    Write-Host "  - $groupDN" -ForegroundColor Gray
                }
            }
        } else {
            Show-WarningMessage "El usuario no pertenece a ningún grupo (excepto 'Domain Users' por defecto)."
        }
        
    } catch {
        Show-ErrorMessage "No se pudieron listar los grupos del usuario: $($_.Exception.Message)"
    }
    
    Pause-Script
}

function Select-User-Interactive {
    param([string]$Action)
    
    try {
        Show-Header
        Write-Host "--- SELECCIÓN DE USUARIO PARA $Action ---" -ForegroundColor Cyan
        Write-Host ""
        
        $searchTerm = Read-Host "Buscar usuario (deja vacío para ver todos)"
        
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $filter = "*"
        } else {
            $filter = "*$searchTerm*"
        }
        
        $users = Get-ADUser -Filter "Name -like '$filter' -or SamAccountName -like '$filter'" -SearchBase $script:BaseDN -Server $script:Server -ErrorAction Stop | Sort-Object SamAccountName
        
        if (-not $users) {
            Show-WarningMessage "No se encontraron usuarios."
            return $null
        }
        
        Write-Host ""
        Write-Host "Usuarios encontrados:" -ForegroundColor Green
        $i = 1
        foreach ($usr in $users) {
            Write-Host "$i) $($usr.SamAccountName) - $($usr.Name)" -ForegroundColor White
            $i++
        }
        
        Write-Host ""
        $selection = Read-Host "Selecciona el número del usuario (0 para cancelar)"
        
        if ($selection -eq '0') {
            return $null
        }
        
        $idx = [int]$selection - 1
        if ($idx -ge 0 -and $idx -lt $users.Count) {
            return $users[$idx]
        } else {
            Show-ErrorMessage "Selección no válida."
            return $null
        }
        
    } catch {
        Show-ErrorMessage "Error al seleccionar usuario: $($_.Exception.Message)"
        return $null
    }
}

# =======================================================
# ===== MENÚ PRINCIPAL =====
# =======================================================

function Show-MainMenu {
    while ($true) {
        try {
            Show-Header
            
            # Si no hay dominio seleccionado
            if (-not $script:Server) {
                Write-Host "--- CONEXIÓN AL DOMINIO ---" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "No hay ningún dominio seleccionado." -ForegroundColor Red
                Write-Host ""
                Write-Host "1. Conectar a un dominio" -ForegroundColor White
                Write-Host "0. Salir" -ForegroundColor White
                Write-Host ""
                
                $option = Read-Host "Selecciona una opción"
                
                switch ($option) {
                    '1' { Select-Domain }
                    '0' { 
                        Write-Host ""
                        Write-Host "¡Hasta pronto!" -ForegroundColor Green
                        Write-Host ""
                        return 
                    }
                    default { 
                        Show-ErrorMessage "Opción no válida."
                        Pause-Script
                    }
                }
                continue
            }
            
            # Menú principal con dominio conectado
            Write-Host "--- MENÚ PRINCIPAL ---" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "1. Gestionar Unidades Organizativas (OU)" -ForegroundColor White
            Write-Host "2. Gestionar Grupos" -ForegroundColor White
            Write-Host "3. Gestionar Usuarios" -ForegroundColor White
            Write-Host "4. Cambiar de dominio" -ForegroundColor White
            Write-Host "0. Salir" -ForegroundColor White
            Write-Host ""
            
            $option = Read-Host "Selecciona una opción"
            
            switch ($option) {
                '1' { Menu-OU }
                '2' { Menu-Groups }
                '3' { Menu-Users }
                '4' { 
                    $script:Server = $null
                    $script:BaseDN = $null
                    $script:DomainUPN = $null
                    Select-Domain 
                }
                '0' { 
                    Write-Host ""
                    Write-Host "¡Hasta pronto!" -ForegroundColor Green
                    Write-Host ""
                    return 
                }
                default { 
                    Show-ErrorMessage "Opción no válida. Por favor, selecciona una opción del menú."
                    Pause-Script
                }
            }
            
        } catch {
            Show-ErrorMessage "Error inesperado en menú principal: $($_.Exception.Message)"
            Pause-Script
        }
    }
}

# =======================================================
# ===== INICIO DEL SCRIPT =====
# =======================================================

try {
    Show-MainMenu
} catch {
    Write-Host ""
    Write-Host "ERROR CRÍTICO: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Pause-Script
}
