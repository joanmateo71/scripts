function HideFromAddressList_OU {
    $OU = Read-Host "Please enter the OU (e.g., OU=Users,DC=YourDomain,DC=com)"
    $users = Get-ADUser -Filter * -SearchBase $OU -Properties msExchHideFromAddressLists

    Write-Host "The following users will be checked:" -ForegroundColor Yellow
    $users | ForEach-Object { Write-Host $_.SamAccountName }

    $confirm = Read-Host "Do you want to proceed with checking and hiding these users from the address list if needed? (y/n)"
    
    if ($confirm -eq 'y') {
        foreach ($user in $users) {
            if ($user.msExchHideFromAddressLists -eq $true) {
                Write-Host "User $($user.SamAccountName) is already hidden from the address list." -ForegroundColor Yellow
            } else {
                Set-ADUser -Identity $user.SamAccountName -Add @{msExchHideFromAddressLists=$true}
                Write-Host "Updated user: $($user.SamAccountName)" -ForegroundColor Green
            }
        }
        Write-Host "Operation completed for all users in the OU." -ForegroundColor Cyan
    } else {
        Write-Host "Operation cancelled. No changes were made." -ForegroundColor Red
    }
}

function HideFromAddressList_User {
    $userSamAccountName = Read-Host "Please enter the SamAccountName of the user"

    $user = Get-ADUser -Identity $userSamAccountName -Properties msExchHideFromAddressLists

    if ($user.msExchHideFromAddressLists -eq $true) {
        Write-Host "User $($user.SamAccountName) is already hidden from the address list." -ForegroundColor Yellow
    } else {
        $confirm = Read-Host "Do you want to hide this user from the address list? (y/n)"
        if ($confirm -eq 'y') {
            Set-ADUser -Identity $user.SamAccountName -Add @{msExchHideFromAddressLists=$true}
            Write-Host "Updated user: $($user.SamAccountName)" -ForegroundColor Green
        } else {
            Write-Host "Operation cancelled. No changes were made." -ForegroundColor Red
        }
    }
}

function Show-Menu {
    Write-Host "Select an option:" -ForegroundColor Cyan
    Write-Host "1: Update users in a specific OU"
    Write-Host "2: Update a specific user"
    Write-Host "3: Exit"
}

do {
    Show-Menu
    $selection = Read-Host "Please enter your selection (1, 2, or 3)"

    switch ($selection) {
        1 { HideFromAddressList_OU }
        2 { HideFromAddressList_User }
        3 { Write-Host "Exiting..." -ForegroundColor Cyan }
        default { Write-Host "Invalid selection, please try again." -ForegroundColor Red }
    }
} while ($selection -ne 3)

Write-Host "Script execution completed." -ForegroundColor Cyan