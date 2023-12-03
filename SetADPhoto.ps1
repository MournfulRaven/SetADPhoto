# Script created by Valeriy Bachilo

$pathToConvert = '\\path_to_storage\UserPhotos\ToConvert\'
$pathToBackup = '\\path_to_storage\UserPhotos\Backup\'
$domain = 'example.com'

Function Menu {
	Write-Host "
############################################" -ForegroundColor White -BackgroundColor Red
	Write-Host "###" -ForegroundColor White -BackgroundColor Red -NoNewLine; 
	Write-Host " Changing Active Directory User Photo " -ForegroundColor White -BackgroundColor Red -NoNewLine; 
	Write-Host "###" -ForegroundColor White -BackgroundColor Red;
	Write-Host "############################################
" -ForegroundColor White -BackgroundColor Red
	Write-Host "Select Function:" -ForegroundColor Green
	Write-Host "
	1. Convert and set photos to user's accounts
	2. Delete user's photo
	3. Export user's photo from AD
	
	0. Exit
	"
	$choice = MenuChoice
}

Function MenuChoice {
	$choice = Read-Host "Item"
	Switch ($choice) {
		0 { Write-Host "Thanks for using this script!" -ForegroundColor Green; Start-Sleep 1; Exit }
		1 { SetADPhoto; Start-Sleep 1; Menu }
		2 { DeleteUserPhoto; Start-Sleep 1; Menu }
		Default
		{ 
			While ($attempt -gt 2) { $attempt = 0; Write-Host ""; Start-Sleep 1; Write-Host "Error..." -ForegroundColor Red; Start-Sleep 1; Write-Host "Check the menu!"; Menu }
		$attempt = $attempt + 1; Write-Host ""; Start-Sleep 1; Write-Host "Error..." -ForegroundColor Red; Start-Sleep 1; Write-Host "Please type only one number from proposed!"; MenuChoice }
	}
}

Function SetADPhoto {
    if (!(Test-Path "ToConvert:\") -or !(Test-Path "ToBackup:\")) {
        New-PSDrive -Name ToConvert -PSProvider FileSystem -Root $pathToConvert | Out-Null
        New-PSDrive -Name ToBackup -PSProvider FileSystem -Root $pathToBackup | Out-Null
    }
    $files = Get-ChildItem -Path 'ToConvert:\'
    if ($files.Count -eq $null) {
        Write-Host "No file has been found... Exiting" -ForegroundColor Green
        Exit
    }
    else {
        ExchangeAzure_NewSession
	    Write-Host $files.count "files found!" -ForegroundColor Green

        foreach ($file in $files) {
            $samaccountname = $file.BaseName
            $fileName = $file.Name
            Write-Host ""
            Write-Host "Backing up file" $fileName
            try { Copy-Item "ToConvert:\$fileName" -Destination 'ToBackup:\' -Force } catch { Write-Host "An error occured while backing up file" $fileName; Exit-PSHostProcess }
            
            if ((Get-ADUser $samaccountname) -eq $null) {
                Write-Host "User" $samaccountname "not found"
                Exit-PSHostProcess
            }
            
            # setting thumbnail attribute photo
            Write-Host "Converting file" $fileName " and setting thumbnailPhoto..."
            $thumbnailPhoto = ( [byte[]]( $( ResizePhoto "ToConvert:\$fileName" 64 95 ) ) )
            Set-ADUser -Identity $samaccountname -Replace @{ thumbnailPhoto = $thumbnailPhoto } -ErrorVariable error
            if ($error -ne $null) {
                Write-Host "Error while setting thumbnailPhoto for user " $samaccountname
                $error = $null
                Exit-PSHostProcess
			}
            $thumbnailPhoto = $null

            # setting jpeg attribute photo
            Write-Host "Converting file" $fileName " and setting jpegPhoto..."
            $jpegPhoto = [byte[]]( $( ResizePhoto "ToConvert:\$fileName" 648 95 ) )
            Set-ADUser -Identity $samaccountname -Replace @{jpegPhoto = $jpegPhoto } -ErrorVariable error
            if ($error -ne $null) {
                Write-Host "Error while setting jpegPhoto for user " $samaccountname
                $error = $null
                Exit-PSHostProcess
			}
            
            # setting exchange photo
            Write-Host "Converting file" $fileName " and setting webPhoto..."
            Set-UserPhoto -Identity $samaccountname -PictureData $jpegPhoto -Confirm:$False -ErrorVariable error
            if ($error -ne $null) {
                Write-Host "Error while setting webPhoto for user " $samaccountname
                $error = $null
                Exit-PSHostProcess
			}
            $jpegPhoto = $null
            
            Remove-Item "ToConvert:\$fileName" -Force
            Write-Host "Jobs for user" $samaccountname "finished!" -ForegroundColor Green
            Write-Host ""
        }
    }
    Write-Host "Done!"
}

Function DeleteUserPhoto {
	Write-Host "Please, enter login for user you want to delete photo: " -NoNewLine -ForegroundColor Green
	$user = Read-Host
	Set-ADUser -Identity $user -Clear thumbnailPhoto
    Set-ADUser -Identity $user -Clear jpegPhoto
	Write-Host "Photo successfully removed!" -ForegroundColor Green

}

Function ResizePhoto(){
    Param (
        [Parameter(Mandatory=$True)][ValidateNotNull()] $imageSource,
        [Parameter(Mandatory=$True)][ValidateNotNull()] $canvasSize,
        [Parameter(Mandatory=$True)][ValidateNotNull()] $quality 
    )

    #if ( !(Test-Path $imageSource) ) { throw("File not found") }
    #if ( ($canvasSize -lt 10) -or ($canvasSize -gt 1000) ) { throw("Size parameter should be in range of 10 -- 10000") }
    #if ( ($quality -lt 0) -or ($quality -gt 100) ) { throw("Quality parameter should be in range of 0 -- 100") }

    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $imageBytes = [byte[]](Get-Content $imageSource -Encoding byte)
    $ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
    $ms.Write($imageBytes, 0, $imageBytes.Length);

    $bmp = [System.Drawing.Image]::FromStream($ms, $true)

    # size of picture after converting
    $canvasWidth = $canvasSize
    $canvasHeight = $canvasSize

    # quality of picture
    $myEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $quality)
    #Получаем тип картинки
    $myImageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | where {$_.MimeType -eq 'image/jpeg'}

    # counting multiplicity
    $ratioX = $canvasWidth / $bmp.Width;
    $ratioY = $canvasHeight / $bmp.Height;
    $ratio = $ratioY
    if( $ratioX -le $ratioY ) {
        $ratio = $ratioX
    }

    # creating empty picture
    $newWidth = [int]( $bmp.Width * $ratio )
    $newHeight = [int]( $bmp.Height * $ratio )
    $bmpResized = New-Object System.Drawing.Bitmap( $newWidth, $newHeight )
    $graph = [System.Drawing.Graphics]::FromImage( $bmpResized )

    $graph.Clear([System.Drawing.Color]::White)
    $graph.DrawImage($bmp,0,0 , $newWidth, $newHeight)

    # creting empty stream
    $ms = New-Object IO.MemoryStream
    $bmpResized.Save( $ms, $myImageCodecInfo, $($encoderParams) )
    
    # cleaning
    $bmpResized.Dispose()
    $bmp.Dispose()
    
    return $ms.ToArray()
}

Function ExchangeAzure_NewSession {
    $UPN = "$env:username@$domain"
    Write-Host "Connecting to Exchange Online... " -ForegroundColor DarkGray
    Get-PSSession | Remove-PSSession
    try {
        Import-Module ExchangeOnlineManagement| Out-Null
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress:$false -ShowBanner:$false -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Connection to Local Exchange session was unsuccessfull... Exiting" -ForegroundColor DarkRed
    }
    Write-Host "Connected" -ForegroundColor DarkGray
}

Menu

