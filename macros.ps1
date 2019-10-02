
 function testmacros {

    $store = $null
    $Macros = $null
    $array = @()
    $chemin= Read-Host "entrer le chemin du dossier à analyser "
    Get-ChildItem $chemin -Recurse | where {$_.extension -like "*.xls*" } | ForEach-Object -Begin {
    $thisThread = [System.Threading.Thread]::CurrentThread
    $originalCulture = $thisThread.CurrentCulture
    
    #demarrer Excel / Open Excel
    $excel = new-object -comobject excel.application
    
    #Ignorer les messages de demandes à l'utilisateur  / Skip message
    $excel.DisplayAlerts = $false
    #$excel.ScreenUpdating = $false
    $excel.Visible = $false
    $excel.Interactive = $false
    $excel.UserControl = $false
       
} -Process {   
    # Ouvrir le fichier Excel / open excel files
    
    $workbook = $excel.workbooks.Open($_.FullName,$false,$true) 
    Write-Host on traite le : $_.Name 


    # Verifier la presence de macros / check contains macros
    if ($workbook.HasVBProject -notlike "false") {
    
    if ($workbook.HasVBProject -like "true") {
    $Macros = "present"
     } else  {

     #$Macros = "Fichiers pas verifier /files not check" 

     } 


    $row = new-object psobject
    $row | add-member -type NoteProperty -name 'Fichier' -Value $_.Name
    $row | add-member -type NoteProperty -name 'Macros' -Value $Macros
    $array += $row
    }
    
       
    #Fermeture du fichier Excel / close file Excel
    $workbook.Close()
    

} -end {
    #Fermer le processus Excel /close process
    $excel.Quit()
    $thisThread.CurrentCulture = $originalCulture
    
    }


    Stop-Process -Name EXCEL 
    $array | Out-GridView -Title "Verification des Macros"

        
  }

  testmacros -verbose
