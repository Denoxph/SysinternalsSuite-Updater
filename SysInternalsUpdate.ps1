##############################Function##############################
#   Dieser bereich gilt den für dasScript benötigten Funktionen    #
####################################################################

#region functions

#region Formatierung zu Tabellen Format

    function Format ($splitArray){

        $ret = New-Object System.Collections.ArrayList 

        $splitArray | ForEach-Object {
        
            $split = $_.Split(' ') | Where-Object {$_ -ne ''}
        
            If($split -ne $null){

                $namefrag = $split[8].Remove(0, $split[8].IndexOf("/") + 1)

                $name = $namefrag.Remove($namefrag.IndexOf("`""), $namefrag.Length - $namefrag.IndexOf("`""))

                If($name.Split(".").Length -gt 1){

                    $release = $split[0] + " " + $split[1] + " " + $split[2] + " " + $split[3] + " " + $split[4] + " " + $split[5]

                    $object = New-Object PSObject -Property @{
                            name = $name
                            release = $release
                            URL = "$uri$name"
                        }

                    $ret += $object

                    $namefrag = $null
                    $name = $null
                    $release = $null
                    $object = $null

                }

            }
        
        }

        return $ret
    }

#endregion


#region Funktion für Update Prüfung + Download von Updates und neuen Tools

    Function ChkUpdate($splitArray){
    
        $log = New-Object System.Collections.ArrayList

        $newData = Format $splitArray

        $oldData = Import-Csv -Path "$path\SysinternalsSuite\Update Log\Versions.csv" -Delimiter ";"

        $newData | ForEach-Object{
        
            #region Vergleichen Release Daten + Laden von Updates

                If($oldData.name -ccontains $_.name){
            
                    $newVersion = $_

                    $oldVersion = $oldData | Where-Object { $_.name -eq $newVersion.name }

                    If( $oldVersion.release -ne $newVersion.release ){

                        $url = $_.URL
                        $release = $_.release
                        $name = $_.name

                        Invoke-WebRequest -Uri $url -OutFile "$path\SysinternalsSuite\$name" -ErrorAction Ignore

                        $object = New-Object PSObject -Property @{
                                name = $name
                                release = $release
                                URL = $url
                                Type = "Update"
                            }

                        $log += $object

                    }
        
                }

            #endregion

            #region Download neuer Tools

                Else{

                    $url = $_.URL
                    $release = $_.release
                    $name = $_.name

                    Invoke-WebRequest -Uri $url -OutFile "$path\SysinternalsSuite\$name" -ErrorAction Ignore

                    $object = New-Object PSObject -Property @{
                            name = $name
                            release = $release
                            URL = $url
                            Type = "NEW"
                        }

                    $log += $object
        
                }

            #endregion
    
        }

        If($log[0] -ne $null){

            $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        
            $log | Select-Object -Property "name", "URL", "release", "Type" | Export-Csv -Path "$path\SysinternalsSuite\Update Log\Log_$date.csv" -Delimiter ";"
        
        }

        $newData | Select-Object -Property "name", "URL", "release" | Export-Csv -Path "$path\SysinternalsSuite\Update Log\Versions.csv" -Delimiter ";"

    }

    #region Vergleich runtergeladener Tools zu Version.csv

        function ChkInstalled(){

            $log = New-Object System.Collections.ArrayList
        
            $toolList = Import-Csv -Path "$path\SysinternalsSuite\Update Log\Versions.csv" -Delimiter ";"

            $installed = Get-ChildItem -Path "$path\SysinternalsSuite" -File

            $toolList | ForEach-Object {
            
                If( !($installed.Name -ccontains $_.name)){
                    
                    $url = $_.URL
                    $release = $_.release
                    $name = $_.name

                    Write-Host "Missing: $name - Trying Reinstall"

                    Try{

                        Invoke-WebRequest -Uri $url -OutFile "$path\SysinternalsSuite\$name" -ErrorAction Ignore
                        
                        $object = New-Object PSObject -Property @{
                            name = $name
                            release = $release
                            URL = $url
                            Type = "NEW"
                        }

                        $log += $object

                    }

                    Catch{
                    
                        Write-Host "Failed"

                    }
            
                }
            
            }

            If($log[0] -ne $null){

                $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        
                $log | Select-Object -Property "name", "URL", "release", "Type" | Export-Csv -Path "$path\SysinternalsSuite\Update Log\Log_$date.csv" -Delimiter ";"
        
            }
        
        }

    #endregion

    #region Entfernen alter log Dateien

        function RemoveOldLogs(){

            $date = Get-Date
        
            $Logs = Get-ChildItem -Path "$path\SysinternalsSuite\Update Log" -File

            $Logs | Where-Object { $Logs.Name.Split("_")[0] -eq "Log" } | ForEach-Object {
            
                If( $date.Subtract($_.CreationTime).Days -ge 14 ){
                
                    Remove-Item $_.FullName
                
                }
            
            }
        
        }

    #endregion

#endregion functions

###############################Skript###############################
#        Dieser bereich beinhaltet das auszuführende Script        #
####################################################################

#region script

#region Variablen die für das script benötigt werden

    # $path je nach bedarf anpassen
    $path = "C:\Users\Phillip Hansen\OneDrive\Desktop\Test"
    $uri = "https://live.sysinternals.com/"
    $newFolder = $false

#endregion

#region Sammeln der Webdaten + initiale Formatierung

    $list = Invoke-WebRequest -Uri $uri

    $wip = $list.RawContent.Remove(0,$list.RawContent.IndexOf("<pre") + 5)

    $wip = $wip.Remove($wip.IndexOf('</pre'), $wip.Length - $wip.IndexOf("</pre"))

    $splitArray = $wip -Split"<br>"

#endregion

#region Pfad Prüfung + Initiale Downloads

    If(!(Test-Path "$path\SysinternalsSuite\Update Log")){
    
        If(!(Test-Path "$path\SysinternalsSuite\Update Log")){

            New-Item -Path "$path" -Name "SysinternalsSuite" -ItemType Directory

            $dFiles = Format $splitArray 
        
            $dFiles | ForEach-Object {
            
                $url = $_.URL
                $name = $_.name

                Try{
                    Invoke-WebRequest -Uri $url -OutFile "$path\SysinternalsSuite\$name" -ErrorAction Ignore
                }

                Catch{
                
                    Write-Host "Failed to load $name"
                
                }
            }

            $newFolder = $true

        }

        New-Item -Path "$path\SysinternalsSuite" -Name "Update Log" -ItemType Directory

        If($newFolder){
    
            $dFiles | Select-Object -Property "name", "URL", "release" | Export-Csv -Path "$path\SysinternalsSuite\Update Log\Versions.csv" -Delimiter ";"
    
        }

    }

#endregion

#region Prüfen auf Updates + Updates und Neue Sysinternals

    If(!$newFolder){

        ChkUpdate $splitArray

    }

#endregion

#region Vergleich runtergeladener Tools zu Version.csv

    ChkInstalled    

#endregion

#region Entfernen von alten Log Dateien

    RemoveOldLogs

#endregion script

Start-Sleep -Seconds 5