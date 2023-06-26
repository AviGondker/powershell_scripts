Get-Content joiner_data.txt | Where-Object {$_.length -gt0} | Where-Object {!$_.StartsWith("#")} | ForEach-Object {
    $var = $_.Split(':',2).Trim()
        New-Variable -Name $var[0] -Value $var[1]
                }

        echo $var[1], $var[2];
        #echo $Name2