# Get the contents of the file
$text = Get-Content -raw -Encoding "UTF8" '.\psDocument.md'
# Get the array of footnote refs to be converted
$fnRefs = $text | Select-String -Pattern '\s\[\^([^\s\]]+)]:' -AllMatches |
    % { $_.matches } |
    # % { $_.groups[1].value } |
    % { $_.groups[1].value } 
forEach ($ref in $fnRefs) { 
    $text = $text -ireplace "\[\^($ref)]([\s\S]+?)\[\^\1]:\s*((?:\S|\s(?!\s+\[\^\d+]))+)\s+", '^[$3]$2' 
}
	
$text | Out-File -Encoding "UTF8" '.\psDocument_replaced.md'
    
