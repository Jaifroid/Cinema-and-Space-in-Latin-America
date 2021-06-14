# DEV: To reload this script after editing, issue ". $profile" at the PS commandline
if (Get-Command 'panprocessor' -erroraction silentlycontinue) { "Reloading panprocessor..." }
if (Get-Command 'Start-Pomodoro' -erroraction silentlycontinue) { "Reloading Start-Pomodoro..." }

function Start-Pomodoro {

    Param (
        [int]$Minutes = 25
    )
 
    $seconds = $Minutes*60
    $delay = 15 #seconds between ticks
 
    for($i = $seconds; $i -gt 0; $i = $i - $delay)
    {
        $percentComplete = 100-(($i/$seconds)*100)
        Write-Progress -SecondsRemaining $i `
                       -Activity "Pomodoro" `
                       -Status "Time remaining:" `
                       -PercentComplete $percentComplete
        Start-Sleep -Seconds $delay
    }
 
   # $player = New-Object System.Media.SoundPlayer "C:\Users\me\Dropbox\Music\CTU.wav"
   # 1..6 | %{ $player.Play() ; sleep -m 1400 }
    
   1..10 | %{ [System.Media.SystemSounds]::Exclamation.Play() ; sleep -m 750 }
}


function panprocessor { 

param (
    [string]$filename = "",
	[int]$offset = 0,
	[int]$steps = 5,
	[switch]$pad = $false,
    [switch]$renumber = $false,
    [switch]$dirlevel = $false,
	[switch]$tomaster = $false,
	[switch]$frommaster = $false,
	[string]$quotes = "",
	[switch]$split = $false,
	[switch]$nosplit = $false,
	[switch]$linknotes = $false,
	[switch]$todocx = $false,
	[switch]$bib = $false,
	# Insert the name of your library file, if any, here
	[string]$library = "C:\Users\geoff\OneDrive\Books\My Library.json",
	[string]$csl = "C:\Users\geoff\Bibliographies\Zotero database\styles\chicago-fullnote-bibliography.csl",
	#[string]$csl = "C:/Users/geoff/OneDrive/Books/chicago-note-bibliography.csl",
	[switch]$notoc = $false,
	[switch]$tomd = $false,
	[switch]$nobom = $false,
	[switch]$addbom = $false,
	[switch]$dumpargs = $false,
	[switch]$help = $false
)

$dump = ''
if ($dumpargs) { $dump='--dump-args' } 
# Deal with cases where no directory or filename is entered
if (($filename -eq "") -and (! $help)) { 
    if ((! $nobom) -and (! $addbom)) {
        $filename = Read-Host "Enter the filename or directory for processing, or type ? for help: "
    } else {
        " "
        ls -r -name . *.md
        if ($nobom) {
            $input = Read-Host "`nAll the above files will have the BOM (if any) removed!`nProceed? (Y/N)"
        } else {
            $input = Read-Host "`nAll the above files will have a BOM added (if necessary)!`nProceed? (Y/N)"
        }
        if ($input -eq "Y") {
            if ($nobom) {
                "Removing Byte Order Mark ..."
                ls -r . *.md | % { [System.IO.File]::WriteAllLines($_.FullName, ((Get-Content $_.FullName) -replace "^\xEF\xBB\xBF", ""))}
            } else {
                "Adding Byte Order Mark ..."
                ls -r . *.md | % {
                    $document = Get-Content -encoding "UTF8" $_.FullName
                    if ($document -match "^(?!\xEF\xBB\xBF)") { 
                        $document | Set-Content -encoding "UTF8" $_.FullName 
                    }
                }
            }
            "Done."
        } else {
            "Operation cancelled!"
        }
        exit
    }
}

# Display help options if user requested
if ((! $filename) -or ($filename -eq "?") -or ($help)) {
@"

Usage: .\panprocessor FILENAME or DIRECTORY [-offset num] [-steps num] [-pad]
    [-renumber] [-dirlevel] [-tomaster | -frommaster] [-quotes single|double|no] 
    [-split | -nosplit] [-linknotes] [-todocx | -tomd] [-bib] [-nobom] [-help] | [-addbom] 

Uses pandoc to convert a Word file to markdown and optionally split it into
numbered sections. Can also be used to compile a set of split markdown files
either into a single markdown document, or into a Word document. Can be used to
renumber all the files in a directory and subdirectories. Can also "tidy" a
single markdown file. If the script is invoked with no options, it will prompt
you to specify them as needed.

FILENAME or      the filename or directory on which to operate (can use absolute
  DIRECTORY      or relative path). If entering a filename, the script will 
                 assume you wish to convert it to markdown, unless it is named
                 *_master.md, in which case it will be used as a master list of
                 files to compile. If a directory, it will assume you wish to
                 compile split files to a document.
-offset num      when splitting files, provide an offset number to start from,
                 otherwise it will start from 0
-steps num       when spliiting files, specify the numbering step; this is
                 useful if you need to re-order files manually later; default 5
-pad             [switch]: if present, numbers will be padded with zeros, like
                 01, 02, 03, or 010, 020, 030, etc., as needed; useful if your
                 filesystem orders files ASCIIbetically rather than naturally
-renumber        renumbers files in directory or subdirectory using -offset,
                 -steps, -pad and -dirlevel
-dirlevel        in combination with -renumber, this indicates that numbering
                 of files should restart in each directory                 
-tomaster        instead of compiling, this switch creates a master document or
                 table of contents containing links to all of the markdown files
                 in the specified directory and subdirectories
-frommaster      uses the specified master table of contents, instead of a
                 directory, to compile the final markdown or Word document                                      
-quotes          "d[ouble]", all quotes will be converted to
                 double; if "s[ingle]" all quotes will be converted to single
-split           causes output files to be split into sections at secondary
                 headers ##; files will be named by the first few letters of
                 the header title; ignored if -nosplit is set
-nosplit         suppresses questions about splitting (and takes precedence)
-linknotes       if set, footnotes will be converted to ^[This is a footnote]
                 style for ease of editing
-todocx          specifies that a markdown document will be compiled to Word
-bib             specifies that pandoc should attempt to format bibtex references
                 using the specified bibtex file or a deault specified below
-notoc           suppresses output of a Table of Contents to a Word document
-tomd            specifies that a markdown document will be compiled to markdown
-nobom           writes utf8 files without a Byte Order Mark (BOM); on its own,
                 removes the BOM from all files in current directory tree
-addbom          adds a BOM to all files in current directory tree
-dumpargs        show the arguments pandoc is using
-help or ?       prints these usage details
                 
"@
    $input = Read-Host "Press any key to exit..."
    exit
}

# If we are extracting from a master ToC, we need to check the file exists and extract the directory source
if ($frommaster) {
    if (Test-Path $filename -PathType leaf) {
        $master = Get-Content -raw -Encoding "UTF8" $filename
        if ($master -match '#[^\r\n_]+_([^\r\n_]+)_[\r\n]') {
            $filename = $matches[1]
            # Get array of documents to compile
            $sourcedocs = $master | Select-String -Pattern ']\(((?:[^\)\r\n]|\)[^\r\n])*)\)[\r\n]' -AllMatches |
                % { $_.matches } |
                % { $_.groups[1].value }
        } else {
            "Master document does not contain valid directory information!"
            exit
        }
    } else {
        "Specified master document was not found! Please check your typing."
        exit
    }
}

# If the path is a file, convert it or split it 
if (Test-Path $filename -PathType leaf) {
    # Convert the file if it's a Word or other word processor document
    "Path is a file..."
    if ($filename -imatch '(.*)\.(?:docx?|rtf|md|markdown|mmd)$') {
        if ($todocx) { 
            $input = Read-Host "Do you want to clean up this document before conversion? (Y/N): "
        } 
        if ((! $todocx) -or ( $input -eq "Y")) {
            $result = $matches[1] + '_temp.md'
            $mediadir = $matches[1] + '_files'
            "Converting to $result with pandoc ..."
            $args = @('-o', $result, '-s', '-t', 'markdown-smart', '--wrap=none', "--extract-media=$mediadir", '--atx-headers', '--reference-location=section')
            # Write-Host pandoc $filename $args
            & pandoc $filename $args $dump
            $originalfile = $filename
            $filename = $result
        } else { $originalfile = $filename }
    }

# Get the contents of the document
$document = Get-Content -raw -Encoding "UTF8" $filename

    if ((! $todocx) -or ($input -eq "Y")) {
        ## Do some post-conversion cleanup ##
        "Doing post-conversion cleanup ..."
        # Replace title meta-data
        $document = $document -ireplace "-{3,}\r?\ntitle:\s*'?([^\r\n]+?)'?\r?\n-{3,}", '# $1'
        # Replace mid-document top-level headers with second-level
        $document = $document -ireplace '(\r?\n)#([^#\r\n]+)', '$1##$2'
        # Remove empty headers
        $document = $document -ireplace '\r?\n#{1,6}\s*?\r?\n', ''
        # Double blocking to single blocking
        $document = $document -ireplace '(\r?\n>\s?)>\s?([^\r\n])', '$1$2'

        if (! $todocx) {
            $input = ""
            if ( ($linknotes)) { 
                $input = Read-Host "Do you want to convert footnotes to inline notes for ease of editing? (Y/N): "
            } else { "Converting footnotes to inline notes ..." }
            if ($linknotes -or ($input -eq "Y")) {
                # Get the array of footnote refs to be converted
                $fnRefs = $document | Select-String -Pattern '\s\[\^([^\s\]]+)]:' -AllMatches |
                    % { $_.matches } |
                    % { $_.groups[1].value }
                # Replace footnotes with inline notes, but note more than one para is not supported by pandoc    
                forEach ($ref in $fnRefs) { 
                    $document = $document -ireplace "\[\^($ref)]([\s\S]+?)\[\^\1]:\s*((?:\S|\s(?![\r\n]+[[#]))+)\s+", '^[$3]$2' 
                }
            }
        }
        if (! $quotes) { 
            $quotes = Read-Host "Do you want to convert quotation marks? (S[ingle]/D[ouble]/N)"
        }
        if ($quotes -imatch '^[^n]') {
            "Converting quotation marks: $quotes ..." 
            # We first convert all quotes to double, even if user requested single, for consistency
            # First-pass conversion: single to double (assumes pandoc has already converted quotes to smart)
            $document = $document -creplace "([\W\D])‘(?!\d\ds)((?:[^’]|’(?=[\w\d])|’(?=[\s])(?<=\b\w\S+s’)(?=[^\n‘]*’[\W\D]))*)’", '$1“$2”'
            # $document | Out-File -Encoding "UTF8" 'mytestfile.md'
            # Second-pass conversion: looks for quotes-in-quotes
            $document = $document -replace '(“[^“”]*)“([^“”]*)”', "`$1‘`$2’"
        }
        if ($quotes -imatch '^s') {
            # User requested conversion to single
            # Conserve already found quotes-in-quotes with placeholders
            $document = $document -creplace "([\W\D])‘(?!\d\ds)((?:[^’]|’(?=[\w\d])|’(?=[\s])(?<=\b\w\S+s’)(?=[^\n‘]*’[\W\D]))*)’", '$1“@@@$2”@@@'
            $document = $document -replace '“(?!@@@)', "‘"
            $document = $document -replace '”(?!@@@)', "’"
            $document = $document -replace '([“”])@@@', '$1' 
        }
    }
    $input = ""
    if ((! $todocx) -and (! $split) -and (! $nosplit)) {
        $input = Read-Host @"
Do you want to split the file into smaller files on major headings 
[starting at $offset, in steps of $steps, with padding: $pad]? (Y/N): 
"@
    }
    if ((! $todocx) -and (! $nosplit) -and ($split -or ($input -eq "Y"))) {
        "Splitting file into sections ..."
        # Adding temp markers DO NOT EDIT SPACING BELOW
        $document = $document -ireplace '[\r\n]+(##)', @'

§§§$1
'@
        # Adding section break DO NOT EDIT SPACING
        $document = $document -ireplace '[\r\n]*$', @'


[SECTION_BREAK]
'@
        # Removing any doubling of section break
        $document = $document -creplace '\[SECTION_BREAK][\r\n]*(\[SECTION_BREAK])', '$1'
        # This split uses a positive lookahead to avoid selecting and splitting out the '§§§'
        $document = $document -split '(?=§§§)'
        $padlength = ([string]($document.count * $steps + $offset)).length
        $i = 0
        forEach ($section in $document) {
            # Removing section placeholders
            $section = $section -ireplace '^§§§', ''
            $prefix = " "
            #if ($section -imatch '^##\s+([^\r\n.:\\/*?"<>|]{2,30})')
            if ($section -imatch '^##\s+([^\r\n]{2,30})') 
            {
                $prefix = ($prefix + $matches[1])
                # Remove incompatible characters
                $prefix = $prefix -replace '[.:\\/*?"<>|()[\]]', '-'
                # Try to cut off at a space
                $prefix = $prefix -ireplace '\s\S{0,4}$', ''
            } else {
                $prefix = " Introduction"
            }
            # Format filenames
            # Use multiplier of $steps so that sections can be easily re-arranged
            $num = $i * $steps + $offset
            # Add leading zeros if user requested
            if ( $pad ) { $num = ([string]$num).PadLeft($padlength, "0") }
            $sectionFilename = $originalfile -ireplace '^((?:[^/\\]*?[/\\])*)(.*?)\.[^.]*$', "`$1§$num$prefix`_`$2.md"
            if ($nobom) {
                "Writing $sectionFilename with no BOM ..."
                [System.IO.File]::WriteAllLines(((Get-Item -Path ".\").FullName + ($sectionFilename -replace '^\.', '')), $section)
            } else {
                "Writing $sectionFilename ..."
                $section | Out-File -Encoding "UTF8" ($sectionFilename)
            }
            # Save changes to master document
            $document[$i] = $section
            $i = $i + 1 
        }
    } 
    if ( $todocx) {
        $toc = ""
        $input = ""
        if (! $notoc) {
            $input = Read-Host "Do you want to add a table of contents to the converted Word document? (Y/N): "
            if ($input -eq "Y") { $toc = '--toc' }
        }
        $input = ""
        if ($bib) {
            if ($library -eq "") { 
                $input = Read-Host "Enter the path and filename for the bibtex library: "
                if ($input -ne "") { $library = $input } 
            }
            if ($library -ne "") {
                $biblio = '--bibliography=' + "$library"
                "Writing $biblio..." 
            }
        } else {
           if ($library -ne "") {
                $input = Read-Host "Do you want to format references using bibtex library? (Y/N): "
                if ($input -eq "Y") { 
                    $biblio = '--bibliography=' + "$library"
                    "Writing $biblio..." 
                }
           } 
        }
        $outfile = $originalfile -ireplace '\.[^.]+$', '.docx'
        $filter = 'docx'
        # $shiftheaders = '--shift-heading-level-by=-1' 
        $args1 = @('-o', $outfile, '-s', '-t', $filter, '--wrap=none', '--reference-doc=template.docx', $biblio, '--csl="' + $csl + '"')
        "Writing output to $outfile ..."
        # Write-Host pandoc $filename $args1 $toc
        & { $OutputEncoding = [Text.Encoding]::Utf8; $document | pandoc $args1 $toc $dump }
        $input = Read-Host "[1] Do you want to open the Word document? (Y/N): "
        if ($input -eq "Y") { 
            "Launching..."
            & $outfile 
        }
    } else {
        $filename = $originalfile
        $input = ""
        if ( $filename -imatch '(.*)\.md$') { 
            $input = Read-Host "Do you want to overwrite $filename ? (Y/N): "
        } else { $input = "N" }
        if ( $input -ne "Y") {
            $filename = $filename -ireplace '(\.[^.]*)$', '_compiled.md'
        }
        if ($nobom) {
            "Writing $filename with no BOM ..."
            [System.IO.File]::WriteAllLines(((Get-Item -Path ".\").FullName + ($filename -replace '^\.', '')), $document)
        } else {
            "Writing $filename ..."
            $document | Out-File -Encoding "UTF8" ($filename)
        }
    }
    "Done."

## END OF Test-Path filename leaf

} elseif (Test-Path $filename -PathType container) {
    # Path is a directory, so either renumber, create a master or recompile and convert the document
    "Path is a directory ..."
    $resourcepath = '--resource-path="' + ($pwd.Path + ($filename -ireplace '^\.', '') -ireplace '\\', '\\') + '"'
    "Setting $resourcepath"
    # Check if the user wants to renumber the files
    $input = ""
    if ( $renumber ) {
        $input = Read-Host @"
Please confirm you wish to renumber in steps of $steps, starting with $offset, all markdown documents in $filename ?
(Pad with leading zeros: $pad; restart each dir: $dirlevel)? (Y/N): 
"@
	}
    if ( $renumber -and ($input -eq "Y")) {
        # Get array of files to rename in natural sort order
        $outfile = $filename -ireplace '[\\/]+$', ''
        $outpath = $outfile + '\*.md'
        $filearray = @(ls -r $outpath | Sort-Object { [regex]::Replace($_.FullName, '\d+', { $args[0].Value.PadLeft(20) } ) } | % { $_.FullName })
        #$filearray
        $padlength = ([string]($filearray.count * $steps + $offset)).length
        # "FlearrayCount = " + ($filearray.count) + "; Padlength = $padlength"
        $i = 0
        $n = 0
        forEach ( $file in $filearray) {
            # If dirlevel is set, restart numbering if directory has changed
            if ( $dirlevel -and $n ) {
                if ( ($file -replace '[^\\]+$', '') -ne ($filearray[$n-1] -replace '[^\\]+$', '') ) { $i = 0 }
            }
            $num = $i * $steps + $offset
            # Add leading zeros if user requested
            if ( $pad ) { $num = ([string]$num).PadLeft($padlength, "0") }
            # Add new numbering (replacing any old numbering)
            $newname = $file -ireplace '(?=[^\\]+$)(?:[^\d\n\r]*\d+\s*)?(.*\.md)$', "`§$num `$1" 
            # Remove parentheses and brackets that would break markdown hyperlinking
            $newname = $newname -replace '[()[\]]', '-'
            Write-Host $file "-->" ($newname -replace '[^\\]*\\', '')
            rename-item $file $newname
            $i = $i + 1 
            $n = $n + 1
        }
        
    } elseif ( ! $renumber) {
        # Start compile
        # Remove trailing slash(es)
        $outfile = $filename -ireplace '[\\/]+$', ''
        $outpath = $outfile + '\*.md'
        $outtype = ''
        $filter = ''
        $input = ''
        if ((! $todocx) -and (! $tomd) -and (! $tomaster)) { 
            if (! $frommaster) { 
                $input = Read-Host "Do you want to convert markdown files in this directory to Word ( $filename )? (Y/N): " 
            } else { 
                $input = Read-Host @"
Do you want to convert markdown files listed in the master table of contents to Word?
(If you answer N, then they will be compiled to markdown format) (Y/N): 
"@  
            }
        } else { $input = "Y" }
        if (($tomd) -or ($tomaster) -or ($input -ne "Y")) { 
            if (! $tomaster) { 
                "Compiling to single markdown file ..." 
                $outtype = '_compiled.md'
                $filter = 'markdown-smart'
            } else { 
                "Creating a master document ..."
                $outtype = '_master.md' 
            }
        } elseif (($todocx) -or ($input -eq "Y")) {
            "Compiling to single Word file ..."
            $outtype = '.docx'
            $filter = 'docx' 
            $input = ""
            if ($bib) {
                if ($library -eq "") { 
                    $input = Read-Host "Enter the path and filename for the bibtex library: "
                    if ($input -ne "") { $library = $input } 
                }
                if ($library -ne "") {
                    $biblio = '--bibliography="' + $library + '"'
                    "Writing $biblio..." 
                }
            } else {
               if ($library -ne "") {
                    $input = Read-Host "Do you want to format references using bibtex library? (Y/N): "
                    if ($input -eq "Y") { 
                        $biblio = '--bibliography=' + "'" + $library + "'"
                        "Writing $biblio..." 
                    }
               } 
            }
        }
        #$mediadir = $outfile + '_files'
        $outfile = $outfile + $outtype
        if (! $notoc) { $toc = '--toc' }
        #if ($outtype -imatch '\.docx$') { $shiftheaders = '--shift-heading-level-by=-1' }
        $args1 = @('-o', $outfile, '-s', '-t', $filter, '--wrap=none', '--atx-headers', '--reference-location=section', $biblio, "--csl=$csl", '--top-level-division=chapter', '--reference-doc=template.docx') #"--extract-media=$mediadir"
        "Writing output to $outfile ..."
        if ($tomaster) {
            $table = @(ls -r -name ($outpath -replace '\*\.md', '') *.md | Sort-Object { [regex]::Replace($_, '\d+', { $args[0].Value.PadLeft(20) }) })
            $thatprefix = ''
            $tableofcontents = forEach ($entry in $table) {
                $snippet = ""
                $filecontent = Get-Content -raw -Encoding "UTF8" ($filename + '\' + $entry)
                if ($filecontent -match '#\s+([^\r\n]+)') {
                    $mainheader = $matches[1]
                } else {
                    $mainheader = $entry -replace '[^\\]+\\', ''
                }
                $thisprefix = ''
                if ($entry -match '(?:[^\\]+\\)') { $thisprefix = $matches[0] }
                if ( $thisprefix -ne $thatprefix) { 
                    '   * ' + ($thisprefix -replace '\\', '')
                }
                # Escape any parentheses
                $entry = $entry 
                "      + [$mainheader](" + (($entry -replace '.+\\([^\\]+)\.md$', '$1') -replace '([()])', '\$1') + ')'
                # Find a text snippet
                # Remove carriage returns except in headings
                # $filecontent = $filecontent -replace '(?<!#[^\n]+)\r?\n(?!#)', ' '
                # Remove any inline footnotes
                $filecontent = $filecontent -replace '\^\[[^\]]+]', ''
                if ($filecontent -cmatch ('[\r\n]+([A-Z¡¿*_[(\\][^#\r\n]+?\.)\s+[A-Z#>\\[(¿¡"' + "']")) {
                    $snippet = $matches[1]
                    # Remove any quotation markup
                    $snippet = $snippet -replace '>\s?', ''
                    # Format snippet for best display
                    '          - ' + $snippet -replace '(.{80}\S*)\s', '$1
               '
                }
                $thatprefix = $thisprefix
            }
            # Write the full table to the master file
            $document = ("# Table of Contents for `_$filename`_`n`n" + ( $tableofcontents | % { $_ + "`n" }))
            if ($nobom) {
                "Writing with no BOM ..."
                [System.IO.File]::WriteAllLines(((Get-Item -Path ".\").FullName + $outfile), $document)
            } else {
                $document | Out-File -Encoding "UTF8" ($outfile)
            }
        } elseif ($frommaster) {
            & pandoc @($sourcedocs | % { ls -r $filename ($_ + '.md') | % { $_.FullName } }) $args1 $toc $dump
        } else { 
            # NB Sorting by .Name as below causes them to be sorted without taking into account the path! 
            # Write-Host @(ls -r $outpath | Sort-Object { [regex]::Replace($_.FullName, '\d+', { $args[0].Value.PadLeft(20) }) } | % { $_.FullName }) $args1
            & pandoc @(ls -r $outpath | Sort-Object { [regex]::Replace($_.FullName, '\d+', { $args[0].Value.PadLeft(20) }) } | % { $_.FullName }) $args1 $toc $resourcepath $dump
            if ($outtype -imatch '\.docx') {
                $input = Read-Host "[2] Do you want to open the Word document? (Y/N): "
                if ($input -eq "Y") { 
                    "Launching..."
                    & $outfile 
                }
            }
        }
        "Done."
    } else { "Renumber operation aborted!" }
} 


}