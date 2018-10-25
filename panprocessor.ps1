param (
    [string]$filename = "",
	[int]$offset = 0,
	[int]$steps = 5,
	[switch]$pad = $false,
	[switch]$renumber = $false,
	[switch]$quotes = $false,
	[switch]$split = $false,
	[switch]$linknotes = $false,
	[switch]$todocx = $false,
	[switch]$tomd = $false,
	[switch]$help = $false
)

if (($filename -eq "") -and (! $help)) { 
    $filename = Read-Host "Enter the filename or directory for processing, or type ? for help: "
}

# Display help options if user requested
if (($filename -eq "?") -or ($help)) {
@"

Usage: .\panprocessor FILENAME or DIRECTORY [-offset num] [-steps num] [-pad]
                [-renumber] [-quotes] [-split] [-linknotes] [-todocx] [-help]

Uses pandoc to convert a Word file to markdown and optionally split it into
numbered sections. Can also be used to compile a set of split markdown files
either into a single markdown document, or into a Word document. Can be used to
renumber all the files in a directory and subdirectories. Can also "tidy" a
single markdown file. If the script is invoked with no options, it will prompt
you to specify them as needed.

FILENAME or      the filename or directory on which to operate (can use absolute
  DIRECTORY      or relative path); if entering a filename, the script will 
                 assume you wish to convert it to markdown; if a directory, it
                 will assume you wish to compile split files to a document
-offset num      when splitting files, provide an offset number to start from,
                 otherwise it will start from 0
-steps num       when spliiting files, specify the numbering step; this is
                 useful if you need to re-order files manually later; default 5
-pad             [switch]: if present, numbers will be padded with zeros, like
                 01, 02, 03, or 010, 020, 030, etc., as needed; useful if your
                 filesystem orders files ASCIIbetically rather than naturally
-renumber        just renumbers files in directory or subdirectory using -offset
                 -steps and -pad
-quotes          if set, all single quotes will be converted to double
-split           causes output files to be split into sections at secondary
                 headers ##; files will be named by the first few letters of
                 the header title
-linknotes       if set, footnotes will be converted to ^[This is a footnote]
                 style for ease of editing
-todocx          specifies that a markdown document will be compiled to Word
-tomd            specifies that a markdown document will be compiled to markdown
-help or ?       prints these usage details
                 
"@
    $input = Read-Host "Press any key to exit..."
    exit
}

# If the path is a file, convert it or split it 
if (Test-Path $filename -PathType leaf) {
    # Convert the file if it's a Word document
    "Path is a file..."
    if ($filename -imatch '(.*)\.(?:docx?|rtf|md|markdown|mmd)$') 
    {
        $result = $matches[1] + '_temp.md'
        "Converting to $result with pandoc ..."
        $args = @('-o', $result, '-s', '-t', 'markdown-smart', '--wrap=none', '--extract-media=.', '--atx-headers', '--reference-location=section')
        & pandoc $filename $args
        $originalfile = $filename
        $filename = $result
    }

    # Get the contents of the document
    $document = Get-Content -raw -Encoding "UTF8" $filename

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

    if ( ! ($linknotes)) { 
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
    if ( ! ($quotes)) { 
        $input = Read-Host "Do you want to convert any single quotation marks to double? (Y/N): "
    } else { "Converting single quotation marks to double ..." }
    if ($quotes -or ($input -eq "Y")) {
        $document = $document -ireplace "([\s([])[''‘’](?!\d\ds)(.*?)[''‘’](?!\w)([\W\D])", '$1“$2”$3'
    }
    if ( ! ($split)) {
        $input = Read-Host @"
Do you want to split the file into smaller files on major headings 
[starting at $offset, in steps of $steps, with padding: $pad]? (Y/N): 
"@
    } else { "Splitting file into sections ..."}
    if ($split -or ($input -eq "Y")) {
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
            if ($section -imatch '^##\s*([^\r\n.:\\/*?"<>|]{2,30})') 
            {
                $prefix = ($prefix + $matches[1])
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
            "Writing $sectionFilename ..."
            $section | Out-File -Encoding "UTF8" ($sectionFilename)
            # Save changes to master document
            $document[$i] = $section
            $i = $i + 1 
        }
    } 
    $filename = $originalfile
    if ( $filename -imatch '(.*)\.md$') { 
        $input = Read-Host "Do you want to overwrite $filename ? (Y/N): "
    } else { $input = "N" }
    if ( $input -ne "Y") {
        $filename = $filename -ireplace '(\.[^.]*)$', '_compiled.md'
    }
    "Writing $filename..."
    $document | Out-File -Encoding "UTF8" ($filename)

## END OF Test-Path filename leaf

} elseif (Test-Path $filename -PathType container) {
    # Path is a directory, so either renumber or recompile and convert the document
    "Path is a directory ..."
    
    # Check if the user wants to renumber the files
    if ( $renumber ) {
        $input = Read-Host @"
Please confirm you wish to renumber in steps of $steps, starting with $offset, all markdown documents in $filename ?
(Pad with leading zeros: $pad)? (Y/N): 
"@
	}
    if ( $renumber -and ($input -eq "Y")) {
        # Get array of files to rename in natural sort order
        $outfile = $filename -ireplace '[\\/]+$', ''
        $outpath = $outfile + '\*.md'
        $filearray = @(ls -r $outpath | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20) } ) } )
        $padlength = ([string]($filearray.count * $steps + $offset)).length
        # "FlearrayCount = " + ($filearray.count) + "; Padlength = $padlength"
        $i = 0
        forEach ( $file in $filearray) {
            $num = $i * $steps + $offset
            # Add leading zeros if user requested
            if ( $pad ) { $num = ([string]$num).PadLeft($padlength, "0") }
            Write-Host ($file -replace ".*?(§\d+.*$)", '$1') "-->" ($file -replace ".*?§\d+(.*$)", "`§$num`$1")
            rename-item $file ($file -replace "§\d+", "`§$num")
            $i = $i + 1 
        }
        
    } elseif ( ! $renumber) {
        # Start compile
        # Remove trailing slash(es)
        $outfile = $filename -ireplace '[\\/]+$', ''
        $outpath = $outfile + '\*.md'
        $outtype = ''
        $filter = ''
        if ((! $todocx) -or (! $tomd)) { 
            $input = Read-Host "Do you want to convert markdown files in this directory to Word ( $filename )? (Y/N): "
        } else { $input = "Y" }
        if (($tomd) -or ($input -ne "Y")) { 
            "Compiling to single markdown file ..."
            $outtype = '_compiled.md'
            $filter = 'markdown-smart'
        } elseif (($todocx) -or ($input -eq "Y")) {
            "Compiling to single Word file ..."
            $outtype = '.docx'
            $filter = 'docx' 
        }
        
        $outfile = $outfile + $outtype
        $args1 = @('-o', $outfile, '-s', '-t', $filter, '--wrap=none', '--extract-media=.', '--atx-headers', '--reference-location=section', '--top-level-division=chapter', '--toc', '--reference-doc=template.docx')
        "Writing output to $outfile ..."
        # Sort-Object below ensures Natural Order instead of ASCIIbetical - see https://stackoverflow.com/questions/5427506
        # & Write-Host @(ls -r $outpath | Sort-Object { [regex]::Replace($_, '§\d+', { $args[0].Value.PadLeft(20) }) } | % { $_ }) $args1
        # NB Sorting by .Name as below causes them to be sorted without taking into account the path! 
        # & Write-Host @(ls -r $outpath | Sort-Object { [regex]::Replace($_.Name, '§\d+', { $args[0].Value.PadLeft(20) }) } | % { $_ }) $args1
        & pandoc @(ls -r $outpath | Sort-Object { [regex]::Replace($_, '\d+', { $args[0].Value.PadLeft(20) }) } | % { $_ }) $args1
        "Done."
    } else { "Renumber operation aborted!" }
} 
