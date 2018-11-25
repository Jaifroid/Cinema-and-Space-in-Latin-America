#!/bin/bash
opDir="$1"
if [ -z "$opDir" ]; then
    echo "You must enter a directory to operate on!"
    exit
fi
find "$opDir" -iname \*.md -type f
echo
echo 'This script will remove the BOM (if found) from all the above files in "'"$opDir"'" and all subdirectories'
echo -n "Are you sure you want to proceed? (y/n)"
read answer
echo
if [ "$answer" != "${answer#[Yy]}" ]; then
    echo "Processing..."
    find "$opDir" -iname \*.md -type f -exec sed -i '1s/^\(\xef\xbb\xbf\)\?//' {} +
    echo "Done!"
else
    echo "Command aborted!"
fi
echo
