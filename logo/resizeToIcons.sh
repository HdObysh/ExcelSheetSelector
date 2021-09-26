#!/bin/zsh
#########################################
# [USAGE]
# ./resizeToIcons.sh [source file path]
#
# THIS SCRIPT ONLY WORKS ON MACOS!
#########################################

script_dir=$(cd $(dirname $0); pwd)
inputFile=$1
sizes=(16 32 64 80)
for size in $sizes; do
    sips -Z $size $inputFile -o $script_dir/../assets/icon-$size.png
done
