#!/bin/bash
MSG=${1:-"Update nested-sort-matrix"}
git add -A
git commit -m "$MSG"
git push
echo "Pushed: $MSG"
