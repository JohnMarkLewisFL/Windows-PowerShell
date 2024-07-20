# This script will automatically close Microsoft Teams and clear the cache files/folders


# Stops the Teams process and opens a confirmation dialog box before proceeding
Stop-Process -Name *teams -Confirm

# Deletes the Teams AppData files as specified in the path and name
Get-ChildItem "C:\Users\*\AppData\Roaming\Microsoft\Teams\*" -directory | Where-Object name -in ('application cache','blob storage','databases','GPUcache','IndexedDB','Local Storage','tmp') | ForEach-Object{Remove-Item $_.FullName -Recurse -Force}