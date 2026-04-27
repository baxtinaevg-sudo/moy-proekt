# TaskComplete Hook - Sound notification when task is completed
# Plays Asterisk system sound to notify user that Cline finished work
[System.Media.SystemSounds]::Asterisk.Play()
Start-Sleep -Milliseconds 500