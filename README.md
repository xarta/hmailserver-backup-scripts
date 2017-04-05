# hmailserver-backup-scripts
I'm working on VBScripts, batch scripts, and a needed PowerShell shim, for my hMailServer installation on my Compute Stick (with a back-up installation in a VM). They'll sit in a protected area (physical subnet partitioned by pfSense), with another two hMailServer installations in the DMZ for incoming relays with anti-spam and anit-virus etc.  I'm trying to grow some general use scripts and will include more for set-up and management over time, but mostly it's about back-up at the moment.

It's a work in progress that I can only come back to rarely / occasionally, and is particularly for my personal circumstances.

Basically, I use a mounted drive for hMailServer including Datafile "G: drive", a mounted drive for MySQL "F: drive" and want to back-up data, MySQL, and hMailServer settings (incl. domains) to encrypted 7zip files on a HooToo Travel Router with a USB flash drive mounted as a Samba Share (independent of domain etc.).

For MySQL dumps, my script generates a my.cnf file for the credentials.

To get around "logged on or not" task schedular issues with VBScripts/Batch files, my VBScript calls a PowerShell script with Execute Policy set to bypass which in turns calls my batch file for the 7zip.

I'm trying to work toward having all credentials, paths, and relevant settings for my scripts in one JSON file which I can protect with NTFS settings.

The hMailServer service is set to run as a normal user.  A separate local admin account is used to run the scripts - valid on that machine only.

Lots to do ... i.e .deleting/managing back-ups etc. to tie into a back-up management process running on a different machine.  Tidying-up (lots) ... just lots of work.

Soon I hope to document some of this along with hMailServer installation and configuration on my blog: blog.xarta.co.uk

