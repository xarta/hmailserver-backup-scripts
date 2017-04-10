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

Xarta.json as of 8th Apr 2017 with passwords removed:

```json
{
	"bkupKeep": {
		"keepYears": 7,
		"keepMonths": 24,
		"keepWeeks": 26,
		"keepDays": 60
	},
	"tasks": {
		"ApprovedDeleteOldBackUps": {
			"TN": "hMailServer DO NOT RUN HERE",
			"SC": "ONCE",
			"D": "01/01/2050",
			"ST": "06:00"
		},
		"ScheduledDeleteOldBackUps": {
			"TN": "hMailServer Prune old back-ups",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "04:00"
		},
		"BkUpHMSsettings": {
			"TN": "hMailServer Settings Backup",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "01:00"
		},
		"BkUpMySql": {
			"TN": "hMailServer mySql Dump",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "01:15"
		},
		"BkUpHMSdata": {
			"TN": "hMailServer 7zip data",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "01:25"
		},
		"DeleteHMSsettings": {
			"TN": "hMailServer Delete tmp settings bkup",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "02:00"
		},
		"DeleteSqlDump": {
			"TN": "hMailServer Delete MySQL Dump",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "02:30"
		},
		"CopyHMSsettings": {
			"TN": "hMailServer 7zip settings",
			"SC": "WEEKLY",
			"D": "MON,TUE,WED,THU,FRI,SAT,SUN",
			"ST": "01:05"
		}
	},
	"hMailServer": {
		"User": "Administrator",
		"Password": "BLAH"
	},
	"mySQL": {
		"backup": {
			"User": "dump",
			"Password": "BLAH"
		},
		"hmailserver": {
			"User": "hMailServer",
			"Password": "BLAH"
		},
		"test": {
			"User": "test",
			"Password": "test"
		}
	},
	"7zip": {
		"Password": "BLAH",
		"test": "test"
	},
	"network": {
		"User": "admin",
		"Password": "BLAH"
	},
	"windowsAccounts": {
		"scheduler": {
			"User": "XartaTask",
			"Password": "BLAH",
			"Group": "Administrators",
			"Fullname": "XartaTasks admin",
			"Description": "Admin for scheduler tasks when XartaMail not logged on"
		},
		"mailservice": {
			"User": "XartaMail",
			"Password": "BLAH",
			"Group": "Users",
			"Fullname": "hMailServer User",
			"Description": "Less priviledged user for hMailServer"
		},
		"testonly": {
			"User": "XartaTest",
			"Password": "BLAH",
			"Group": "Users",
			"Fullname": "Mr Xarta Test",
			"Description": "Just for test use in scripting"
		}
	},
	"paths": {
		"mysqlexe": "C:\\Program Files (x86)\\MySQL\\MySQL Server 5.7\\bin\\mysql.exe",
		"mysqlini": "F:\\sql\\prog\\my.ini",
		"mysqldumpexe": "C:\\Program Files (x86)\\MySQL\\MySQL Server 5.7\\bin\\mysqldump.exe",
		"mysqlcheckexe": "C:\\Program Files (x86)\\MySQL\\MySQL Server 5.7\\bin\\mysqlcheck.exe",
		"mysqldumpoutput": "G:\\mysql_dump",
		"mysqldumpdefaultsextrafile": "F:\\sql\\prog\\my.cnf",
		"hmdata": "G:\\hMailServer\\Data",
		"uncServer": "\\\\XWIFI02",
		"uncPath": "\\USBDisk1_Volume1",
		"hmsettingsbkup": "G:\\settings_backup",
		"hmcertificates": "G:\\certificates",
		"hmini": "G:\\hMailServer\\Bin\\hMailServer.INI"
	}
}
```

Little "servers" go in this cupboard ... https://github.com/xarta/hmailserver-backup-scripts/tree/master/pics/cupboard

![Picture of plain cupboard doors](/pics/cupboard/20170405_162808.jpg?raw=true "Cupboard for my little servers etc.")
![Distance shot of open cupboard with little servers](/pics/cupboard/20170405_162147.jpg?raw=true "Cupboard for my little servers etc.")
