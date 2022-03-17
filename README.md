# Excelerate
This project aims to fix "broken" Excel files that load extremely slowly due to broken references.  It does that by searching for broken references and removing them from the Excel file.


## Running
This project is intended to be run on PowerShell Core.  All plugins should ideally run on all OSes.


## Building
This project may be built using `make`.  To build, simply run `make` then `make test` to run the Pester tests.

For ease of testing, you may also run `make shell` to dump you into a shell with the module imported.


## Installing
Note that `make install` is **NOT** currently implemented.  Installation is currently manual.  In the future we will try to implement `make install`.

To install, copy all the files in the "build" directory into a new folder in your PowerShell modules directory.


## Public functions
The following functions are publicly exported from this module for use.  These functions may be used either manually, or in automated jobs.

* Repair-XLSX
  This function will attempt to repair an XLSX file by removing broken references
