# Introduction
A simple script that can Convert xlsx to csv.
Support the xlsx cells tha have formulas.

# Usage
- Download the release version , use the terminal. (Recommend)

USAGE:
    xlsx2csv [OPTIONS] <input> [output]

ARGS:
    <input>     Input XLSX file
    <output>    Output CSV file (optional). if no value , use sheet name as output

OPTIONS:
    -h, --help              Print help information
    -s, --sheet <sheet>     Sheet name (optional)
    -u, --use-sheet-name    Use sheet name as output (optional)

Example:
on windows, make a .bat file, copy the code below, you can change "test.xlsx" as you file name.
```bat
@Echo Off
%1 mshta vbscript:createobject("shell.application").shellexecute("%~s0","::","","runas",1)(window.close)&exit
cd /d %~dp0

xlsx2csv test.xlsx

pause
```

- You can also use download the code and build it . Make sure that you are familiar with rust.
```powershell
cargo build --release
cargo run examples/test.xlsx 
```

# License
MIT LICENSE

Fell free to use the script. 

If the script help you , fell free to give it a star.

