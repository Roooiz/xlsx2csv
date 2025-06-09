This project created for help converted .xlsx to .csv files

# Launch
1. Clone this repository
2. While on the blade level with the xlsx2csv.go file, execute the command `go build`
3. Launch programm `xlsx2csv` on Linux

```bash
./xlsx2csv -input=path/to/dir/or/file -outdir=path/to/dir/to/save/csv -enc=utf8
```
## Launch options
 -  `-input` - path to XLSX file or folder with XLSX files 
 - `-outdir` - the path to the folder to save CSV (if the entrance is the folder)
 - `-enc` - coding CSV (UTF-8 by default, UTF-16le, UTF-16BE, Windows-1251, ISO-8859-1)
 - `-sheet` - the name of the sheet for conversion (by default is the first)
 - `-delim` - Field separator in CSV (by default ';')
