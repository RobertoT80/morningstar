This is a powershell script useful to batch retrieve values from a list of funds you want to monitor in the Morningstar website.
The values are retrieved in the html code obtained via basic http requests on port 80 so no need to enable API keys or open firewalls.

The ID of the funds have to be provided in a text file, each in one line.
The values will be written to the current directory in a csv file named 'fund_list_<date>.csv'.

To run it, simply execute it in your CWD and pass your file list as a parameter, the csv will be output there.
.\get-fundValues.ps1 -listFile 'C:\test\fund_list.txt'
