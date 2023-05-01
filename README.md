# Usage

## Prequisites
* Install python on your machine from - https://www.python.org/downloads/
* Clone this repository, or download this repository as a zip and unzip it to a location
* Make sure you have at least firefox or chrome installed. The script current supports just these two browsers
* Open your cmd / terminal and go the location where you have unzipped this directory 
    * `cd <location_of_the_directory>`
* You will notice an `uscis_info.xlsx` file. Add your entries to this file and save the file. **Close `Microsoft Excel` before running the script**
* Run the script as
```
python uscis_form_submitter.py -b firefox 
```

Script takes the following arguments
```
-b <either firefox or chrome> (required field)
-c <specify a value between 5 or 10> (optional) -> Controls the number of tabs to be opened at a time. Default is 5
-f <absolute path of the file> (optional) -> Make sure the excel file follows the template in this repository. Default is uscis_info.xlsx
```