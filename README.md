# Usage

## Prequisites
* Install python on your machine from - https://www.python.org/downloads/
* Clone this repository, or download this repository as a zip and unzip it to a location
* Make sure you have at least firefox or chrome installed. The script current supports just these two browsers
* Open your cmd / terminal and go the location where you have unzipped this directory 
    * `cd <location_of_the_directory>`
* Run (only once)
    ```
    pip install -r requirements.txt
    ```
* You will notice an `uscis_info.xlsx` file. Add your entries to this file and save the file. 

Some important notes before running the script

1. **Close `Microsoft Excel` before running the script**

2. **Do not close the browser window or tabs after you have finished submitting , the script will handle that for you. Manually closing the tabs will interfere with the logic of the script**

3. **There will be always be a blank tab , do not close that tab and leave it as is**


Run the script with (These are just example commands, tweak them according to your usage).
```
python uscis_form_submitter.py -b firefox -s "Consolidated Sheet" -fname foo -lname bar -e foobar@bar.com -p 123456789
```
Note if you are using mac you may need to use

```
python3 uscis_form_submitter.py -b firefox -s "Consolidated Sheet" -fname foo -lname bar -e foobar@bar.com -p 123456789 
```

Script takes the following arguments
```
-b <either firefox or chrome> (required field)
-s <sheet name within the excel file> <required>
-c <specify a value between 5 or 10> (optional) -> Controls the number of tabs to be opened at a time. Default is 5
-f <absolute path of the file> (optional) -> Make sure the excel file follows the template in this repository. Default is uscis_info.xlsx
-fname <Your first name> (optional)
-lname <Your last name> (optional)
-p <phone number> (optional)
-e <email> (optinal)
```