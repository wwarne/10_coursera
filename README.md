# Coursera Dump

Script grabs information from the [Coursera](https://coursera.org/) website and structurize it to Excel file.
It collects the following:
1. Title of the course
2. Languages
3. Rating
4. Start date 
5. Duration of the course
6. Course url

# How to use

1. Install requirements `pip install -r requirements.txt`
2. Run command `python coursera.py -f [FILEPATH] -n [NUMBER OF COURSES TO PROCESS]` 

Without additional arguments script will process 20 random courses and save information to a file called `courses.xlsx`. 

## For example
`python coursera.py` will save information about 20 random courses to a file `courses.xlsx` placed in the same directory where the script is.

`python coursera.py -f /home/user/my_courses.xlsx -n 100` will save information about 100 courses to a file `/home/user/my_courses.xlsx`

If you need information about all available courses just provide a big number like `python coursera.py -f courses.xlsx -n 10000`

# Build with

* [requests](http://docs.python-requests.org/en/master/) 
* [openpyxl](https://openpyxl.readthedocs.io/en/default/) 
* [lxml](http://lxml.de/)

# Project Goals

The code is written for educational purposes. Training course for web-developers - [DEVMAN.org](https://devman.org)
