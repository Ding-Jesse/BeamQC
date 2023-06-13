# RC checker webpage

## Overview

Input: XS-PLAN, XS-BEAM, XS-COL, and choose the things you want to check, such as whether there are something miss, whether the size is same in the drawing. 

Output: The result with text file and mark-on drawing file. 

## Install

Requires Python 3.7+ and Pip. 

Clone this repo: 
```
git clone https://github.com/bamboochen92518/BeamQC
```
### Require Package (Update in 2023/06/13)
- using pip to install fpdf2 
```
conda install python=3.9 matplotlib=3.7.1 pandas=1.5.3 numpy=1.24.3 flask=2.2.2 waitress=2.0.0 openpyxl=3.0.10 flask-session requests xlsxwriter
conda install -c conda-forge flask-mail
conda install -c anaconda pywin32
pip install fpdf2 greenlet 
```

## File Structure
### `nginx`
```
|__ /
  |__ nginx
    |__ nginx.conf
```

Click here to know how to install nginx ---> https://hackmd.io/@bamboochen92518/BJBguM5A5

### `flask`
```
|__ /BeamQC
  |__ /INPUT   <--┐
  |__ /OUTPUT  <--| You need to add these directories by yourself or change the directory in python code. 
  |__ /result  <--┘
  |__ /static
  |  |__ /css
  |    |__ rubic.css
  |  |__ lots of files or pictures for webpage
  |__ /templates 
  |  |__ base.html
  |  |__ lots of html files
  |__ app.py
  |__ main.py
  |__ plan_to_beam.py
  |__ plan_to_col.py
  |__ wsgi.py
```

## Getting Started

### Start server

#### development
```
cd [PATH_TO_BEAMQC]
set FLASK_ENV=development
flask run
```
Then you can visit the webpage from http://localhost:[port_number]

#### production(using waitress as WSGI server)
```
cd /nginx
start nginx.exe
cd [PATH_TO_BEAMQC]
python wsgi.py
```
Then you can visit the webpage from http://localhost

## Others

Visit our webpage and try it by yourself!!!

Link: http://www.freerccheck.com.tw
