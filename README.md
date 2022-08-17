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

## Start server

### development
```
cd BeamQC
set FLASK_ENV=development
flask run
```
### production(using waitress as WSGI server)
```
python wsgi.py
```

## File structure
### `nginx`
```
|__ /
  |__ nginx
    |__ nginx.conf
```
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