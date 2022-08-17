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

## Other directory