# Auto-Join-Zoom-Meeting
Join a zoom meeting with out filling in meeting id's or passcodes, one button for it all! 

## Setup

See attached [excel document](https://github.com/JareBear12418/Auto-Join-Zoom-Meeting/blob/main/List.xlsx). MAKE sure it's filled out with the correct credentials.


### Sample 
| MEETING NAME | MEETING ID | MEETING PASSCODE (NOT REQUIRED) 
| --- | --- | --- |
| work meeting at 5 | 1234567890 | this_is_password |
| school | 3948329482 |  |

You may need to build it your self. Windows doesn't like unlicensed software with the ability to control inputs being distributed.

## NOTICE

Zoom must be running for this program to function.

![image](https://user-images.githubusercontent.com/25397800/151075856-bd2c694b-12b6-45fe-8582-805b43cf1fba.png)

## Demo 
<p align="center">
    <img height="400" src="https://user-images.githubusercontent.com/25397800/151076170-dea2944b-8bab-4085-8620-0ce658a5f16a.png" />
</p>

## Build

1. Create a virtual environment: `virtualenv venv`
2. Activate it: `venv/Scripts/Activate.ps1`
3. Install requirements: `pip install -r requirements.txt`
4. To compile: `pyinstaller -F --noconsole --icon=images/icon.ico main.py --name="Auto Join Zoom Meeting" `
5. `.exe` will be located in a folder called `dist`
