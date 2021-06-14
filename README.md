# ExecJS
This Android utility simulates a webpage on your Android Chrome Browser in which you can type javascript commands and press Enter to execute them. The result of each command will appear immediately after executing.

**Note:** As of 14 June 2021, the Desktop version of ExecJS is also available. To use it, download the file ExecJS.js from this repository, and run it by double-clicking it in File Explorer.

## Installation requirements:
1. Termux
2. Android Chrome Browser
3. Termux git and python package. If you don't have git and python installed, install them by
```
pkg install git python
```
**This app does NOT require root access**

## Installation and usage instructions:
1. Clone repository
```
git clone http://github.com/javadbayat/ExecJS
```
2. To launch the utility, first jump into the directory
```
cd ExecJS
```
3. Then, execute the following:
```
python ExecJS.py
```  

4. The Chrome Browser will pop up. You may see an error message in Chrome due to technical issues. If so, refresh the page after a couple of seconds.
Then the app will soon show up with a black screen. To execute commands, tap on the screen and start typing with your keyboard.
5. To quit the app, close Chrome browser, go back to Termux, and press Ctrl+C.
