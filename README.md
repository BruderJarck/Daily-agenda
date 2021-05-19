# Daily-agenda
Checks your outlook calendar and shows your appointments(+its duration) during a given day in a simple table. 

## disclaimer
The content of this readme focuses on Windows(win 10 tested only). 

Latest docs and exe can be found in artifacts of last ci run.

## dependencies
- python 3.7 https://www.python.org/downloads/release/python-379/
- git installed https://git-scm.com/downloads
- pip installed https://pip.pypa.io/en/stable/installing/
- content of requirements.txt

## installation 

in your terminal:

  ```
  git clone https://github.com/JackWolf24/Daily-agenda.git
  ```
  
  cd in repo
  
  ```
  pip install -r requirements.txt
  ```
  
  ## build exe

  cd into daily_agenda/
  
  ```
  pyinstaller daily_agenda.spec
  ```
  
  after process beeing finished, the exe file is located in 
  
  ```
  daily_agenda/dist
  ```
