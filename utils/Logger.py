#
#     ____                  __________
#    / __ \_   _____  _____/ __/ / __ \_      __
#   / / / / | / / _ \/ ___/ /_/ / / / / | /| / /
#  / /_/ /| |/ /  __/ /  / __/ / /_/ /| |/ |/ /
#  \____/ |___/\___/_/  /_/ /_/\____/ |__/|__/
# 
#  The copyright indication and this authorization indication shall be
#  recorded in all copies or in important parts of the Software.
# 
#  @link https://github.com/0verfl0w767
#
import datetime
import sys

class Logger:
  def __init__(self):
    pass
  
  def getTime(self) -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  
  def logo(self) -> None:
    logo = f"""

     .d8888b.  888     888 888       888 8888888 888b    888  .d8888b.   .d8888b.        .d8888b.   .d88888b.  888b    888 888     888 8888888888 8888888b. 88888888888 8888888888 8888888b.  
    d88P  Y88b 888     888 888   o   888   888   8888b   888 d88P  Y88b d88P  Y88b      d88P  Y88b d88P" "Y88b 8888b   888 888     888 888        888   Y88b    888     888        888   Y88b 
    Y88b.      888     888 888  d8b  888   888   88888b  888 888    888 Y88b.           888    888 888     888 88888b  888 888     888 888        888    888    888     888        888    888 
     "Y888b.   888     888 888 d888b 888   888   888Y88b 888 888         "Y888b.        888        888     888 888Y88b 888 Y88b   d88P 8888888    888   d88P    888     8888888    888   d88P 
        "Y88b. 888     888 888d88888b888   888   888 Y88b888 888  88888     "Y88b.      888        888     888 888 Y88b888  Y88b d88P  888        8888888P"     888     888        8888888P"  
          "888 888     888 88888P Y88888   888   888  Y88888 888    888       "888      888    888 888     888 888  Y88888   Y88o88P   888        888 T88b      888     888        888 T88b   
    Y88b  d88P Y88b. .d88P 8888P   Y8888   888   888   Y8888 Y88b  d88P Y88b  d88P      Y88b  d88P Y88b. .d88P 888   Y8888    Y888P    888        888  T88b     888     888        888  T88b  
     "Y8888P"   "Y88888P"  888P     Y888 8888888 888    Y888  "Y8888P88  "Y8888P"        "Y8888P"   "Y88888P"  888    Y888     Y8P     8888888888 888   T88b    888     8888888888 888   T88b
     
    """
    sys.stdout.write(logo + "\n")
  
  def info(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [INFO] " + text + "\033[0m" + "\n")
  
  def warning(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [WARNNING] " + text + "\033[0m" + "\n")
  
  def debuggerInfo(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [DEBUG] " + text + "\033[0m" + "\n")
  
