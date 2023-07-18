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
#  @author 0verfl0w767
#  @link https://github.com/0verfl0w767
#  @license MIT LICENSE
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
                                           __                                         _           __ 
              _______  ____  __      _____/ /___ ___________      ____  _________    (_)__  _____/ /_
             / ___/ / / / / / /_____/ ___/ / __ `/ ___/ ___/_____/ __ \/ ___/ __ \  / / _ \/ ___/ __/
            (__  ) /_/ / /_/ /_____/ /__/ / /_/ (__  |__  )_____/ /_/ / /  / /_/ / / /  __/ /__/ /_  
           /____/\__, /\__,_/      \___/_/\__,_/____/____/     / .___/_/   \____/_/ /\___/\___/\__/  
                /____/                                        /_/              /___/                 
          
          Unofficial su-wings (SAHMYOOK UNIV.) lecture information system.
          Github: https://github.com/syu-kr/lecture.syu.kr-everytime
          Author: 0verfl0w767 (https://github.com/0verfl0w767)
          Version: 1.0v
          License: MIT LICENSE
          
    """
    sys.stdout.write(logo + "\n")
  
  def info(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [INFO] " + text + "\033[0m" + "\n")
  
  def warnning(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [WARNNING] " + text + "\033[0m" + "\n")
  
  def debuggerInfo(self, text: str) -> None:
    sys.stdout.write("[" + self.getTime() + "] [DEBUG] " + text + "\033[0m" + "\n")
  