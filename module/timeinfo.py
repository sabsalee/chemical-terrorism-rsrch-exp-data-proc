from datetime import datetime, timedelta

class TimeInfo:
    def __init__(self, duration=600) -> None:
        self.__exist = False
        self.__starttime = None
        self.__duration = duration
        self.timedatalist = None

    def set_time(self, time: datetime|list, only_time=False):
        if not only_time:
            if time.second >= 30:
                time = time.replace(second=0)
                time = time + timedelta(minutes=1)
            else:
                time = time.replace(second=0)

            self.__starttime = time
            self.__exist = True
        else:
            self.timedatalist = time
            self.__exist = False

    
    def get_time(self) -> datetime:
        return self.__starttime

    
    def set_duration(self, du:int) -> None:
        self.__duration = du

    
    def get_duration(self) -> int:
        return self.__duration


    def isExist(self) -> bool:
        return self.__exist