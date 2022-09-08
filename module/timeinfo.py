from datetime import datetime, timedelta

class TimeInfo:
    def __init__(self, duration=600) -> None:
        self.__exist = False
        self.__starttime = None
        self.__endtime = None
        self.__duration = duration
        self.__stab_period = 0
        self.timedatalist = None

    def set_time(self, time: datetime|list, only_time=False):
        if not only_time:
            if time.second >= 30:
                time = time.replace(second=0)
                time = time + timedelta(minutes=1)
            else:
                time = time.replace(second=0)

            self.__starttime = time
            self.__endtime = time + timedelta(seconds=self.get_duration())
            self.__exist = True
        else:
            self.timedatalist = time
            self.__exist = False

    def get_time(self) -> datetime:
        return self.__starttime

    def get_end_time(self) -> datetime:
        return self.__endtime
        
    

    def set_stab_period(self, sec:int) -> None:
        self.__stab_period = sec

    def get_stab_period(self) -> int:
        return self.__stab_period
    


    def set_duration(self, du:int) -> None:
        self.__duration = du
    
    def get_duration(self) -> int:
        return self.__duration


    def isExist(self) -> bool:
        return self.__exist