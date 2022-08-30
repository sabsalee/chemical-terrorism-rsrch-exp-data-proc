from datetime import datetime, timedelta

class TimeInfo:
    def __init__(self) -> None:
        self.__exist = False
        self.__starttime = None
        self.__endtime = None

    def set_time(self, type, time: datetime):
        if time.second >= 30:
            time = time.replace(second=0)
            time = time + timedelta(minutes=1)
        else:
            time = time.replace(second=0)

        if type == "start":
            self.__starttime = time
            self.__endtime = time + timedelta(minutes=10)
            self.__exist = True
        elif type == "end":
            if self.__starttime == None:
                raise NeverAssignStartTime
            self.__endtime = time
        else:
            raise ExpectedTimeTypeError
    
    def get_time(self, type) -> datetime:
        if type == "start":
            return self.__starttime
        elif type == "end":
            return self.__endtime
        else:
            raise ExpectedTimeTypeError

    def isExist(self) -> bool:
        return self.__exist

class ExpectedTimeTypeError(Exception):
    def __str__(self) -> str:
        return "시간 타입이 누락되었습니다. start 또는 end로 시간 타입을 입력하여야 합니다."

class NeverAssignStartTime(Exception):
    def __str__(self) -> str:
        return "시작 시간이 선언된 적이 없습니다. 시작 시간 선언 없이 끝 시간을 먼저 선언할 수 없습니다."