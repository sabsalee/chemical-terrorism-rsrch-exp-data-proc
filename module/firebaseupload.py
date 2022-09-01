import firebase_admin
from firebase_admin import db, credentials
from module.timeinfo import TimeInfo
from datetime import datetime

class Database:
    def __init__(self) -> None:
        cred = credentials.Certificate('key/chemical-terrorism-research-firebase-rtdb-key.json')
        firebase_admin.initialize_app(cred, {
            'databaseURL': 'https://chemical-terrorism-research-default-rtdb.firebaseio.com/'
        })

    def upload_dummy(self, time:TimeInfo, order:int, dataDict:dict, __type:str):
        orderRef = db.reference(f'order/{order}')
        dataRef = db.reference(f'order/{order}/data')

        if __type == 'total':
            __isExistVailidateRef = db.reference(f'order/{order}/data/DUMMY_ID1/0')
            __isExistVailidateValue = __isExistVailidateRef.get()
        elif __type == 'CO2':
            __isExistVailidateRef = db.reference(f'order/{order}/data/CO2_1/0') # TODO: 키값만 받아올 수 있는 API가 있다면?
            __isExistVailidateValue = __isExistVailidateRef.get()
        elif __type == 'He':
            __isExistVailidateRef = db.reference(f'order/{order}/data/He_1/0')
            __isExistVailidateValue = __isExistVailidateRef.get()
        
        # 중복된 값이 있다면 
        if __isExistVailidateValue != None:
            print(f'\r이미 [{order}차] 실험 결과가 인터넷 데이터베이스에 있습니다. 그래도 올리시겠습니까?')
            while True:
                cmd = input('(Y/N) >')
                if cmd == "Y" or cmd == "y" or cmd == "N" or cmd == "n":
                    break
                print('잘못 입력하셨습니다. Y 또는 N 을 입력해주세요.')
            if cmd == "Y" or cmd == "y":
                self.__backup_data(order)
                self.__upload(time, orderRef, dataRef, dataDict, update=True)
            else:
                print('\r\r데이터 업로드를 건너뜁니다.\n||||||||||||', end='')
        else:
            self.__upload(time, orderRef, dataRef, dataDict)    

    # 업로드 내부함수
    def __upload(self, time: TimeInfo, orderRef: db.Reference, dataRef: db.Reference, dataDict: db.Reference, update=False):
        orderRef.update({'date':time.get_time('start').strftime('%Y-%m-%d')})
        if update:
            orderRef.update({'updatedAt': datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
            print('||||||||||||', end='')
        else:
            orderRef.update({'createdAt': datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
        dataRef.update(dataDict)

    # 업로드 할 때 이미 업로드 된 차수가 있다면 그 차수를 백업시키는 함수
    def __backup_data(self, order: int):
        date = db.reference(f'order/{order}/date').get()
        uploadedAt = db.reference(f'order/{order}/updatedAt').get()
        if uploadedAt == None:
            uploadedAt = db.reference(f'order/{order}/createdAt').get()
        backupAt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data = db.reference(f'order/{order}/data').get()

        backupRef = db.reference(f'order/{order}/backup')
        backedUpData = backupRef.get()
        if backedUpData == None:
            backupRef.update({
                0:{
                    'date':date,
                    'uploadedAt': uploadedAt,
                    'backupAt': backupAt,
                    'data':data
                }
            })
        else:
            backedUpCount = len(backedUpData) # 이것은 개수이고 우리는 인덱스를 이용해야해서 +1 하지 않고 바로 사용한다.
            backupRef.update({
                backedUpCount:{
                    'date':date,
                    'uploadedAt': uploadedAt,
                    'backupAt': backupAt,
                    'data':data
                }
            })