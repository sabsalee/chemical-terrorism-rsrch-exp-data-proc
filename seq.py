from module.processing import *
from module.firebaseupload import *
import pickle, os
import configparser
from tqdm import tqdm



def clear():
    if os.name in ('nt', 'dos'):
        os.system('cls')
    else:
        os.system('clear')

def main():
    clear()
    conf = configparser.ConfigParser()
    conf.read('version.ini', encoding='utf-8')
    version = conf['DEFAULT']['version']
    print(f'''
    화학테러 논문조 더미 및 계측기 데이터 가공 프로그램 ({version})

    - 이 프로그램을 통해 CSV파일을 엑셀파일로 가공하게 됩니다.
    - 실험 결과 정리를 위해서 이 프로그램으로 가공된 데이터는 서버에 업로드 됩니다.
    - 컴퓨터가 인터넷에 연결되지 않았다면 데이터는 서버에 업로드되지 못하므로 인터넷 연결을 확인해주세요.


    중요!!!: 폴더 안에 가공된 파일이 이미 있다면 덮어씌워집니다.
    주의!! : 동시에 실험했던 데이터만 같은 폴더에 있어야 합니다. 그렇지 않으면 시간 정보가 꼬입니다.
    참고   : 엑셀 내 차트 디자인은 프로그램에서 적용할 수 없습니다. 엑셀에서 직접 서식파일을 선택하세요.''')

    try:
        rtdb = Database()
        print('\n\n[정보] 인터넷 연결 성공')
    except:
        rtdb = None
        print('\n\n[경고] 인터넷 연결 실패 - 가공된 데이터가 인터넷에 업로드되지 않습니다.')

    while True:
        print("\n\n\n처리해야되는 폴더의 경로를 입력해주세요.")
        while True:
            try:
                isTimeinfoFileExist, lists = filter_csv_from_dir(input("> "))
                break
            except FileNotFoundError:
                print('[오류] 해당 경로가 존재하지 않습니다. 다시 시도하세요.\n')
            except NotADirectoryError:
                print('[오류] 해당 경로는 폴더가 아닙니다. 다시 시도하세요.\n')
                
        clear()
        time = TimeInfo()
        errCount = 0

        # 실험 차수 수집
        try:
            with open('order.pkl', 'rb') as f:
                order = pickle.load(f)
        except:
            order = None

        if order != None:
            try:
                print('[ 실험 차수 확인 ]\n')
                print(f'이전에 데이터를 가공했던 실험의 차수가 [{order}차] 입니다.')
                print(f'이번에 가공할 데이터의 차수는 [{order+1}차]가 맞습니까?')
                print('\n맞으면 엔터를, 아니라면 실험 차수를 숫자로 입력해주세요.')
                order = int(input('> '))
                try:
                    with open('order.pkl', 'wb') as f:
                        pickle.dump(order, f)
                except:
                    pass
            except:
                order += 1
                try:
                    with open('order.pkl', 'wb') as f:
                        pickle.dump(order, f)
                except:
                    pass
        else:
            while True:
                try:
                    print('실험이 "몇차" 실험인지 "차수"를 "숫자"로 입력해주세요.')
                    order = int(input('> '))
                    try:
                        with open('order.pkl', 'wb') as f:
                            pickle.dump(order, f)
                    except:
                        pass
                    break
                except:
                    print('[오류] 숫자를 입력하셔야 합니다.')

        if isTimeinfoFileExist == False:
            print("\n\n\n[ 실험 시작시간 설정 ]\n")
            print("현재 선택한 폴더 내에 실험 시작시간 정보를 가진 파일이 없는 것 같습니다.")
            while True:
                try:
                    print("\n실험 시작시간을 직접 입력해주세요 ( 예시[14시 8분] -> 14:8 )")
                    _time = input("> ")
                    _time = list(map(int, _time.split(':')))
                    time.set_time(_time, only_time=True)
                    break
                except:
                    print('\n잘못 입력하신 것 같습니다.')
            

        try:
            print('\n\n\n[ 실험 지속시간(진행시간) 설정 ]\n')
            print(f'기본 설정되어있는 실험 지속시간(진행시간)은 [{time.get_duration()}초] 입니다.')
            print('\n맞으면 엔터를, 아니라면 지속시간(단위:초)을 숫자로 입력해주세요.')
            du = int(input('> '))
            time.set_duration(du)
        except:
            pass

        clear()
        print(f'이제 실험 데이터 가공을 시작하겠습니다. 아래의 내용을 확인해주세요.\n')
        print(f'> {order}차 실험 데이터')
        print(f'> {"실험 시작시간: 파일에서 추출 예정" if isTimeinfoFileExist else f"실험 시작시간: {time.timedatalist[0]}시 {time.timedatalist[1]}분"}')
        print(f'> 실험 지속시간: {time.get_duration()}초')
        input("\n>> 진행하려면 엔터키를 눌러주세요.\n\n")


        for i, e in enumerate(lists):
            try:
                bar = tqdm(total = 100, desc=f'[{i+1}/{len(lists)}] {e.name}', bar_format='{desc:30.28}{percentage:3.0f}%|{bar:30}|', smoothing=1)
                # print(f'[{i+1}/{len(lists)}] {e.name}')

                df = preprocess_csv_to_df(e, time)
                # print('|||', end='')
                bar.update(20)

                wb = dataframe_to_excel(df)
                # print('|||', end='')
                bar.update(10)

                wb = formular_process(wb)
                # print('|||', end='')
                bar.update(10)

                if e.type != "dummy" and rtdb != None: # dummy 데이터 외(dummy는 total로만) 백업하기
                    dataDict = firebase_process(wb, e, e.type)
                    # print('|||', end='')
                    bar.update(10)

                    rtdb.upload_dummy(time, order, dataDict, e.type)
                    # print('|||', end='')
                    bar.update(20)
                else:
                    # print('||||||', end='')
                    bar.update(30)

                wb = expression_process(wb, e)
                # print('|||', end='')
                bar.update(10)

                chart_process(wb, e)
                # print('|||', end='')
                bar.update(10)

                wb.save(f"{e.dirPath}/{e.name[:-4]}.xlsx")
                # print('||| OK!', end='\n\n')
                bar.update(10)
                bar.close()

            except Exception as err:
                bar.close()
                errCount += 1
                print('\r[오류] ', end='')
                print(err)
                print('[오류] ' + e.name + "의 파일 변환 중 오류가 발생하여 가공하지 않고 넘어갑니다.\n")

        print(f'\n총 {len(lists)}개의 파일 중에서 {len(lists)-errCount}개의 파일 가공에 성공하였습니다.')
        print("새로운 폴더에서 가공을 시작하시려면 'Y'를, 종료하시려면 'N'을 입력해주세요.")
        while True:
            isLoop = input('> ')
            if isLoop == "Y" or isLoop == "y" or isLoop == "N" or isLoop == "n":
                break
            print("\n새로운 폴더에서 가공을 시작하시려면 'Y'를, 종료하시려면 'N'을 입력해주세요.")

        if isLoop == 'N' or isLoop == "n":
            break

if __name__ == '__main__':
    pass