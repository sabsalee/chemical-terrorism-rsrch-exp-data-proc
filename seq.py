from module.processing import *
# from module.firebaseupload import *
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

    중요!!!: 폴더 안에 가공된 파일이 이미 있다면 덮어씌워집니다.
    주의!! : 동시에 실험했던 같은 종류의 데이터만 한 폴더에 넣으세요.
    참고   : 엑셀 내 차트 디자인은 프로그램에서 적용할 수 없습니다. 엑셀에서 직접 서식파일을 선택하세요.''')

    # try:
    #     rtdb = Database()
    #     print('\n\n[정보] 인터넷 연결 성공')
    # except:
    #     rtdb = None
    #     print('\n\n[경고] 인터넷 연결 실패 - 가공된 데이터가 인터넷에 업로드되지 않습니다.')



    # DIALOG

    while True:
        time = TimeInfo()
        errCount = 0

        # 실험 차수 수집
        # try:
        #     with open('order.pkl', 'rb') as f:
        #         order = pickle.load(f)
        # except:
        #     order = None

        # if order != None:
        #     try:
        #         print('[ 실험 차수 확인 ]\n')
        #         print(f'이전에 데이터를 가공했던 실험의 차수가 [{order}차] 입니다.')
        #         print(f'이번에 가공할 데이터의 차수는 [{order+1}차]가 맞습니까?')
        #         print('\n맞으면 엔터를, 아니라면 실험 차수를 숫자로 입력해주세요.')
        #         order = int(input('> '))
        #         try:
        #             with open('order.pkl', 'wb') as f:
        #                 pickle.dump(order, f)
        #         except:
        #             pass
        #     except:
        #         order += 1
        #         try:
        #             with open('order.pkl', 'wb') as f:
        #                 pickle.dump(order, f)
        #         except:
        #             pass
        # else:
        #     while True:
        #         try:
        #             print('실험이 "몇차" 실험인지 "차수"를 "숫자"로 입력해주세요.')
        #             order = int(input('> '))
        #             try:
        #                 with open('order.pkl', 'wb') as f:
        #                     pickle.dump(order, f)
        #             except:
        #                 pass
        #             break
        #         except:
        #             print('[오류] 숫자를 입력하셔야 합니다.')
        
        print('\n\n\n[ 데이터 종류 설정 ]\n')
        print('가공할 데이터 종류를 선택해주세요.')
        print('1. 이산화탄소 계측기')
        print('2. 헬륨 계측기')
        print('3. 산소 계측기')
        print('4. 더미')
        while True:
            n = input('번호를 입력해주세요 > ')
            sel = {'1':'CO2', '2':'He', '3':'O2', '4':'dummy'}
            if n in sel:
                file_type = sel[n]
                break
            print('잘못 입력하셨습니다.')

        
        print("\n\n\n[ 폴더 경로 설정 ]\n")
        print("처리해야되는 폴더의 경로를 입력해주세요.")
        while True:
            try:
                lists = filter_csv_from_dir(input("> "), file_type=file_type)
                break
            except FileNotFoundError:
                print('[오류] 해당 경로가 존재하지 않습니다. 다시 시도하세요.\n')
            except NotADirectoryError:
                print('[오류] 해당 경로는 폴더가 아닙니다. 다시 시도하세요.\n')

        print("\n\n\n[ 실험 시작시간 설정 ]\n")
        while True:
            try:
                print("실험 시작시간을 직접 입력해주세요 ( 예시[14시 8분] -> 14:8 )")
                _time = input("> ")
                _time = list(map(int, _time.split(':')))
                time.set_time(_time, only_time=True)
                break
            except:
                print('\n잘못 입력하신 것 같습니다.')

        # print("\n\n\n[ 안정화 시간 설정 ]\n")
        # while True:
        #     try:
        #         print('실험 시작 후 반응시간을 계산하기 위해서 안정화 시간을 입력해야합니다.')
        #         print('실험 안정화 시간을 초로 입력해주세요.')
        #         __stabililze_period = int(input('>' ))
        #         time.set_stab_period(__stabililze_period)
        #         break
        #     except:
        #         print('\n잘못 입력하셨습니다.')
        # 임시로 0초 설정
        time.set_stab_period(0)
        
        try:
            print('\n\n\n[ 실험 지속시간(진행시간) 설정 ]\n')
            print(f'기본 설정되어있는 실험 지속시간(진행시간)은 [{time.get_duration()}초] 입니다.')
            print('\n맞으면 엔터를, 아니라면 지속시간(단위:초)을 숫자로 입력해주세요.')
            du = int(input('> '))
            time.set_duration(du)
        except:
            pass



        # PROCESS

        clear()
        print(f'이제 실험 데이터 가공을 시작하겠습니다. 아래의 내용을 확인해주세요.\n')
        # print(f'> {order}차 실험 데이터')
        print(f'> 데이터 종류: {file_type}')
        print(f'> 실험 시작시간: {time.timedatalist[0]}시 {time.timedatalist[1]}분')
        print(f'> 실험 지속시간: {time.get_duration()}초')
        input("\n>> 진행하려면 엔터키를 눌러주세요.\n\n")

        for i, e in enumerate(lists):
            try:
                bar = tqdm(total = 100, desc=f'[{i+1}/{len(lists)}] {e.name}', bar_format='{desc:30.28}{percentage:3.0f}%|{bar:30}|', smoothing=1)

                df = preprocess_to_df(e, time)
                bar.update(20)

                # 그래프 분석을 위한 데이터 생성
                if e.type == 'CO2': #TODO: 계측기 데이터 늘어나면 맞게 수정한다.
                    calculated_dict = calculate_df(df, time.get_stab_period())
                else:
                    calculated_dict = None

                wb = dataframe_to_excel(df)
                bar.update(10)

                wb = formular_process(wb, e.type)
                bar.update(10)

                # if e.type != "dummy" and rtdb != None: # dummy 데이터 외(dummy는 total로만) 백업하기
                #     dataDict = firebase_process(wb, e, e.type)
                #     bar.update(10)

                #     rtdb.upload_dummy(time, order, dataDict, e.type)
                #     bar.update(20)
                # else:
                #     bar.update(30)
                bar.update(30)

                wb = expression_process(wb, e)
                bar.update(10)

                if calculated_dict != None:
                    wb = insert_calculated_values(wb, calculated_dict)

                chart_process(wb, e)
                bar.update(10)

                wb.save(f"{e.dirPath}/{e.name[:-4]} 정리.xlsx")
                bar.update(10)
                bar.close()

            except Exception as err:
                bar.close()
                errCount += 1
                print('\r[오류] ', end='')
                print(err)
                print('[오류] ' + e.name + "의 파일 변환 중 오류가 발생하여 가공하지 않고 넘어갑니다.\n")



        # FINISH

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
    main()