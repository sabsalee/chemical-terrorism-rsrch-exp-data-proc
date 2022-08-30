from module.processing import *
from module.firebaseupload import *
import pickle, os




def clear():
    if os.name in ('nt', 'dos'):
        os.system('cls')
    else:
        os.system('clear')




clear()
print('''
화학테러 논문조 더미 및 계측기 데이터 가공 프로그램

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
            lists = filter_csv_from_dir(input("> "))
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
            print(f'이전에 데이터를 가공했던 실험의 차수가 [{order}차] 입니다.')
            print(f'이번에 가공할 데이터의 차수는 [{order+1}차]가 맞습니까?')
            print('\n맞으면 Y를 아니라면 실험 차수를 숫자로 입력해주세요.')
            order = int(input('>'))
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
                order = int(input('>'))
                try:
                    with open('order.pkl', 'wb') as f:
                        pickle.dump(order, f)
                except:
                    pass
                break
            except:
                print('[오류] 숫자를 입력하셔야 합니다.')

    clear()
    print(f'[{order}차] 실험 데이터 가공을 시작하겠습니다.')
    input(">> 진행하려면 엔터키를 눌러주세요.\n\n")


    for i, e in enumerate(lists):
        try:
            print(f'[{i+1}/{len(lists)}] {e.name}')

            df = preprocess_csv_to_df(e, time)
            print('|||', end='')

            wb = dataframe_to_excel(df)
            print('|||', end='')

            wb = formular_process(wb, e)
            print('|||', end='')

            if e.type == "total" and rtdb != None:
                dataDict = firebase_process(wb, e.type)
                print('|||', end='')

                rtdb.upload_dummy(time, order, dataDict)
                print('|||', end='')
            # elif e.type == "CO2" and rtdb != None: # FIXME: 여기 계측기 늘어나면 수정
            #     dataDict = firebase_process(wb, e.type)
            #     print('|||', end='')

            #     rtdb.upload_dummy(time, order, dataDict)
            #     print('|||', end='')
            else:
                print('||||||', end='')

            wb = expression_process(wb, e)
            print('|||', end='')

            chart_process(wb, e)
            print('|||', end='')

            wb.save(f"{e.dirPath}/{e.name[:-4]}.xlsx")
            print('||| OK!', end='\n\n')

        except Exception as err:
            errCount += 1
            print('\r[오류] ', end='')
            print(err)
            print('[오류] ' + e.name + "의 파일 변환 중 오류가 발생하여 가공하지 않고 넘어갑니다.\n")

    print(f'''
    \n총 {len(lists)}개의 파일 중에서 {len(lists)-errCount}개의 파일 가공에 성공하였습니다.
    새로운 폴더에서 가공을 시작하시려면 'Y'를, 종료하시려면 'N'을 입력해주세요.
    ''')
    while True:
        isLoop = input('> ')
        if isLoop == "Y" or isLoop == "y" or isLoop == "N" or isLoop == "n":
            break
        print("\n새로운 폴더에서 가공을 시작하시려면 'Y'를, 종료하시려면 'N'을 입력해주세요.")

    if isLoop == 'N' or isLoop == "n":
        break