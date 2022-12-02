import updater
import seq
from self_destruction import self_destruction

def main():
    self_destruction()
    res = updater.main()
    if res:
        print('프로그램이 업데이트 되었습니다. 원활한 실행을 위해 종료 후 다시 실행해 주세요.')
        input('엔터를 누르면 프로그램이 종료됩니다.')
        quit()
    seq.main()

if __name__ == '__main__':
    main()