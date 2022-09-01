import requests, os
import configparser
import zipfile
import shutil
import time
from datetime import datetime
from tqdm import tqdm



def clear():
    if os.name in ('nt', 'dos'):
        os.system('cls')
    else:
        os.system('clear')

def extract(path, filename) -> dict:
    with zipfile.ZipFile(f'{path}/{filename}', 'r') as zf:
        for member in tqdm(zf.infolist(), desc='[업데이트] 받은 파일 압축 해제하는 중... '):
         try:
             zf.extract(member, path)
         except zipfile.error as e:
             pass
        file_list = zf.namelist()
        orig_folder_name = file_list[0]
        folders_path = [f for f in file_list if f[-1] in '/' ]
        orig_file_path_list = [f for f in file_list if f not in folders_path]
        dest_file_path_list = [f.replace(orig_folder_name, '') for f in orig_file_path_list]
        return {
            'orig_folder_name': orig_folder_name,
            'folders_path': folders_path,
            'orig_file_path_list': orig_file_path_list,
            'dest_file_path_list': dest_file_path_list,
        }

# https://gist.github.com/yanqd0/c13ed29e29432e3cf3e7c38467f42f51
def download(url: str, fname: str, dir='', desc=None, chunk_size=1024):
    resp = requests.get(url, stream=True)
    total = int(resp.headers.get('content-length', 0))
    if desc == None:
        desc = fname
    if dir != '':
        dir += '/'
    with open(dir + fname, 'wb') as file, tqdm(
        desc=desc,
        total=total,
        unit='iB',
        unit_scale=True,
        unit_divisor=1024,
    ) as bar:
        for data in resp.iter_content(chunk_size=chunk_size):
            size = file.write(data)
            bar.update(size)


isErrorRaised = False
clear()
try:
    # Get program's version
    conf = configparser.ConfigParser()
    print('[업데이트] 프로그램의 버전을 확인하는 중...')
    try:
        conf.read('version.ini', encoding='utf-8')
        version_from_local = conf['DEFAULT']['version']
    except:
        print('[업데이트] 경고! 프로그램 버전 정보가 없습니다! 최신 버전으로 업데이트 합니다.')
        version_from_local = 'v0.0.0'
        conf['DEFAULT'] = {}
        conf['DEFAULT']['version'] = version_from_local

    # Get latest version from github releases
    print('[업데이트] 인터넷에서 최신 버전의 업데이트 확인 중...')
    response = requests.get("https://api.github.com/repos/sabsalee/chemical-terrorism-rsrch-exp-data-proc/releases/latest")
    version_from_releases = response.json()['tag_name']
    zipLink = response.json()['zipball_url']
    
    # Start update process
    try:
        if version_from_local != version_from_releases:
            print(f'[업데이트] 새로운 업데이트가 발견되어 업데이트를 진행합니다. ({version_from_local} -> {version_from_releases})\n')
            print(f'''============================================
            response.json()['body']
            ============================================''')
            print('\n10초 뒤 업데이트가 시작됩니다.')
            time.sleep(10)
            os.makedirs('temp', exist_ok=True)
            download(zipLink, f'{version_from_releases}.zip', 'temp', f'\n\n[업데이트] 새로운 버전({version_from_releases})을 다운로드하는 중... ')
            result = extract('temp', f'{version_from_releases}.zip')
            file_list = result['orig_file_path_list']
            for i, f in enumerate(tqdm(file_list, desc='[업데이트] 새로운 버전을 설치하는 중... ')):
                shutil.move(f'temp/{f}', f'{result["dest_file_path_list"][i]}')
            print('[업데이트] 업데이트 마무리하는 중...')
            shutil.rmtree('temp', ignore_errors=True)
            conf['DEFAULT']['version'] = version_from_releases
            conf['DEFAULT']['lastUpdateCheck'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conf['DEFAULT']['recentUpdateDate'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open('version.ini', 'w', encoding='utf-8') as cf:
                conf.write(cf)
            print('\n[업데이트] 업데이트 완료!')
            print('10초 뒤 이동합니다.')
            time.sleep(10)
        else:
            conf['DEFAULT']['lastUpdateCheck'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print('\n[업데이트] 최신 버전의 프로그램을 사용 중입니다.')
            print('5초 뒤 이동합니다.')
            time.sleep(5)
    except:
        print('[업데이트] 업데이트 진행 중 오류 발생')
        isErrorRaised = True
except:
    print('[업데이트] 업데이트 확인 중 오류 발생')
    isErrorRaised = True
finally:
    if isErrorRaised:
        print('5초 뒤 이동합니다.')
        time.sleep(5)