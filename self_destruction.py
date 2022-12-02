import shutil, os

def self_destruction():
    shutil.rmtree('module', ignore_errors=True)
    shutil.rmtree('key', ignore_errors=True)
    os.remove('updater.py')
    os.remove('seq.py')

