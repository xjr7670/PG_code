#-*- coding:utf-8 -*-

import os

if os.path.exists('config.inc'):
    work_dir = raw_input("Use last folder(just press Enter) or a new one(enter that)?")
    if work_dir == '':
        with open('config.inc') as f:
            config_str = f.read()
            config = eval(config_str)
            work_dir = config['work_dir']
else:
    work_dir = raw_input("Please enter the working directory: ")
    d = {}
    while not os.path.exists(work_dir):
        work_dir = raw_input("Please enter the completed folder path, start with the C/D/E etc., "
                             "Case insensitive: ")
    else:
        d["work_dir"] = work_dir
    with open('config.inc', 'w') as f:
        f.write(str(d))

print work_dir