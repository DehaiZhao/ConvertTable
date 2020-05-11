# ConvertTable

### Note:

If you run the code through command line, please run `convert_table.py`.

If you compile the code, please use `convert_table_v2.spec`.

See the details below.

## Reqirment
`pip3 install pandas openpyxl`

## Run
`python3 convert_table.py /path/to/excel /output/path`
![](https://github.com/DehaiZhao/ConvertTable/blob/master/Images/WechatIMG108.png)

## Complie
`pip3 install pyinstaller`

`pyinstaller -F convert_table_v2.py`

### Double click this file to run
![](https://github.com/DehaiZhao/ConvertTable/blob/master/Images/WechatIMG109.png)

![](https://github.com/DehaiZhao/ConvertTable/blob/master/Images/WX20200511-215727%402x.png)

## Issues
No model named 'pkg_resources.py2_warn'

1. open `convert_table_v2.spec`
2. modify `hiddenimports=['pkg_resources.py2_warn']`
3. compile again `pyinstaller -F convert_table_v2.spec`. Use `.spec` instead of `.py`
