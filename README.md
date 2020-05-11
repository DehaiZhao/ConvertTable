# ConvertTable

## Reqirment
`pip3 install pandas openpyxl`

## Run
`python3 convert_table.py /path/to/excel /output/path`

## Complie
`pip3 install pyinstaller`
`pyinstaller -F convert_table_v2.py`

## Issues
No model named 'pkg_resources.py2_warn'

1. open `convert_table_v2.spec`
2. modify `hiddenimports=['pkg_resources.py2_warn']`
3. `pyinstaller -F convert_table_v2.spec`
