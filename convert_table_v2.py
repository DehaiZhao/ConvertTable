import pandas as pd
import os

EXTS = ['.xls', '.xlsx']

def create_df(row_set, col_set, dict_data):
	df_data = {}
	for col in col_set:
		df_data[col] = []
		for row in row_set:
			df_data[col].append(dict_data[row + str(int(col))])
	df_data = pd.DataFrame(df_data)
	return df_data

def convert(filename, output_path):
	sheets = pd.read_excel(filename, None)
	with pd.ExcelWriter(os.path.join(output_path, os.path.splitext(os.path.basename(filename))[0] + '_output.xlsx')) as writer:
		for key in sheets:
			rows = sheets[key]['Row'].values
			cols = sheets[key]['Column'].values
			values = sheets[key]['Nuclei Valid - Number of Objects'].values
			row_set = sorted(set(rows))
			col_set = sorted(set(cols))
			data = []
			dict_data = {}
			for i in range(len(rows)):
				data.append({'row': rows[i], 'col': cols[i], 'val': values[i]})
			for item in data:
				dict_data[item['row'] + str(int(item['col']))] = item['val']
			df_data = create_df(row_set, col_set, dict_data)
			df_data.to_excel(writer, startrow=int(col_set[0])-1, startcol=ord(row_set[0])-65, index=False, header=False, sheet_name=key)

def main():
	input_path = input('Please enter file path:').strip()
	output_path = input('Please enter output path:').strip()
	if os.path.isdir(input_path):
		file_list = os.listdir(input_path)
		excel_list = []
		for f in file_list:
			if os.path.splitext(f)[1] in EXTS:
				excel_list.append(os.path.join(input_path, f))
	elif os.path.isfile(input_path):
		excel_list = [input_path]
	else:
		print('Invalid path')
		exit(1)
	for excel in excel_list:
		convert(excel, output_path)

if __name__ == '__main__':
	main()