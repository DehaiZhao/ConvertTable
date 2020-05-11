import pandas as pd
import os

file_name = '190802_siRNA-Drug_AJG_new_analysis.xlsx'

def create_df(row_set, col_set, dict_data):
	df_data = {}
	for col in col_set:
		df_data[col] = []
		for row in row_set:
			df_data[col].append(dict_data[row + str(int(col))])
	df_data = pd.DataFrame(df_data)
	return df_data

def main():
	sheets = pd.read_excel(file_name, None)
	with pd.ExcelWriter('test.xlsx') as writer:
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

if __name__ == '__main__':
	main()