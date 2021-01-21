import pandas as pd
import os

def remove_weirdness(cell):
	replace_dict = {" - " : [";"],
					"'" : ['"', '“', '”', '„', '"', '«', '»', '“', '”']}
	for key, value in replace_dict.items():
		for i in value:
			cell = cell.replace(i, key)
	return cell

path = os.path.dirname(os.path.realpath(__file__))
filename = "input.xlsx"
df = pd.read_excel(str(path+"/"+filename))

df["New_contentText"]  = [remove_weirdness(s) for s in df["contentText"].tolist()]
df["New_contentTitle"]  = [remove_weirdness(s) for s in df["contentTitle"].tolist()]

df.to_excel(str(path+"/"+"resolved.xlsx"), index = False)
print("Finished")
