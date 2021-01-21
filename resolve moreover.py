import pandas as pd
import requests
import os
import time


def main():
	path = os.path.dirname(os.path.realpath(__file__))
	filename = "input.xlsx"
	df = pd.read_excel(str(path+"/"+filename))
	urls = df["URL"].tolist()
	new_urls = []



	for counter, url in enumerate(urls):
		try:
			print(str(str(counter + 1)  + "/" + str(len(urls))))
			new_url = requests.get(url, headers = {'User-Agent': 'Mozilla/5.0'}, timeout = 7)
			new_url = new_url.url 
			
		except:
			new_url = "Failed redirect"
		new_urls.append(new_url)
		time.sleep(2)
	df["resolved"] = new_urls
	df.to_excel(str(path+"/"+"resolved.xlsx"), index = False)
	print("Finished")

if __name__ == '__main__': 
    main()
