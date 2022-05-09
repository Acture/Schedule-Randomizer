import pandas as pd
import random
import string
import datetime


general_list = [['公司重大业务沟通', 1],
				['董事会会议', 0.5],
				['海外跨所合作会议', 0.5],
				['XX公司跨平台合作会议', 5 ],
				['XX项目上所沟通会议', 2],
				['海外业务拓展会议', 1],
				['XX（海外国家）业务拓展沟通会', 2],
				['XX公司领投会议', 7],
				['XX公司投资会议', 5],
				['XX主题项目投资评审', 2]]
theme_list = ['NFT',
			  'GameFi',
			  'web3',
			  "Defi",
			  "元宇宙",
			  "DAO",
			  "DCEP",
			  "以太坊2.0",
			  "Layer2",
			  "波卡"
			  ]

country_list = ['瑞士','挪威','美国','新加坡','澳大利亚','丹麦','瑞典','爱尔兰','荷兰','英国','加拿大',
                '奥地利','芬兰','德国','比利时','法国','新西兰','日本','意大利','韩国','西班牙']

morning_points = [ "10:30","11:00", "11:30", "12:00"]
afternoon_points = [ "12:30", "13:00", "13:30", "14:00","14:30", "15:00", "15:30", "16:00", "16:30","17:00", "17:30", "18:00", "18:30", "19:00","19:30"]
evening_points = ["20:00", "20:30", "21:00", "21:30", "22:00", "22:30", "23:00", "23:30","24:00"]
time_points = [morning_points,afternoon_points,evening_points]

fixed_event = ['高管早会', '基金投资IC会议','午餐','晚餐']
fixed_time = ["09:30-10:00", "10:00-10:30", "12:00-12:30","19:30-20:00"]

def event():
	result = []
	for thing in general_list:
		if thing[1] == 0.5:
			if random.randrange(0, 100, 1) > 50:
				result.append(thing[0])
			thing[1] = 0
		if thing[1] >= 1:
			if 'XX项目' in thing[0]:
				for _ in range(thing[1]):
					result.append(thing[0].replace(
						'XX项目', "".join(random.choice(theme_list))+'项目'))

			if 'XX公司' in thing[0]:
				for _ in range(thing[1]):
					result.append(thing[0].replace('XX公司', "".join(
						random.sample(string.ascii_uppercase, 2))+'公司'))

			if 'XX（海外国家）' in thing[0]:
				for _ in range(thing[1]):
					result.append(thing[0].replace(
						'XX（海外国家）', "".join(random.choice(country_list))))
			if 'XX' not in thing[0]:
				for _ in range(thing[1]):
					result.append(thing[0])
	random.shuffle(result)
	while len(result) > 23:
		del result[random.randrange(0,len(time_points)-1,1)]

	return result

def time():
	result = []

	for m in range(0,3):
		for n in range(len(time_points[m])-1):
			result.append(time_points[m][n]+"-"+time_points[m][n+1])

	print(result)

	j = random.randrange(1,len(time_points)-1,1)
	k = random.randrange(1,len(time_points[j])-1,1)
	del time_points[j][k]


	a = random.randrange(1,len(time_points)-1,1)
	b = random.randrange(1,len(time_points[a])-1,1)
	del time_points[a][b]

	result = []

	for m in range(0,3):
		for n in range(len(time_points[m])-1):
			result.append(time_points[m][n]+"-"+time_points[m][n+1])
	

	return result

def add_fixed_event(df:pd.DataFrame)->pd.DataFrame:
	lines = []
	to_day =  datetime.datetime.today().strftime('%A')
	if to_day == "Saturday":
		fixed_event[0] = '高管周会'
	for x,y in zip(fixed_event,fixed_time):
		lines.append({'今日事项':x,'时间段':y})
	lines = pd.DataFrame(lines)
	df = df.append(lines).sort_values(by=['时间段']).reset_index(drop=True)
	return df

def main():
	event_today = event()
	time_period = time()
	print(event_today)
	print(time_period)
	df = pd.DataFrame({
		'今日事项': event_today,
		'时间段': time_period
	})
	df = add_fixed_event(df)
	df.to_excel('output.xlsx', sheet_name="Sheet1")


if __name__ == '__main__':
	main()
	
