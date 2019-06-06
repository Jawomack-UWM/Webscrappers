import requests, os, bs4, re, openpyxl

wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
sheet['A1'] = 'Job'
sheet['b1'] = 'Ranking'
sheet['c1'] = 'Early Pay'
sheet['d1'] = 'Middle Pay'
sheet['e1'] = 'Pay Growth'
sheet['f1'] = 'High Meaning'
sheet['g1'] = 'High Satisfaction'
sheet['h1'] = 'Low Stress'
sheet['i1'] = 'Employment Projection'
r = 2

#2-46 for best #47 - 91 for worst (+1 because computers start at 0)
for i in range(2,91):
	print i
	
	url = 'http://www3.forbes.com/leadership/the-best-and-worst-masters-degrees-for-jobs-in-2017/%d/' %i
	res = requests.get(url)
	page_grab = bs4.BeautifulSoup(res.text, 'html.parser')
	page_2_txt = str(page_grab)

	testfile = open('test.txt', 'w')
	testfile.write(page_2_txt)
	testfile.close()
	#I don't think I have to close and reopen.  Check later

	example = open('test.txt')
	ex = bs4.BeautifulSoup(example.read(), 'html.parser')

	scb = 0
	holder = []
	final = []

	for j in ex.findAll('b'):
		scb +=1

		if scb == 2:
			break
	
		for k in j.parent.next_siblings:
			holder.append(k)

	for item in holder:
	
		k = re.sub(r'\D', '', str(item))
	
		final.append(k)

	for j in final:

		if j == '':
			final.remove(j)

	early = ex.find('b').next_sibling
	elems = ex.select('p strong')
	k = re.sub(r'\D', '', str(elems))

	name = str(elems[0])
	job = re.sub('<[^<]+?>', '', name)


	job_title = job[3:]
	ranking = k
	early_pay = early
	middle_pay = final[0]
	pay_growth = float(final[1])/100
	High_meaning = float(final[2])/100
	High_Satisfaction = float(final[3])/100
	low_stress = float(final[4])/100
	employment_projection =float(final[5])/100

	final = [job_title, ranking, early_pay, middle_pay, pay_growth, High_meaning, High_Satisfaction,low_stress,employment_projection]
	
	for j in range(0,len(final)):
		
		sheet.cell(row = r, column = j+1).value = final[j]
	r += 1


wb.save('forbes.xlsx')
#Justin Womack, 12/20/17
#First name of both lists doesn't list in excel.  This is due to the website using a title Bold tag
#it would take more time to code a solution up then just simply typing these in  
#Need updated Regex code.  any job title with & in it keeps the &amp tag.


















