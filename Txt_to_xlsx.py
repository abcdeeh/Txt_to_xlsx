from Txt_to_xlsx_programs import Copy_and_Paste,paste1
import xlsxwriter,os,sys
import numpy as np

files = []
i = 0
while i < 3:
	print("Drag and drop data directory")
	path =input()
	files2=np.array(files)
	a=str(path[0])
	if a=="\'" or a=='\"':
		path=str(path[1:-2])

	del a
	while i < 3:
		if os.path.exists(path):
			break
		else:
			print("<error:The path does not exist.>")
			print("Drag and drop data's directory")
			path =input()


	try:
		f=open(path,'r')
		f.close()
		base, ext = os.path.splitext(path)
		if ext=='.txt' or ext=='' or ext=='.text':
			break
	except:
		pass
	try:
		for f in os.listdir(path):
			if os.path.isfile(os.path.join(path, f)):
				base, ext = os.path.splitext(f)
				if ext=='.txt' or ext=='' or ext=='.text':
					files2=np.append(files2,f)
		if files2.shape[0]!=0:
			break
	except:
		pass





kkk=np.array(path.split("\\"))
pqr=np.array(kkk[kkk.shape[0]-1].split("/"))
del kkk
asds=pqr.shape[0]
ands=pqr[asds-1]
file_name= ands +".xlsx"
print("Output file`s name;"+file_name)
print("sheet's name")
if files2.shape[0]!=0:
	path1=os.path.join(path, files2[i])
else:
	path1=path
base, ext = os.path.splitext(path1)
kkk=np.array(base.split("\\"))
pqr=np.array(kkk[kkk.shape[0]-1].split("/"))

asds=pqr.shape[0]
Sheet=pqr[asds-1]
if len(Sheet)>30:
    Sheet=Sheet[0:28]
book = xlsxwriter.Workbook(file_name, {'constant_memory': True})
book.add_worksheet(Sheet)
Copy_and_Paste.copy(book,Sheet,path1)
i+=1
p=2
while i < files2.shape[0]:
	path1=os.path.join(path, files2[i])
	base, ext = os.path.splitext(path1)
	kkk=np.array(base.split("\\"))
	pqr=np.array(kkk[kkk.shape[0]-1].split("/"))
	asds=pqr.shape[0]
	Sheet=pqr[asds-1]
	if len(Sheet)>30:
		Sheet=Sheet[0:28]
	try:
		book.add_worksheet(Sheet)
	except:
		try:
			Sheet1=Sheet+"1"
			book.add_worksheet(Sheet1)
			Sheet=Sheet1
			del Sheet1
			p=2
		except:
			while p<100:
				try:
					Sheet1=Sheet+str(p)
					book.add_worksheet(Sheet1)
					Sheet=Sheet1
					del Sheet1
					break

				except:
					p+=1


	if p==100:
		print("Reduce the number of files to be read")
		input()
		book.close()
		sys.exit()
	Copy_and_Paste.copy(book,Sheet,path1)
	i+=1
book.close()
