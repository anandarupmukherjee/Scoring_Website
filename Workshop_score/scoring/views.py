from django.shortcuts import render
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
from django.conf import settings
from django.core.files.storage import FileSystemStorage
import os

plt.rcdefaults()
sns.set(style='whitegrid', rc={"grid.linewidth": 1.2, "xtick.labelsize" : 14, 
                               "ytick.labelsize" : 14, "font.family":'Times New Roman', "boxplot.boxprops.color": "black"})


font1 = {'family': 'Times New Roman',
         'weight': 'bold',
         'size': 18}

font2 = {'family': 'Times New Roman',
         'weight': 'bold',
         'size': 16}

font3 = {'family': 'Times New Roman',
         'weight': 'bold',
         'size': 14}

# Create your views here.
#def home(request):
#   return render(request, 'home.html')

def home(request):
    
    print("Cache cleared")
    if request.method == 'POST' and request.FILES['myfile']:
        os.system("rm media/*")
        myfile = request.FILES['myfile']
        fs = FileSystemStorage()
        filename = fs.save(myfile.name, myfile)
        uploaded_file_url = fs.url(filename)
        return render(request, 'home.html', {
            'uploaded_file_url': uploaded_file_url
        })
    return render(request, 'home.html')



def index(request):

    #excel_file = request.FILES["/media/Priority_areas.xlsx"]
    #wb = openpyxl.load_workbook("workshop/media/Priority_areas.xlsx")
    p=os.listdir('media')
    print(p)
    

    xls = pd.ExcelFile('media/Priority_areas.xlsx', engine='openpyxl')
    print("success")
    df1 = pd.read_excel(xls, 'Quality', names=["Solutions","Priority", "Interesting"])
    df2 = pd.read_excel(xls, 'Agility', names=["Solutions","Priority", "Interesting"])
    df3 = pd.read_excel(xls, 'Price', names=["Solutions","Priority", "Interesting"])
    df4 = pd.read_excel(xls, 'Delivery', names=["Solutions","Priority", "Interesting"])
    df5 = pd.read_excel(xls, 'Sustainability', names=["Solutions","Priority", "Interesting"])
    df6 = pd.read_excel(xls, 'People and Information', names=["Solutions","Priority", "Interesting"])
    df7 = pd.read_excel(xls, 'Plant and Equipment', names=["Solutions","Priority", "Interesting"])
    df8 = pd.read_excel(xls, 'Supply Chain', names=["Solutions","Priority", "Interesting"])
    df9 = pd.read_excel(xls, 'Demand Management', names=["Solutions","Priority", "Interesting"])
    df10 = pd.read_excel(xls, 'Cash Flow', names=["Solutions","Priority", "Interesting"])

    print('stage-1 clear')

    data = [df1, df2, df3, df4, df5, df6, df7, df8, df9, df10]
    master_df = pd.concat(data, ignore_index=True, sort=True)
    master_df['Solutions'].apply(lambda x: x.replace("'",""))

    new_master=master_df[master_df.duplicated(['Solutions'], False)]
    new_master=new_master.sort_values(['Solutions'])
    gk=master_df.groupby("Solutions").sum()
    gk=gk.sort_values(["Priority"], ascending=[False])

    print('stage-2 clear')

    pr=gk['Priority'].iloc[:10]
    intr=gk['Interesting'].iloc[:10]
    labels=gk.index[:10]
    width=0.35
    intr2=pd.concat([pd.DataFrame(pr),pd.DataFrame(intr)], axis=1,join="inner")
    intr2=intr2.sort_values(by=["Priority","Interesting"])
    intr3=intr2.iloc[::-1] 

    print('stage-3 clear')


    plt.figure(figsize=(8,8))
    plt.bar(intr3.index, intr3["Interesting"], width, bottom=pr,label='Interesting',alpha=0.65, edgecolor='black', color='orange')
    plt.bar(intr3.index, intr3["Priority"], width, label='Priority',alpha=0.65, edgecolor='black', color='green')
    plt.ylabel("No. of selections", **font1)
    plt.xlabel("Solutions",**font1)
    plt.xticks(rotation=90, **font2)
    plt.yticks(rotation=90, **font1)
    plt.legend()
    plt.savefig("media/votes.png",dpi=400, bbox_inches = 'tight')
    plt.show()

    return render(request, 'index.html')
