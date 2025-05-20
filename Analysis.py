#------------------------
#---> المكتبات
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns 
from matplotlib.backends.backend_pdf import PdfPages

#----> اعدادات شكل الرسومات
sns.set(style='whitegrid')
#---> قراءة البيانت من ملف الاكسيل
file_bath = (r"C:\Users\Abdo  Mahmoud\Desktop\مشاريع التحليل و بليثون لي GitHub\Proj-power and python\sales_project_data.xlsx")
sales_df   = pd.read_excel(file_bath,  sheet_name="Sales")
returns_df = pd.read_excel(file_bath,  sheet_name="Returns")
review_df  = pd.read_excel(file_bath,  sheet_name="Reviews")
#----> نظرة علي البيانات (اول 5 صفوف)
print("\n مبيعات : ")
print( sales_df .head())        #----> يظهر اول صفوف في شيت المبيعات
print("\n مرتجعات : ")
print( returns_df .head())      #---->  يظهر اول صفوف في شيت المرتجعات
print("\n مراجعات : ")
print( review_df .head())       #---->  يظهر اول صفوف في شيت المرجعات

#--------------------------------------
#----> معالجة البيانات (Data Type)
#---------------------------------
#--> معاجة التواريخ في كل شيت
sales_df['Sale_Date'] = pd.to_datetime(sales_df["Sale_Date"])
returns_df['Return_Date'] = pd.to_datetime(returns_df["Return_Date"])
review_df['Review_Date'] = pd.to_datetime(review_df["Review_Date"])
#---->  اضافة عمود للشهور في المبيعات لاستخدام  في التحليل و الفلتر
sales_df['manth'] = sales_df['Sale_Date'].dt.to_period('M')

#----------------------------
#----> (EDA) تحليل استكشافي
#-------------------------------------
#---> اعلي المنتجات مبيعا
Top_Product = sales_df.groupby('Product_Name')['Quantity_Sold'].sum().sort_values(ascending=False).head(10)
print("\n اعلب المنتجات مبيعا : ")
print(Top_Product)
#---> توزيع التصفيات
category_distribution  = sales_df['Category'].value_counts()
#---> تاثير الخصومات
discount_effect = sales_df.groupby('Discount')['Quantity_Sold'].sum()
#---> متوسط تقيم المنتجات
avg_rating = review_df.groupby('Product_ID')['Rating'].mean()
#----------------------
#---> انشاء ملف تلخيص تحليلي
#------------------------
pdf_path = (r"C:\Users\Abdo  Mahmoud\Desktop\مشاريع التحليل و بليثون لي GitHub\Proj-power and python/Sales_Analysis_report.pdf")
with PdfPages(pdf_path) as pdf :
    #---> ملخص الكود علي شكل جداول
    fig, ax = plt.subplots(figsize= (8.5,11))
    ax.axis('off')
    text = F"""
      تقرير تحليلي :
      
      * اعلي المنتجات مبيعا : {Top_Product.to_string()}
      * توزرع التصنيفات : {category_distribution.to_string()}
      * تاثير الخصومات  :{discount_effect.to_string()}
      * متوسط تقيم المنتجات : {avg_rating.sort_values(ascending=False).head(5).to_string()}
      """
    ax.text(0.05, 0.95, text, va='top', fontsize= 10, family= 'omnospace' )
    pdf.savefig()
    plt.close()

    #---> المنجات الاعلي مبيعا
    plt.figure(figsize=(10,6))
    Top_Product.plot(kind='bar' , color='skyblue')
    plt.title("اعلي المنتجات مبيعا")
    plt.ylabel('الكمية المباعة')
    plt.xticks(rotation = 45 ,ha = 'right')
    plt.tight_layout()
    plt.savefig()
    plt.close()

#------------------
print("\n\u2714\ufe0f PDF : Sales_Analysis_report.pdf ")
pdf_path









