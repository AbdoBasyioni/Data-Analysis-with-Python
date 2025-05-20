
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
# إعدادات عامة
sns.set(style="whitegrid")
plt.rcParams['font.family'] = 'Arial'  # يدعم اللغة العربية لو مثبت

# تحميل البيانات
file_path = (r"C:\Users\Abdo  Mahmoud\Desktop\مشاريع التحليل و بليثون لي GitHub\Task_2\بيانات_مبيعات_وهمية.xlsx")
xls = pd.ExcelFile(file_path)
output_excel = "نتائج_تحليل_العملاء.xlsx"
results_dir = "نتائج_التحليل"
os.makedirs(results_dir, exist_ok=True)

# تحميل البيانات
xls = pd.ExcelFile(file_path)
df = xls.parse(xls.sheet_names[0])

# عرض أسماء الشيتات واختيار أول شيت
print("Sheets:", xls.sheet_names)
df = xls.parse(xls.sheet_names[0])

# ============== التحليلات ============== #

# 2. أشهر طريقة دفع
if 'طريقة الدفع' in df.columns:
    top_payment = df['طريقة الدفع'].value_counts().idxmax()
    payment_counts = df['طريقة الدفع'].value_counts()
    plt.figure(figsize=(6,4))
    sns.countplot(data=df, x='طريقة الدفع', order=payment_counts.index, palette='Set2')
    plt.title('توزيع طرق الدفع')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{results_dir}/طرق_الدفع.png")
    plt.close()

# 3. أعلى العملاء تقييمًا
if 'تقييم العميل' in df.columns and 'اسم العميل' in df.columns:
    top_rated = df[['اسم العميل', 'تقييم العميل']].drop_duplicates('اسم العميل')
    top_rated_sorted = top_rated.sort_values(by='تقييم العميل', ascending=False).head(5)

# 4. توزيع التقييمات
if 'تقييم العميل' in df.columns:
    plt.figure(figsize=(6,4))
    sns.histplot(df['تقييم العميل'], bins=10, kde=True, color='orange')
    plt.title('توزيع تقييمات العملاء')
    plt.tight_layout()
    plt.savefig(f"{results_dir}/توزيع_التقييمات.png")
    plt.close()

# 5. أكثر المدن أو الفروع مبيعًا
city_branch_analysis = {}
for col in ['المدينة', 'الفرع']:
    if col in df.columns:
        top_items = df[col].value_counts().head(5)
        city_branch_analysis[col] = top_items
        plt.figure(figsize=(6,4))
        sns.barplot(x=top_items.values, y=top_items.index, palette='Blues_d')
        plt.title(f'أكثر {col} من حيث النشاط')
        plt.xlabel('عدد المعاملات')
        plt.tight_layout()
        plt.savefig(f"{results_dir}/تحليل_{col}.png")
        plt.close()

# حفظ النتائج إلى ملف Excel
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name="بيانات_نظيفة", index=False)
    if 'طريقة الدفع' in df.columns:
        payment_counts.to_frame(name="عدد المعاملات").to_excel(writer, sheet_name="طرق_الدفع")
    if 'تقييم العميل' in df.columns and 'اسم العميل' in df.columns:
        top_rated_sorted.to_excel(writer, sheet_name="أعلى_العملاء_تقييما", index=False)
    for key, value in city_branch_analysis.items():
        value.to_frame(name='عدد المعاملات').to_excel(writer, sheet_name=f"تحليل_{key}")

print("✅ تم الانتهاء من تحليل البيانات وحفظ النتائج والصور.")
