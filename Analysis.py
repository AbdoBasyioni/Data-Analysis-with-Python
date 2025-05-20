
import pandas as pd
import random
from faker import Faker

fake = Faker('ar_EG')

# بيانات ثابتة
products = [
    ('هاتف محمول', 'إلكترونيات'),
    ('لاب توب', 'إلكترونيات'),
    ('سماعات', 'إكسسوارات'),
    ('كرسي مكتب', 'أثاث'),
    ('مكيف هواء', 'أجهزة منزلية'),
    ('غسالة', 'أجهزة منزلية'),
    ('حقيبة ظهر', 'أزياء'),
    ('ساعة ذكية', 'إلكترونيات'),
    ('ميكروويف', 'أجهزة منزلية'),
    ('لوحة مفاتيح', 'إكسسوارات'),
]

branches = ['القاهرة', 'الإسكندرية', 'طنطا', 'المنصورة', 'أسيوط']
payment_methods = ['نقدًا', 'بطاقة ائتمان', 'تحويل بنكي', 'محفظة إلكترونية']

data = []
for _ in range(5000):
    order_date = fake.date_between(start_date='-1y', end_date='today')
    customer_name = fake.name()
    product, category = random.choice(products)
    quantity = random.randint(1, 10)
    price = round(random.uniform(50, 5000), 2)
    total_price = round(quantity * price, 2)
    branch = random.choice(branches)
    rating = random.randint(1, 5)
    payment = random.choice(payment_methods)
    
    data.append({
        'تاريخ الطلب': order_date,
        'اسم العميل': customer_name,
        'اسم المنتج': product,
        'الفئة': category,
        'الكمية': quantity,
        'السعر': price,
        'الفرع': branch,
        'إجمالي السعر': total_price,
        'تقييم العميل': rating,
        'طريقة الدفع': payment
    })

df = pd.DataFrame(data)
df.to_excel('بيانات_مبيعات_وهمية.xlsx', index=False)
print(r"C:\Users\Abdo  Mahmoud\Desktop\مشاريع التحليل و بليثون لي GitHub\Task_2\تم حفظ الملف باسم: بيانات_مبيعات_وهمية.xlsx")
