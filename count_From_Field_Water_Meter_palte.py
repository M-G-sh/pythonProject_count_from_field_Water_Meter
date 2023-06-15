import arcpy
import pandas as pd
from openpyxl import Workbook


def process_data(feature_class, field_name, second_field, output_excel):
    # إنشاء طبقة نتيجة الاختيار
    selection_layer = arcpy.management.MakeFeatureLayer(feature_class, "تحديد")

    # استعلام لاستخراج جميع القيم الموجودة في الحقل الأول
    values = set([row[0] for row in arcpy.da.SearchCursor(feature_class, field_name)])

    results = []

    for value in values:
        # تشكيل الشرط بناءً على القيمة الحالية
        expression = f"'{value}'"
        where_clause = f"{field_name} = {expression} AND ({second_field} IS NOT NULL AND {second_field} <> 'Closed' AND {second_field} <> 'Not_Clear' AND {second_field} <> '0' AND {second_field}  <> 'مغلق ' )"

        # تنفيذ الاختيار بناءً على الشرط المحدد
        arcpy.management.SelectLayerByAttribute(selection_layer, "NEW_SELECTION", where_clause)

        # عرض عدد الكائنات المحددة للقيمة الحالية
        result = arcpy.GetCount_management(selection_layer)
        count = int(result.getOutput(0))
        print(f"تم تحديد {count} كائنات لقيمة {value} في الحقل {field_name}.")

        # حذف الاختيار السابق للاستعداد للاختيار التالي
        arcpy.management.SelectLayerByAttribute(selection_layer, "CLEAR_SELECTION")

        # إضافة النتائج إلى قائمة النتائج
        results.append((value, count))

    # حذف الطبقة المؤقتة للاختيار
    arcpy.management.Delete(selection_layer)

    # تحويل قائمة النتائج إلى DataFrame
    df = pd.DataFrame(results, columns=["Value", "Palate_No"])

    # إنشاء ملف Excel
    writer = pd.ExcelWriter(output_excel, engine='openpyxl')
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # تعديل اسم العمود في ورقة العمل الأولى
    workbook = writer.book
    worksheet = workbook['Sheet1']
    worksheet.cell(row=1, column=2).value = 'Water_Meter_No'

    # حفظ الملف Excel
    writer.save()

    print(f"تم حفظ النتائج في ملف {output_excel}.")


# تحديد مسار الـ feature class
feature_class = "D:/Work_Project/داتا شركة المياه الوطنية/part10/29حى لملوم/DB/احياء تجميع 29 حى اصلى/Al_DAta.gdb/AGP_METERINGPOINT"

# تحديد الحقول المطلوبة ومسار ملف الإكسل الناتج
field_name_1 = "area"
second_field_1 = "water_meter_no"
output_excel_1 = "D:/Work_Project/داتا شركة المياه الوطنية/part10/29حى لملوم/TEST1.xlsx"

field_name_2 = "area"
second_field_2 = "nwc_barcode_plate"
output_excel_2 = "D:/Work_Project/داتا شركة المياه الوطنية/part10/29حى لملوم/TEST2.xlsx"

# تنفيذ الدالة الأولى
process_data(feature_class, field_name_1, second_field_1, output_excel_1)

# تنفيذ الدالة الثانية
process_data(feature_class, field_name_2, second_field_2, output_excel_2)