import xlwings as xw
from win32com.client import DispatchEx

# Connect to the Excel application
app = xw.App(visible=False)  # Set visible=True if you want to see Excel
# Open the workbook (replace 'your_workbook.xlsx' with the actual file path)
workbook = xw.Book("G:\ITB guns\PENTINGGGGG!!\Portofolio\Source\PycharmProjects\Wellness PHM\input_file_period2.xlsx")
# Select Sheet2
sheet = workbook.sheets["Sheet2"]
# Your Python list (example)
name_list = ['Rasyid supriadi','Bima Adie Nugraha','Erlan Misdiono ','Wawan Prawiardy','Bukti Hamonangan','Delano Kawilarang ','Eko Suretno','Andrie Saputra','Pradito Mahayani','Hadrianus Septiwiratmiarto ','Ahmad Budi Wicaksono ','Marlon Epafras Sium Simanjuntak','Tahan Bati','Andrianto Wahyuda','Muslimin Md','Umar Rohim ','Hariyanto','YOFI FAUZY','Sandi','Gusti wildan','DIMAS FAJAR ADE NOFAN','M ali akbar','Edilaman ','Budiman','Jumardin side','Juwairi','M.Wahyu','D A WIDAYAT','Herwin','Muhammad rizky maulana','Abdul Nasser','Jumansyah','Syamsuddin Halim','Imam Nurcahya','Verdy Dekker']
site_list = ['BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP','BKP']
selected_list = [271,373,375,376,377,382,383,384,385,386,387,389,391,392,393,394,395,396,397,399,401,403,404,405,407,408,409,410,411,412,415,416,417,418,41]
iteration = 0
# Use 'for in list' for specific output
for i in selected_list:
# for i in range (0,430): for non spesific output
# for i in range (0,430):
    # Write the first value from the list to cell A1
    # sheet.range("A1").value = i+1 #for non specific output
    sheet.range("A1").value = i #for specific output
    # Save changes (optional)
    workbook.save()
    # Convert to PDF
    excel_app = DispatchEx("Excel.Application")
    workbook_pdf = excel_app.Workbooks.Open("G:\ITB guns\PENTINGGGGG!!\Portofolio\Source\PycharmProjects\Wellness PHM\input_file_period2.xlsx")
    worksheet = workbook_pdf.Worksheets("HealthProfile")
    # pdf_path = f"G:\ITB guns\PENTINGGGGG!!\Portofolio\Source\PycharmProjects\Wellness PHM\Wellness Profile_{name_list[i]}_{site_list[i]}.pdf"
    pdf_path = f"G:\ITB guns\PENTINGGGGG!!\Portofolio\Source\PycharmProjects\Wellness PHM\Wellness Profile_{name_list[iteration]}_{site_list[iteration]}.pdf" #for specific output
    worksheet.ExportAsFixedFormat(0, pdf_path)
    workbook_pdf.Close(False)
    iteration +=1 #for specific output
# Close the workbook and Excel application
workbook.close()
app.quit()
