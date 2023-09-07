import pdfplumber, os , sys, openpyxl

 
 
def write_excel_xlsx(path, sheet_name, value):
    index = len(value)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['C'].width = 40
    sheet.title = sheet_name
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.cell(row=i+1, column=j+1, value=str(value[i][j]))
    workbook.save(path)
    print("xlsx格式表格写入数据成功！")
 
 
xlsx_file_name = 'summary.xlsx'
 
sheet_name_xlsx = '发票类型及金额'
 
content = [["条目", "金额", "文件名称", "合计", "总计"]]

summary_map = {}

def mapToContent(summary_map) :
    aggregate_money = 0.0
    for key, value_list in summary_map.items() :
        first = True
        catgory_money = 0.0
        for value in value_list :
            catgory_money = catgory_money + value[0]
            if first :
                value.insert(0, key)
                value.append("")
                value.append("")
                content.append(value)
            else :
                value.insert(0, "")
                value.append("")
                value.append("")
                content.append(value)
            first = False
        total = ["", "", "", catgory_money, ""]
        content.append(total)
        aggregate_money = aggregate_money + catgory_money
    aggregate = ["", "", "", "", aggregate_money]
    content.append(aggregate)
def getFileName(file) :
    dir_name, full_file_name = os.path.split(file);
    return full_file_name;

def getCatgoryMoney(file_path) :
    pdf =  pdfplumber.open(file_path) 

    first_page = pdf.pages[0]

    text = first_page.extract_text()
    
    text_list = text.split('\n')
    index = 0
    for line in text_list :
        if line.find(u'货物或应税劳务') != -1 :
            catgory = text_list[index + 1].split(' ')[0]
        if line.find(u'价税合计') != -1 :
            money = float(line.split('¥')[1])
        index = index + 1
    return [catgory, money]

def main() :
    pdf_file_list = [];
    current_dir = os.getcwd();
    
    for root, dirs, files in os.walk(current_dir, topdown=False):
        for name in files:
            if name.endswith(".pdf") :
                pdf_file_list.append(os.path.join(root, name));
    for file in pdf_file_list :
        catgoryWithMoney = getCatgoryMoney(file)
        file_name = getFileName(file)
        if len(catgoryWithMoney[0]) == 0 :
            print(file , " has error")
            input("按下任意键退出程序")
            exit(-1)
        catgoryWithMoney.append(file_name)
        item_list = catgoryWithMoney
        summary_map.setdefault(item_list[0], []).append(item_list[1 : ])

    mapToContent(summary_map)        
    write_excel_xlsx(xlsx_file_name, sheet_name_xlsx, content)
    input("按下任意键退出程序")
    
if(__name__ == '__main__') :
    main();   