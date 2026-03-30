import pdfplumber, os, sys, openpyxl, logging, re

logger = logging.getLogger(__name__)


def _silence_third_party_loggers():
    """避免 basicConfig(DEBUG) 打开 pdfminer 等库的刷屏英文 DEBUG。"""
    for name in (
        "pdfminer",
        "pdfminer.psparser",
        "pdfminer.pdfinterp",
        "pdfminer.pdfpage",
        "pdfminer.converter",
        "pdfplumber",
        "PIL",
    ):
        logging.getLogger(name).setLevel(logging.WARNING)


def _ensure_utf8_stdio():
    """Windows 控制台默认编码常导致中文日志乱码或看起来像“没有中文”。"""
    if sys.platform != "win32":
        return
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def configure_logging(verbose=False):
    """
    只为本脚本配置日志：中文 INFO/WARNING/ERROR 正常输出；
    verbose=True 或环境变量 PDF_INVOICE_DEBUG=1 时输出本脚本的 DEBUG（如第一页全文）。
    """
    _ensure_utf8_stdio()
    _silence_third_party_loggers()

    env_debug = os.environ.get("PDF_INVOICE_DEBUG", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    level = logging.DEBUG if (verbose or env_debug) else logging.INFO

    logger.handlers.clear()
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    h = logging.StreamHandler(sys.stdout)
    h.setLevel(level)
    h.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    )
    logger.addHandler(h)


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


def _invoice_line_compact(s):
    """去掉空白后再比对关键词；pdfplumber 常把「价税合计」拆成「价税 合计」。"""
    if not s:
        return ""
    return re.sub(r"\s+", "", s)


def _parse_money_after_yen_symbols(line):
    """从一行中按 ￥/¥ 截取，取其后第一个金额（支持千分位逗号、¥ 后空格）。"""
    if not line:
        return None
    for sym in ("￥", "¥"):
        if sym not in line:
            continue
        for part in line.split(sym)[1:]:
            part = part.strip().replace(",", "")
            m = re.match(r"^(\d+(?:\.\d+)?)", part)
            if m:
                try:
                    return float(m.group(1))
                except ValueError:
                    continue
    return None


def getCatgoryMoney(file_path):
    catgory = ""
    money = None
    logger.info("处理 PDF: %s", file_path)

    pdf = pdfplumber.open(file_path)
    try:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        if text is None:
            logger.warning("第一页 extract_text() 返回 None，可能为扫描件或无法提取文字")
            text = ""
        else:
            logger.debug("第一页文本长度: %d", len(text))

        logger.info("第一页 extract_text 解析内容:\n%s", text if text else "(无文本)")

        text_list = text.split("\n")
        logger.debug("第一页按行数: %d", len(text_list))

        index = 0
        for line in text_list:
            line = line or ""
            if "货物或应税劳务" in line:
                if index + 1 < len(text_list):
                    catgory = text_list[index + 1].split(" ")[0]
                    logger.debug("行 %d 匹配「货物或应税劳务」，类目: %r", index, catgory)
                else:
                    logger.warning("行 %d 匹配「货物或应税劳务」但无下一行", index)
            if "项目名称" in line:
                if index + 1 < len(text_list):
                    catgory = text_list[index + 1].split(" ")[0]
                    logger.debug("行 %d 匹配「项目名称」，类目: %r", index, catgory)
                else:
                    logger.warning("行 %d 匹配「项目名称」但无下一行", index)
            if "价税合计" in _invoice_line_compact(line):
                logger.debug("行 %d 匹配「价税合计」(去空白后): %r", index, line)
                parsed = _parse_money_after_yen_symbols(line)
                if parsed is None:
                    logger.warning("「价税合计」行未能从 ￥/¥ 后解析金额: %r", line)
                else:
                    money = parsed
                    logger.debug("解析金额: %s", money)
            index += 1

        if not catgory:
            logger.warning("未匹配到「货物或应税劳务」或「项目名称」后的类目，catgory 为空")
        if money is None:
            logger.error(
                "未解析到金额（第一页无「价税合计」或去空白后仍不匹配，或 ￥/¥ 后数字解析失败）。"
                " 电子票常见为「价税 合计」被拆字加空格；上文 INFO 已打印本页全文。"
            )

        return [catgory, money if money is not None else 0.0]
    finally:
        pdf.close()

def main():
    configure_logging(verbose=False)
    pdf_file_list = []
    current_dir = os.getcwd()
    logger.info("工作目录: %s", current_dir)

    for root, dirs, files in os.walk(current_dir, topdown=False):
        for name in files:
            if name.endswith(".pdf"):
                pdf_file_list.append(os.path.join(root, name))
    logger.info("共发现 %d 个 PDF 文件", len(pdf_file_list))

    for file in pdf_file_list:
        file_name = getFileName(file)
        print(file_name)
        logger.info("开始: %s", file_name)
        try:
            catgoryWithMoney = getCatgoryMoney(file)
        except Exception:
            logger.exception("解析失败: %s", file)
            raise
        logger.debug("结果: catgory=%r, money=%r", catgoryWithMoney[0], catgoryWithMoney[1])
        if len(catgoryWithMoney[0]) == 0:
            logger.error("类目为空，终止: %s", file)
            print(file, " has error")
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