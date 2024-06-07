# Only for China's Invoice

FILE_FOLDER = r'D:/your/file/path/'
# 找出folder里面所有pdf，这里特指发票pdf

for file_path in glob.glob(os.path.join(FILE_FOLDER,'*.pdf')):
#使用pdfplumber 0.11 版本分析内容，用这个model原因是试了很多还是觉得这个比较好用
    with pdfplumber.open(file_path) as pdf:
#读取pdf文字内容，这里使用simple方法和tolerance结合，tolerance=3是读取间隔，允许3个空格以内不换行，具体查看pdfplumber文档
        t1 = pdf.pages[0].extract_text_simple(x_tolerance=3, y_tolerance=3)
#找到中文字符
        if re.search(r'([\u4e00-\u9fa5])\1{2,}',t1):
            #定位汉字及中文符号，这里删除pdfplumber解释出某些发票出现叠字重复字例如**发发发票票票**问题
            t1 = re.sub(r'([\u4e00-\u9fa5\u3000-\u303f\u3105-\u312f\u31a0-\u31bf\uff00-\uffef])\1{2,}',r'\1',t1)
# 一般购销的两个公司名都会包含在这一行内容，**购 名称: xx公司 销 名称: xx公司**，这里用个条件兼容一些没有**名称**两字的，需要判断**购...销**
        if not '名称' in t1[t1.index('购'):t1.index('销')]:
            print(re.findall(r'(?<=购|销)\s?[\u4e00-\u9fa5()（）]+', t1))
        else:
            print(re.findall(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%re.search(r'称\s*[\s:：]|$',t1).group(), t1))
# 获取发票号码
        InvoiceCode = re.search(r'(?<=发票号码[:：]).*|(?<=发票号码 [:：]).*|$', t1).group().strip()
# 获取金额，金额数字提取在**小写和¥￥**之后
        money1 = re.search(r'(?<=小写).*|$', t1).group().strip()
        money1 = re.search(r'(?<=[¥￥]).*|$', money1).group().strip()
# 提取日期
        xx = re.search(r'(?<=日期).*日|$', t1).group().strip()
        datestr1 = re.sub(r'[:：\s*]','',xx)
# 发票商品内容, 某些发票 截取内容从：纳税人识别号or身份证号.... 至：普通发票代码
        idx_t1 = re.search('纳税.{1}识别号',t1)
        if not idx_t1:
            idx_t1 = re.search('身份证号',t1)
        idx_t1 = idx_t1.span()[0]
        idx_t2 = t1.index('小写')
        Content = t1[idx_t1:idx_t2] if not t_page2 else t_page2[t_page2.index('普通发票代码'):]
