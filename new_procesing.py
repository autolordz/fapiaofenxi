import pdfplumber,os,glob,re,pathlib,datetime,string,shutil,time
import pandas as pd
from itertools import chain

t0 = time.time()
print('Start Move All'.center(30,'*'))

list_tags = dict( 
     # 以下类目有先后顺序
     油费 = '汽油|柴油|加油',
     交通 = '客运|交通|地铁|滴滴|出行|摩拜|骑安|客运服务',
     差旅 = '差旅|酒店|住宿|寄存|旅游|旅店|宾馆|旅行',
     文娱 = '广播影视服务|文化|电影|影视|娱乐',
     服饰 = '纺织产品|服装|服饰|衣服|衫|裤|裙|袜|鞋|饰品|运动服|盖璞|迅销|飒拉',
     办公 = '纸制品|印刷|文具|打印纸|笔|文件夹|胶水|回形针|剪刀|纸刀|订书|书桌垫',
     数码设备 = '电线电缆|配电控制设备|移动通信设备|数码|设备|电脑|手机|麦克风|耳机|相机|USB|转换器|插座|路由|显示器|键盘|鼠标|灯',
     家具 = '家具|照明装置',
     家电 = '非电力家用器具',
     医疗 = '医疗服务|宠物|美容|护肤|自疗|理疗|健康',
     教育服务 = '教育服务|培训|考试|报名|课程',
     技术服务 = '信息技术服务|增值服务',
     餐饮 = '餐饮服务', 
     食品 = '植物油|茶|饲料|蔬菜|水果|饮料|食品|糖|糖果|焙烤食品|肉及肉制品|方便食品|谷物|调味品|乳制品|水产加工品|果类加工品|营养保健食品|酒|海水产品|加工盐|熟肉制品', 
     物业管理 = '售电|管理费|物业管理|租赁|保洁|垃圾费|租金|供电|水冰雪',
     运输服务 = '运输服务',
     物流服务 = '物流',
     现代服务 = '设计服务|现代服务',
     生活百货 = '预付|美容护肤品|日用杂品|其他化学制品|洗涤剂',
     经营 = '无形资产|服务费|体育用品|轴承|五金|包装',
     详见销货清单 = '详见销货清单'
     )


zhon_pun = "！？｡＂＃＄％＆＇（）＊＋，－／：；＜＝＞＠［＼］＾＿｀｛｜｝～｟｠｢｣､、〃》「」『』【】〔〕〖〗〘〙〚〛〜〝〞〟〰〾〿–—‘’‛“”„‟…‧﹏."
en_pun = f'{(string.whitespace+string.punctuation)}：；'

PAT1 = rf'[{zhon_pun+en_pun}]'

# Default return Catalog will be '其他'
def catalogf(content):
    keys = re.findall(rf'\*[{zhon_pun}\u4e00-\u9fa5\s\n]+\*', content)
    ll = [len([x for x in keys if re.search(tag,x)]) for tag in list_tags.values()]
    catalog = '其他'
    if max(ll) > 0:
        catalog = list(list_tags.keys())[ll.index(max(ll))]
    return catalog

# 读取的内容过短不够字符数的合并到下一行
def list_combine(list1):
    it = iter(list1)
    for x in it:
        x = x.translate({ord(c): None for c in '\n'})
        # print(x)
        if len(x) < 6:
            try:
                x += ' ' + next(it)
            except StopIteration:
                pass
        yield x

#表头格式
RECORD_HEADER=['序号','发票号码','卖家','卖家号码','买家','买家号码','开票日期','月份','金额','类别','内容']

# 建立记录到指定目录
FILE_FOLDER=r'D:\xx'
RECORD_PATH = os.path.join(FILE_FOLDER, 'iv_records.xlsx')

if not os.path.exists(RECORD_PATH):
    df = pd.DataFrame(columns=RECORD_HEADER)
else:
    df = pd.read_excel(RECORD_PATH,dtype=object,na_filter=False)
df1 = df.copy()

# 索取目录下所有pdf文件
p4 = [str(p1) for p1 in pathlib.Path(FILE_FOLDER).glob('*') if str(p1).endswith('.pdf')]

number1 = 1
for file_path in p4: #[:10]
    # print(file_path)
    try:
        with pdfplumber.open(file_path) as pdf:
            t1 = pdf.pages[0].extract_text_simple(x_tolerance=4, y_tolerance=3)
            # datestr1 = re.search(r'(?<=%s)\d{4}年\d{2}月\d{2}日|$'%re.search(r'开票日期[\s:：]|$',t1).group().strip(),t1).group().strip()
            datestr1 = re.search(r'(?<=%s).*|$'%re.search(r'开票日期[\s:：]|$',t1).group().strip(),t1).group().strip()
            datestr1 = re.sub(PAT1,'', datestr1)
            date2 = datetime.datetime.strptime(datestr1, '%Y年%m月%d日').strftime('%Y年%m月')
            Invoicecode = re.search(r'(?<=%s).*|$'%re.search(r'发票号码[\s:：]|$',t1).group().strip(),t1).group().strip()
            Invoicecode = re.sub(PAT1,'', Invoicecode)
    
            t2 = pdf.pages[0].extract_table(table_settings={"text_tolerance": 4})
            t5 = list(list_combine(filter(None, chain(*t2))))
            t4 = t5.copy()
            for i,x in enumerate(t4):
                if '购买' in x:
                    Buyer = re.search(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%re.search(r'称\s*[\s:：]|$',x).group().strip(), x).group().strip()
                    Buyercode = re.sub(PAT1,'',re.search(r'(?<=%s).*|$'%re.search(r'纳税.{1}识别号|$',x).group().strip(),x).group())
                    t4.pop(i)
                    break
            for i,x in enumerate(t4):
                if '销售' in x:
                    Seller = re.search(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%re.search(r'称\s*[\s:：]|$',x).group().strip(), x).group().strip()
                    Sellercode = re.sub(PAT1,'',re.search(r'(?<=%s).*|$'%re.search(r'纳税.{1}识别号|$',x).group().strip(),x).group())
                    t4.pop(i)
                    break
            for i,x in enumerate(t4):
                if '小写' in x:
                    Price1 = re.search(r'(?<=[¥￥]).*|$', re.search(r'(?<=小写).*|$', x).group().strip()).group().strip()
                    # money2 = x[x.index]
                    t4.pop(i)
                    break
                
            if len(pdf.pages) > 1:
                # 有大于一页发票
                content1 = pdf.pages[1].extract_text_simple(x_tolerance=4, y_tolerance=3)
            else:
                for i,x in enumerate(t4):
                    if '名称' in x and '*' in x:
                        content1 = x
                        t4.pop(i)
                        break
            catalog1 = catalogf(content1)
            row1 = [number1,Invoicecode,Seller,Sellercode,Buyer,Buyercode,datestr1,date2,Price1,catalog1,content1]
            df.loc[len(df),:] = row1
            number1 +=1
            
        # 1 生成文件夹
        FolderDate = datetime.datetime.strptime(datestr1, '%Y年%m月%d日').strftime('%Y %m月')
        folder4 = os.path.join(FILE_FOLDER,Buyer,FolderDate)
        # print(folder4)
        os.makedirs(folder4,exist_ok=1)
        # 2 改名 Move the file into folder
        fname = f'{Seller}-{catalog1}-{Price1}-{datestr1}-{Invoicecode}.pdf'
        new_file = os.path.join(folder4,fname)
        shutil.move(file_path,new_file)
        print(new_file)
        
    except Exception as e:
        print(f'Error occured !',e)

df.drop_duplicates(['发票号码'],keep='last',inplace=True)
df = df.reset_index(drop=True)
df['发票号码'] = df['发票号码'].apply(str)
df['金额'] = df['金额'].apply(float)

if not df1.iloc[:,1:].equals(df.iloc[:,1:]):
    # df.to_csv(RECORD_PATH,encoding='utf-8-sig',index=0)
    df.to_excel(RECORD_PATH,index=0) # encoding='utf-8-sig'
    
os.system('start excel %s'%RECORD_PATH)

print(('End Move All time: %.3f'%(time.time() - t0)).center(30,'*'))
