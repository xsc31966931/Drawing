# -*- coding: utf-8 -*-
"""生成《给爸妈看的图纸讲解.docx》——口语化、可打印"""
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"D:\code\Drawing\给爸妈看的图纸讲解.docx"
doc = Document()

# ---------- 页面 A4 + 2.5cm ----------
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.5)

# ---------- 样式工具 ----------
def set_east(style, cn, en, size, bold=False, color=(0, 0, 0)):
    style.font.name = en
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(*color)
    rpr = style.element.get_or_add_rPr()
    old = rpr.find(qn('w:rFonts'))
    if old is not None:
        rpr.remove(old)
    rf = OxmlElement('w:rFonts')
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(attr), en)
    rf.set(qn('w:eastAsia'), cn)
    rpr.insert(0, rf)

set_east(doc.styles['Normal'], '宋体', 'Times New Roman', 12)
normal = doc.styles['Normal']
normal.paragraph_format.line_spacing = 1.5
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after = Pt(0)

set_east(doc.styles['Title'], '黑体', 'Times New Roman', 18, bold=True)
set_east(doc.styles['Heading 1'], '黑体', 'Times New Roman', 16, bold=True)
set_east(doc.styles['Heading 2'], '黑体', 'Times New Roman', 14, bold=True)
set_east(doc.styles['Heading 3'], '黑体', 'Times New Roman', 12, bold=True)
for nm, sb, sa in (('Heading 1', 14, 6), ('Heading 2', 11, 5), ('Heading 3', 9, 4)):
    st = doc.styles[nm]
    st.paragraph_format.space_before = Pt(sb)
    st.paragraph_format.space_after = Pt(sa)
    st.paragraph_format.keep_with_next = True

def indent2(p, chars=200):
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind'); pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars))
    ind.set(qn('w:firstLine'), '480')

def body(text, indent=True):
    p = doc.add_paragraph(style='Normal')
    if indent:
        indent2(p)
    p.add_run(text)
    return p

def labeled(label, text):
    """加粗小开头 + 正文，同一段"""
    p = doc.add_paragraph(style='Normal')
    indent2(p)
    r = p.add_run(label)
    r.bold = True
    p.add_run(text)
    return p

def h1(t): 
    p = doc.add_paragraph(t, style='Heading 1'); return p
def h2(t): 
    p = doc.add_paragraph(t, style='Heading 2'); return p
def h3(t): 
    p = doc.add_paragraph(t, style='Heading 3'); return p

def room(name, duty, size, note):
    h3(name)
    labeled('这是干嘛的：', duty)
    labeled('有多大：', size)
    labeled('要注意：', note)

# ---------- 页脚页码 ----------
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run()
fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
run.font.name = 'Times New Roman'; run.font.size = Pt(10.5)

# ================= 正文 =================
t = doc.add_paragraph('咱家房子图纸大白话讲解', style='Title')
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub = doc.add_paragraph('（给爸妈看的 · 照着施工图一张一张说）', style='Normal')
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
for r in sub.runs:
    r.font.size = Pt(12)
doc.add_paragraph()

# 一、名词
h1('一、先把图纸上的几个词说明白')
body('咱们不用学看图纸，只要先弄明白下面几个词，后面说房间就都听得懂了。')
labeled('标高：', '就是“这个地方有多高”。图纸上拿一层屋里的地面当起点，写作 ±0.000，单位是米。比它高的写正数，比它低的写负数。比如 −0.700 的意思就是比一层地面低 70 公分，3.200 就是比一层地面高 3 米 2。')
labeled('层高：', '从这层地面到上一层地面的高度，简单说就是“这一层有多高”。咱们家一层层高 3 米 2，二层 3 米，三层 3 米 3，都比普通楼房（一般 2 米 8）要高，住着不憋屈。')
labeled('开间和进深：', '开间就是一间屋左右有多宽，进深就是前后有多长。图纸上的数字单位是毫米，1000 毫米就是 1 米。')
labeled('面积：', '就是一间屋地面有多大，单位平方米（㎡）。说“这间屋 10 平方”，大概就是能放开一张大床加两个床头柜那么大。')

# 二、整体
h1('二、咱家房子整体什么样')
body('咱们家是一套联排别墅，东西宽约 6 米，南北长约 11 米，简单记就是“6 米宽、11 米长”。')
body('房子从下到上一共是：最底下半地下一层（车库层，地面标高 −6.000），往上是第一层（±0.000）、第二层（3.200）、第三层（6.200），最顶上是坡屋顶，屋脊最高处标高 12.900。坡屋顶下面还有一层闷顶，能放杂物。')
body('各层地面的标高串起来是这样：车库 −6.000 → 一层 ±0.000 → 二层 3.200 → 三层 6.200 → 露台 7.200 → 屋脊 12.900。')
body('每层的北侧都有一部家用电梯的位置（井道已经留好，约 1 米 3 见方），旁边是楼梯。楼梯是“三跑”的，就是每上一层要转两个弯、分三段走，走起来不累。')
body('各层面积加起来大约 230 平方米：车库层约 65 平方，一层约 55 平方，二层约 50 平方，三层约 61 平方。')

# 三、逐层逐间
h1('三、一层一层、一间一间地说')

h2('（一）车库层（半地下，地面标高 −6.000）')
body('这一层在最底下，从北边走下沉车道、再上坡道进去。顶板就是一层的地面，所以屋里很高，差不多 6 米。')
room('1. 车库',
     '停车用的，在这一层最南边。北墙上有一个大卷帘门（宽 2 米 5、高 2 米 2），车从这里进出，门外接下沉车道和坡道。',
     '约 6 米 × 4 米 5，27 平方米，停一辆车很宽裕，旁边还能放东西。',
     '一是坡道和门口要做好排水，下雨天水不能倒灌；二是提前留好新能源车充电的插座；三是这层特别高，以后可以打高货架，甚至考虑搭个小夹层（要另找专业人看）。')
room('2. 工具间',
     '放工具、杂物、换季东西的屋子，在车库北边。',
     '约 6 米 × 3 米，18 平方米。',
     '灯要亮、插座要够；因为在地下，注意通风，东西都垫高一点放，别直接贴地。')
room('3. 储藏/设备间',
     '最北边一间，放不常用的杂物；以后热水器、新风这类大设备也可以放在这儿。',
     '约 6 米 × 3 米 3，将近 20 平方米。',
     '放设备的话要提前留好水电和检修的地方；同样注意通风、防潮。')
labeled('4. 折返楼梯：', '从车库层上到一层，楼梯要折好几折，一共约 39 级，中间有几个休息平台，走累了可以歇。')

h2('（二）第一层（一层，地面 ±0.000，层高 3.2 米）')
body('这一层是全家的“门面”和活动区：进门是客厅、餐厅，北边是厨房和卫生间。')
room('1. 客厅',
     '全家待得最多的地方，放沙发、电视，在一层西南角，和餐厅是敞开连通的，显得特别大。',
     '约 5 米 7 × 4 米 1，23 平方米。',
     '南面是一整面落地窗（宽 3 米 3），西墙上也有一扇窗，采光很好；要提前想好窗帘怎么做、落地窗里面加不加防护栏杆。')
room('2. 餐厅',
     '放餐桌吃饭的地方，紧挨着客厅和厨房，从厨房端菜几步就到。',
     '约 3 米 8 × 4 米，15 平方米，放一张六人餐桌没问题。',
     '餐边柜的位置和插座要提前留；因为和客厅通着，灯可以和客厅一起设计。')
room('3. 厨房',
     '做饭的地方，在一层最北角，北墙上有窗户，油烟好往外排。橱柜沿着北墙和西墙摆，是个 L 形；南边是推拉门，对着餐厅。',
     '约 2 米 6 × 1 米 3，3.3 平方米，属于“小而紧凑”，一个人转身做饭正好。',
     '地方小，插座一定要留够（冰箱、微波炉、电饭煲等）；地面和墙角要做防水；推拉门比普通门省地方。')
room('4. 卫生间',
     '一层的厕所和洗澡间，在北角、紧挨着厨房。这样两家的水管、热水器共用一面墙，既省事又省钱。',
     '约 2 米 3 × 1 米 3，将近 3 平方米，里面有洗手台、马桶和淋浴。',
     '尽量做干湿分离，洗澡水才不会弄得到处湿；防水和排风一定要做好；因为挨着厨房，热水来得快。')
room('5. 采光井',
     '在客厅南边、一层楼板上开的一个天井（约 2 米 7 × 1 米 3），从这儿能直接看到车库，作用是给底下车库透光、透气。',
     '天井本身不占房间，洞口约 3.5 平方米。',
     '这是个“楼上能看到楼下”的洞口，边上护栏或者玻璃顶一定要做结实，家里有小孩尤其要当心，不能翻越、踩空。')
labeled('6. 入户大门：', '南面是双开大门，进门先有一个凹廊挡风遮雨。注意门外地面是 −0.700，比屋里低 70 公分，所以要上几级台阶才进门，家里老人进出慢一点。')
labeled('7. 家用电梯：', '在一层北侧，门朝南开，以后坐电梯可以直达每一层，搬东西、老人上下楼都方便。')

h2('（三）第二层（二层，地面标高 3.200，层高 3.0 米）')
body('这一层是“休息区”，一共有三间卧室、两个卫生间、一个衣帽间，南边还有阳台。')
room('1. 西南卧',
     '二层最大的卧室，在西南角。南面是落地窗，带一扇推拉门直接通阳台，西墙上也有窗，通风采光都好。可以做老人房或者儿童房。',
     '约 3 米 × 4 米，12.3 平方米，放床、衣柜、书桌都够。',
     '如果给老人住，这间最合适——离卫生间近、朝南暖和；通阳台的推拉门门槛要做平，别绊脚。')
room('2. 东南卧',
     '东南角的卧室，南墙上有窗户。可以做卧室，也可以做书房。',
     '约 2 米 1 × 4 米，8.6 平方米。',
     '房间偏窄，家具顺着墙摆更省地方；书桌靠窗放光线好。')
room('3. 西北卧',
     '西北角的小卧室，北墙上有窗户。',
     '约 2 米 9 × 2 米 4，6.8 平方米。',
     '面积不大，做单人卧室、书房或者儿童房都行；定制家具比买成品更贴合。')
room('4. 西卫',
     '二层靠西的卫生间，西墙上开了窗，是个“明卫”，白天不用开灯、潮气散得快。',
     '约 2 米 × 2 米 4，4.75 平方米，里面有洗手台、马桶、淋浴。',
     '有窗户通风好，不容易发霉；防水层墙面要刷到 1 米 8 高。')
room('5. 东北公卫',
     '东北角的公共卫生间，挨着西北卧，北墙上有窗。',
     '约 2 米 5 × 1 米 3，3.25 平方米。',
     '房间比较窄，洗手台、马桶尺寸要量好再买；同样做好防水和通风。')
room('6. 衣帽间',
     '在西卫和卧室之间的一间小屋，专门做衣柜、放衣服，相当于一个走入式衣柜。',
     '约 1 米 2 × 2 米 7，3.2 平方米。',
     '地方不大，建议三面都做柜子，最能装；里面可以加感应灯，开门就亮。')
room('7. 阳台',
     '在西南卧南边的露天小阳台，晾晒衣服、晒被子、放个小椅子晒太阳。',
     '约 2 米 7 × 0 米 9，2.4 平方米。',
     '地面要做防水和排水坡度，不然下雨积水；晾衣架提前留好位置和电源。')

h2('（四）第三层（三层，地面标高 6.200，层高 3.3 米）')
body('这一层最安静、视野最好，按剖面图的推断，从南到北是大主卧、大卫生间和一个大露台。')
room('1. 主卧',
     '三层南边整一层都是主人房，南面一整面落地窗，西墙上还有窗，又亮堂又安静。',
     '约 5 米 6 × 4 米 5，25 平方米，是全家最大的卧室，床、衣柜、梳妆台、小沙发都放得下。',
     '落地窗多，窗帘和防护要提前考虑；床头开关做成双控，躺下不用起来关灯。')
room('2. 主卫',
     '主卧北边的大卫生间，西墙上有小窗。可以放双台盆、马桶、淋浴，还放得下一个浴缸。',
     '约 5 米 6 × 3 米，将近 17 平方米，比一般人家的卧室还大。',
     '面积大、用水多，防水是重中之重，做完要做 48 小时闭水试验；浴缸位置要提前留好上下水和承重。')
room('3. 露台',
     '最北边的露天大平台，从主卫推门出去就是。可以乘凉、看景、晒被子、种花，也能放桌椅喝茶。',
     '约 5 米 6 × 3 米 3，18.6 平方米。',
     '露天的地方最怕漏水，地面防水和排水一定要做好；四周的围挡（女儿墙）高度要够、要结实；出门有一两步台阶，夜里走要注意。')

h2('（五）坡屋顶和顶上的闷顶')
body('三层顶板（标高约 9.5 米）再往上就是坡屋顶，像人字一样斜上去，最高的屋脊标高 12.900 米。屋檐矮的地方形成一圈闷顶，可以放杂物。')
body('要注意：坡屋顶夏天被太阳直晒会比较热，顶面隔热要做好；上闷顶留一个检修口，拿东西方便。')

# 四、重点事项
h1('四、装修时要重点盯着的几件事')
items = [
 ('防水：', '四个卫生间、厨房、阳台和露台都要做防水，做完必须做 48 小时闭水试验，确认不漏了再贴砖。'),
 ('水电：', '插座、开关、灯位在动工前就要一间一间想清楚，尤其是厨房和电视墙；电梯井道、车库充电桩提前预留。水电完工封起来之前，要现场量数、拍照留底。'),
 ('楼梯：', '车库那段折返梯相对陡一些，老人上下要扶好扶手；所有楼梯扶手要装牢，踏步面别打滑。'),
 ('采光井：', '一层那个直通车库的天井，护栏或玻璃顶必须结实，这是安全大事。'),
 ('门窗：', '落地窗又大又亮，要想好怎么开启、怎么加防护；默认沿用开发商原配的窗户，要换好窗户得另外加一笔钱。'),
 ('尺寸：', '图纸上的房间尺寸是按图推算的，和实际盖好可能差 10 公分左右，买家具前以现场实际量的尺寸为准。'),
 ('预算：', '别墅装修隐蔽工程多，建议在预算之外再留 8% 到 10% 的备用钱；家用电梯、换窗户、大家电和活动家具都不包含在基础预算里，要单独准备。'),
]
for i, (lab, txt) in enumerate(items, 1):
    p = doc.add_paragraph(style='Normal')
    indent2(p)
    r = p.add_run(f'{i}. '); r.bold = True
    r2 = p.add_run(lab); r2.bold = True
    p.add_run(txt)

doc.add_paragraph()
end = doc.add_paragraph('特别说明：三层“主卧—主卫—露台”的布局，以及车库外面的下沉车道和坡道，是按照剖面图推断的；最终以实际房子和正式施工图为准。', style='Normal')
indent2(end)
for r in end.runs:
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

doc.save(OUT)
print('saved:', OUT)
