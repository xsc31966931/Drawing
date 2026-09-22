# -*- coding: utf-8 -*-
"""生成《全屋装修预算清单_原木风.xlsx》"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = r"D:\code\Drawing\全屋装修预算清单_原木风.xlsx"

# 行：楼层, 房间, 科目, 项目, 单位, 工程量, 单价低, 单价高, 备注
ROWS = [
# ================= B1 车库层（-6.000，通高至 ±0） =================
("B1 车库层", "车库", "拆改", "墙面基层清理/界面剂", "㎡", 116, 4, 8, "毛坯混凝土面，无铲腻子"),
("B1 车库层", "车库", "水电", "照明及充电桩线路（人工+辅材）", "项", 1, 1200, 2800, "含新能源车充电位预留"),
("B1 车库层", "车库", "泥瓦", "地面找平+耐磨地坪漆", "㎡", 27, 45, 85, ""),
("B1 车库层", "车库", "泥瓦", "排水沟/集水井（含盖板）", "项", 1, 800, 1800, "配合下沉车道排水"),
("B1 车库层", "车库", "油漆", "墙面/顶面防霉乳胶漆（通高约5.8m）", "㎡", 143, 28, 45, "墙116+顶27"),
("B1 车库层", "车库", "主材安装", "车库卷帘门（含电机安装）", "樘", 1, 2500, 6000, "北侧车辆门 2.5×2.2"),
("B1 车库层", "工具间", "水电", "照明/插座线路（人工+辅材）", "项", 1, 800, 1800, ""),
("B1 车库层", "工具间", "泥瓦", "地面找平+地坪漆", "㎡", 18, 45, 85, ""),
("B1 车库层", "工具间", "油漆", "墙顶防霉乳胶漆（通高）", "㎡", 122, 28, 45, "墙104+顶18"),
("B1 车库层", "工具间", "木工", "简易工作台（现场制作）", "项", 1, 600, 1500, ""),
("B1 车库层", "储藏/设备间", "水电", "照明/插座线路（人工+辅材）", "项", 1, 800, 1800, ""),
("B1 车库层", "储藏/设备间", "泥瓦", "地面找平+地坪漆", "㎡", 19.8, 45, 85, ""),
("B1 车库层", "储藏/设备间", "油漆", "墙顶防霉乳胶漆（通高）", "㎡", 128, 28, 45, "墙108+顶20"),
("B1 车库层", "储藏/设备间", "主材安装", "成品储物货架（含安装）", "组", 2, 600, 1500, ""),
("B1 车库层", "楼梯/公共", "油漆", "楼梯间墙顶乳胶漆", "㎡", 80, 28, 45, "折返梯周边墙面"),
# ================= F1 一层（±0.000，层高3.2m） =================
("F1 一层", "客厅", "拆改", "铲墙皮/界面处理", "㎡", 51, 7, 16, ""),
("F1 一层", "客厅", "水电", "强弱电/插座点位（人工+辅材）", "㎡", 23.2, 80, 120, "电视墙、落地灯位"),
("F1 一层", "客厅", "泥瓦", "地面找平", "㎡", 23.2, 25, 40, ""),
("F1 一层", "客厅", "木工", "石膏板平顶+窗帘盒", "㎡", 23.2, 100, 180, ""),
("F1 一层", "客厅", "油漆", "墙面腻子+乳胶漆", "㎡", 51, 30, 45, ""),
("F1 一层", "客厅", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 24.4, 260, 450, "含5%损耗"),
("F1 一层", "客厅", "主材安装", "窗台石/小五金", "项", 1, 600, 1400, ""),
("F1 一层", "餐厅", "拆改", "铲墙皮/界面处理", "㎡", 41, 7, 16, ""),
("F1 一层", "餐厅", "水电", "强弱电/插座点位（人工+辅材）", "㎡", 15, 80, 120, "餐边柜电位"),
("F1 一层", "餐厅", "泥瓦", "地面找平", "㎡", 15, 25, 40, ""),
("F1 一层", "餐厅", "木工", "石膏板平顶+窗帘盒", "㎡", 15, 100, 180, ""),
("F1 一层", "餐厅", "油漆", "墙面腻子+乳胶漆", "㎡", 41, 30, 45, ""),
("F1 一层", "餐厅", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 15.8, 260, 450, "含5%损耗"),
("F1 一层", "厨房", "拆改", "铲墙皮/界面处理", "㎡", 20.7, 7, 16, ""),
("F1 一层", "厨房", "水电", "给排水+强弱电（人工+辅材）", "㎡", 3.3, 150, 240, "厨电专线"),
("F1 一层", "厨房", "泥瓦", "防水（地面+墙面上返0.3m）", "㎡", 5.6, 70, 110, ""),
("F1 一层", "厨房", "泥瓦", "地面找平", "㎡", 3.3, 25, 40, ""),
("F1 一层", "厨房", "泥瓦", "墙面拉毛+贴墙砖（人工）", "㎡", 11, 45, 65, ""),
("F1 一层", "厨房", "泥瓦", "地面贴砖（人工）", "㎡", 3.3, 40, 60, ""),
("F1 一层", "厨房", "木工", "铝扣板吊顶", "㎡", 3.3, 110, 180, ""),
("F1 一层", "厨房", "主材安装", "厨房墙地砖（材料）", "㎡", 15.7, 80, 160, "含损耗"),
("F1 一层", "厨房", "主材安装", "原木地柜（含石英石台面）", "延米", 3.5, 1200, 2000, "北墙+西墙"),
("F1 一层", "厨房", "主材安装", "原木吊柜", "延米", 2.5, 500, 900, ""),
("F1 一层", "厨房", "主材安装", "水槽龙头/拉篮五金", "套", 1, 1200, 3000, ""),
("F1 一层", "厨房", "主材安装", "烟机灶具", "套", 1, 3500, 8000, "也可归家电采购"),
("F1 一层", "卫生间", "拆改", "铲墙皮/界面处理", "㎡", 20.8, 7, 16, ""),
("F1 一层", "卫生间", "水电", "给排水+电位（人工+辅材）", "㎡", 2.9, 180, 280, "干湿分区电位"),
("F1 一层", "卫生间", "泥瓦", "防水（地面+墙面1.8m）", "㎡", 15.5, 70, 110, "闭水试验48h"),
("F1 一层", "卫生间", "泥瓦", "墙地面找平/拉毛", "㎡", 6, 30, 50, ""),
("F1 一层", "卫生间", "泥瓦", "贴墙砖（人工）", "㎡", 20.8, 45, 65, ""),
("F1 一层", "卫生间", "泥瓦", "贴地砖（人工）", "㎡", 2.9, 40, 60, ""),
("F1 一层", "卫生间", "木工", "铝扣板吊顶", "㎡", 2.9, 110, 180, ""),
("F1 一层", "卫生间", "主材安装", "墙地砖（材料）", "㎡", 26, 90, 170, "含损耗"),
("F1 一层", "卫生间", "主材安装", "原木浴室柜（含台盆）", "套", 1, 900, 1800, ""),
("F1 一层", "卫生间", "主材安装", "坐便器", "套", 1, 1200, 3500, ""),
("F1 一层", "卫生间", "主材安装", "淋浴花洒/淋浴房", "套", 1, 1200, 3500, ""),
("F1 一层", "卫生间", "主材安装", "五金挂件/地漏", "套", 1, 500, 1200, ""),
("F1 一层", "入户/公共", "拆改", "入户区域基层处理", "㎡", 20, 7, 16, ""),
("F1 一层", "入户/公共", "水电", "配电箱/弱电箱改造", "项", 1, 2000, 5000, "含漏保升级"),
("F1 一层", "入户/公共", "泥瓦", "入户地砖铺贴（含找平）", "㎡", 8, 55, 85, ""),
("F1 一层", "入户/公共", "主材安装", "入户地砖（材料）", "㎡", 9, 100, 200, ""),
("F1 一层", "入户/公共", "主材安装", "入户钢木大门（含门套，如更换）", "樘", 1, 2500, 6000, "原门可用则不计"),
# ================= F2 二层（3.200，层高3.0m） =================
("F2 二层", "西北卧", "拆改", "铲墙皮/界面处理", "㎡", 26, 7, 16, ""),
("F2 二层", "西北卧", "水电", "强弱电/点位（人工+辅材）", "㎡", 6.8, 80, 120, ""),
("F2 二层", "西北卧", "泥瓦", "地面找平", "㎡", 6.8, 25, 40, ""),
("F2 二层", "西北卧", "木工", "石膏顶角线/窗帘盒", "m", 10.5, 15, 28, ""),
("F2 二层", "西北卧", "油漆", "墙面腻子+乳胶漆", "㎡", 26, 30, 45, ""),
("F2 二层", "西北卧", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 7.1, 260, 450, "含损耗"),
("F2 二层", "西北卧", "主材安装", "实木复合门（含套锁五金）", "樘", 1, 1300, 2200, ""),
("F2 二层", "西南卧", "拆改", "铲墙皮/界面处理", "㎡", 34, 7, 16, ""),
("F2 二层", "西南卧", "水电", "强弱电/点位（人工+辅材）", "㎡", 12.3, 80, 120, ""),
("F2 二层", "西南卧", "泥瓦", "地面找平", "㎡", 12.3, 25, 40, ""),
("F2 二层", "西南卧", "木工", "石膏顶角线/窗帘盒", "m", 14.2, 15, 28, ""),
("F2 二层", "西南卧", "油漆", "墙面腻子+乳胶漆", "㎡", 34, 30, 45, ""),
("F2 二层", "西南卧", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 12.9, 260, 450, "含损耗"),
("F2 二层", "西南卧", "主材安装", "实木复合门（含套锁五金）", "樘", 1, 1300, 2200, ""),
("F2 二层", "西南卧", "主材安装", "阳台原木框推拉门（含安装）", "㎡", 4, 800, 1500, "1.65×2.4"),
("F2 二层", "东南卧", "拆改", "铲墙皮/界面处理", "㎡", 33, 7, 16, ""),
("F2 二层", "东南卧", "水电", "强弱电/点位（人工+辅材）", "㎡", 8.6, 80, 120, ""),
("F2 二层", "东南卧", "泥瓦", "地面找平", "㎡", 8.6, 25, 40, ""),
("F2 二层", "东南卧", "木工", "石膏顶角线/窗帘盒", "m", 12.4, 15, 28, ""),
("F2 二层", "东南卧", "油漆", "墙面腻子+乳胶漆", "㎡", 33, 30, 45, ""),
("F2 二层", "东南卧", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 9.1, 260, 450, "含损耗"),
("F2 二层", "东南卧", "主材安装", "实木复合门（含套锁五金）", "樘", 1, 1300, 2200, ""),
("F2 二层", "西卫", "拆改", "铲墙皮/界面处理", "㎡", 22, 7, 16, ""),
("F2 二层", "西卫", "水电", "给排水+电位（人工+辅材）", "㎡", 4.75, 180, 280, ""),
("F2 二层", "西卫", "泥瓦", "防水（地面+墙面1.8m）", "㎡", 20.5, 70, 110, ""),
("F2 二层", "西卫", "泥瓦", "找平/拉毛", "㎡", 8, 30, 50, ""),
("F2 二层", "西卫", "泥瓦", "贴墙砖（人工）", "㎡", 22, 45, 65, ""),
("F2 二层", "西卫", "泥瓦", "贴地砖（人工）", "㎡", 4.75, 40, 60, ""),
("F2 二层", "西卫", "木工", "铝扣板吊顶", "㎡", 4.75, 110, 180, ""),
("F2 二层", "西卫", "主材安装", "墙地砖（材料）", "㎡", 29.5, 90, 170, "含损耗"),
("F2 二层", "西卫", "主材安装", "原木浴室柜（含台盆）", "套", 1, 900, 1800, ""),
("F2 二层", "西卫", "主材安装", "坐便器", "套", 1, 1000, 3000, ""),
("F2 二层", "西卫", "主材安装", "淋浴花洒", "套", 1, 1000, 3000, ""),
("F2 二层", "西卫", "主材安装", "五金挂件/地漏", "套", 1, 400, 1000, ""),
("F2 二层", "东北公卫", "拆改", "铲墙皮/界面处理", "㎡", 19, 7, 16, ""),
("F2 二层", "东北公卫", "水电", "给排水+电位（人工+辅材）", "㎡", 3.25, 180, 280, ""),
("F2 二层", "东北公卫", "泥瓦", "防水（地面+墙面1.8m）", "㎡", 17, 70, 110, ""),
("F2 二层", "东北公卫", "泥瓦", "找平/拉毛", "㎡", 6, 30, 50, ""),
("F2 二层", "东北公卫", "泥瓦", "贴墙砖（人工）", "㎡", 19, 45, 65, ""),
("F2 二层", "东北公卫", "泥瓦", "贴地砖（人工）", "㎡", 3.25, 40, 60, ""),
("F2 二层", "东北公卫", "木工", "铝扣板吊顶", "㎡", 3.25, 110, 180, ""),
("F2 二层", "东北公卫", "主材安装", "墙地砖（材料）", "㎡", 24, 90, 170, "含损耗"),
("F2 二层", "东北公卫", "主材安装", "原木浴室柜（含台盆）", "套", 1, 900, 1800, ""),
("F2 二层", "东北公卫", "主材安装", "坐便器", "套", 1, 1000, 3000, ""),
("F2 二层", "东北公卫", "主材安装", "淋浴花洒", "套", 1, 1000, 3000, ""),
("F2 二层", "东北公卫", "主材安装", "五金挂件/地漏", "套", 1, 400, 1000, ""),
("F2 二层", "衣帽间", "拆改", "铲墙皮/界面处理", "㎡", 21.7, 7, 16, ""),
("F2 二层", "衣帽间", "水电", "照明/点位（人工+辅材）", "㎡", 3.24, 80, 120, "感应灯"),
("F2 二层", "衣帽间", "泥瓦", "地面找平", "㎡", 3.24, 25, 40, ""),
("F2 二层", "衣帽间", "油漆", "墙面腻子+乳胶漆", "㎡", 21.7, 30, 45, ""),
("F2 二层", "衣帽间", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 3.4, 260, 450, "含损耗"),
("F2 二层", "衣帽间", "主材安装", "实木复合门（含套锁）", "樘", 1, 1100, 1900, ""),
("F2 二层", "阳台", "泥瓦", "防水+排水坡度", "㎡", 2.43, 70, 110, ""),
("F2 二层", "阳台", "泥瓦", "贴地砖（人工）", "㎡", 2.43, 50, 75, ""),
("F2 二层", "阳台", "主材安装", "阳台仿古砖（材料）", "㎡", 2.7, 90, 160, ""),
("F2 二层", "阳台", "主材安装", "电动晾衣架", "套", 1, 600, 1500, ""),
("F2 二层", "楼梯/公共", "油漆", "楼梯间墙面乳胶漆", "㎡", 60, 30, 45, ""),
# ================= F3 三层（6.200，层高3.3m，布局按剖面推断） =================
("F3 三层", "主卧", "拆改", "铲墙皮/界面处理", "㎡", 54, 7, 16, ""),
("F3 三层", "主卧", "水电", "强弱电/点位（人工+辅材）", "㎡", 25.4, 80, 120, "床头双控"),
("F3 三层", "主卧", "泥瓦", "地面找平", "㎡", 25.4, 25, 40, ""),
("F3 三层", "主卧", "木工", "石膏板吊顶+窗帘盒", "㎡", 25.4, 100, 160, ""),
("F3 三层", "主卧", "油漆", "墙面腻子+乳胶漆", "㎡", 54, 30, 45, ""),
("F3 三层", "主卧", "主材安装", "多层实木地板（原木色，含安装）", "㎡", 26.7, 260, 450, "含损耗"),
("F3 三层", "主卧", "主材安装", "实木复合门（含套锁五金）", "樘", 1, 1300, 2200, ""),
("F3 三层", "主卫", "拆改", "铲墙皮/界面处理", "㎡", 52, 7, 16, ""),
("F3 三层", "主卫", "水电", "给排水+电位（人工+辅材）", "㎡", 16.9, 160, 250, "双台盆+浴缸"),
("F3 三层", "主卫", "泥瓦", "防水（地面+墙面1.8m）", "㎡", 48, 70, 110, "大面防水重点验收"),
("F3 三层", "主卫", "泥瓦", "找平/拉毛", "㎡", 20, 30, 50, ""),
("F3 三层", "主卫", "泥瓦", "贴墙砖（人工）", "㎡", 52, 45, 65, ""),
("F3 三层", "主卫", "泥瓦", "贴地砖（人工）", "㎡", 16.9, 40, 60, ""),
("F3 三层", "主卫", "木工", "防水石膏板/铝扣板吊顶", "㎡", 16.9, 120, 200, ""),
("F3 三层", "主卫", "主材安装", "墙地砖（材料）", "㎡", 75, 90, 170, "含损耗"),
("F3 三层", "主卫", "主材安装", "原木双台盆浴室柜", "套", 1, 1800, 4000, ""),
("F3 三层", "主卫", "主材安装", "坐便器", "套", 1, 1500, 4000, ""),
("F3 三层", "主卫", "主材安装", "独立浴缸+龙头", "套", 1, 4000, 12000, ""),
("F3 三层", "主卫", "主材安装", "淋浴花洒/玻璃隔断", "套", 1, 1500, 4000, ""),
("F3 三层", "主卫", "主材安装", "五金挂件/地漏", "套", 1, 800, 1800, ""),
("F3 三层", "主卫", "主材安装", "实木玻璃露台门（含安装）", "㎡", 2.1, 900, 1600, "0.8×2.1，门外有台阶"),
("F3 三层", "露台", "泥瓦", "防水+排水（地面7.200）", "㎡", 18.6, 80, 120, "外露面必须做耐候防水"),
("F3 三层", "露台", "泥瓦", "铺地砖（人工）", "㎡", 18.6, 55, 80, ""),
("F3 三层", "露台", "主材安装", "户外地砖/防腐木地板（材料）", "㎡", 19.5, 100, 220, ""),
("F3 三层", "露台", "油漆", "女儿墙内侧外墙漆/真石漆", "㎡", 25, 45, 80, "女儿墙7.6/8.0/8.6阶梯"),
("F3 三层", "露台", "主材安装", "户外家具/花池", "项", 1, 1500, 4000, "可后期添置"),
# ================= 全屋项目 =================
("全屋", "全屋", "主材安装", "开关插座面板+安装（约140个）", "个", 140, 25, 55, "原木/哑光面板"),
("全屋", "全屋", "主材安装", "全屋灯具（原木风，含安装）", "项", 1, 12000, 30000, "含客餐厅主灯、卧室灯、筒灯"),
("全屋", "全屋", "主材安装", "全屋窗帘（13窗，含轨道纱帘）", "项", 1, 8000, 20000, ""),
("全屋", "全屋", "主材安装", "楼梯原木踏步板（含安装）", "踏", 102, 250, 500, "车库折返梯+三跑梯合计"),
("全屋", "全屋", "主材安装", "楼梯木扶手（含安装）", "m", 37, 180, 380, ""),
("全屋", "全屋", "其他", "设计费", "㎡", 232, 80, 150, ""),
("全屋", "全屋", "其他", "垃圾清运（全屋）", "项", 1, 3000, 6000, ""),
("全屋", "全屋", "其他", "开荒保洁", "㎡", 232, 8, 15, ""),
("全屋", "全屋", "其他", "成品保护/材料搬运", "项", 1, 3000, 7000, ""),
]

FLOORS = ["B1 车库层", "F1 一层", "F2 二层", "F3 三层", "全屋"]
CATS = ["拆改", "水电", "泥瓦", "木工", "油漆", "主材安装", "其他"]

# ---------- 样式 ----------
C_DARK = "6B4F2E"   # 深棕（表头）
C_FLOOR = "E8DCC8"  # 楼层分隔行浅米
C_SUB = "F4EEE2"    # 小计行
thin = Side(style="thin", color="C9BCA3")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_TITLE = Font(name="微软雅黑", size=15, bold=True, color="3A2C17")
F_HEAD = Font(name="微软雅黑", size=10.5, bold=True, color="FFFFFF")
F_FLOOR = Font(name="微软雅黑", size=11, bold=True, color="5A4326")
F_BODY = Font(name="微软雅黑", size=10, color="333333")
F_SUB = Font(name="微软雅黑", size=10.5, bold=True, color="5A4326")
F_NOTE = Font(name="微软雅黑", size=9.5, color="808080")
A_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
A_R = Alignment(horizontal="right", vertical="center")

wb = Workbook()

# ================= Sheet1 预算总表 =================
ws = wb.active
ws.title = "预算总表"
HEADERS = ["楼层", "房间", "科目", "项目说明", "单位", "工程量", "单价低(元)", "单价高(元)", "合价低(元)", "合价高(元)", "备注"]
NCOL = len(HEADERS)

r = 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOL)
c = ws.cell(r, 1, "全屋装修预算清单（原木风 · 保定中档口径）")
c.font = F_TITLE; c.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[r].height = 30
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOL)
c = ws.cell(r, 1, "编制日期：2026-09-22　建筑面积约232㎡（套内）　价格为参考区间，最终以实际报价为准")
c.font = F_NOTE; c.alignment = Alignment(horizontal="center", vertical="center")
r += 1

HEAD_ROW = r
for j, h in enumerate(HEADERS, 1):
    cell = ws.cell(r, j, h)
    cell.font = F_HEAD; cell.alignment = A_C; cell.border = BORDER
    cell.fill = PatternFill("solid", fgColor=C_DARK)
ws.row_dimensions[r].height = 26
r += 1

DATA_START = r
current_floor = None
floor_ranges = {}   # floor -> [start, end] data rows
for row in ROWS:
    floor = row[0]
    if floor != current_floor:
        # 楼层分隔行
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOL)
        cell = ws.cell(r, 1, "▼ " + floor)
        cell.font = F_FLOOR; cell.alignment = A_L
        cell.fill = PatternFill("solid", fgColor=C_FLOOR)
        for j in range(1, NCOL + 1):
            ws.cell(r, j).border = BORDER
        r += 1
        current_floor = floor
        floor_ranges.setdefault(floor, [r, None])
    vals = list(row)
    for j, v in enumerate(vals, 1):
        cell = ws.cell(r, j, v)
        cell.font = F_BODY; cell.border = BORDER
        if j in (4, 11):
            cell.alignment = A_L
        elif j in (6, 7, 8, 9, 10):
            cell.alignment = A_R
        else:
            cell.alignment = A_C
    # 合价公式
    ws.cell(r, 9).value = f"=ROUND(F{r}*G{r},0)"
    ws.cell(r, 10).value = f"=ROUND(F{r}*H{r},0)"
    for j in (7, 8, 9, 10):
        ws.cell(r, j).number_format = "#,##0"
    floor_ranges[floor][1] = r
    r += 1
DATA_END = r - 1

# 合计行
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
cell = ws.cell(r, 1, "全屋合计（不含暂列金；不含项见“说明”表）")
cell.font = F_SUB; cell.alignment = Alignment(horizontal="right", vertical="center")
cell.fill = PatternFill("solid", fgColor=C_SUB)
for j in range(1, 9):
    ws.cell(r, j).border = BORDER
    ws.cell(r, j).fill = PatternFill("solid", fgColor=C_SUB)
ws.cell(r, 9).value = f"=SUM(I{DATA_START}:I{DATA_END})"
ws.cell(r, 10).value = f"=SUM(J{DATA_START}:J{DATA_END})"
for j in (9, 10):
    cc = ws.cell(r, j); cc.font = F_SUB; cc.border = BORDER
    cc.number_format = "#,##0"; cc.alignment = A_R
    cc.fill = PatternFill("solid", fgColor=C_SUB)
ws.cell(r, 11).border = BORDER
ws.cell(r, 11).fill = PatternFill("solid", fgColor=C_SUB)
TOTAL_ROW = r
r += 1
# 暂列金行
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
cell = ws.cell(r, 1, "不可预见费（暂列金，按合计 8%–10%）")
cell.font = F_BODY; cell.alignment = Alignment(horizontal="right", vertical="center")
for j in range(1, 9):
    ws.cell(r, j).border = BORDER
ws.cell(r, 9).value = f"=ROUND(I{TOTAL_ROW}*0.08,0)"
ws.cell(r, 10).value = f"=ROUND(J{TOTAL_ROW}*0.1,0)"
for j in (9, 10):
    cc = ws.cell(r, j); cc.font = F_BODY; cc.border = BORDER
    cc.number_format = "#,##0"; cc.alignment = A_R
ws.cell(r, 11).border = BORDER
RESERVE_ROW = r
r += 1
# 含暂列金总计
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
cell = ws.cell(r, 1, "预计总投入区间（含暂列金）")
cell.font = F_SUB; cell.alignment = Alignment(horizontal="right", vertical="center")
for j in range(1, 9):
    ws.cell(r, j).border = BORDER
    ws.cell(r, j).fill = PatternFill("solid", fgColor=C_FLOOR)
ws.cell(r, 9).value = f"=I{TOTAL_ROW}+I{RESERVE_ROW}"
ws.cell(r, 10).value = f"=J{TOTAL_ROW}+J{RESERVE_ROW}"
for j in (9, 10):
    cc = ws.cell(r, j); cc.font = F_SUB; cc.border = BORDER
    cc.number_format = "#,##0"; cc.alignment = A_R
    cc.fill = PatternFill("solid", fgColor=C_FLOOR)
ws.cell(r, 11).border = BORDER
ws.cell(r, 11).fill = PatternFill("solid", fgColor=C_FLOOR)
GRAND_ROW = r

widths = [11, 11, 9, 30, 6, 8, 10, 10, 11, 11, 22]
for j, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = f"A{HEAD_ROW + 1}"
ws.sheet_view.showGridLines = False

# ================= Sheet2 汇总 =================
ws2 = wb.create_sheet("汇总")
ws2.sheet_view.showGridLines = False
r = 1
ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws2.cell(r, 1, "预算汇总（数据自动取自“预算总表”）")
c.font = F_TITLE; c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[r].height = 28
r += 2

def block_header(r, titles):
    for j, t in enumerate(titles, 1):
        cell = ws2.cell(r, j, t)
        cell.font = F_HEAD; cell.alignment = A_C; cell.border = BORDER
        cell.fill = PatternFill("solid", fgColor=C_DARK)
    return r + 1

T = "预算总表"
# --- 按楼层 ---
r = block_header(r, ["楼层", "合价低(元)", "合价高(元)", "占比(按中值)"])
floor_start = r
for f in FLOORS:
    ws2.cell(r, 1, f).font = F_BODY
    ws2.cell(r, 1).alignment = A_L; ws2.cell(r, 1).border = BORDER
    ws2.cell(r, 2).value = f'=SUMIF(\'{T}\'!A:A,A{r},\'{T}\'!I:I)'
    ws2.cell(r, 3).value = f'=SUMIF(\'{T}\'!A:A,A{r},\'{T}\'!J:J)'
    for j in (2, 3):
        cc = ws2.cell(r, j); cc.font = F_BODY; cc.border = BORDER
        cc.number_format = "#,##0"; cc.alignment = A_R
    ws2.cell(r, 4).border = BORDER
    r += 1
floor_end = r - 1
# 合计
ws2.cell(r, 1, "合计").font = F_SUB
ws2.cell(r, 1).fill = PatternFill("solid", fgColor=C_SUB)
ws2.cell(r, 1).border = BORDER; ws2.cell(r, 1).alignment = A_L
ws2.cell(r, 2).value = f"=SUM(B{floor_start}:B{floor_end})"
ws2.cell(r, 3).value = f"=SUM(C{floor_start}:C{floor_end})"
for j in (2, 3):
    cc = ws2.cell(r, j); cc.font = F_SUB; cc.border = BORDER
    cc.number_format = "#,##0"; cc.alignment = A_R
    cc.fill = PatternFill("solid", fgColor=C_SUB)
ws2.cell(r, 4).border = BORDER; ws2.cell(r, 4).fill = PatternFill("solid", fgColor=C_SUB)
floor_total_row = r
for rr in range(floor_start, floor_end + 1):
    ws2.cell(rr, 4).value = f"=IFERROR((B{rr}+C{rr})/(B${floor_total_row}+C${floor_total_row}),0)"
    ws2.cell(rr, 4).number_format = "0.0%"
    ws2.cell(rr, 4).font = F_BODY; ws2.cell(rr, 4).alignment = A_R
r += 2

# --- 按科目 ---
r = block_header(r, ["科目", "合价低(元)", "合价高(元)", "占比(按中值)"])
cat_start = r
for ct in CATS:
    ws2.cell(r, 1, ct).font = F_BODY
    ws2.cell(r, 1).alignment = A_L; ws2.cell(r, 1).border = BORDER
    ws2.cell(r, 2).value = f'=SUMIF(\'{T}\'!C:C,A{r},\'{T}\'!I:I)'
    ws2.cell(r, 3).value = f'=SUMIF(\'{T}\'!C:C,A{r},\'{T}\'!J:J)'
    for j in (2, 3):
        cc = ws2.cell(r, j); cc.font = F_BODY; cc.border = BORDER
        cc.number_format = "#,##0"; cc.alignment = A_R
    ws2.cell(r, 4).border = BORDER
    r += 1
cat_end = r - 1
ws2.cell(r, 1, "合计").font = F_SUB
ws2.cell(r, 1).fill = PatternFill("solid", fgColor=C_SUB)
ws2.cell(r, 1).border = BORDER; ws2.cell(r, 1).alignment = A_L
ws2.cell(r, 2).value = f"=SUM(B{cat_start}:B{cat_end})"
ws2.cell(r, 3).value = f"=SUM(C{cat_start}:C{cat_end})"
for j in (2, 3):
    cc = ws2.cell(r, j); cc.font = F_SUB; cc.border = BORDER
    cc.number_format = "#,##0"; cc.alignment = A_R
    cc.fill = PatternFill("solid", fgColor=C_SUB)
ws2.cell(r, 4).border = BORDER; ws2.cell(r, 4).fill = PatternFill("solid", fgColor=C_SUB)
cat_total_row = r
for rr in range(cat_start, cat_end + 1):
    ws2.cell(rr, 4).value = f"=IFERROR((B{rr}+C{rr})/(B${cat_total_row}+C${cat_total_row}),0)"
    ws2.cell(rr, 4).number_format = "0.0%"
    ws2.cell(rr, 4).font = F_BODY; ws2.cell(rr, 4).alignment = A_R
r += 2

# --- 总价区 ---
r = block_header(r, ["项目", "金额低(元)", "金额高(元)", "折合每㎡(元)"])
items = [
    ("基装合计（拆改/水电/泥瓦/木工/油漆/其他）",
     f'=SUM(B{cat_start}:B{cat_start+4})+B{cat_start+6}',
     f'=SUM(C{cat_start}:C{cat_start+4})+C{cat_start+6}'),
    ("主材合计（主材安装）",
     f'=B{cat_start+5}', f'=C{cat_start+5}'),
    ("工程合计", f"='{T}'!I{TOTAL_ROW}", f"='{T}'!J{TOTAL_ROW}"),
    ("暂列金（8%–10%）", f"='{T}'!I{RESERVE_ROW}", f"='{T}'!J{RESERVE_ROW}"),
    ("预计总投入区间", f"='{T}'!I{GRAND_ROW}", f"='{T}'!J{GRAND_ROW}"),
]
sum_start = r
for name, lo, hi in items:
    ws2.cell(r, 1, name).font = F_BODY if "总投入" not in name else F_SUB
    ws2.cell(r, 1).alignment = A_L; ws2.cell(r, 1).border = BORDER
    ws2.cell(r, 2).value = lo; ws2.cell(r, 3).value = hi
    for j in (2, 3):
        cc = ws2.cell(r, j); cc.border = BORDER; cc.number_format = "#,##0"; cc.alignment = A_R
        cc.font = F_SUB if "总投入" in name else F_BODY
    ws2.cell(r, 4).value = f"=ROUND((B{r}+C{r})/2/232,0)"
    cc = ws2.cell(r, 4); cc.border = BORDER; cc.number_format = "#,##0"; cc.alignment = A_R
    cc.font = F_BODY
    if "总投入" in name:
        for j in range(1, 5):
            ws2.cell(r, j).fill = PatternFill("solid", fgColor=C_FLOOR)
    r += 1

for j, w in enumerate([34, 14, 14, 15], 1):
    ws2.column_dimensions[get_column_letter(j)].width = w

# ================= Sheet3 说明 =================
ws3 = wb.create_sheet("说明")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 100
NOTES = [
("T", "预算编制说明"),
("B", "一、编制依据"),
("N", "1. 依据设计师施工图：别墅一层平面图、二层平面图、1-1剖面图（图签 L29# 1-1剖面图 1:200）。"),
("N", "2. 风格口径：原木风（多层实木地板、原木/实木复合橱柜与定制柜、实木复合门、木饰面），中档定位。"),
("N", "3. 价格口径：2026年9月河北保定市场参考区间，人工参考土巴兔保定价格指南及2026年工程人公开报价，"),
("N", "   主材参考土巴兔2026地板价格指南、淘宝/1688在售材料价；仅作预算参考，正式签约前请至少取得3家报价。"),
("B", "二、工程量口径"),
("N", "1. 原始平面图未标注数字尺寸链，房间尺寸按剖面图轴网（4500+3000+3300=10800mm）反推，精度约±100mm。"),
("N", "2. 面积为房间净面积；地板、瓷砖工程量已含约5%损耗；墙面面积按周长×层高并扣除门窗洞口估算。"),
("N", "3. B1车库层（-6.000）为通高空间，顶板即一层±0.000地面，墙顶面按通高约5.8m计；若后期加设夹层，费用另计。"),
("N", "4. F3三层布局（主卧/主卫/露台）系按1-1剖面图推断；露台地面标高7.200，女儿墙呈7.6/8.0/8.6阶梯状。"),
("N", "5. 车库北侧下沉通道及坡道为示意做法（图纸未画车行路径），坡道与挡墙土建费用未包含在本预算内。"),
("B", "三、不含项目（需另行采购/报价）"),
("N", "1. 家用电梯设备采购与安装（井道已在北侧预留，约1.27×1.27m，设备价格以厂家报价为准）。"),
("N", "2. 外门窗更换（默认沿用开发商原配窗；如需换系统窗另行计价）。"),
("N", "3. 外墙土建改造、下沉庭院/坡道挡墙等土建工程。"),
("N", "4. 大家电：冰箱、电视、洗衣机、空调/中央空调、新风、净水、热水器等（烟机灶具已列入，也可转家电采购）。"),
("N", "5. 活动家具（沙发、餐桌椅、床、床垫）、床品布艺、装饰摆件等软装。"),
("N", "6. 采暖系统（地暖/暖气片）及燃气相关工程。"),
("B", "四、使用建议"),
("N", "1. 建议另留工程合计8%–10%不可预见费（表中已单列），老房/别墅隐蔽工程变数较多。"),
("N", "2. 水电按实际米数/点位结算，封槽前现场计量并拍照留档；防水做完做48小时闭水试验。"),
("N", "3. 付款按节点（开工/水电/泥木/油漆/竣工）分期，尾款建议留5%左右、验收合格后结清。"),
("N", "4. 原木风注意木材含水率与环保等级（建议ENF/E0级），保定冬季干燥供暖，实木类需预留伸缩缝。"),
]
r = 1
for kind, text in NOTES:
    cell = ws3.cell(r, 1, text)
    if kind == "T":
        cell.font = Font(name="微软雅黑", size=14, bold=True, color="3A2C17")
        ws3.row_dimensions[r].height = 26
    elif kind == "B":
        cell.font = Font(name="微软雅黑", size=11.5, bold=True, color="5A4326")
        ws3.row_dimensions[r].height = 22
    else:
        cell.font = Font(name="微软雅黑", size=10.5, color="333333")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    r += 1

wb.save(OUT)
print("saved:", OUT)
