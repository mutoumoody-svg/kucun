"""商品名称关键词分类规则，被各个 clean_*.py 整理脚本共用。

两套独立判断：
- classify_material(name): 商品 / 包装物料 —— 是否是可销售实物还是辅料包材
- classify_brand(name): 品牌 —— STTOKE / 慕咖 / 食品 / 周边商品 / 未分类

两者正交：同一品牌下既有商品也有包装物料（比如"慕咖杯密封胶条套装"是
慕咖品牌的包装物料，"STTOKE奢华黑16OZ标签"是STTOKE品牌的包装物料）。

品牌判断优先级：先看是否命中"STTOKE"关键词。"慕咖STTOKE不锈钢...杯"这种命名里
会同时出现"慕咖"和"STTOKE"，但这类杯型本质是STTOKE系列（夜光海洋限定款、分享壶、
硅胶勺等），按品牌应归到STTOKE，所以STTOKE优先于慕咖判断。

extract_group_keyword(name): 从名字里提取规格关键词（目前识别容量，如12OZ/16OZ/20OZ），
同一规格的产品（不同颜色/款式）排在一起，方便对着看。没有识别到规格的归入"其他规格"。
"""

import re

SIZE_PATTERN = re.compile(r"(\d+)\s*OZ", re.IGNORECASE)

PACKAGING_KEYWORDS = [
    "标签", "贴纸", "胶条", "磨砂袋", "展示架", "腰圈", "腰封", "文件物品", "物流卡",
    "内衬", "外盒", "纸箱", "礼袋", "卡片", "折页", "宣传册", "合格证", "声明",
    "警示贴", "快递盒", "白皮箱", "气泡信封袋", "牛皮纸袋", "内托", "PDQ",
    "绒布袋", "挂耳咖啡盒白卡纸", "单只外盒", "新册子", "条形码",
    "外箱", "帆布袋", "opp袋", "不干胶贴", "挂耳*10盒装",
    "礼盒", "木勺", "红包",
]

# 注意：STTOKE关键词在品牌判断里优先级最高，详见 classify_brand
STTOKE_KEYWORDS = ["STTOKE", "盈旋杯", "联名礼盒套装", "符合性声明", "茶水冲泡器", "量勺夹", "茶漏"]
# "挂耳咖啡1片"单片装挂耳咖啡是慕咖自己的咖啡线（MoodyCoffee），"好有水瓶"是慕咖的
# 水瓶产品线，这两类名字里没有"慕咖"/"Mood"字样，但归类应该算慕咖品牌
MOODY_KEYWORDS = ["慕咖", "Mood", "MoodyCoffee", "MoodCone", "月亮杯", "挂耳咖啡1片", "好有水瓶"]
FOOD_KEYWORDS = ["巴恩天然", "蜂蜜", "苹果醋", "挂耳咖啡", "康蜜乐"]
PERIPHERAL_KEYWORDS = ["Alpaka", "Hario", "Delter", "杯刷", "颜色随机玻璃杯", "Baby Towel"]

BRAND_ORDER = ["STTOKE", "慕咖", "食品", "周边商品", "未分类"]

# 库存状态行底色（对应 turnover_lookup.py 给的"库存状态"列）：
# 整体走蓝色系（颜色深浅表示快慢），只有"积压"单独用粉红色突出标出
STATUS_COLORS = {
    "快速": "DDEBF7",
    "正常": "BDD7EE",
    "偏慢": "9DC3E6",
    "未知": "C9D6E8",
    "积压": "FFC7CE",
}

# 已知改版后货号变化（旧批次记录用的SKU -> 当前库存里的新SKU）
SKU_ALIAS = {"BNMGO550-250": "NEWBNMGO550-250"}

# 产品名称修正表（SKU -> 正确名称）：仅用于ERP文件里名称本身就错的情况。
# 大多数名称会由 name_registry.json 从上传的ERP文件里自动维护，不要在这里重复。
NAME_OVERRIDE = {
    # 盈旋杯单只7OZ
    "S7SCUPHSGR": "STTOKE盈旋杯【单只】莫吉托薄荷7OZ",
    "S7SCUPHSPK": "STTOKE盈旋杯【单只】覆盆子玫瑰7OZ",
    "S7SCUPHSBM": "STTOKE盈旋杯【单只】玛格丽特蓝7OZ",
    "S7SCUPHSYL": "STTOKE盈旋杯【单只】含羞草柔黄7OZ",
    "S7SCUPHSBL": "STTOKE盈旋杯【单只】朗姆经典黑7OZ",
    "STTOKECG7":  "STTOKE盈旋杯 【莫尼耶香槟7OZ】",
    "STTEWHI12":  "STTOKE不锈钢咖啡随行杯【时尚白12OZ】",
    "STTESwirl Cup": "STTOKE家用杯（SET）",
    # 夜光海洋系列
    "STTOKELOSS12": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【鲨鱼 12OZ密封旋盖】",
    "STTOKELOSO12": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【章鱼 12OZ密封旋盖】",
    "STTOKELOSS16": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【鲨鱼 16OZ密封旋盖】",
    "STTOKELOSO16": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【章鱼 16OZ密封旋盖】",
    "STTOKELOSS20": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【鲨鱼 20OZ吸管杯】",
    "STTOKELOSO20": "慕咖STTOKE不锈钢随行杯夜光海洋系列限定款【章鱼 20OZ吸管杯】",
    # 咖啡随行杯20OZ
    "STTOKELESBP20": "慕咖STTOKE不锈钢咖啡随行杯【粉玫酿 20OZ 吸管杯】",
    "STTOKELESIG20": "慕咖STTOKE不锈钢咖啡随行杯【绿橄榄 20OZ 吸管杯】",
    # 分享壶20OZ
    "STTOKESJHCB": "慕咖STTOKE不锈钢分享壶 20OZ【晨雾白】",
    "STTOKESJHEG": "慕咖STTOKE不锈钢分享壶 20OZ【暮色灰】",
    # 吸管套
    "STTOKESTIG": "慕咖STTOKE 20oz吸管套绿橄榄",
    "STTOKESTBP": "慕咖STTOKE 20oz吸管套粉玫酿",
}

# ============================================================
# 各页面SKU排列顺序（来源：260630库存整理产品顺序排列标准.xlsx）
# 新增产品不在列表里的，自动排在对应页面末尾。
# ============================================================

SKU_ORDER_GOODS_STTOKE = [
    "STTOKEMM7", "STTOKERR7", "STTOKELB7", "STTOKEMY7",
    "S7SCUPHSGR", "S7SCUPHSYL", "S7SCUPHSBM", "S7SCUPHSPK",
    "STTOKESX76",
    "STTOKELCSAW08", "STTOKELCSMT08", "STTOKELCSLB08", "STTOKELCSBB08", "STTOKELCSBP08",
    "STTOKEBLA12", "STTOKEWHI12", "STTOKEGRE12", "STTOKEBLU12", "STTOKEHG12", "STTOKEMT12",
    "STTOKEUPL12", "STTOKEDS12", "STTOKEWG12", "STTOKEESBB12", "STTOKEESDY12",
    "STTOKELESBP12", "STTOKELESIG12", "STTOKEFSKB12", "STTOKEFSP12", "STTOKEFLD12",
    "STTOKEDC12", "STTOKEMM12", "STTOKEMSG12", "STTOKELFRG12", "STTOKELFGB12",
    "STTEWHALE12", "STTOKEOSD12", "STTOKEOSMRS12", "STTOKELOSS12", "STTOKELOSO12",
    "STTEBLA16", "STTEWHI16", "STTOKEGRE16", "STTOKEBLU16", "STTOKEHG16", "STTOKEMT16",
    "STTOKEUPL16", "STTOKEDS16", "STTOKEWG16", "STTOKEESBB16", "STTOKEESDY16",
    "STTOKELESBP16", "STTOKELESIG16", "STTOKEFSKB16", "STTOKEFSP16", "STTOKEFLD16",
    "STTOKEDC16", "STTOKEMM16", "STTOKEMSG16", "STTOKELFRG16", "STTOKELFGB16",
    "STTEWHALE16", "STTOKEOSD16", "STTOKEOSMRS16", "STTOKELOSS16", "STTOKELOSO16",
    "STTOKEPSBLU16", "STTOKEMB16",
    "STTOKEHG20", "STTOKEDS20", "STTOKEESBB20", "STTEBLA20", "STTEWHI20",
    "STTOKELESIG20", "STTOKELESBP20", "STTOKELOSS20", "STTOKELOSO20",
    "STTOKETAKMT12", "STTOKETAKMT16",
    "STTOKEMOOMINW16", "STTOKEMOOMINBS16", "STTOKEMOOMINS20",
    "STTOKEMOOMINKSA", "STTOKEMOOMINKM", "STTOKEMOOMINKS",
    "STTOKESJHCB", "STTOKESJHEG",
    "STTOKEYSG", "STTOKEHEBB", "STTOKEHEMT", "STTOKEHEBP",
    "STTOKETEAINF",
    "STTEBTTBSX", "STTEBTLISX", "STTEBTGTSX", "STTOKEBLSX",
    "STTOKESDS", "STTOKESSG", "STTOKESBB", "STTOKESLB", "STTOKESMT",
    "STTOKESTMT", "STTOKESTDS", "STTOKESTSG", "STTOKESTIG", "STTOKESTBP", "STTOKESTLB", "STTOKESTBB",
    "NEWSTTOKEBG2", "NEWSTTOKEBG2-BB", "NEWSTTOKEBG2-DY", "NEWSTTOKEBG2-SR", "STTOKEBG2",
    "STTOKECBSL20",
    "STTEBTTBSXNB", "STTEBTGTSXNB", "STTEBTLISXNB",
    "STTOKEGJS", "STTOKEBG3", "STDZGFBOX", "STTOKEZJW", "STSZ", "STTOKEZJB", "HARIO",
]

SKU_ORDER_GOODS_MOODY = [
    "MC-FOREST*1", "MC-HONEYPROCESS*1", "MCZG-NATURAL*1", "MCTZ-NATURAL*1",
    "MKSYBB02", "MKSYBWO1", "MKSYBO04", "MKSYBGO3", "MKSYBSO6", "MKSYBPO5",
    "MSYBDZ001", "MKSYBS-BLACK", "MKSYBS-WHITE",
    "MCITYG", "MCITYGLG", "MCITYWLG",
    "MCBOTBLWHI", "BOTBLWHI", "MCBOTBLAM",
    "MCBWSPTEAM", "MCBWSPBF", "MCBOTHDRUN", "MCBOTHDWSH",
    "MVCPKL", "MVCPKSLG", "MVCPKLLG", "MVCYESLG", "MVCYELLG",
    "MVCGYS", "MVCGYSLG", "MVCGNLLG", "MVCGNSLG",
    "MCBLU12", "MCGRE12",
    "MKWQG", "MBCMH12", "MBWY12",
]

SKU_ORDER_GOODS_FOOD = [
    "BNPRHY", "BNMGO30", "BNMGO100", "BNMGO100-500", "BNMGO300-250",
    "BNMGO850-500", "BNMGO1050-500",
    "NEWBNMGO100", "NEWNBNMGO300-250", "NEWBNMGO550-250", "NEWBNMGO850-250", "NEWBNMGO1050-250",
    "BN300LH", "barnescv", "barneshonyyv", "CPHCANDY", "CKPH340-1",
]

SKU_ORDER_GOODS_PERIPHERAL = [
    "LSBS", "MFFBT", "STALPAKA", "STDPRESS", "STHARIO", "MOODYBLBSJ",
]

SKU_ORDER_PKG_STTOKE = [
    "SETSwirlBox", "STTEBTBOX", "KJJZX-6", "KJJZX-7", "STTOKECD", "STTOKEXCC",
    "STTOKERBH01", "STNCMSD", "STTOKE12ozSM", "STTOKE16ozSM",
    "STNCGP1", "STNCGP5", "STNCGP3", "STNCGP2", "STNCGP4", "STTOKEFBD",
    "STTOKELCSBP08BQ", "STTOKELCSLB08BQ", "STTOKELCSBB08BQ", "STTOKELCSAW08BQ", "STTOKELCSMT08BQ",
    "STTOKEBQDC12", "STTOKEUPL12BQ", "STBQFLD12", "STBQFSP12", "STTOKEMT12BQ",
    "STTOKEBQWHI12", "STTOKEBQGRE12", "STBQWG12", "STTOKEBQBLU12",
    "STTOKELOSO12BQ", "STTOKELOSS12BQ", "STTOKEST12BQ", "STTOKEYEKK12BQ",
    "STBQFSKB12", "STTOKEBQMSG12", "STTOKEESBB12BQ", "STTOKEESDY12BQ",
    "STTOKEBQMM12", "STTOKETAKMT12BQ", "STTOKEWH12BQ", "STTOKEBQBLA12",
    "STTOKELESBP12BQ", "STTOKELESIG12BQ", "STTOKELFRG12BQ",
    "STTOKEOSD12BQ", "STTOKEOSMRS12BQ", "STBQDS12", "STTOKEHG12BQ",
    "STTOKETIF12BQ", "STTOKEPINK12BQ", "STTOKELFGB12BQ", "STBQBLU12",
    "STTOKBQEDC16", "STTOKEBQMM16", "STBQFSKB16", "STBQFSP16", "STBQWG16",
    "STTOKEBQBLU16", "STTOKEUPL16BQ", "STBQFLD16",
    "STTOKELOSO16BQ", "STTOKELOSS16BQ", "STTOKEHG16BQ", "STTOKEYELL16BQ",
    "STBQWHI16", "STTOKEBQGRE16", "STTOKEST16BQ", "STTOKEWH16BQ",
    "STTOKEBQMSG16", "STTOKTIF16BQ", "STTOKEPSBLU16BQ",
    "STTOKELESBP16BQ", "STTOKELESIG16BQ", "STTOKELFGB16BQ", "STTOKELFRG16BQ",
    "STTOKEMB16BQ", "STBQBLA16", "STTOKEMT16BQ", "STTOKEPINK16BQ",
    "STTOKEESDY16BQ", "STTOKEESBB16BQ", "STTOKETAKMT16BQ", "STTOKEMOOMINW16BQ",
    "STTOKESJHEGBQ", "STTOKESJHCBBQ", "STTOKEMOOMINS20BQ",
    "STTEBTTBNBBQ", "STTEBTLINBBQ", "STTEBTGTNBBQ", "STTOKEBLSXBQ",
    "STTOKEESBB20BQ", "STTEWHI20BQ", "STTOKEHG20BQ", "STTEBLA20BQ", "STTOKEDS20BQ",
    "STTOKELOSO20BQ", "STTOKELOSS20BQ", "STTOKELESBP20BQ", "STTOKELESIG20BQ",
    "STTEBQYSG", "STTEBQTEAINF", "STTOKESX76BQ",
    "STTOKEHEBBBQ", "STTOKEHEBPBQ", "STTOKEHEMTBQ",
    "STTOKEMOOMINKLMBQ", "STTOKEMOOMINKSBQ", "STTOKEMOOMINKMBQ", "STTOKEMOOMINKMTXM",
    "STTOKEswirlhgz", "S7SCUPHSGYBQ", "STTOKEBQMM7", "STTOKEBQCG7",
    "STTOKEMOOMINKSABQ", "STTOKEBQRR7",
    "ST12BZWHLK", "ST16BZWHLK", "STDZGFBAG", "STGQGFBAG", "STTOKERR7Single",
]

SKU_ORDER_PKG_MOODY = [
    "MKTZ", "MKNCMSD", "MCFourCD", "MKWJWP",
    "NATURAL*10BGJ", "MCTZ*10BGJ", "YULIN*10BGJ", "MEISHI*10BGJ", "FOREST*10BGJ",
    "MKMFTZH",
    "FORESTBGJ", "MCTZBGJ", "YULINBGJ", "NATURALBGJ",
    "MOODYGIFT", "MOODYFBD", "MOODYCD", "MCBOTBLABQ",
    "FOREST10BOX", "MKZJ", "MoodyCoffeeCARD", "MCTZ10BOX",
]

SKU_ORDER_PKG_FOOD = [
    "BNHHBAG", "BAENKP", "BAENLD", "BAENLH", "BAENNC2", "BAENNC3", "BNMNBAG",
    "MKYNMSGEB10BQ", "BN100LHBQX", "BAEN850YF", "BAEN1050YF", "COFFEfOX10",
    "BNSCGIFTBAG", "BAENBOX", "BNbagNEW", "BNHHBOX", "BNMNBOX",
    "BNSCGIFT", "BNgfboxsnew", "BNgfboxs", "BNgfboxsnew500g*2", "BNMS", "BNRedPacket-SNAKE",
]

SKU_ORDER_PKG_UNCLASSIFIED = [
    "sttokeopp", "STTOKE20ozJST", "BNWJXN", "bannescs", "BNKDH-8",
    "QIPAO1113", "BPX-3", "BPX-4", "BPX-5", "BPX-6", "BPX-7", "BPX-8", "BPX-10",
    "MGTMGEWHNW10", "BNPDQ", "HNPZD", "BPX-BP1", "BPX-BP2",
]

# 慕咖商品页最上面置顶的4个SKU（挂耳咖啡单片），与下方其他慕咖商品之间空两行隔开
MOODY_TOP_ITEMS_SKU = ["MC-FOREST*1", "MC-HONEYPROCESS*1", "MCZG-NATURAL*1", "MCTZ-NATURAL*1"]

# 向后兼容旧引用
MOODY_TOP_ITEMS = ["云森林挂耳咖啡1片", "云雨林挂耳咖啡1片", "云之光挂耳咖啡1片"]

# 按页面类型选对应的SKU顺序表
_SKU_ORDER_MAP = {
    ("商品", "STTOKE"):   SKU_ORDER_GOODS_STTOKE,
    ("商品", "慕咖"):     SKU_ORDER_GOODS_MOODY,
    ("商品", "食品"):     SKU_ORDER_GOODS_FOOD,
    ("商品", "周边商品"): SKU_ORDER_GOODS_PERIPHERAL,
    ("包装物料", "STTOKE"):   SKU_ORDER_PKG_STTOKE,
    ("包装物料", "慕咖"):     SKU_ORDER_PKG_MOODY,
    ("包装物料", "食品"):     SKU_ORDER_PKG_FOOD,
    ("包装物料", "未分类"):   SKU_ORDER_PKG_UNCLASSIFIED,
}


def sku_sort_key(sku: str, material: str, brand: str) -> int:
    """按该页面的SKU顺序表排序；不在列表里的新SKU排到末尾。"""
    order = _SKU_ORDER_MAP.get((material, brand), [])
    try:
        return order.index(sku)
    except ValueError:
        return len(order) + 1


def classify_material(name: str) -> str:
    """返回 商品 / 包装物料"""
    if any(k in name for k in PACKAGING_KEYWORDS):
        return "包装物料"
    return "商品"


def classify_brand(name: str) -> str:
    """返回品牌：STTOKE / 慕咖 / 食品 / 周边商品 / 未分类"""
    if any(k in name for k in STTOKE_KEYWORDS):
        return "STTOKE"
    if any(k in name for k in MOODY_KEYWORDS):
        return "慕咖"
    if any(k in name for k in FOOD_KEYWORDS):
        return "食品"
    if any(k in name for k in PERIPHERAL_KEYWORDS):
        return "周边商品"
    return "未分类"


def classify(name: str) -> tuple[str, str]:
    """返回 (类型: 商品/包装物料, 品牌)，保留给旧调用方使用"""
    return classify_material(name), classify_brand(name)


def extract_group_keyword(name: str) -> str:
    """提取规格分组关键词（容量，如12OZ/16OZ/20OZ），没识别到归入"其他规格" """
    m = SIZE_PATTERN.search(name)
    return f"{int(m.group(1))}OZ" if m else "其他规格"


def group_sort_key(label: str) -> int:
    """规格分组的排序权重，按容量从小到大，"其他规格"排在最后"""
    m = re.match(r"(\d+)OZ$", label)
    return int(m.group(1)) if m else 9999
