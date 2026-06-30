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
    "绒布袋", "挂耳咖啡盒白卡纸", "单只外盒", "吸管套", "新册子", "条形码",
    "外箱", "帆布袋", "opp袋", "不干胶贴", "挂耳*10盒装",
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

# 食品品牌商品表的固定展示顺序（人工指定，不是按规格/到期日/库存量算出来的）。
# 没出现在这个列表里的食品SKU排在列表后面，按库存量从大到小排。
FOOD_CUSTOM_ORDER = [
    "巴恩天然澳洲纯蜂蜜",
    "巴恩天然活性麦努卡蜂蜜倒立装 MGO30+ 250克",
    "巴恩天然活性麦努卡蜂蜜 MGO100+ 250克",
    "巴恩天然活性麦努卡蜂蜜 MGO100+ 500克",
    "巴恩天然活性麦努卡蜂蜜 MGO300+ 250克",
    "巴恩天然活性麦努卡蜂蜜 MGO 850+ 500克",
    "巴恩天然活性麦努卡蜂蜜 MGO 1050+ 500克",
    "新-巴恩天然新西兰麦卢卡蜂蜜 MGO100+250克",
    "新-巴恩天然新西兰麦卢卡蜂蜜 MGO300+ 250克",
    "新-巴恩天然新西兰麦卢卡蜂蜜 MGO 550+ 250克",
    "新-巴恩天然新西兰麦卢卡蜂蜜 MGO 850 + 250克",
    "新-巴恩天然新西兰麦卢卡蜂蜜 MGO 1050+ 250克",
    "巴恩天然?麦卢卡蜂蜜(MGO300+)马年礼盒",
    "巴恩天然2025-花花礼盒",
    "巴恩天然2025-马年礼盒",
    "抽拉款-巴恩天然礼盒-客户定制款（SC）250g*2",
    "抽拉款-巴恩天然蜂蜜礼盒250g*2",
    "巴恩天然蜂蜜礼盒",
    "NEW巴恩天然蜂蜜礼盒500g*2",
    "巴恩天然原浆浓缩苹果醋500毫升",
    "巴恩天然原浆浓缩有机蜂蜜苹果醋500毫升",
    "巴恩天然蜂蜜木勺",
    "巴恩天然蛇年红包",
    "康蜜乐麦卢卡蜂蜜硬糖(柠姜味)",
    "康蜜乐CAPILANO儿童蜂蜜 倒立装340g赠品",
]


def food_sort_key(name: str) -> int:
    """食品品牌的固定排序权重，没在FOOD_CUSTOM_ORDER里的排在最后"""
    return FOOD_CUSTOM_ORDER.index(name) if name in FOOD_CUSTOM_ORDER else len(FOOD_CUSTOM_ORDER) + 1


# 慕咖商品表最上面单独置顶的几项（跟下面其他慕咖商品之间空两行隔开）
MOODY_TOP_ITEMS = ["云森林挂耳咖啡1片", "云雨林挂耳咖啡1片", "云之光挂耳咖啡1片"]

# 慕咖商品表（置顶项之外）的固定展示顺序（人工指定）。
# 没出现在这个列表里的慕咖SKU排在列表后面，按库存量从大到小排。
MOODY_CUSTOM_ORDER = [
    "好有水瓶 暗夜黑（无logo 光板）",
    "好有水瓶 皓月白（无logo 光板）",
    "好有水瓶mini 暗夜黑（无logo 光板）",
    "好有水瓶mini 皓月白（无logo 光板）",
    "慕咖MOODY BOTTLE 好有水瓶 皓月白",
    "慕咖MOODY BOTTLE 好有水瓶mini 暗夜黑",
    "慕咖Moody×海底小纵队儿童保温水瓶-Team",
    "慕咖Moody×海底小纵队儿童保温水瓶-缤纷",
    "慕咖moody海底小纵队暗夜黑-奔跑",
    "慕咖moody海底小纵队皓月白-Say Hi",
    "慕咖Mood Cup半糖女孩保温杯（L）",
    "慕咖Mood Cup半糖女孩保温杯（L）带logo",
    "慕咖Mood Cup半糖女孩保温杯（S）带logo",
    "慕咖Mood Cup梦幻随行杯",
    "慕咖Mood Cup柠檬汽水保温杯（L）带logo",
    "慕咖Mood Cup柠檬汽水保温杯（S）带logo",
    "慕咖Mood Cup微雨随行杯",
    "慕咖Mood Cup元气骑士保温杯（S）",
    "慕咖Mood Cup元气骑士保温杯（S）带logo",
    "慕咖Mood Cup元气森屿保温杯（L）带logo",
    "慕咖Mood Cup元气森屿保温杯（S）带logo",
    "慕咖MoodCone蛋筒便携杯（冰雪藍）",
    "慕咖MoodCone蛋筒便携杯（苏打绿）",
]


def moody_sort_key(name: str) -> int:
    """慕咖品牌（置顶项之外）的固定排序权重，没在MOODY_CUSTOM_ORDER里的排在最后"""
    return MOODY_CUSTOM_ORDER.index(name) if name in MOODY_CUSTOM_ORDER else len(MOODY_CUSTOM_ORDER) + 1


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
