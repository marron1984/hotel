# -*- coding: utf-8 -*-
"""
ホテル収支計画(Proforma)モデルの「入力セル」定義。

元のサンプルExcel (template.xlsx) は数式で相互参照された精緻なモデルであり、
ここではユーザーが実際に入力すべき前提値(ドライバー)のセルだけを定義する。
Web フォームの描画と、Excel への書き戻しの両方でこの定義を共有する (DRY)。

field の属性:
  sheet   : シート名
  cell    : セル座標 (例 "D95")
  label   : 画面表示ラベル(日本語)
  kind    : "num" 数値 / "pct" 割合(セルは小数=fraction, 画面は%表示) /
            "int" 整数 / "text" 文字列
  help    : 補足説明(任意)
  spread  : 同じ値を書き込む追加セルのリスト(任意。月別セルなどに横展開)
"""

# ---- 客室タイプ (Room Types シート 7〜16 行) --------------------------------
# Excel上の対応:
#   Room Types!C{r}=名称, D{r}=室数, F{r}/G{r}=広さ(sqm)
#   Rooms!G{35+i}=稼働率(タイプ別), Rooms!G{53+i}=客室単価(タイプ別)
ROOM_TYPE_ROWS = list(range(7, 17))  # 7..16 (最大10タイプ)


def room_type_fields():
    fields = []
    for i, r in enumerate(ROOM_TYPE_ROWS):
        fields.append({
            "idx": i,
            "name_cell":  ("Room Types", f"C{r}"),
            "rooms_cell": ("Room Types", f"D{r}"),
            "size_cells": [("Room Types", f"F{r}"), ("Room Types", f"G{r}")],
            "occ_cell":   ("Rooms", f"G{35 + i}"),
            "rate_cell":  ("Rooms", f"G{53 + i}"),
        })
    return fields


# ---- 平坦な入力フィールド群 (セクション単位) --------------------------------
SECTIONS = [
    {
        "id": "general",
        "title": "基本情報",
        "fields": [
            {"sheet": "Room Types", "cell": "A1", "label": "ホテル名", "kind": "text",
             "help": "全シートのタイトルに反映されます"},
            {"sheet": "Rooms", "cell": "E14", "label": "1室あたり宿泊人数", "kind": "num",
             "help": "稼働客室1室あたりの平均宿泊者数"},
        ],
    },
    {
        "id": "growth",
        "title": "成長率の前提 (2〜5年目)",
        "fields": [
            {"sheet": "Proforma", "cell": "F14", "label": "稼働率の伸び 2年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "I14", "label": "稼働率の伸び 3年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "L14", "label": "稼働率の伸び 4年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "O14", "label": "稼働率の伸び 5年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "F15", "label": "ADRの伸び 2年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "I15", "label": "ADRの伸び 3年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "L15", "label": "ADRの伸び 4年目", "kind": "pct"},
            {"sheet": "Proforma", "cell": "O15", "label": "ADRの伸び 5年目", "kind": "pct"},
        ],
    },
    {
        "id": "rooms_cost",
        "title": "客室部門 その他費用",
        "fields": [
            {"sheet": "Rooms", "cell": "J173", "label": "客室清掃費 (1室あたり/円)", "kind": "num"},
            {"sheet": "Rooms", "cell": "J175", "label": "リネン費 (宿泊者1人あたり/円)", "kind": "num"},
            {"sheet": "Rooms", "cell": "J176", "label": "客室消耗品 (宿泊者1人あたり/円)", "kind": "num"},
            {"sheet": "Rooms", "cell": "J177", "label": "その他費用率 (客室売上比)", "kind": "pct"},
            {"sheet": "Rooms", "cell": "M174", "label": "海外OTA比率", "kind": "pct"},
            {"sheet": "Rooms", "cell": "N174", "label": "海外OTA経由率", "kind": "pct"},
            {"sheet": "Rooms", "cell": "O174", "label": "海外OTA手数料率", "kind": "pct"},
            {"sheet": "Rooms", "cell": "N175", "label": "国内OTA経由率", "kind": "pct"},
            {"sheet": "Rooms", "cell": "O175", "label": "国内OTA手数料率", "kind": "pct"},
        ],
    },
    {
        "id": "restaurant",
        "title": "レストラン部門",
        "fields": [
            {"sheet": "Restaurant", "cell": "D22", "label": "朝食 席数", "kind": "int"},
            {"sheet": "Restaurant", "cell": "D24", "label": "朝食 宿泊客カバー率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "D26", "label": "朝食 客単価 (円)", "kind": "num"},
            {"sheet": "Restaurant", "cell": "D28", "label": "朝食 ウォークイン回転率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "D41", "label": "昼食 席数", "kind": "int"},
            {"sheet": "Restaurant", "cell": "D43", "label": "昼食 宿泊客カバー率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "D45", "label": "昼食 客単価 (円)", "kind": "num"},
            {"sheet": "Restaurant", "cell": "D47", "label": "昼食 ウォークイン回転率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "D60", "label": "夕食 席数", "kind": "int"},
            {"sheet": "Restaurant", "cell": "D62", "label": "夕食 宿泊客カバー率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "D64", "label": "夕食 客単価 (円)", "kind": "num"},
            {"sheet": "Restaurant", "cell": "D66", "label": "夕食 ウォークイン回転率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "H170", "label": "食材・飲料 原価率", "kind": "pct"},
            {"sheet": "Restaurant", "cell": "H171", "label": "その他費用率", "kind": "pct"},
        ],
    },
    {
        "id": "mod",
        "title": "その他収益部門 (MOD)",
        "fields": [
            {"sheet": "MOD", "cell": "E20", "label": "MODチャージ (宿泊者1人あたり/円)", "kind": "num",
             "spread": [f"{c}20" for c in "FGHIJKLMNOP"] + ["Q20"]},
            {"sheet": "MOD", "cell": "E24", "label": "MOD原価率", "kind": "pct"},
            {"sheet": "MOD", "cell": "D32", "label": "テナント1 面積 (坪)", "kind": "num"},
            {"sheet": "MOD", "cell": "D33", "label": "テナント1 賃料 (坪/月/円)", "kind": "num"},
            {"sheet": "MOD", "cell": "D34", "label": "テナント1 稼働率", "kind": "pct"},
            {"sheet": "MOD", "cell": "D39", "label": "テナント2 面積 (坪)", "kind": "num"},
            {"sheet": "MOD", "cell": "D40", "label": "テナント2 賃料 (坪/月/円)", "kind": "num"},
            {"sheet": "MOD", "cell": "D41", "label": "テナント2 稼働率", "kind": "pct"},
            {"sheet": "MOD", "cell": "D46", "label": "テナント3 面積 (坪)", "kind": "num"},
            {"sheet": "MOD", "cell": "D47", "label": "テナント3 賃料 (坪/月/円)", "kind": "num"},
            {"sheet": "MOD", "cell": "D48", "label": "テナント3 稼働率", "kind": "pct"},
            {"sheet": "MOD", "cell": "D56", "label": "駐車場 台数", "kind": "int"},
            {"sheet": "MOD", "cell": "D57", "label": "駐車料金 (1日/円)", "kind": "num"},
            {"sheet": "MOD", "cell": "D58", "label": "駐車場 稼働率", "kind": "pct"},
            {"sheet": "MOD", "cell": "H97", "label": "MOD 原価率 (PL)", "kind": "pct"},
            {"sheet": "MOD", "cell": "H98", "label": "MOD その他費用率", "kind": "pct"},
        ],
    },
    {
        "id": "undist",
        "title": "非配賦費用",
        "fields": [
            {"sheet": "Undist Expense", "cell": "H130", "label": "クレジットカード利用率", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "J130", "label": "クレジットカード手数料率", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "H131", "label": "通信費 (1室あたり/月/円)", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H132", "label": "一般管理 その他経費率", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "E133", "label": "システム関連費用 (年額/円)", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H134", "label": "業務委託費(本社経費) 月額/円", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H137", "label": "広告宣伝費 (GOR比)", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "H138", "label": "S&Mフィー (GOR比)", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "H139", "label": "営業マーケティング その他 (GOR比)", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "H140", "label": "S&M委託料 月額/円", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H143", "label": "施設管理費 (月坪単価/円)", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "P143", "label": "施設面積 (坪)", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "P144", "label": "修繕費 月額/円", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H145", "label": "保険(運営) 年額/円", "kind": "num"},
            {"sheet": "Undist Expense", "cell": "H146", "label": "施設維持 その他率", "kind": "pct"},
            {"sheet": "Undist Expense", "cell": "H149", "label": "水道光熱費 (稼働室1泊あたり/円)", "kind": "num"},
        ],
    },
    {
        "id": "fees",
        "title": "マネジメントフィー・オーナー経費",
        "fields": [
            {"sheet": "Proforma", "cell": "R87", "label": "ベースマネジメントフィー (GOR比)", "kind": "pct"},
            {"sheet": "Proforma", "cell": "R92", "label": "インセンティブフィー (GOP比)", "kind": "pct"},
            {"sheet": "Proforma", "cell": "D95", "label": "固定資産税 (年額/円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "R96", "label": "損害保険料 (1室あたり/円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "R97", "label": "FF&Eリザーブ (GOR比)", "kind": "pct"},
            {"sheet": "Proforma", "cell": "R98", "label": "Capexリザーブ (GOR比)", "kind": "pct"},
        ],
    },
    {
        "id": "project",
        "title": "プロジェクトコスト",
        "fields": [
            {"sheet": "Proforma", "cell": "T8",  "label": "取得価格 (FF&E込/円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "T9",  "label": "新規サイネージ等 (円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "T10", "label": "OSE等 (円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "T11", "label": "リブランディング費用 (円)", "kind": "num"},
            {"sheet": "Proforma", "cell": "V11", "label": "リブランディング月数", "kind": "num"},
            {"sheet": "Proforma", "cell": "V12", "label": "運転資金 月数", "kind": "num"},
            {"sheet": "Proforma", "cell": "T13", "label": "オペレーター技術料 (円)", "kind": "num"},
        ],
    },
]


# ---- 人員計画 (給与・人数) --------------------------------------------------
# 各部門ごとに「役職ラベル / 月給セル / 人数セル」を対応付ける。
# 月給・人数は元モデルでは同じ並び順の行に格納されている。
def _pairs(sheet, col, items):
    """items = [(label, salary_row, head_row), ...]"""
    return {"sheet": sheet, "col": col, "roles": [
        {"label": l, "salary_cell": f"{col}{sr}", "head_cell": f"{col}{hr}"}
        for (l, sr, hr) in items
    ]}


STAFFING = [
    {"id": "staff_rooms", "title": "人員計画 — 客室部門", **_pairs("Rooms", "G", [
        ("Administration / Operation Manager", 110, 131),
        ("Front / Manager", 111, 132),
        ("Front / Assistant Manager", 112, 133),
        ("Front / Team Leader", 113, 134),
        ("Front / Guest Service Officer", 114, 135),
        ("Front / Guest Service Officer (PT)", 115, 136),
    ])},
    {"id": "staff_rest", "title": "人員計画 — レストラン部門", **_pairs("Restaurant", "E", [
        ("All Day Dining / Manager", 123, 138),
        ("All Day Dining / Assistant Manager", 124, 139),
        ("All Day Dining / Team Leader", 125, 140),
        ("All Day Dining / Waiter (R&F)", 126, 141),
        ("All Day Dining / Waiter (PT)", 127, 142),
        ("Culinary / Chef de Cuisine", 128, 143),
        ("Culinary / Sous Chef", 129, 144),
        ("Culinary / Chef de Partie", 130, 145),
        ("Culinary / Commis (R&F)", 131, 146),
        ("Culinary / Commis (PT・朝食)", 132, 147),
    ])},
    {"id": "staff_mod", "title": "人員計画 — その他収益部門", **_pairs("MOD", "E", [
        ("Other Revenue / Manager", 65, 75),
        ("Other Revenue / Assistant Manager", 66, 76),
        ("Other Revenue / Team Leader 1", 67, 77),
        ("Other Revenue / Team Leader 2", 68, 78),
        ("Other Revenue / Part Time", 69, 79),
    ])},
    {"id": "staff_undist", "title": "人員計画 — 非配賦部門", **_pairs("Undist Expense", "E", [
        ("A&G / Project Leader", 32, 65),
        ("A&G / Hotel Manager", 33, 66),
        ("A&G / Assistant to GM", 34, 67),
        ("A&G / Dir of Admin & Finance", 35, 68),
        ("Finance / Manager", 36, 69),
        ("Finance / Assistant Manager", 37, 70),
        ("Finance / Team Leader", 38, 71),
        ("Finance / Officer", 39, 72),
        ("IS / Manager", 40, 73),
        ("IS / Assistant Manager", 41, 74),
        ("Material / Manager", 42, 75),
        ("Material / Assistant Manager", 43, 76),
        ("Material / Officer", 44, 77),
        ("HR / Manager", 45, 78),
        ("HR / Assistant Manager", 46, 79),
        ("HR / Officer", 47, 80),
        ("Training / Manager", 48, 81),
        ("Training / Assistant Manager", 49, 82),
        ("S&M / Marketing Manager", 52, 85),
        ("Sales / Assistant Manager", 53, 86),
        ("S&M / Coordinator", 54, 87),
        ("Engineering / Manager", 57, 90),
        ("Engineering / Officer", 58, 91),
    ])},
]


# =========================================================================== #
#  出力フォーマット選択
# =========================================================================== #
# detailed: 既存の詳細モデル(7シート・数式)。 nara: 1枚もの収支サマリー。
OUTPUTS = [
    {"id": "detailed", "label": "詳細モデル（全シート・数式付き）", "file": "template.xlsx"},
    {"id": "nara", "label": "サマリー形式（1枚もの・Naraレイアウト）", "file": "template_nara.xlsx"},
]


# --- Nara サマリーの入力フィールド（シート名は "Sheet1"） -------------------
def _f(cell, label, kind, help=None, spread=None):
    d = {"sheet": "Sheet1", "cell": cell, "label": label, "kind": kind}
    if help:
        d["help"] = help
    if spread:
        d["spread"] = spread
    return d


NARA_SECTIONS = [
    {
        "id": "n_general",
        "title": "基本情報",
        "fields": [
            _f("A1", "ホテル名", "text"),
            _f("D11", "客室数", "int", "全年共通（2年目以降に自動展開）"),
            _f("D10", "年間日数", "int"),
            _f("D14", "稼働率（1年目）", "pct"),
            _f("D15", "ADR（1年目／円）", "num"),
            _f("D18", "1室あたり宿泊人数", "num"),
            _f("C30", "サービスチャージ率", "pct", "レストラン＋MOD売上に対して"),
        ],
    },
    {
        "id": "n_growth",
        "title": "成長率の前提（2〜5年目）",
        "fields": [
            _f("F14", "稼働率の伸び 2年目", "pct"), _f("I14", "稼働率の伸び 3年目", "pct"),
            _f("L14", "稼働率の伸び 4年目", "pct"), _f("O14", "稼働率の伸び 5年目", "pct"),
            _f("F15", "ADRの伸び 2年目", "pct"), _f("I15", "ADRの伸び 3年目", "pct"),
            _f("L15", "ADRの伸び 4年目", "pct"), _f("O15", "ADRの伸び 5年目", "pct"),
            _f("F28", "レストラン売上の伸び 2年目", "pct"), _f("I28", "同 3年目", "pct"),
            _f("L28", "同 4年目", "pct"), _f("O28", "同 5年目", "pct"),
            _f("F29", "MOD売上の伸び 2年目", "pct"), _f("I29", "同 3年目", "pct"),
            _f("L29", "同 4年目", "pct"), _f("O29", "同 5年目", "pct"),
        ],
    },
    {
        "id": "n_rev",
        "title": "売上・原価（1年目）",
        "fields": [
            _f("D28", "レストラン売上（1年目／円）", "num"),
            _f("D29", "MOD + ルーフトップバー売上（1年目／円）", "num"),
            _f("E35", "レストラン原価率", "pct", spread=["H35", "K35", "N35", "Q35"]),
            _f("E36", "MOD原価率", "pct", spread=["H36", "K36", "N36", "Q36"]),
            _f("E46", "客室 その他費用率（対客室売上）", "pct", spread=["H46", "K46", "N46", "Q46"]),
            _f("E47", "レストラン その他費用率", "pct", spread=["H47", "K47", "N47", "Q47"]),
            _f("E48", "MOD その他費用率", "pct", spread=["H48", "K48", "N48", "Q48"]),
        ],
    },
    {
        "id": "n_personnel",
        "title": "部門人件費（年額・全年共通）",
        "fields": [
            _f("D40", "客室 人件費（円/年）", "num"),
            _f("D41", "レストラン 人件費（円/年）", "num"),
            _f("D42", "MOD 人件費（円/年）", "num"),
        ],
    },
    {
        "id": "n_undist",
        "title": "非配賦費用",
        "fields": [
            _f("D66", "一般管理 人件費（円/年）", "num"),
            _f("D67", "一般管理 その他（1年目/円）", "num"),
            _f("F67", "一般管理 その他の伸び 2年目", "pct"),
            _f("I67", "同 3年目", "pct"),
            _f("L67", "同 4年目", "pct"),
            _f("O67", "同 5年目", "pct"),
            _f("D69", "営業・マーケ 人件費（円/年）", "num"),
            _f("D70", "営業・マーケ その他（1年目/円）", "num"),
            _f("H70", "営業・マーケ その他（2年目以降・対総売上）", "pct",
               spread=["K70", "N70", "Q70"]),
            _f("D72", "施設維持 人件費（円/年）", "num"),
            _f("D73", "施設維持 その他（1年目/円）", "num"),
            _f("F73", "施設維持 その他の伸び 2年目", "pct"),
            _f("I73", "同 3年目", "pct"),
            _f("L73", "同 4年目", "pct"),
            _f("O73", "同 5年目", "pct"),
            _f("C74", "水道光熱費（稼働室1泊あたり/円）", "num"),
        ],
    },
    {
        "id": "n_fees",
        "title": "マネジメントフィー・リザーブ",
        "fields": [
            _f("R90", "ベースマネジメントフィー（対GOR）", "pct"),
            _f("R95", "インセンティブフィー（対GOP）", "pct"),
            _f("R101", "FF&Eリザーブ（対GOR）", "pct"),
            _f("R102", "Capexリザーブ（対GOR）", "pct"),
        ],
    },
    {
        "id": "n_owner",
        "title": "オーナー経費（年額・全年共通）",
        "fields": [
            _f("D98", "固定資産税（建物のみ／円）", "num"),
            _f("D99", "借地料 Ground Lease（円/年）", "num"),
            _f("D100", "損害保険料（円/年）", "num"),
        ],
    },
    {
        "id": "n_project",
        "title": "プロジェクトコスト",
        "fields": [
            _f("T22", "土地関係（円）", "num"),
            _f("T23", "建物関係（円）", "num"),
            _f("T24", "開発経費（dhp都市開発／円）", "num"),
            _f("T25", "Deposit（円）", "num"),
            _f("T26", "FF&E（円）", "num"),
            _f("T27", "OSE等（円）", "num"),
            _f("T28", "Pre-Opening（円）", "num"),
            _f("T29", "Working Capital（円）", "num"),
            _f("T30", "HMJ技術料（円）", "num"),
            _f("T31", "Upfront Moxy Party Fees（円）", "num"),
        ],
    },
]
