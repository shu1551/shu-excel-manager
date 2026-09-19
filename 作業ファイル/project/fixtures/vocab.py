# -*- coding: utf-8 -*-
"""お題の表の語を種で選ぶ（2026-09-18）。

鍛えた 27 本が「課名・支出額・T支出明細・課マスタ」を書き込み、職場の表で動かなかった。お題の表が全部同じ語だったので
試験で落ちなかった＝表ごとに見出しの語・シート名・テーブル名・項目の名前を変え、語を書き込んだマクロが必ず落ちる試験にする。

  v = Vocab(seed)
  v.word('ka')            … 項目の列の見出し（課名・所属・部署名…）。同じ Vocab の中では役ごとに 1 つに決まり、役どうし重ならない
  v.items(n)              … 項目の名前 n 個（課・地区・施設・事業のどれかの組）
  v.kamokus(n)            … 科目の名前 n 個（科目・費目・節のどれかの組）
  v.sheet('meisai')       … シート名（支出明細・明細・データ・Sheet1…）
  v.table('meisai')       … テーブル名（T支出明細・tbl明細・データ表…）
"""
import random

HEADS = {
    'ka': ['課名', '所属', '部署名', '担当課', '所属名', '部署', '課'],
    'amount': ['支出額', '金額', '支払額', '決算額', '支出金額', '金額（円）'],
    'kamoku': ['科目', '費目', '節', '科目名', '費目名'],
    'date': ['支出日', '支払日', '日付', '伝票日付', '支出年月日'],
    'tekiyo': ['摘要', '内容', '件名', '備考', '支出内容'],
    'yosan': ['予算額', '予算現額', '当初予算', '予算'],
    'shikko': ['執行額', '支出済額', '執行済額', '実績額'],
    'rate': ['執行率', '執行割合', '進捗率', '消化率'],
    'code': ['課コード', '所属コード', 'コード', '部署コード'],
    'no': ['伝票番号', '整理番号', 'No', '番号', '管理番号'],
    'prev': ['前年度', '前年度決算額', 'R7年度', '前年度実績'],
    'cur': ['今年度', '当年度', 'R8年度', '本年度実績'],
    'ninzu': ['人数', '職員数', '人員'],
    'name': ['氏名', '職員名', '名前'],
    'kihon': ['基本給', '給料', '本俸'],
    'teate': ['手当', '諸手当', '手当額'],
}

# 項目の名前の組（課・地区・施設・事業）
ITEMS = {
    '課': ['総務課', '財政課', '企画課', '税務課', '福祉課', '建設課', '農林課', '教育総務課', '観光課', '環境課', '会計課',
          '議会事務局', '商工課', '保健課', '子育て支援課', '市民課', '防災課', '都市計画課', '上下水道課', '文化スポーツ課'],
    '地区': ['北地区', '南地区', '東地区', '西地区', '中央地区', '山間地区', '沿岸地区', '駅前地区', '新町地区', '本町地区',
           '旭地区', '若葉地区', '緑ヶ丘地区', '港地区', '川西地区', '川東地区', '高台地区', '桜地区', '松原地区', '泉地区'],
    '施設': ['中央公民館', '市民体育館', '図書館', '郷土資料館', '児童センター', '文化会館', '保健センター', '福祉会館', '斎場',
           '清掃センター', '浄水場', '学校給食センター', '市民プール', '陸上競技場', '野球場', '勤労会館', '交流センター',
           '美術館', '観光案内所', '道の駅'],
    '事業': ['道路維持事業', '除雪事業', '広報事業', '健診事業', '保育事業', '観光振興事業', '農地整備事業', '防災訓練事業',
           '図書購入事業', '公園管理事業', '移住支援事業', '商店街支援事業', '学校改修事業', '水道更新事業', 'ごみ収集事業',
           '敬老事業', '子育て相談事業', '文化祭事業', 'スポーツ大会事業', '空き家対策事業'],
}

# 項目の列の見出し（項目の組に合わせる）
ITEM_HEADS = {
    '課': HEADS['ka'],
    '地区': ['地区', '地区名', '地域', '地域名'],
    '施設': ['施設名', '施設', '施設の名称'],
    '事業': ['事業名', '事業', '細事業名'],
}

KAMOKUS = {
    '科目': ['需用費', '役務費', '委託料', '使用料及び賃借料', '工事請負費', '備品購入費', '負担金補助及び交付金', '旅費', '報償費',
           '原材料費'],
    '費目': ['消耗品費', '燃料費', '食糧費', '印刷製本費', '光熱水費', '修繕料', '通信運搬費', '手数料', '保険料', '賃借料'],
    '節': ['報酬', '給料', '職員手当等', '共済費', '災害補償費', '賃金', '交際費', '公課費', '補償補塡及び賠償金', '積立金'],
}

KAMOKU_HEADS = {'科目': ['科目', '科目名'], '費目': ['費目', '費目名'], '節': ['節', '節名', '節区分']}

SHEETS = {
    'meisai': ['支出明細', '明細', 'データ', 'Sheet1', '4月分', '支出一覧', '台帳'],
    'shukei': ['課別支出', '集計', '集計表', 'Sheet1', '一覧', 'まとめ', '実績'],
    'getsuji': ['月次集計', '月別', '推移', 'Sheet1', '月別実績', '月報'],
    'master': ['課マスタ', 'マスタ', 'コード表', '対応表', '所属一覧'],
    'meibo': ['職員名簿', '名簿', '職員一覧', 'Sheet1', '人事データ'],
    'yosan': ['予算執行', '執行状況', '予算', 'Sheet1', '執行管理'],
}

TABLE_PREFIX = ['T', 'tbl', 'テーブル_', '']


class Vocab:
    """種で語を決める。役ごとに 1 語・役どうし重ならない（支出額と執行額の役が同じ「執行額」にならない）。"""

    def __init__(self, seed, item_kind=None, kamoku_kind=None):
        self.seed = seed
        self.rng = random.Random(f"vocab-{seed}")
        self._words = {}
        self._sheets = {}
        self.item_kind = item_kind or self.rng.choice(sorted(ITEMS))
        self.kamoku_kind = kamoku_kind or self.rng.choice(sorted(KAMOKUS))
        self.prefix = self.rng.choice(TABLE_PREFIX)

    def word(self, role):
        if role not in self._words:
            used = set(self._words.values())
            pool = (ITEM_HEADS[self.item_kind] if role == 'ka' else KAMOKU_HEADS[self.kamoku_kind] if role == 'kamoku'
                    else HEADS[role])
            cands = [w for w in pool if w not in used]
            self._words[role] = self.rng.choice(cands)
        return self._words[role]

    def items(self, n, rng=None):
        r = rng or self.rng
        pool = ITEMS[self.item_kind]
        if n <= len(pool):
            return r.sample(pool, n)
        return [f"{r.choice(pool)}{i + 1}" for i in range(n)]

    def kamokus(self, n, rng=None):
        r = rng or self.rng
        pool = KAMOKUS[self.kamoku_kind]
        return r.sample(pool, min(n, len(pool)))

    def sheet(self, role):
        if role not in self._sheets:
            used = set(self._sheets.values())
            self._sheets[role] = self.rng.choice([s for s in SHEETS[role] if s not in used])
        return self._sheets[role]

    def table(self, role):
        base = self.sheet(role)
        name = f"{self.prefix}{base}" if self.prefix else f"{base}表"
        name = name.replace('Sheet1', 'データ').replace(' ', '')
        if not name or name[0].isdigit():
            name = 'T' + name                 # テーブル名は数字で始められない（Excel がブックを開けなくなる）
        return name


if __name__ == '__main__':
    for s in (1, 2, 3):
        v = Vocab(s)
        print(s, v.word('ka'), v.word('amount'), v.word('kamoku'), v.sheet('meisai'), v.table('meisai'), v.items(3), v.kamokus(2))
