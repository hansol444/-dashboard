"""
post_processor.py — Q2: 번역 후처리 규칙
각 규칙은 독립 함수로 분리되어 있으며, config에서 on/off 가능합니다.
"""

import re
from dataclasses import dataclass, field


@dataclass
class PostProcessResult:
    text: str
    changes: list[str] = field(default_factory=list)

    def changed(self, rule: str, before: str, after: str):
        self.changes.append(f"[{rule}] '{before}' → '{after}'")


# ─── 규칙 함수들 ────────────────────────────────────────────────────────────────

def fix_billion(translated: str, original: str = "") -> PostProcessResult:
    """
    한국어 원문의 억/조 단위를 기반으로 번역 결과의 수치 오류를 교정.
    예: 원문 "14억" → 번역 "14B" → 교정 "1.4B"
        1억 = 100M, 10억 = 1B, 1조 = 1T
    """
    result = PostProcessResult(text=translated)

    # 원문에서 숫자+억/조 패턴 추출
    oku_matches = re.findall(r'(\d+(?:\.\d+)?)\s*억', original)
    jo_matches = re.findall(r'(\d+(?:\.\d+)?)\s*조', original)

    text = result.text
    for num in oku_matches:
        # 1억 = 100M → 숫자B 형태가 나오면 교정
        ko_val = float(num)
        correct_val = ko_val / 10  # 14억 → 1.4B
        # 번역에서 "14B" 패턴
        pattern = rf'\b{re.escape(num.rstrip("0").rstrip("."))}B\b'
        correct_str = f"{correct_val:g}B"
        new_text = re.sub(pattern, correct_str, text)
        if new_text != text:
            result.changed("fix_billion", f"{num}B", correct_str)
            text = new_text

    for num in jo_matches:
        ko_val = float(num)
        correct_val = ko_val  # 1조 = 1T
        pattern = rf'\b{re.escape(num.rstrip("0").rstrip("."))}T\b'
        correct_str = f"{correct_val:g}T"
        new_text = re.sub(pattern, correct_str, text)
        if new_text != text:
            result.changed("fix_billion", f"{num}T", correct_str)
            text = new_text

    result.text = text
    return result


def fix_duplicates(translated: str, original: str = "") -> PostProcessResult:
    """
    연속 중복 단어/구문 제거.
    - 3회 이상 연속 반복: 무조건 제거 (예: "JK JK JK" → "JK")
    - 2회 연속 반복: 짧은 약어/코드(1~5자)만 제거 (예: "JK JK" → "JK")
      긴 단어의 의도적 반복은 유지 (예: "very very" 유지)
    """
    result = PostProcessResult(text=translated)
    text = result.text

    # 3회 이상 연속 반복 제거
    pattern_3plus = r'\b(\w+)(\s+\1){2,}\b'
    new_text = re.sub(pattern_3plus, r'\1', text)
    if new_text != text:
        result.changed("fix_duplicates", text, new_text)
        text = new_text

    # 2회 연속: 짧은 토큰(1~5자)만 제거 — 약어/코드 오류 대응
    def dedup_short(m):
        word = m.group(1)
        if len(word) <= 5:
            result.changed("fix_duplicates", m.group(0), word)
            return word
        return m.group(0)  # 긴 단어는 유지

    text = re.sub(r'\b(\w+)\s+\1\b', dedup_short, text)

    result.text = text
    return result


def fix_month_abbrev(translated: str, original: str = "") -> PostProcessResult:
    """
    원문에 월 표기(1월~12월)가 있는데 번역에서 숫자+M으로 잘못 번역된 경우 교정.
    예: 원문 "1월" → 번역 "1M" → 교정 "January"
    """
    MONTHS = {
        1: "January", 2: "February", 3: "March", 4: "April",
        5: "May", 6: "June", 7: "July", 8: "August",
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    result = PostProcessResult(text=translated)

    # 원문에서 N월 패턴 추출
    month_matches = re.findall(r'(\d{1,2})월', original)
    text = result.text

    # 금액 컨텍스트 판별용 통화 키워드
    CURRENCY_CONTEXT = r'(?:KRW|USD|AUD|EUR|JPY|GBP|CNY|SGD|\$|£|€|¥)'

    for m in month_matches:
        num = int(m)
        if 1 <= num <= 12:
            # 번역에서 "1M", "12M" 같은 잘못된 패턴 찾기
            # 단, 금액 컨텍스트(통화코드 인접)에 있으면 건너뜀
            pattern = rf'(?<!\d){num}M(?!\w)'
            for match in re.finditer(pattern, text):
                start, end = match.start(), match.end()
                surrounding = text[max(0, start-10):min(len(text), end+10)]
                if re.search(CURRENCY_CONTEXT, surrounding):
                    continue  # 금액 단위 M → 건너뜀
                correct = MONTHS[num]
                text = text[:start] + correct + text[end:]
                result.changed("fix_month_abbrev", f"{num}M", correct)
                break  # 한 번에 하나씩 교체 (인덱스 변동 방지)

    result.text = text
    return result


def fix_currency_order(translated: str, original: str = "") -> PostProcessResult:
    """
    통화 코드가 숫자 뒤에 오는 경우 앞으로 이동.
    예: "100M KRW" → "KRW 100M"
    """
    CURRENCIES = ["KRW", "USD", "AUD", "EUR", "JPY", "GBP", "CNY", "SGD"]
    result = PostProcessResult(text=translated)
    text = result.text

    for currency in CURRENCIES:
        # "숫자(단위) 통화코드" 패턴
        pattern = rf'(\d+(?:\.\d+)?(?:[KMBT])?)\s+({currency})\b'
        match = re.search(pattern, text)
        if match:
            new_text = re.sub(pattern, rf'{currency} \1', text)
            result.changed("fix_currency_order", match.group(0), f"{currency} {match.group(1)}")
            text = new_text

    result.text = text
    return result


def fix_australian_spelling(translated: str, original: str = "") -> PostProcessResult:
    """
    미국식 철자를 호주식으로 교정.
    """
    SPELLING_MAP = {
        r'\borganization\b': 'organisation',
        r'\borganizations\b': 'organisations',
        r'\banalyze\b': 'analyse',
        r'\banalyzes\b': 'analyses',
        r'\banalyzing\b': 'analysing',
        r'\bcolor\b': 'colour',
        r'\bcolors\b': 'colours',
        r'\bcenter\b': 'centre',
        r'\bcenters\b': 'centres',
        r'\bbehavior\b': 'behaviour',
        r'\bbehaviors\b': 'behaviours',
        r'\blabor\b': 'labour',
        r'\bprogram\b': 'programme',
        r'\bprograms\b': 'programmes',
        r'\blicense\b': 'licence',
        r'\brecognize\b': 'recognise',
        r'\brecognizes\b': 'recognises',
        r'\bspecialize\b': 'specialise',
        r'\bspecializes\b': 'specialises',
        r'\bmaximize\b': 'maximise',
        r'\bminimize\b': 'minimise',
        r'\bprioritize\b': 'prioritise',
        r'\bstandardize\b': 'standardise',
        r'\bcustomize\b': 'customise',
        r'\boptimize\b': 'optimise',
        r'\bsummarize\b': 'summarise',
        r'\bfavorite\b': 'favourite',
        r'\bfavor\b': 'favour',
        r'\bcatalog\b': 'catalogue',
        r'\bdialog\b': 'dialogue',
        r'\bcheck\b': 'check',  # 유지 (같음)
    }

    result = PostProcessResult(text=translated)
    text = result.text

    def case_preserving_replace(match, replacement):
        """원래 단어의 대소문자 패턴을 유지한 채 교체."""
        original = match.group(0)
        if original.isupper():
            return replacement.upper()
        if original[0].isupper():
            return replacement[0].upper() + replacement[1:]
        return replacement

    for pattern, replacement in SPELLING_MAP.items():
        # 패턴의 핵심 단어와 replacement가 같으면 건너뜀 (예: check → check)
        core_word = pattern.replace(r'\b', '')
        if core_word == replacement:
            continue
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            fixed = case_preserving_replace(match, replacement)
            new_text = re.sub(pattern, lambda m: case_preserving_replace(m, replacement), text, flags=re.IGNORECASE)
            if new_text != text:
                result.changed("fix_australian_spelling", match.group(0), fixed)
                text = new_text

    result.text = text
    return result


# ─── PostProcessor 클래스 ───────────────────────────────────────────────────────

RULE_FUNCTIONS = {
    "fix_billion": fix_billion,
    "fix_duplicates": fix_duplicates,
    "fix_month_abbrev": fix_month_abbrev,
    "fix_currency_order": fix_currency_order,
    "fix_australian_spelling": fix_australian_spelling,
}

DEFAULT_RULES = list(RULE_FUNCTIONS.keys())


class PostProcessor:
    def __init__(self, enabled_rules: list[str] = None):
        self.enabled_rules = enabled_rules if enabled_rules is not None else DEFAULT_RULES

    def process(self, translated: str, original: str = "") -> tuple[str, list[str]]:
        """
        활성화된 규칙을 순서대로 적용.
        반환: (최종 번역문, 변경 로그 리스트)
        """
        text = translated
        all_changes = []

        for rule_name in self.enabled_rules:
            fn = RULE_FUNCTIONS.get(rule_name)
            if fn is None:
                continue
            result = fn(text, original)
            text = result.text
            all_changes.extend(result.changes)

        return text, all_changes


if __name__ == "__main__":
    pp = PostProcessor()

    tests = [
        ("14B revenue", "매출 14억"),
        ("JK JK joined the team", "JK가 팀에 합류"),
        ("1M results were strong", "1월 실적이 좋았다"),
        ("100M KRW budget", "예산 100M KRW"),
        ("The organization will analyze the color data", ""),
    ]

    for translated, original in tests:
        result, changes = pp.process(translated, original)
        print(f"입력: {translated}")
        print(f"결과: {result}")
        if changes:
            for c in changes:
                print(f"  {c}")
        print()
