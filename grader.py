"""可解释的英语作文批改核心。

项目只使用 Python 标准库，适合离线课堂演示。系统把词汇统计、篇章结构和
可复现的语法规则组合为五维评分，并为每条问题保留定位与修改建议。
"""

from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass, asdict
from typing import Iterable


WORD_RE = re.compile(r"[A-Za-z]+(?:'[A-Za-z]+)?")
SENTENCE_RE = re.compile(r"[^.!?]+[.!?]?", re.MULTILINE)
STOP_WORDS = {
    "a", "an", "and", "are", "as", "at", "be", "because", "been", "but", "by",
    "can", "do", "for", "from", "had", "has", "have", "he", "her", "his", "how",
    "i", "if", "in", "is", "it", "its", "me", "my", "of", "on", "or", "our",
    "she", "should", "so", "that", "the", "their", "them", "there", "they", "this",
    "to", "us", "was", "we", "were", "what", "when", "which", "who", "will", "with",
    "you", "your", "write", "essay", "about", "whether", "more", "most", "very",
}
TRANSITIONS = {
    "first", "firstly", "second", "secondly", "finally", "however", "therefore",
    "moreover", "furthermore", "besides", "although", "instead", "meanwhile",
    "consequently", "in addition", "for example", "for instance", "in conclusion",
    "to sum up", "on the one hand", "on the other hand", "as a result",
}
ADVANCED_WORDS = {
    "beneficial", "significant", "essential", "efficient", "convenient", "sustainable",
    "perspective", "responsibility", "opportunity", "challenge", "consequence", "motivate",
    "participate", "independent", "environment", "communication", "technology", "academic",
    "appropriate", "considerable", "potential", "nevertheless", "increasingly", "crucial",
    "maintain", "improve", "contribute", "balance", "effective", "reasonable", "development",
}


@dataclass
class Issue:
    category: str
    level: str
    message: str
    original: str
    suggestion: str
    sentence: int | None = None


COMMON_REPLACEMENTS = [
    (re.compile(r"\bI am agree\b", re.I), "I agree", "语法", "agree 是动词，前面不使用 am"),
    (re.compile(r"\bpeople is\b", re.I), "people are", "主谓一致", "people 表示复数，应使用 are"),
    (re.compile(r"\bstudents is\b", re.I), "students are", "主谓一致", "复数主语 students 应搭配 are"),
    (re.compile(r"\bthey is\b", re.I), "they are", "主谓一致", "they 应搭配 are"),
    (re.compile(r"\bhe have\b", re.I), "he has", "主谓一致", "第三人称单数应使用 has"),
    (re.compile(r"\bshe have\b", re.I), "she has", "主谓一致", "第三人称单数应使用 has"),
    (re.compile(r"\bit have\b", re.I), "it has", "主谓一致", "第三人称单数应使用 has"),
    (re.compile(r"\bdoesn't has\b", re.I), "doesn't have", "动词形式", "助动词 does 后使用动词原形"),
    (re.compile(r"\bdidn't went\b", re.I), "didn't go", "动词形式", "助动词 did 后使用动词原形"),
    (re.compile(r"\bcan to\b", re.I), "can", "动词形式", "情态动词 can 后直接接动词原形"),
    (re.compile(r"\bmore better\b", re.I), "better", "比较级", "better 已经是比较级"),
    (re.compile(r"\bmuch students\b", re.I), "many students", "限定词", "可数名词复数使用 many"),
    (re.compile(r"\bmany information\b", re.I), "much information", "限定词", "information 是不可数名词"),
    (re.compile(r"\badvice(s)?\b", re.I), "advice", "词形", "advice 通常作不可数名词"),
    (re.compile(r"\binformations\b", re.I), "information", "词形", "information 是不可数名词"),
    (re.compile(r"\bhomeworks\b", re.I), "homework", "词形", "homework 是不可数名词"),
    (re.compile(r"\benvironmently\b", re.I), "environmentally", "拼写", "副词应拼写为 environmentally"),
    (re.compile(r"\bdefinately\b", re.I), "definitely", "拼写", "正确拼写为 definitely"),
    (re.compile(r"\brecieve\b", re.I), "receive", "拼写", "正确拼写为 receive"),
    (re.compile(r"\bwich\b", re.I), "which", "拼写", "正确拼写为 which"),
    (re.compile(r"\bteh\b", re.I), "the", "拼写", "正确拼写为 the"),
]


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def words(text: str) -> list[str]:
    return [item.lower() for item in WORD_RE.findall(text)]


def sentences(text: str) -> list[str]:
    return [m.group(0).strip() for m in SENTENCE_RE.finditer(text) if m.group(0).strip()]


def prompt_keywords(prompt: str) -> set[str]:
    return {word for word in words(prompt) if word not in STOP_WORDS and len(word) > 2}


def _sentence_number(text: str, position: int) -> int:
    return max(1, len(re.findall(r"[.!?]+", text[:position])) + 1)


def detect_issues(text: str) -> list[Issue]:
    found: list[Issue] = []
    for pattern, replacement, category, reason in COMMON_REPLACEMENTS:
        for match in pattern.finditer(text):
            found.append(Issue(category, "中", reason, match.group(0), replacement, _sentence_number(text, match.start())))

    for index, sentence in enumerate(sentences(text), 1):
        stripped = sentence.strip()
        first = re.search(r"[A-Za-z]", stripped)
        if first and first.group(0).islower():
            found.append(Issue("大小写", "中", "句首字母应大写", stripped[:28], stripped[: first.start()] + first.group(0).upper() + stripped[first.end() :], index))
        if stripped and stripped[-1] not in ".!?":
            found.append(Issue("标点", "低", "完整句末应添加标点", stripped[-28:], stripped + ".", index))

    for match in re.finditer(r"\bi\b", text):
        found.append(Issue("大小写", "中", "第一人称代词 I 必须大写", match.group(0), "I", _sentence_number(text, match.start())))
    for match in re.finditer(r"\b([A-Za-z]+)\s+\1\b", text, re.I):
        found.append(Issue("重复", "低", "相邻单词重复", match.group(0), match.group(1), _sentence_number(text, match.start())))
    for match in re.finditer(r"\s+([,.;!?])", text):
        found.append(Issue("标点", "低", "英文标点前不留空格", match.group(0), match.group(1), _sentence_number(text, match.start())))
    for match in re.finditer(r"([!?.,])\1{1,}", text):
        found.append(Issue("标点", "低", "避免连续使用相同标点", match.group(0), match.group(1), _sentence_number(text, match.start())))
    for match in re.finditer(r"\b(a)\s+([aeiou][A-Za-z]*)", text, re.I):
        found.append(Issue("冠词", "中", "元音音素开头的词前通常使用 an", match.group(0), f"an {match.group(2)}", _sentence_number(text, match.start())))
    for match in re.finditer(r"\b(an)\s+([bcdfgjklmnpqrstvwxyz][A-Za-z]*)", text, re.I):
        found.append(Issue("冠词", "中", "辅音音素开头的词前通常使用 a", match.group(0), f"a {match.group(2)}", _sentence_number(text, match.start())))

    # 同一位置可能被多个宽泛规则命中，保留类别、原文和句号唯一的记录。
    unique: dict[tuple[str, str, int | None], Issue] = {}
    for issue in found:
        unique[(issue.category, issue.original.lower(), issue.sentence)] = issue
    return list(unique.values())[:30]


def corrected_text(text: str) -> str:
    result = text
    for pattern, replacement, _, _ in COMMON_REPLACEMENTS:
        result = pattern.sub(replacement, result)
    result = re.sub(r"\bi\b", "I", result)
    result = re.sub(r"\b([A-Za-z]+)\s+\1\b", r"\1", result, flags=re.I)
    result = re.sub(r"\s+([,.;!?])", r"\1", result)
    result = re.sub(r"([!?.,])\1{1,}", r"\1", result)
    result = re.sub(r"\b(a)\s+([aeiou][A-Za-z]*)", r"an \2", result, flags=re.I)
    result = re.sub(r"\b(an)\s+([bcdfgjklmnpqrstvwxyz][A-Za-z]*)", r"a \2", result, flags=re.I)
    parts = re.split(r"([.!?]+\s+|\n+)", result)
    for index in range(0, len(parts), 2):
        parts[index] = re.sub(r"^([\s\"']*)([a-z])", lambda m: m.group(1) + m.group(2).upper(), parts[index])
    result = "".join(parts).strip()
    if result and result[-1] not in ".!?":
        result += "."
    return result


def _content_score(prompt: str, essay_words: list[str], sentence_count: int) -> tuple[float, float, list[str]]:
    keys = prompt_keywords(prompt)
    essay_set = set(essay_words)
    overlap = len(keys & essay_set) / max(1, len(keys))
    score = 7.0 + 13.0 * overlap
    lower = " ".join(essay_words)
    evidence_markers = sum(marker in lower for marker in ("for example", "for instance", "such as", "because", "this shows"))
    score += min(3.0, evidence_markers * 1.2)
    if sentence_count >= 5:
        score += 2.0
    notes = [f"题目关键词覆盖率 {overlap * 100:.0f}%"]
    if overlap < 0.35:
        notes.append("与题目核心词呼应不足，建议在主题句中直接回应题目")
    if evidence_markers == 0:
        notes.append("缺少例证或原因说明，可加入 for example / because")
    return clamp(score, 0, 25), overlap, notes


def _organization_score(text: str, sent: list[str]) -> tuple[float, list[str], list[str]]:
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    lower = text.lower()
    used = sorted(marker for marker in TRANSITIONS if re.search(rf"\b{re.escape(marker)}\b", lower))
    score = 3.0
    score += min(5.0, len(paragraphs) * 1.7)
    score += min(6.0, len(used) * 1.3)
    if sent and re.search(r"\b(I think|In my opinion|Nowadays|It is (clear|important))\b", sent[0], re.I):
        score += 2.5
    if sent and re.search(r"\b(in conclusion|to sum up|therefore|all in all)\b", sent[-1], re.I):
        score += 2.5
    lengths = [len(words(item)) for item in sent] or [0]
    if len(sent) >= 5 and max(lengths) <= 35:
        score += 1.0
    notes: list[str] = [f"共 {len(paragraphs)} 段，识别到 {len(used)} 种衔接表达"]
    if len(paragraphs) < 3:
        notes.append("建议按“引言—主体—结论”划分至少三段")
    if len(used) < 2:
        notes.append("衔接词偏少，可使用 however、moreover、therefore 等")
    return clamp(score, 0, 20), used, notes


def _vocabulary_score(tokens: list[str]) -> tuple[float, float, list[str]]:
    content = [word for word in tokens if word not in STOP_WORDS]
    unique = set(content)
    diversity = len(unique) / max(1, len(content))
    advanced = sorted(unique & ADVANCED_WORDS)
    long_ratio = sum(len(word) >= 8 for word in content) / max(1, len(content))
    repeated = Counter(content).most_common(1)[0][1] / max(1, len(content)) if content else 1.0
    score = 5 + diversity * 7 + min(3.0, len(advanced) * 0.75) + min(2.0, long_ratio * 12)
    score += clamp((0.14 - repeated) * 18, 0, 3)
    notes = [f"实词多样度 {diversity * 100:.0f}%，高级词汇 {len(advanced)} 个"]
    if diversity < 0.48:
        notes.append("词汇重复较明显，可替换高频词并使用同义表达")
    if not advanced:
        notes.append("可适量加入更准确的学术或主题词汇")
    return clamp(score, 0, 20), diversity, notes


def _grammar_score(issues: Iterable[Issue], sentence_count: int) -> tuple[float, list[str]]:
    issue_list = list(issues)
    weights = {"高": 3.0, "中": 2.0, "低": 1.0}
    penalty = sum(weights.get(item.level, 1.0) for item in issue_list)
    density = len(issue_list) / max(1, sentence_count)
    score = clamp(25 - penalty, 0, 25)
    notes = [f"检测到 {len(issue_list)} 处可解释问题，平均每句 {density:.2f} 处"]
    if issue_list:
        notes.append("建议优先修正中等级问题，再统一检查大小写与标点")
    else:
        notes.append("未触发内置规则，仍建议人工复核复杂句法")
    return score, notes


def _length_score(word_count: int) -> tuple[float, str]:
    if 120 <= word_count <= 260:
        return 10.0, "篇幅位于建议区间（120—260 词）"
    if 90 <= word_count < 120 or 260 < word_count <= 320:
        return 7.0, "篇幅接近建议区间，可适当调整"
    if 60 <= word_count < 90 or 320 < word_count <= 380:
        return 4.0, "篇幅与建议区间差距较大"
    return 1.0, "篇幅明显不足或过长，建议控制在 120—260 词"


def grade_essay(prompt: str, essay: str, title: str = "") -> dict[str, object]:
    prompt = (prompt or title).strip()
    essay = essay.strip()
    if len(essay) < 20:
        raise ValueError("作文内容过短，请至少输入 20 个字符")
    tokens = words(essay)
    sent = sentences(essay)
    if len(tokens) < 5:
        raise ValueError("有效英文单词过少，请输入完整作文")

    issue_list = detect_issues(essay)
    content, relevance, content_notes = _content_score(prompt, tokens, len(sent))
    organization, transitions, organization_notes = _organization_score(essay, sent)
    vocabulary, diversity, vocabulary_notes = _vocabulary_score(tokens)
    grammar, grammar_notes = _grammar_score(issue_list, len(sent))
    length_score, length_note = _length_score(len(tokens))
    dimensions = {
        "内容与切题": round(content, 1),
        "结构与衔接": round(organization, 1),
        "词汇运用": round(vocabulary, 1),
        "语法与规范": round(grammar, 1),
        "篇幅与格式": round(length_score, 1),
    }
    total = round(sum(dimensions.values()), 1)
    if total >= 90:
        band, summary = "A", "主题明确、表达成熟，已达到优秀水平"
    elif total >= 80:
        band, summary = "B", "整体完成度较好，少量细节仍可优化"
    elif total >= 70:
        band, summary = "C", "基本完成写作任务，结构或语言需要加强"
    elif total >= 60:
        band, summary = "D", "能够表达主要意思，但问题较集中"
    else:
        band, summary = "E", "内容与语言基础仍需重点完善"

    recommendations: list[str] = []
    ranked = sorted(dimensions.items(), key=lambda item: item[1] / {"内容与切题": 25, "结构与衔接": 20, "词汇运用": 20, "语法与规范": 25, "篇幅与格式": 10}[item[0]])
    advice_map = {
        "内容与切题": "先用一句清晰的中心论点回应题目，再补充两条理由或例证。",
        "结构与衔接": "采用三段式结构，并用衔接词明确展示观点之间的关系。",
        "词汇运用": "替换重复词，优先选择准确自然的搭配，避免为复杂而复杂。",
        "语法与规范": "逐句检查主谓一致、动词形式、冠词、大小写和句末标点。",
        "篇幅与格式": "将全文调整到 120—260 词，并用空行划分自然段。",
    }
    for name, _ in ranked[:3]:
        recommendations.append(advice_map[name])

    return {
        "title": title.strip(),
        "prompt": prompt,
        "score": total,
        "band": band,
        "summary": summary,
        "dimensions": dimensions,
        "statistics": {
            "words": len(tokens),
            "sentences": len(sent),
            "paragraphs": len([p for p in re.split(r"\n\s*\n", essay) if p.strip()]),
            "relevance": round(relevance * 100, 1),
            "vocabulary_diversity": round(diversity * 100, 1),
            "transitions": transitions,
            "issue_count": len(issue_list),
        },
        "notes": {
            "内容与切题": content_notes,
            "结构与衔接": organization_notes,
            "词汇运用": vocabulary_notes,
            "语法与规范": grammar_notes,
            "篇幅与格式": [length_note],
        },
        "issues": [asdict(item) for item in issue_list],
        "corrected_text": corrected_text(essay),
        "recommendations": recommendations,
        "disclaimer": "本结果由离线规则与文本统计生成，适合学习反馈，不替代教师最终评分。",
    }

