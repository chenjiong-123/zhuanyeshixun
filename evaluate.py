"""在带参考分的测试集上评估评分稳定性。"""

from __future__ import annotations

import json
from pathlib import Path

from grader import grade_essay
from repository import load_samples


def score_band(score: float) -> str:
    if score >= 90:
        return "A"
    if score >= 80:
        return "B"
    if score >= 70:
        return "C"
    if score >= 60:
        return "D"
    return "E"


def evaluate() -> dict[str, object]:
    rows = []
    errors = []
    band_hits = 0
    for sample in load_samples():
        result = grade_essay(str(sample["prompt"]), str(sample["essay"]), str(sample["title"]))
        predicted = float(result["score"])
        reference = float(sample["reference_score"])
        error = abs(predicted - reference)
        errors.append(error)
        expected_band = score_band(reference)
        band_hits += result["band"] == expected_band
        rows.append({"id": sample["id"], "title": sample["title"], "reference": reference, "predicted": predicted, "absolute_error": round(error, 1), "expected_band": expected_band, "predicted_band": result["band"]})
    metrics = {
        "sample_count": len(rows),
        "mae": round(sum(errors) / max(1, len(errors)), 2),
        "within_10_points": round(sum(error <= 10 for error in errors) / max(1, len(errors)) * 100, 1),
        "band_accuracy": round(band_hits / max(1, len(rows)) * 100, 1),
    }
    return {"metrics": metrics, "details": rows}


if __name__ == "__main__":
    report = evaluate()
    target = Path(__file__).resolve().parent / "docs" / "evaluation_results.json"
    target.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"评估结果已保存：{target}")

