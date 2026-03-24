"""개인정보 스캔 — 문서 텍스트에서 민감 정보를 정규식으로 감지.

지원 감지 항목:
- 주민등록번호 (risk: critical)
- 전화번호 (risk: high)
- 이메일 (risk: medium)
- 계좌번호 (risk: high)
- 여권번호 (risk: high)

패턴 매칭 순서가 중요: 주민번호 → 전화번호 → 이메일 → 계좌번호 → 여권번호
이미 매칭된 위치는 후속 패턴에서 제외하여 오탐 방지.
"""
import re


# 매칭 순서가 우선순위: 먼저 매칭된 것이 확정, 겹치는 위치는 스킵
_PATTERNS = [
    {
        "type": "주민등록번호",
        "pattern": r"\b(\d{6})\s*[-–]\s*([1-4]\d{6})\b",
        "risk": "critical",
        "mask": lambda m: m.group(1) + "-" + m.group(2)[0] + "******",
    },
    {
        "type": "전화번호",
        "pattern": r"\b(0\d{1,2})[-.\s]?(\d{3,4})[-.\s]?(\d{4})\b",
        "risk": "high",
        "mask": lambda m: m.group(1) + "-" + "****" + "-" + m.group(3),
    },
    {
        "type": "이메일",
        "pattern": r"\b([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\b",
        "risk": "medium",
        "mask": lambda m: m.group(1)[:2] + "***@" + m.group(2),
    },
    {
        "type": "계좌번호",
        "pattern": r"\b(\d{3,6})[-](\d{2,6})[-](\d{4,8})\b",
        "risk": "high",
        "mask": lambda m: m.group(1) + "-****-" + m.group(3)[-2:],
    },
    {
        "type": "여권번호",
        "pattern": r"\b([A-Z]{1,2})(\d{7,8})\b",
        "risk": "high",
        "mask": lambda m: m.group(1) + "*" * len(m.group(2)),
    },
]


def _ranges_overlap(start1, end1, start2, end2):
    """두 범위가 겹치는지 확인."""
    return start1 < end2 and start2 < end1


def scan_privacy(text):
    """텍스트에서 개인정보를 스캔하여 결과 반환.

    패턴 우선순위 순서로 매칭하며, 이미 매칭된 위치와 겹치는 후속 매칭은 제외.

    Returns: {
        "found": bool,
        "total_findings": int,
        "findings": [{type, value, masked_value, risk, position}, ...],
        "risk_summary": {critical: N, high: N, medium: N, low: N},
    }
    """
    if not isinstance(text, str) or not text:
        return {"found": False, "total_findings": 0, "findings": [],
                "risk_summary": {"critical": 0, "high": 0, "medium": 0, "low": 0},
                "recommendation": "검사할 텍스트가 없습니다."}

    findings = []
    risk_summary = {"critical": 0, "high": 0, "medium": 0, "low": 0}
    matched_ranges = []  # (start, end) 튜플 목록

    for pat_info in _PATTERNS:
        for m in re.finditer(pat_info["pattern"], text):
            start, end = m.start(), m.end()

            # 이미 매칭된 범위와 겹치면 스킵
            if any(_ranges_overlap(start, end, s, e) for s, e in matched_ranges):
                continue

            matched_ranges.append((start, end))
            masked = pat_info["mask"](m)
            findings.append({
                "type": pat_info["type"],
                "value": masked,  # 원본 대신 마스킹된 값만 노출
                "masked_value": masked,
                "risk": pat_info["risk"],
                "position": start,
            })
            risk_summary[pat_info["risk"]] = risk_summary.get(pat_info["risk"], 0) + 1

    # 위치 순 정렬
    findings.sort(key=lambda f: f["position"])

    recommendation = ""
    if risk_summary["critical"] > 0:
        recommendation = "주민등록번호가 포함되어 있습니다. 즉시 마스킹 처리가 필요합니다."
    elif risk_summary["high"] > 0:
        recommendation = "민감 개인정보가 포함되어 있습니다. 마스킹을 권장합니다."
    elif findings:
        recommendation = "개인정보가 일부 포함되어 있습니다. 확인이 필요합니다."
    else:
        recommendation = "개인정보가 감지되지 않았습니다."

    return {
        "found": len(findings) > 0,
        "total_findings": len(findings),
        "findings": findings,
        "risk_summary": risk_summary,
        "recommendation": recommendation,
    }
