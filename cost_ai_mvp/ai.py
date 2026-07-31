from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import Any


def generate_openai_compatible(base_url: str, api_key: str, model: str, prompt: str, temperature: float = 0.2, timeout: int = 60) -> str:
    url = base_url.rstrip("/") + "/v1/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "你是工程造价数据助手。只基于用户给出的结构化摘要解释，不编造价格。"},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result: dict[str, Any] = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise RuntimeError(f"AI 接口调用失败：{exc}") from exc
    return result["choices"][0]["message"]["content"]

