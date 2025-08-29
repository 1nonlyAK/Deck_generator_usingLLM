import os
import json
import re
from typing import List, Dict, Any, Optional

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from groq import Groq


# ---------------- Web Fetch Helper ---------------- #

def fetch_web_facts(topic: str, num_results: int = 3) -> list:
    url = "https://duckduckgo.com/html/"
    params = {"q": topic}
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, params=params, headers=headers, timeout=10)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    results = []
    for result in soup.find_all("div", class_="result", limit=num_results):
        title_elem = result.find("a", class_="result__a")
        snippet_elem = result.find("a", class_="result__snippet")
        if title_elem and snippet_elem:
            results.append(f"{title_elem.get_text(strip=True)}: {snippet_elem.get_text(strip=True)}")
    return results


# ---------------- JSON Helper ---------------- #

def safe_json_parse(raw: str) -> dict:
    if not raw or not isinstance(raw, str):
        return {"title": "Parsing Failed", "overview": "", "slides": [], "conclusion": ""}

    try:
        parsed = json.loads(raw)
        if isinstance(parsed, str) and parsed.strip().startswith("{"):
            return json.loads(parsed)
        return parsed
    except Exception:
        match = re.search(r"\{.*\}", raw, re.S)
        if match:
            candidate = re.sub(r",(\s*[}\]])", r"\1", match.group(0))
            try:
                return json.loads(candidate)
            except Exception:
                pass
        return {"title": "Parsing Failed", "overview": raw[:500], "slides": [], "conclusion": ""}


# ---------------- LLM Client ---------------- #

def _init_groq_client() -> Groq:
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY is not set. Add it to your .env file.")
    return Groq(api_key=api_key)


def _call_llm_and_parse(client, system: str, user: str, model: str, max_tokens: int = 2000) -> Dict[str, Any]:
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user + "\n\nReturn ONLY valid JSON."}],
        temperature=0.4,
        max_tokens=max_tokens,
    )
    try:
        raw = resp.choices[0].message.content.strip()
    except Exception as e:
        return {"title": "Parsing Failed", "overview": f"Bad LLM response: {e}", "slides": [], "conclusion": ""}
    return safe_json_parse(raw)


# ---------------- Polishing Pass ---------------- #

def _polish_content(client, draft: Dict[str, Any], model: str = "llama3-8b-8192", web_facts: Optional[list] = None) -> Dict[str, Any]:
    system = (
        "You are a senior editor at a consulting firm. "
        "Polish the JSON deck for clarity, conciseness, and tone. "
        "Keep schema identical: title, overview, slides[type,title,topics[subtitle,bullets,sources,chart,table]], conclusion. "
        "Each topic may optionally include 'chart' or 'table'. "
        "If included, keep data small (2–6 points). "
        "Every topic must have 'sources'."
    )

    fact_context = ""
    if web_facts:
        fact_context = "\nWeb facts you may cite:\n" + "\n".join([f"- {fact}" for fact in web_facts])

    user = f"""
Here is a draft slide deck in JSON.
Polish wording, ensure sources, and keep schema intact.
{fact_context}

{json.dumps(draft, indent=2)}
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        temperature=0.2,
        max_tokens=1800,
    )
    try:
        raw = resp.choices[0].message.content.strip()
    except Exception:
        return draft

    match = re.search(r"\{.*\}", raw, re.S)
    if match:
        raw = match.group(0)

    polished = safe_json_parse(raw)

    # Ensure sources exist everywhere
    if isinstance(polished, dict) and "slides" in polished:
        for slide in polished["slides"]:
            topics = slide.get("topics", [])
            if isinstance(topics, dict):
                topics = [topics]
            elif not isinstance(topics, list):
                topics = []
            for t in topics:
                if "sources" not in t or not isinstance(t["sources"], list):
                    t["sources"] = web_facts[:2] if web_facts else ["General industry reports"]
            slide["topics"] = topics

    return polished

    # 🔧 Auto-fix chart schemas
    if isinstance(polished, dict) and "slides" in polished:
        for slide in polished["slides"]:
            for topic in slide.get("topics", []):
                if "chart" in topic:
                    chart = topic["chart"]
                    values = chart.get("values", [])
                    categories = chart.get("categories", [])

                    # If values exist but categories missing → generate Item 1..N
                    if values and not categories:
                        chart["categories"] = [f"Item {i+1}" for i in range(len(values))]

                    # If lengths don’t match → trim to min length
                    if categories and values and len(categories) != len(values):
                        n = min(len(categories), len(values))
                        chart["categories"] = categories[:n]
                        chart["values"] = values[:n]

                    topic["chart"] = chart


# ---------------- Content Generator ---------------- #

def generate_content(
    topic: str,
    *,
    model: str = "llama3-8b-8192",
    depth: int = 3,
    web_facts: Optional[list] = None
) -> Dict[str, Any]:
    load_dotenv()
    if not isinstance(topic, str) or not topic.strip():
        raise ValueError("Topic must be a non-empty string.")
    topic = topic.strip()
    client = _init_groq_client()

    context = ""
    if web_facts:
        context = "\nRecent facts:\n" + "\n".join([f"- {fact}" for fact in web_facts]) + "\n"

    system = (
        "You are a senior strategy consultant. "
        "Output ONLY valid JSON."
    )

    user = f"""
Create a professional 8–12 slide business deck.

Topic: {topic}

{context}

Rules:
- Each slide has 2–4 topics.
- Each topic has: "subtitle", "bullets", "sources".
- Optionally include a small "chart" or "table" if relevant.
- Charts:
   - Must include BOTH "categories" and "values".
   - categories and values must have the same length.
   - Use 2–6 data points maximum.
- Tables:
   - Must include "headers" and 2–5 rows.
   - Keep rows short and numeric/textual.

Schema:
{{
  "title": "...",
  "overview": "...",
  "slides": [
    {{
      "type": "Market Trends",
      "title": "...",
      "topics": [
        {{
          "subtitle": "...",
          "bullets": ["..."],
          "sources": ["..."],
          "chart": {{
            "type": "bar",
            "title": "Example",
            "categories": ["2019","2020","2021"],
            "values": [2.5,3.1,4.0]
          }},
          "table": {{
            "headers": ["Region","Sales"],
            "rows": [["NA","1200"],["EU","950"]]
          }}
        }}
      ]
    }}
  ],
  "conclusion": "..."
}}
"""

    draft = _call_llm_and_parse(client, system, user, model=model)

    polished = _polish_content(client, draft, model=model, web_facts=web_facts)
    return polished
