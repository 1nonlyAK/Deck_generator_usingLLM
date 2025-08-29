import argparse
import sys
from dotenv import load_dotenv

from content_generator import generate_content, fetch_web_facts
from presentation_engine import create_presentation


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Generate a PowerPoint presentation using LLM + Groq.")
    parser.add_argument("topic", type=str, help="Presentation topic")
    parser.add_argument("--model", type=str, default="llama3-8b-8192")
    parser.add_argument("--depth", type=int, default=3)
    parser.add_argument("-o", "--output", type=str, default=None)
    parser.add_argument("--no-web", dest="no_web", action="store_true", help="Disable web fact fetching")
    args = parser.parse_args()

    topic = args.topic.strip()
    if not topic:
        print("Error: empty topic")
        return 2

    output = args.output or f"{topic.lower().replace(' ', '_')}_presentation.pptx"

    try:
        print(f"[1/3] Fetching web facts for: {topic!r}")
        web_facts = [] if args.no_web else fetch_web_facts(topic, num_results=3)
        print(f"Retrieved {len(web_facts)} facts.")

        print("[2/3] Generating content...")
        content = generate_content(topic, model=args.model, depth=args.depth, web_facts=web_facts)
        print("Content generated.")

        print(f"[3/3] Building slide deck: {output}")
        create_presentation(content, output)
        print(f"✔ Done. Saved: {output}")
        return 0
    except Exception as e:
        print(f"ERROR: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
