"""
Main orchestrator: generate a communications recommendations document from raw inputs.

Usage:
    python src/generate_doc.py --case cases/google_health_app
    python src/generate_doc.py --case cases/google_health_app --client-type custom --model your-model-name

Environment:
    OPENROUTER_API_KEY must be set in environment or .env file (for OpenRouter client).
    YOUR_CLIENT_KEY, YOUR_PASS_KEY, ENDPOINT_URL, YOUR_EMAIL (for custom client).
"""
import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

# Add src directory to path for imports when running from project root
sys.path.insert(0, str(Path(__file__).resolve().parent))

from dotenv import load_dotenv

from llm_client import OpenRouterClient, CustomChatClient
from template_parser import get_schema
from doc_populator import populate_template, create_placeholder_template


def load_sources(case_dir: Path) -> str:
    """Read all .txt files from case_dir/sources/ and concatenate."""
    sources_dir = case_dir / "sources"
    if not sources_dir.exists():
        raise FileNotFoundError(f"Sources directory not found: {sources_dir}")

    texts = []
    for txt_file in sorted(sources_dir.glob("*.txt")):
        texts.append(txt_file.read_text(encoding="utf-8"))

    if not texts:
        raise ValueError(f"No .txt source files found in {sources_dir}")

    return "\n\n---\n\n".join(texts)


def load_images(case_dir: Path) -> list:
    """Find all image files in case_dir/images/. Deduplicate by preferring PNG/JPG over AVIF/WebP."""
    images_dir = case_dir / "images"
    if not images_dir.exists():
        return []
    extensions = {".png", ".jpg", ".jpeg", ".avif", ".gif", ".bmp", ".webp"}
    all_images = sorted([p for p in images_dir.iterdir() if p.suffix.lower() in extensions])

    # Deduplicate: group by stem (filename without extension), prefer PNG/JPG over AVIF/WebP
    by_stem: dict[str, Path] = {}
    for img_path in all_images:
        stem = img_path.stem.lower()
        if stem not in by_stem:
            by_stem[stem] = img_path
            continue
        existing = by_stem[stem]
        existing_ext = existing.suffix.lower()
        new_ext = img_path.suffix.lower()
        # Prefer PNG/JPEG over AVIF/WEBP/GIF/BMP
        preferred = {".png", ".jpg", ".jpeg"}
        if existing_ext not in preferred and new_ext in preferred:
            by_stem[stem] = img_path

    return sorted([str(p) for p in by_stem.values()])


def build_llm_prompt(raw_text: str, schema: dict) -> tuple:
    """Build system and user prompts for the LLM."""
    system_prompt = (
        "You are a communications expert assistant. Your task is to read raw incident or launch information "
        "and generate a structured communications recommendations document. "
        "You must output valid JSON only, with no extra commentary. "
        "If a field cannot be determined from the input, use the string 'N/A'. "
        "Use the exact keys specified in the schema."
    )

    schema_json = json.dumps(schema, indent=2)
    user_prompt = f"""TEMPLATE SCHEMA:
{schema_json}

RAW INPUT TEXT:
{raw_text}

INSTRUCTIONS:
1. Map content from the raw text to the template fields. Draft any missing sections (Approach, Materials, Messaging Statement, Social Messaging) based on context.
2. Extract FAQs from Q&A format in the text. If no FAQs are present, return an empty list.
3. talking_points should be a list of concise bullet strings.
4. For the distribution object, infer which checkboxes should be ticked based on the situation type and severity. Use the exact option strings from the schema.
5. For approval_status, infer which approvals are needed based on content sensitivity (e.g., data/privacy issues may need Legal; public-facing issues may need PR).
6. status_date should be today's date: {datetime.now().strftime('%Y-%m-%d')}.
7. resource_links: if no links are mentioned, use 'N/A'.
8. support_channels: default to ["Social", "Community", "Voice Channel", "Chat", "Email"] unless the text suggests otherwise.
9. audience: infer from the text (e.g., "Android users of Google Health App").

OUTPUT STRICT JSON with these exact top-level keys:
- topic
- audience
- support_channels (list of strings)
- situation_overview (string)
- approach (string)
- materials (string)
- messaging_statement (string)
- social_messaging (string)
- talking_points (list of strings)
- faqs (list of {{"question": string, "answer": string}})
- distribution (object with keys: source_channels, distribution_channels, distribution_methods, status_date, resource_links)
- approval_status (list of strings)
- assessment_questions (list of {{"question": string, "answer": string}})
"""
    return system_prompt, user_prompt


def main():
    parser = argparse.ArgumentParser(description="Generate communications recommendations document")
    parser.add_argument("--case", required=True, help="Path to case directory (e.g., cases/google_health_app)")
    parser.add_argument("--template", default=None, help="Path to template .docx (optional, will create placeholder if missing)")
    parser.add_argument("--output", default=None, help="Output file path (optional)")
    parser.add_argument("--model", default="openai/gpt-4o-mini", help="Model name to use for LLM extraction")
    parser.add_argument(
        "--client-type",
        choices=["openrouter", "custom"],
        default="openrouter",
        help="LLM client to use: openrouter (default) or custom",
    )
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent.parent
    load_dotenv(project_root / ".env")

    case_dir = Path(args.case).resolve()
    if not case_dir.exists():
        print(f"Error: case directory not found: {case_dir}")
        sys.exit(1)

    # Load inputs
    raw_text = load_sources(case_dir)
    images = load_images(case_dir)
    print(f"Loaded sources from {case_dir / 'sources'}")
    print(f"Found {len(images)} image(s)")

    # Template
    template_path = args.template
    if template_path is None:
        template_path = project_root / "template" / "communications_recommendations_template.docx"
        if not template_path.exists():
            print("Template not found, creating placeholder template...")
            template_path.parent.mkdir(parents=True, exist_ok=True)
            create_placeholder_template(str(template_path))
            print(f"Created placeholder template at {template_path}")

    template_path = Path(template_path).resolve()
    if not template_path.exists():
        print(f"Error: template not found: {template_path}")
        sys.exit(1)

    # Output path
    output_path = args.output
    if output_path is None:
        output_dir = project_root / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"communications_recommendations_{case_dir.name}.docx"
    else:
        output_path = Path(output_path).resolve()

    # LLM extraction
    if args.client_type == "custom":
        print("Calling Custom Chat LLM for structured extraction...")
        client = CustomChatClient()
    else:
        print("Calling OpenRouter LLM for structured extraction...")
        client = OpenRouterClient()
    schema = get_schema()
    system_prompt, user_prompt = build_llm_prompt(raw_text, schema)

    try:
        structured_data = client.extract_structured_json(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            model=args.model,
        )
    except Exception as e:
        print(f"LLM extraction failed: {e}")
        print("Falling back to mock data for the toy example...")
        structured_data = {
            "topic": "Google Health App Crash - Android v12.0.3",
            "audience": "Android users of Google Health App running version 12.0.3 or below",
            "support_channels": ["Social", "Community", "Voice Channel", "Chat", "Email"],
            "situation_overview": "The Google Health App crashes immediately on launch for Android users running version 12.0.3 or below. Affected users see a white screen for 2–3 seconds before the app force closes with no error message. The issue was first reported on April 28, 2026, and currently impacts approximately 14,000 devices. A hotfix is in development and expected to roll out by May 15, 2026.",
            "approach": "Incident response: engineering team is developing a hotfix expected by May 15, 2026. Customer support is communicating temporary workarounds. Social and community teams are monitoring for escalations. No data breach or privacy risk identified.",
            "materials": "FAQ document, social media response templates, customer service talking points, in-app notification copy",
            "messaging_statement": "We are aware of an issue affecting Google Health App on Android devices running version 12.0.3 or below. Our engineering team is actively working on a hotfix, expected to roll out by May 15, 2026. In the meantime, affected users can access their health data via health.google.com on a mobile browser. No health data has been compromised.",
            "social_messaging": "We're working on a fix for Android users experiencing app crashes. A hotfix is coming by May 15. Your data is safe — access it temporarily at health.google.com. Thanks for your patience!",
            "talking_points": [
                "Issue affects Android devices running Google Health App version 12.0.3 or below only",
                "iOS users are not impacted",
                "Hotfix expected by May 15, 2026, delivered via Play Store automatic update — no reinstall needed",
                "Temporary workaround: use health.google.com in mobile browser",
                "No health data at risk; crash is UI-layer only"
            ],
            "faqs": [
                {"question": "Which Android versions are affected?", "answer": "Only Android devices running Google Health App version 12.0.3 or below are impacted. iOS users are not affected."},
                {"question": "Is my health data safe?", "answer": "Yes. All health data is stored securely on Google's servers and is not at risk. The crash occurs at the UI layer and does not affect data integrity."},
                {"question": "What should I do while waiting for the fix?", "answer": "You can access your health data via health.google.com on a mobile browser as a temporary workaround."},
                {"question": "Will I need to reinstall the app after the update?", "answer": "No. The hotfix will be delivered as an automatic update through the Play Store. No reinstall is necessary."}
            ],
            "distribution": {
                "source_channels": ["Social", "Support S.com", "FAQ"],
                "distribution_channels": ["Social/Community", "Reactive (DM)", "Support Portal"],
                "distribution_methods": ["Support S.com", "Alert Communication (Email)", "Social Post"],
                "status_date": datetime.now().strftime("%Y-%m-%d"),
                "resource_links": "N/A"
            },
            "approval_status": ["CSD Internal", "PR"],
            "assessment_questions": [
                {"question": "Is this issue customer-facing?", "answer": "Yes, the app crash is immediately visible to users on launch."},
                {"question": "Are there any legal or compliance concerns?", "answer": "No data breach or privacy risk identified."},
            ],
        }

    # Debug: print extracted data
    print("\n--- Extracted Data ---")
    print(json.dumps(structured_data, indent=2))
    print("--- End Extracted Data ---\n")

    # Populate document
    print(f"Populating template: {template_path}")
    populate_template(
        template_path=str(template_path),
        output_path=str(output_path),
        data=structured_data,
        images=images,
    )

    print(f"\nSuccess! Generated document: {output_path}")


if __name__ == "__main__":
    main()
