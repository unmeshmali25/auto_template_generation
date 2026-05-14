"""
Template parser: extracts placeholder structure from the Word template.
For Phase 1 we use a hard-coded schema derived from the Google template images.
"""
from typing import Dict, List, Any


# Hard-coded template schema based on the 4 analyzed images.
# In Phase 2 this could be auto-extracted from the .docx.
TEMPLATE_SCHEMA: Dict[str, Any] = {
    "sections": {
        "topic": {
            "description": "Title of the communications recommendation (replaces [Topic to be Inserted])",
            "type": "text",
        },
        "audience": {
            "description": "Target audience description",
            "type": "text",
        },
        "support_channels": {
            "description": "List of support channels (Social / Community / Voice Channel / Chat / Email)",
            "type": "list",
            "default": ["Social", "Community", "Voice Channel", "Chat", "Email"],
        },
        "situation_overview": {
            "description": "High-level summary of the issue or launch",
            "type": "paragraph",
        },
        "approach": {
            "description": "Strategic approach and response plan",
            "type": "paragraph",
        },
        "materials": {
            "description": "List of materials needed (e.g., FAQ, Social post, Talking points)",
            "type": "paragraph",
        },
        "messaging_statement": {
            "description": "Official messaging statement",
            "type": "paragraph",
        },
        "social_messaging": {
            "description": "Social media friendly messaging",
            "type": "paragraph",
        },
        "talking_points": {
            "description": "Bullet points for spokespeople",
            "type": "bullets",
        },
        "faqs": {
            "description": "List of FAQ objects with question and answer",
            "type": "faqs",
        },
        "distribution": {
            "description": "Distribution information table with checkboxes",
            "type": "distribution_table",
            "source_channels": [
                "Social",
                "Voice Call",
                "Chat",
                "Community",
                "Repair Centers",
                "Global / HQ",
            ],
            "distribution_channels": [
                "Social/Community",
                "Reactive (DM)",
                "Proactive",
                "Voice Call",
                "Chat",
                "Repair Centers",
                "Carrier",
                "Retail",
                "E-Com CS",
            ],
            "distribution_methods": [
                "SSAM (AI recommendation engine)",
                "Training (Classroom)",
                "Support S.com (Customer Facing)",
                "Vendor Huddle (Conf Call)",
                "Alert Communication (Email)",
                "System Modification (Pop-Up)",
                "IVR Message (1800 GOOGLE)",
                "Retail Communication (FSM)",
            ],
        },
        "status_date": {
            "description": "Date for status tracking (today's date)",
            "type": "date",
        },
        "resource_links": {
            "description": "Relevant resource links",
            "type": "text",
        },
        "approval_status": {
            "description": "Approval checkboxes",
            "type": "approval_table",
            "options": [
                "CSD Internal",
                "Legal",
                "PR",
                "Product Mkt",
                "Product Mgmt",
                "HQ PR",
            ],
        },
        "assessment_questions": {
            "description": "Assessment questions for the distribution table",
            "type": "faqs",
        },
    }
}


def get_schema() -> Dict[str, Any]:
    """Return the template schema."""
    return TEMPLATE_SCHEMA
