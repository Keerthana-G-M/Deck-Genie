import re
import os
from dotenv import load_dotenv
from io import BytesIO
from typing import Dict, List, Tuple

# Load environment variables
load_dotenv()

# Constants for presentation design
FONTS = {
    "title": "Helvetica Neue",
    "body": "Helvetica Neue",
    "fallback1": "Segoe UI",
    "fallback2": "Arial"
}

COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "blue": "0078D7",
    "light_blue": "B3D7F2",  # Added light blue color
    "red": "E74C3C", 
    "green": "2ECC71",
    "gray": "D3D3D3",
    "light_gray": "F5F5F5",  # Added light gray for subtle backgrounds
    "accent": "4CAF50",  # Added accent color (green)
    "accent_dark": "388E3C"  # Added darker accent color (dark green)
}

FONT_SIZES = {
    "title": 44,
    "title_slide": 48,
    "subtitle": 24,
    "body": 24,
    "call_to_action": 48
}

# Icons for features and slides
FEATURE_ICONS = [
    "✓",  # Checkmark
    "⚙️",  # Gear
    "📊",  # Chart
    "👤",  # User
    "🔒",  # Lock
    "⚡",  # Lightning
    "🔄"   # Sync
]

SLIDE_ICONS = {
    "problem": "⚠️",  # Warning
    "solution": "💡",  # Light bulb
    "features": "⚙️",  # Gear
    "advantage": "🏆",  # Trophy
    "audience": "👥",  # People
    "call_to_action": "🚀",  # Rocket
    "market_slide": "💰", # Market
    "roadmap_slide": "🗓️", # Roadmap
    "team_slide": "🧑‍🤝‍🧑" # Team
}

# Image search keywords for each slide type
IMAGE_KEYWORDS = {
    "problem": ["business challenge", "corporate problem", "business obstacle", "productivity problem", "enterprise difficulty", "digital transformation challenge"],
    "solution": ["business innovation", "enterprise solution", "digital solution", "business technology", "tech innovation", "corporate solution", "saas solution"],
    "features": ["software interface", "technology dashboard", "business analytics", "enterprise software", "saas product", "tech platform", "digital tool"],
    "advantage": ["business growth", "competitive edge", "market leadership", "business success", "enterprise advantage", "performance chart", "business strategy"],
    "audience": ["corporate meeting", "business professionals", "executive team", "business conference", "professional team meeting", "enterprise clients", "b2b meeting"],
    "market analysis": ["market research", "business analytics", "financial growth", "market trends", "data charts"], # Added for market_slide
}

# Layout measurements (in inches)
MARGINS = {
    "left": 1.0,   # Increased to meet 1-inch minimum margin requirement
    "right": 1.0,  # Increased to meet 1-inch minimum margin requirement
    "top": 0.5,
    "bottom": 0.5
}

CONTENT_AREA = {
    "left": MARGINS["left"],
    "top": 2.0,  # Proper spacing below title
    "width": 6.0,  # Increased width for better text layout in 16:9 format
    "height": 4.5  # Reduced height to prevent overlap with bottom margin
}

IMAGE_AREA = {
    "left": 7.5,  # Moved further right to prevent overlap with content
    "top": 2.0,   # Aligned with content area
    "width": 4.8,  # Adjusted for 16:9 format while respecting right margin
    "height": 4.5   # Matching content area height
}


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a string to be used as a filename.
    
    Args:
        filename: The original string
        
    Returns:
        A sanitized version safe for use as a filename
    """
    # Replace spaces with underscores and remove special characters
    sanitized = re.sub(r'[^\w\s-]', '', filename)
    sanitized = re.sub(r'[\s]+', '_', sanitized)
    return sanitized


def truncate_text(text: str, max_length: int, suffix: str = "...") -> str:
    """
    Truncate text if it exceeds the maximum length.
    
    Args:
        text: The text to truncate
        max_length: Maximum allowed length
        suffix: String to append if truncated
        
    Returns:
        Truncated text with suffix if needed
    """
    if len(text) <= max_length:
        return text
    return text[:max_length-len(suffix)] + suffix


def get_api_key(key_name: str) -> str:
    """
    Get API key from environment variables.
    
    Args:
        key_name: Name of the API key in .env file
        
    Returns:
        API key as string
        
    Raises:
        ValueError: If API key is not found
    """
    api_key = os.getenv(key_name)
    if not api_key:
        raise ValueError(f"{key_name} not found in environment variables")
    return api_key


def match_icon_to_feature(feature) -> str:
    """
    Match an appropriate icon to a feature based on content.
    
    Args:
        feature: The feature text (string) or feature object (dict)
        
    Returns:
        An icon character from FEATURE_ICONS
    """
    # Handle both string and dictionary formats
    if isinstance(feature, dict):
        # Extract feature text from dictionary format
        feature_text = feature.get('feature', feature.get('name', feature.get('title', str(feature))))
    else:
        # Handle string format
        feature_text = str(feature)
    
    feature_lower = feature_text.lower()
    
    # Map features to appropriate icons based on keywords
    if any(keyword in feature_lower for keyword in ["secure", "privacy", "compliance", "protect"]):
        return FEATURE_ICONS[4]  # Lock
    elif any(keyword in feature_lower for keyword in ["fast", "speed", "quick", "rapid"]):
        return FEATURE_ICONS[5]  # Lightning
    elif any(keyword in feature_lower for keyword in ["sync", "integrate", "connect"]):
        return FEATURE_ICONS[6]  # Sync
    elif any(keyword in feature_lower for keyword in ["data", "analytics", "report", "insight"]):
        return FEATURE_ICONS[2]  # Chart
    elif any(keyword in feature_lower for keyword in ["user", "customer", "experience", "interface"]):
        return FEATURE_ICONS[3]  # User
    elif any(keyword in feature_lower for keyword in ["feature", "function", "capability", "tool"]):
        return FEATURE_ICONS[1]  # Gear
    else:
        return FEATURE_ICONS[0]  # Default to checkmark