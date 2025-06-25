import requests
from io import BytesIO
from PIL import Image
import base64
import os
import hashlib
from typing import Dict, Any, Optional, Union

# Image cache to avoid refetching images
_image_cache: Dict[str, Any] = {}

def get_image_from_unsplash(query: str, orientation: str = "landscape", use_cache: bool = True) -> Dict[str, Any]:
    """
    Get an image from Unsplash or from the cache if available.
    
    Args:
        query: Search query for the image
        orientation: Image orientation (landscape, portrait, squarish)
        use_cache: Whether to use the image cache
        
    Returns:
        Dictionary with image data (url, data, etc.)
    """
    # Create a cache key from the query and orientation
    cache_key = f"unsplash_{query}_{orientation}"
    
    # Check if we have this image in the cache
    if use_cache and cache_key in _image_cache:
        return _image_cache[cache_key]
    
    # Otherwise, fetch from Unsplash API (your existing code to fetch images)
    # This is just a placeholder - replace with your actual Unsplash API code
    try:
        # Your API call to Unsplash here
        # ...
        
        # For example purposes, let's assume we get back image_data
        image_data = {
            "url": f"https://example.com/{query}.jpg",
            "alt_description": f"Image for {query}",
            # Other image properties...
        }
        
        # Cache the result
        _image_cache[cache_key] = image_data
        
        return image_data
    except Exception as e:
        print(f"Error fetching image from Unsplash: {e}")
        return {}

def get_image_for_slide(slide_type: str, content: Dict[str, Any], use_cache: bool = True) -> Dict[str, Any]:
    """
    Get an appropriate image for a slide type based on its content.
    
    Args:
        slide_type: Type of slide (title, problem, solution, etc.)
        content: Slide content dictionary
        use_cache: Whether to use the image cache
        
    Returns:
        Image data dictionary
    """
    # Define queries for different slide types
    queries = {
        "title": content.get("product_name", ""),
        "problem": "problem challenge issue",
        "solution": "solution innovation success",
        "features": "features product technology",
        "advantage": "competitive advantage success",
        "audience": "target audience people market",
        "cta": "call to action success partnership",
        "market": "market growth opportunity",
        "roadmap": "roadmap strategy journey",
        "team": "team leadership professionals"
    }
    
    # Get the query for this slide type
    query = queries.get(slide_type, "business presentation")
    
    # If the slide type is "title", add the product name to make it more specific
    if slide_type == "title" and "product_name" in content:
        query = f"{content['product_name']} product"
        
    # Get the image using the query
    return get_image_from_unsplash(query, use_cache=use_cache)

def clear_image_cache():
    """Clear the image cache."""
    global _image_cache
    _image_cache = {}

def get_cached_images():
    """Get a copy of the current image cache."""
    return _image_cache.copy()

def set_image_cache(cache_data):
    """Set the image cache with provided data."""
    global _image_cache
    _image_cache = cache_data
