import streamlit as st
from typing import Dict, Any, List, Optional
import io
from ppt_generator import create_presentation

def get_standard_slide_order():
    """Get the standard order of slides in a presentation."""
    return [
        'title_slide', 
        'problem_slide', 
        'solution_slide', 
        'features_slide', 
        'advantage_slide', 
        'audience_slide',
        'market_slide',
        'roadmap_slide',
        'team_slide',
        'cta_slide'
    ]

def get_slide_title(slide_type: str) -> str:
    """Get a human-readable title for a slide type."""
    titles = {
        "title_slide": "Title Slide",
        "problem_slide": "Problem Statement",
        "solution_slide": "Solution Overview",
        "features_slide": "Key Features",
        "advantage_slide": "Competitive Advantage",
        "audience_slide": "Target Audience",
        "cta_slide": "Call to Action",
        "market_slide": "Market Analysis",
        "roadmap_slide": "Product Roadmap",
        "team_slide": "Team Overview"
    }
    return titles.get(slide_type, slide_type.replace('_', ' ').title())

def initialize_slide_order(content: Dict[str, Any]) -> List[str]:
    """
    Initialize the slide order from presentation content, 
    preserving the original presentation order.
    
    Args:
        content: Dictionary containing presentation content
        
    Returns:
        List of slide types in proper order
    """
    # Get standard order as a reference
    standard_order = get_standard_slide_order()
    
    # Extract slide types from content (keys ending with _slide)
    available_slides = [k for k in content.keys() if k.endswith('_slide')]
    
    # Organize slides in the standard order if they exist in available_slides
    ordered_slides = [slide for slide in standard_order if slide in available_slides]
    
    # Add any additional slides not in standard_order but in available_slides
    for slide in available_slides:
        if slide not in ordered_slides:
            ordered_slides.append(slide)
    
    return ordered_slides

def generate_reordered_presentation(
    original_content: Dict[str, Any], 
    custom_slide_order: List[str],
    filename: str,
    image_manager=None
) -> bytes:
    """
    Generate a reordered presentation based on the original content and a custom slide order.
    Uses caching to avoid regenerating the same presentation multiple times.
    
    Args:
        original_content: Dictionary containing the original presentation content
        custom_slide_order: List specifying the new order of slides
        filename: Name for the output file
        image_manager: Optional image manager for custom images
        
    Returns:
        bytes: Binary data for the PowerPoint file
    """
    try:
        # Create a cache key from content hash and slide order
        import hashlib
        import json
        
        # Create a key based on content and slide order
        content_hash = hashlib.md5(json.dumps(original_content, sort_keys=True).encode()).hexdigest()
        order_hash = hashlib.md5(json.dumps(custom_slide_order).encode()).hexdigest()
        cache_key = f"reordered_presentation_{content_hash}_{order_hash}"
        
        # Check if this exact configuration is cached
        if 'reordered_presentation_cache' not in st.session_state:
            st.session_state.reordered_presentation_cache = {}
            
        # Return cached presentation if available
        if cache_key in st.session_state.reordered_presentation_cache:
            return st.session_state.reordered_presentation_cache[cache_key]
        
        # Create a copy of the content to avoid modifying the original
        content_copy = {k: v for k, v in original_content.items()}
        
        # Ensure the metadata is preserved
        if 'metadata' in original_content:
            content_copy['metadata'] = original_content['metadata'].copy()
        
        # Print debug information
        print(f"Generating presentation with custom order: {custom_slide_order}")
        print(f"Available slides in content: {[k for k in content_copy.keys() if k != 'metadata']}")
        
        # Generate the presentation with the custom slide order
        presentation_buffer = create_presentation(
            content_copy,
            filename,
            image_manager=image_manager,
            custom_slide_order=custom_slide_order
        )
        
        if not presentation_buffer:
            raise Exception("Failed to generate presentation buffer")
        
        # Cache the result
        st.session_state.reordered_presentation_cache[cache_key] = presentation_buffer
        
        return presentation_buffer
        
    except Exception as e:
        st.error(f"Error generating reordered presentation: {str(e)}")
        raise e  # Re-raise the exception for proper error handling

def move_slide_up(i: int):
    """Move a slide up in the order."""
    if i > 0:  # Safety check
        st.session_state.slide_order[i], st.session_state.slide_order[i-1] = \
            st.session_state.slide_order[i-1], st.session_state.slide_order[i]
        st.session_state.has_modifications = True
        # Force regeneration of preview
        if 'preview_slides' in st.session_state:
            del st.session_state.preview_slides

def move_slide_down(i: int):
    """Move a slide down in the order."""
    if i < len(st.session_state.slide_order) - 1:  # Safety check
        st.session_state.slide_order[i], st.session_state.slide_order[i+1] = \
            st.session_state.slide_order[i+1], st.session_state.slide_order[i]
        st.session_state.has_modifications = True
        # Force regeneration of preview
        if 'preview_slides' in st.session_state:
            del st.session_state.preview_slides

def render_slide_reordering_ui(content: Dict[str, Any]) -> None:
    """
    Render UI elements for reordering slides, ensuring the initial order
    matches the original presentation. Uses session state callbacks for better performance.
    
    Args:
        content: Dictionary containing presentation content
    """
    # Initialize slide order if not already in session state
    if 'slide_order' not in st.session_state:
        st.session_state.slide_order = initialize_slide_order(content)
    
    # Initialize modifications flag if it doesn't exist
    if 'has_modifications' not in st.session_state:
        st.session_state.has_modifications = False
    
    # Display current slide order with up/down buttons
    for i, slide_type in enumerate(st.session_state.slide_order):
        col1, col2, col3 = st.columns([5, 1, 1])
        
        slide_title = get_slide_title(slide_type)
        with col1:
            st.markdown(f"**{i+1}. {slide_title}**")
        
        # Move Up button
        with col2:
            disabled = i <= 0  # Disable if it's the first slide
            if st.button("↑", key=f"up_{slide_type}_{i}", disabled=disabled, on_click=move_slide_up, args=(i,)):
                pass  # The actual action happens in the callback
            
        # Move Down button
        with col3:
            disabled = i >= len(st.session_state.slide_order) - 1  # Disable if it's the last slide
            if st.button("↓", key=f"down_{slide_type}_{i}", disabled=disabled, on_click=move_slide_down, args=(i,)):
                pass  # The actual action happens in the callback
