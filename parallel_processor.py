"""
Parallel processing utilities for fast slide generation.
This module helps improve performance by generating slides in parallel.
"""
import concurrent.futures
import streamlit as st
from typing import Dict, Any, List, Callable, Tuple
import time

def generate_slides_in_parallel(
    content_generators: Dict[str, Callable],
    slide_types: List[str],
    slide_args: Dict[str, Tuple],
    max_workers: int = 3
) -> Dict[str, Any]:
    """
    Generate slides in parallel to improve performance.
    
    Args:
        content_generators: Dictionary mapping slide types to generator functions
        slide_types: List of slide types to generate
        slide_args: Dictionary mapping slide types to arguments for their generator
        max_workers: Maximum number of parallel workers (default 3 to avoid rate limits)
    
    Returns:
        Dictionary containing generated slides
    """
    results = {}
    
    # Create a cache key based on inputs for each slide
    slide_cache_keys = {}
    for slide_type in slide_types:
        if slide_type in slide_args:
            import hashlib
            import json
            # Create a unique cache key for this slide's inputs
            args_str = json.dumps(slide_args[slide_type], sort_keys=True)
            cache_key = f"{slide_type}_{hashlib.md5(args_str.encode()).hexdigest()}"
            slide_cache_keys[slide_type] = cache_key
    
    # Initialize slide cache if it doesn't exist
    if 'parallel_slide_cache' not in st.session_state:
        st.session_state.parallel_slide_cache = {}
    
    # Determine which slides need to be generated vs retrieved from cache
    slides_to_generate = []
    for slide_type in slide_types:
        if slide_type in slide_cache_keys:
            cache_key = slide_cache_keys[slide_type]
            if cache_key in st.session_state.parallel_slide_cache:
                # Use cached slide
                results[slide_type] = st.session_state.parallel_slide_cache[cache_key]
            else:
                # Mark for generation
                slides_to_generate.append(slide_type)
    
    # Generate required slides in parallel
    if slides_to_generate:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_slide = {}
            for slide_type in slides_to_generate:
                if slide_type in content_generators and slide_type in slide_args:
                    future = executor.submit(
                        content_generators[slide_type],
                        *slide_args[slide_type]
                    )
                    future_to_slide[future] = slide_type
            
            # Collect results as they complete
            for future in concurrent.futures.as_completed(future_to_slide):
                slide_type = future_to_slide[future]
                try:
                    slide_content = future.result()
                    results[slide_type] = slide_content
                    
                    # Cache the result
                    if slide_type in slide_cache_keys:
                        cache_key = slide_cache_keys[slide_type]
                        st.session_state.parallel_slide_cache[cache_key] = slide_content
                        
                except Exception as e:
                    st.error(f"Error generating {slide_type}: {str(e)}")
    
    return results


def update_presentation_with_parallel_slides(
    presentation_content: Dict[str, Any], 
    parallel_slides: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Update a presentation content dictionary with slides generated in parallel.
    
    Args:
        presentation_content: Original presentation content
        parallel_slides: Dictionary of slides generated in parallel
    
    Returns:
        Updated presentation content
    """
    # Create a copy to avoid modifying the original
    updated_content = presentation_content.copy()
    
    # Update with parallel-generated slides
    for slide_type, slide_content in parallel_slides.items():
        updated_content[slide_type] = slide_content
    
    return updated_content
