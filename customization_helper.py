import streamlit as st
from typing import Dict, Any, List, Tuple
from preview_generator import SlidePreviewGenerator

def create_customization_section_from_preview(preview_generator: SlidePreviewGenerator) -> Tuple[List[str], Dict[str, Any], set]:
    """
    Create a customization section that matches exactly what's in the live preview.
    Returns: (slide_order, editor_content, deleted_slides)
    """
    
    # Get current slide order and data from preview
    slide_order = preview_generator.get_current_slide_order_for_customization()
    slides_data = preview_generator.get_preview_slides_data()
    
    if not slide_order or not slides_data:
        # Fallback to session state if preview is empty
        if 'final_slide_order' in st.session_state and 'generated_content' in st.session_state:
            slide_order = st.session_state.final_slide_order.copy()
            slides_data = {k: v for k, v in st.session_state.generated_content.items() if k != 'metadata'}
        else:
            return [], {}, set()
    
    # Initialize session state for customization if not exists
    if 'editor_slide_order' not in st.session_state:
        st.session_state.editor_slide_order = slide_order.copy()
    
    if 'editor_content' not in st.session_state:
        st.session_state.editor_content = slides_data.copy()
    
    if 'deleted_slides' not in st.session_state:
        st.session_state.deleted_slides = set()
    
    # Ensure customization matches current preview
    current_preview_slides = set(slide_order)
    current_editor_slides = set(st.session_state.editor_slide_order)
    
    # If preview has changed, update the editor to match
    if current_preview_slides != current_editor_slides:
        st.session_state.editor_slide_order = slide_order.copy()
        st.session_state.editor_content.update(slides_data)
        # Clear deleted slides that are now back in preview
        st.session_state.deleted_slides = st.session_state.deleted_slides - current_preview_slides
    
    # Display reordering section
    st.markdown("### 🎛️ Customize Your Presentation")
    st.markdown("**Reorder your slides using the buttons below:**")
    
    # Create columns for reordering buttons
    active_slides = [s for s in st.session_state.editor_slide_order if s not in st.session_state.deleted_slides]
    
    if len(active_slides) > 1:
        cols = st.columns(len(active_slides))
        
        for i, slide_type in enumerate(active_slides):
            with cols[i]:
                display_name = preview_generator.get_slide_display_name(slide_type)
                icon = preview_generator.get_slide_icon(slide_type)
                
                st.markdown(f"**{icon} {display_name}**")
                
                # Move up button (disabled for first slide)
                if st.button("↑", key=f"up_{slide_type}", disabled=(i == 0)):
                    if i > 0:
                        # Swap with previous slide
                        current_order = st.session_state.editor_slide_order.copy()
                        prev_slide = active_slides[i-1]
                        
                        # Find positions in the full order
                        pos1 = current_order.index(slide_type)
                        pos2 = current_order.index(prev_slide)
                        
                        # Swap
                        current_order[pos1], current_order[pos2] = current_order[pos2], current_order[pos1]
                        st.session_state.editor_slide_order = current_order
                        st.rerun()
                
                # Move down button (disabled for last slide)
                if st.button("↓", key=f"down_{slide_type}", disabled=(i == len(active_slides) - 1)):
                    if i < len(active_slides) - 1:
                        # Swap with next slide
                        current_order = st.session_state.editor_slide_order.copy()
                        next_slide = active_slides[i+1]
                        
                        # Find positions in the full order
                        pos1 = current_order.index(slide_type)
                        pos2 = current_order.index(next_slide)
                        
                        # Swap
                        current_order[pos1], current_order[pos2] = current_order[pos2], current_order[pos1]
                        st.session_state.editor_slide_order = current_order
                        st.rerun()
                
                # Delete button (disabled for title and CTA slides)
                can_delete = slide_type not in ['title_slide', 'cta_slide']
                if st.button("🗑️", key=f"delete_{slide_type}", disabled=not can_delete):
                    if can_delete:
                        st.session_state.deleted_slides.add(slide_type)
                        st.rerun()
    
    # Show slide count
    st.markdown(f"**Current slide count:** {len(active_slides)} slides")
    
    # Option to restore deleted slides
    if st.session_state.deleted_slides:
        st.markdown("### 🔄 Restore Deleted Slides")
        deleted_list = list(st.session_state.deleted_slides)
        
        cols = st.columns(min(len(deleted_list), 4))
        for i, slide_type in enumerate(deleted_list):
            with cols[i % 4]:
                display_name = preview_generator.get_slide_display_name(slide_type)
                icon = preview_generator.get_slide_icon(slide_type)
                
                if st.button(f"➕ {icon} {display_name}", key=f"restore_{slide_type}"):
                    st.session_state.deleted_slides.remove(slide_type)
                    st.rerun()
    
    return st.session_state.editor_slide_order, st.session_state.editor_content, st.session_state.deleted_slides

def get_final_content_for_export(slide_order: List[str], editor_content: Dict[str, Any], deleted_slides: set, original_metadata: Dict[str, Any] = None) -> Dict[str, Any]:
    """
    Get the final content structure for PowerPoint export based on customization choices.
    """
    final_content = {}
    
    # Add slides in the specified order, excluding deleted ones
    for slide_type in slide_order:
        if slide_type not in deleted_slides and slide_type in editor_content:
            final_content[slide_type] = editor_content[slide_type]
    
    # Add metadata
    if original_metadata:
        final_content['metadata'] = original_metadata
    elif 'generated_content' in st.session_state and 'metadata' in st.session_state.generated_content:
        final_content['metadata'] = st.session_state.generated_content['metadata']
    
    # Update slide count in metadata
    if 'metadata' in final_content:
        final_content['metadata']['slide_count'] = len([s for s in slide_order if s not in deleted_slides])
    
    return final_content

def update_preview_after_customization(preview_generator: SlidePreviewGenerator, slide_order: List[str], editor_content: Dict[str, Any], deleted_slides: set):
    """
    Update the live preview to reflect customization changes.
    """
    # Update preview based on current customization state
    preview_generator.update_preview_from_session_state(editor_content, slide_order, deleted_slides)
