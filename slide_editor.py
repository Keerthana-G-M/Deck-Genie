import streamlit as st
from typing import Dict, Any, List, Optional
import copy
import io
from PIL import Image
# from streamlit_sortables import sort_items

class SlideEditor:
    """Handles slide editing, reordering, and deletion functionality."""
    
    def __init__(self):
        self.slide_titles = [
            "Title Slide",
            "Problem Statement", 
            "Solution Overview",
            "Key Features",
            "Competitive Advantage",
            "Target Audience", 
            "Call to Action",
            "Market Analysis",
            "Product Roadmap",
            "Team Overview"
        ]
        self.slide_keys = [
            "title_slide",
            "problem_slide", 
            "solution_slide",
            "features_slide",
            "advantage_slide",
            "audience_slide", 
            "cta_slide",
            "market_slide",
            "roadmap_slide",
            "team_slide"
        ]
    
    def initialize_editor_state(self, original_content: Dict[str, Any]):
        """Initialize the editor state in session state."""
        if 'editor_content' not in st.session_state:
            st.session_state.editor_content = copy.deepcopy(original_content)
        
        if 'slide_order' not in st.session_state:
            st.session_state.slide_order = self.slide_keys.copy()
        
        # Ensure deleted_slides is always a set of strings
        if 'deleted_slides' not in st.session_state:
            st.session_state.deleted_slides = set()
        elif not isinstance(st.session_state.deleted_slides, set):
            # Fix corrupted data - convert to set
            st.session_state.deleted_slides = set()
        else:
            # Ensure all items in the set are strings
            cleaned_deleted = set()
            for item in st.session_state.deleted_slides:
                if isinstance(item, str):
                    cleaned_deleted.add(item)
            st.session_state.deleted_slides = cleaned_deleted
        
        if 'has_modifications' not in st.session_state:
            st.session_state.has_modifications = False
        
        if 'uploaded_images' not in st.session_state:
            st.session_state.uploaded_images = {}
    
    def render_slide_editor(self, original_content: Dict[str, Any], preview_generator=None):
        """Render the slide editor interface with real-time preview updates."""
        self.initialize_editor_state(original_content)
        self.preview_generator = preview_generator
        
        # Slide reordering section
        st.markdown("### 🔄 Reorder Slides")
        self._render_slide_reordering()
          # Download buttons
        self._render_download_buttons(original_content)
    
    def _ensure_deleted_slides_safe(self):
        """Ensure deleted_slides is always a safe set of strings."""
        if not isinstance(st.session_state.deleted_slides, set):
            st.session_state.deleted_slides = set()
        else:
            # Clean any non-string items
            try:
                st.session_state.deleted_slides = {
                    item for item in st.session_state.deleted_slides 
                    if isinstance(item, str)
                }
            except (TypeError, AttributeError):
                st.session_state.deleted_slides = set()
    
    def _render_slide_reordering(self):
        """Render the slide reordering interface without causing app reruns."""
        # Ensure deleted_slides is safe to use
        self._ensure_deleted_slides_safe()
        
        # Ensure slide_order exists and is properly initialized
        if 'slide_order' not in st.session_state or not st.session_state.slide_order:
            st.session_state.slide_order = self.slide_keys.copy()
        
        # Check if we need to update slide_order with any new slides from editor_content
        if 'editor_content' in st.session_state:
            for slide_key in st.session_state.editor_content:
                if slide_key != 'metadata' and slide_key not in st.session_state.slide_order and slide_key in self.slide_keys:
                    st.session_state.slide_order.append(slide_key)
        
        # Validate slide_order contains valid keys
        valid_slide_order = [key for key in st.session_state.slide_order if key in self.slide_keys]
        if len(valid_slide_order) != len(st.session_state.slide_order):
            st.session_state.slide_order = valid_slide_order if valid_slide_order else self.slide_keys.copy()
        
        # Get current active slides for display
        active_slides = []
        for slide_key in st.session_state.slide_order:
            if slide_key not in st.session_state.deleted_slides and slide_key in self.slide_keys:
                try:
                    slide_index = self.slide_keys.index(slide_key)
                    # Check if this slide is in editor_content to confirm it exists
                    if 'editor_content' not in st.session_state or slide_key in st.session_state.editor_content:
                        slide_title = self.slide_titles[slide_index]
                        active_slides.append((slide_key, slide_title))
                except (ValueError, IndexError):
                    continue
        
        if not active_slides:
            st.info("ℹ️ No slides available for reordering.")
            return
        
        st.info("📝 Use the buttons below to reorder slides (changes are instant):")
        
        # Use columns for better layout
        col_labels, col_buttons = st.columns([5, 2])
        with col_labels:
            st.write("**Current Slide Order:**")
        with col_buttons:
            st.write("**Actions:**")
        
        # Display current order with move buttons
        for i, (slide_key, slide_title) in enumerate(active_slides):
            col1, col2, col3 = st.columns([5, 1, 1])
            
            with col1:
                st.write(f"**{i+1}.** {slide_title}")
            
            with col2:
                # Move up button
                if i > 0:
                    if st.button("⬆️", key=f"move_up_{slide_key}", help="Move up"):
                        self._move_slide_up(slide_key)
                        # Flag that modifications were made
                        st.session_state.has_modifications = True
            
            with col3:
                # Move down button
                if i < len(active_slides) - 1:
                    if st.button("⬇️", key=f"move_down_{slide_key}", help="Move down"):
                        self._move_slide_down(slide_key)
                        # Flag that modifications were made
                        st.session_state.has_modifications = True
    
    def _move_slide_up(self, slide_key: str):
        """Move a slide up one position among active (non-deleted) slides."""
        # Get current active slides in order
        active_slides = [key for key in st.session_state.slide_order 
                        if key not in st.session_state.deleted_slides]
        
        if slide_key not in active_slides:
            return
            
        current_active_index = active_slides.index(slide_key)
        
        # Can't move up if already at the top of active slides
        if current_active_index <= 0:
            return
            
        # Swap with the active slide above in the active list
        active_slides[current_active_index], active_slides[current_active_index - 1] = \
            active_slides[current_active_index - 1], active_slides[current_active_index]
        
        # Rebuild the full order by inserting active slides back while preserving deleted slide positions
        self._rebuild_full_order_from_active(active_slides)
        
        st.session_state.has_modifications = True
        self._update_preview_lightweight_no_rerun()
        
        # Force a clean rerun to prevent scrolling issues
        st.rerun()
    
    def _move_slide_down(self, slide_key: str):
        """Move a slide down one position among active (non-deleted) slides."""
        # Get current active slides in order
        active_slides = [key for key in st.session_state.slide_order 
                        if key not in st.session_state.deleted_slides]
        
        if slide_key not in active_slides:
            return
            
        current_active_index = active_slides.index(slide_key)
        
        # Can't move down if already at the bottom of active slides
        if current_active_index >= len(active_slides) - 1:
            return
            
        # Swap with the active slide below in the active list
        active_slides[current_active_index], active_slides[current_active_index + 1] = \
            active_slides[current_active_index + 1], active_slides[current_active_index]
        
        # Rebuild the full order by inserting active slides back while preserving deleted slide positions
        self._rebuild_full_order_from_active(active_slides)
        
        st.session_state.has_modifications = True
        self._update_preview_lightweight_no_rerun()
          # Force a clean rerun to prevent scrolling issues
        st.rerun()
    
    def _rebuild_full_order_from_active(self, reordered_active_slides):
        """Rebuild the full slide order from reordered active slides, preserving deleted slide positions."""
        new_order = []
        active_index = 0
        
        # Keep track of all available slide keys including those in editor_content
        all_available_keys = set(self.slide_keys)
        if 'editor_content' in st.session_state:
            for key in st.session_state.editor_content:
                if key != 'metadata':
                    all_available_keys.add(key)
                    
        # Go through the original slide order and insert active slides in new order
        for original_key in self.slide_keys:
            if original_key in st.session_state.deleted_slides:                # Keep deleted slides at their original positions
                new_order.append(original_key)
            else:
                # Insert active slides in their new order
                if active_index < len(reordered_active_slides):
                    new_order.append(reordered_active_slides[active_index])
                    active_index += 1
        
        # Make sure we've included all active slides
        while active_index < len(reordered_active_slides):
            new_order.append(reordered_active_slides[active_index])
            active_index += 1
        
        # Make sure all slides from editor_content are included
        if 'editor_content' in st.session_state:
            for key in st.session_state.editor_content:
                if key != 'metadata' and key not in new_order and key not in st.session_state.deleted_slides:
                    new_order.append(key)
        
        # Update session state with the new order
        st.session_state.slide_order = new_order
        
        # Flag that modifications were made to enable download
        st.session_state.has_modifications = True

    def _update_preview_lightweight_no_rerun(self):
        """Update preview order without any reruns or image fetching."""
        if hasattr(self, 'preview_generator') and self.preview_generator:
            try:
                # Simply update the session state preview order
                # The UI will automatically reflect this change on the next render cycle
                if 'preview_slides' in st.session_state and st.session_state.preview_slides:
                    # Create mapping of slide keys to preview data
                    slide_previews = {}
                    slide_type_mapping = {
                        'title_slide': 0, 'problem_slide': 1, 'solution_slide': 2, 'features_slide': 3,
                        'advantage_slide': 4, 'audience_slide': 5, 'cta_slide': 6
                    }
                    
                    # Map existing previews by slide type
                    for preview in st.session_state.preview_slides:
                        original_index = preview.get('index', 0)
                        for key, index in slide_type_mapping.items():
                            if index == original_index:
                                slide_previews[key] = preview
                                break
                    
                    # Reorder according to current slide order
                    reordered = []
                    for slide_key in st.session_state.slide_order:
                        if (slide_key not in st.session_state.deleted_slides and 
                            slide_key in slide_previews):
                            reordered.append(slide_previews[slide_key])
                    
                    # Update preview slides order in session state
                    st.session_state.preview_slides = reordered
                    
            except Exception:
                # Silently handle errors - preview will update on next full render
                pass

    def _render_individual_slide_editors(self):
        """Render editors for individual slides."""
        # Ensure deleted_slides is safe to use
        self._ensure_deleted_slides_safe()
        
        for slide_key in st.session_state.slide_order:
            if slide_key not in st.session_state.deleted_slides:
                slide_index = self.slide_keys.index(slide_key)
                slide_title = self.slide_titles[slide_index]
                
                # Use simple containers instead of nested expanders
                st.markdown(f"#### 📄 {slide_title}")
                with st.container():
                    self._render_single_slide_editor(slide_key, slide_title)
                st.markdown("---")
    
    def _render_single_slide_editor(self, slide_key: str, slide_title: str):
        """Render editor for a single slide."""
        if slide_key not in st.session_state.editor_content:
            st.warning(f"Content for {slide_title} not found.")
            return
        
        slide_content = st.session_state.editor_content[slide_key]
        
        # Delete slide button
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button(f"🗑️ Delete Slide", key=f"delete_{slide_key}"):
                st.session_state.deleted_slides.add(slide_key)
                st.session_state.has_modifications = True
                self._update_preview_if_available()
        
        with col1:
            st.markdown(f"**Editing: {slide_title}**")
        
        # Title editing (for most slides)
        if 'title' in slide_content:
            new_title = st.text_input(
                "Slide Title",
                value=slide_content['title'],
                key=f"title_{slide_key}"
            )
            if new_title != slide_content['title']:
                st.session_state.editor_content[slide_key]['title'] = new_title
                st.session_state.has_modifications = True
                self._update_preview_if_available()
        
        # Content editing based on slide type
        if slide_key == 'title_slide':
            self._edit_title_slide(slide_content, slide_key)
        elif slide_key in ['problem_slide', 'advantage_slide']:
            self._edit_bullet_slide(slide_content, slide_key)
        elif slide_key in ['solution_slide', 'audience_slide']:
            self._edit_paragraph_slide(slide_content, slide_key)
        elif slide_key == 'features_slide':
            self._edit_features_slide(slide_content, slide_key)
        elif slide_key == 'cta_slide':
            self._edit_cta_slide(slide_content, slide_key)
        
        # Image upload section
        self._render_image_upload(slide_key, slide_title)
    
    def _edit_title_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit title slide content."""
        if 'subtitle' in slide_content:
            new_subtitle = st.text_input(
                "Subtitle",
                value=slide_content['subtitle'],
                key=f"subtitle_{slide_key}"
            )
            if new_subtitle != slide_content['subtitle']:
                st.session_state.editor_content[slide_key]['subtitle'] = new_subtitle
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _edit_bullet_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit slides with bullet points."""
        if 'bullets' in slide_content:
            st.markdown("**Bullet Points:**")
            bullets = slide_content['bullets'].copy()
            
            for i, bullet in enumerate(bullets):
                col1, col2 = st.columns([4, 1])
                with col1:
                    new_bullet = st.text_area(
                        f"Bullet {i+1}",
                        value=bullet,
                        height=80,
                        key=f"bullet_{slide_key}_{i}"
                    )
                    if new_bullet != bullet:
                        st.session_state.editor_content[slide_key]['bullets'][i] = new_bullet
                        st.session_state.has_modifications = True
                        self._update_preview_if_available()
                
                with col2:
                    if st.button(f"🗑️", key=f"delete_bullet_{slide_key}_{i}"):
                        st.session_state.editor_content[slide_key]['bullets'].pop(i)
                        st.session_state.has_modifications = True
                        self._update_preview_if_available()
            
            # Add new bullet point
            if st.button(f"➕ Add Bullet Point", key=f"add_bullet_{slide_key}"):
                st.session_state.editor_content[slide_key]['bullets'].append("New bullet point")
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _edit_paragraph_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit slides with paragraph content."""
        if 'paragraph' in slide_content:
            new_paragraph = st.text_area(
                "Content",
                value=slide_content['paragraph'],
                height=100,
                key=f"paragraph_{slide_key}"
            )
            if new_paragraph != slide_content['paragraph']:
                st.session_state.editor_content[slide_key]['paragraph'] = new_paragraph
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _edit_features_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit features slide content."""
        if 'features' in slide_content:
            st.markdown("**Features:**")
            features = slide_content['features'].copy()
            
            for i, feature in enumerate(features):
                col1, col2 = st.columns([4, 1])
                with col1:
                    new_feature = st.text_area(
                        f"Feature {i+1}",
                        value=feature,
                        height=80,
                        key=f"feature_{slide_key}_{i}"
                    )
                    if new_feature != feature:
                        st.session_state.editor_content[slide_key]['features'][i] = new_feature
                        st.session_state.has_modifications = True
                        self._update_preview_if_available()
                
                with col2:
                    if st.button(f"🗑️", key=f"delete_feature_{slide_key}_{i}"):
                        st.session_state.editor_content[slide_key]['features'].pop(i)
                        st.session_state.has_modifications = True
                        self._update_preview_if_available()
            
            # Add new feature
            if st.button(f"➕ Add Feature", key=f"add_feature_{slide_key}"):
                st.session_state.editor_content[slide_key]['features'].append("New feature: Description")
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _edit_cta_slide(self, slide_content: Dict[str, Any], slide_key: str):
        """Edit call to action slide content."""
        if 'call_to_action' in slide_content:
            new_cta = st.text_input(
                "Call to Action",
                value=slide_content['call_to_action'],
                key=f"cta_{slide_key}"
            )
            if new_cta != slide_content['call_to_action']:
                st.session_state.editor_content[slide_key]['call_to_action'] = new_cta
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _render_image_upload(self, slide_key: str, slide_title: str):
        """Render image upload section for a slide."""
        st.markdown("**Custom Image:**")
        
        uploaded_file = st.file_uploader(
            f"Upload custom image for {slide_title}",
            type=['png', 'jpg', 'jpeg'],
            key=f"image_{slide_key}"
        )
        
        if uploaded_file is not None:
            # Store the uploaded image
            st.session_state.uploaded_images[slide_key] = uploaded_file.getvalue()
            st.session_state.has_modifications = True
            self._update_preview_if_available()
            
            # Show preview
            image = Image.open(uploaded_file)
            st.image(image, caption="Uploaded Image Preview", width=200)
        
        # Show current custom image if exists
        elif slide_key in st.session_state.uploaded_images:
            image_data = st.session_state.uploaded_images[slide_key]
            image = Image.open(io.BytesIO(image_data))
            st.image(image, caption="Current Custom Image", width=200)
            
            if st.button(f"🗑️ Remove Custom Image", key=f"remove_image_{slide_key}"):
                del st.session_state.uploaded_images[slide_key]
                st.session_state.has_modifications = True
                self._update_preview_if_available()
    
    def _render_download_buttons(self, original_content: Dict[str, Any]):
        """Render download buttons for original and reordered presentations."""
        st.markdown("---")
        st.markdown("## 📥 Download Options")
        
        col1, col2 = st.columns(2)
        
        # Original presentation download
        with col1:
            st.markdown("### Original Presentation")
            st.markdown("Download the original generated presentation")
            
            from ppt_generator import create_presentation
            from utils import sanitize_filename
            
            try:
                original_filename = f"{sanitize_filename(original_content['metadata']['company_name'])}_{sanitize_filename(original_content['metadata']['product_name'])}_Overview.pptx"
                
                # Cache the original presentation buffer to prevent API calls on every re-render
                cache_key = f"original_buffer_{id(original_content)}"
                if cache_key not in st.session_state:
                    st.session_state[cache_key] = create_presentation(original_content, original_filename)
                original_buffer = st.session_state[cache_key]
                
                st.download_button(
                    label="📥 Download Original",
                    data=original_buffer,
                    file_name=original_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error creating original presentation: {str(e)}")
        
        # Reordered presentation download
        with col2:
            st.markdown("### Reordered Presentation")
            if st.session_state.has_modifications:
                st.markdown("✅ **Slide order changed** - Download your reordered version")
                
                try:
                    modified_content = self._prepare_modified_content(original_content)
                    modified_filename = f"{sanitize_filename(original_content['metadata']['company_name'])}_{sanitize_filename(original_content['metadata']['product_name'])}_Reordered.pptx"
                    
                    # Create reordered presentation
                    from image_manager import ImageManager
                    image_manager = ImageManager(
                        st.session_state.get('original_images_cache', {}),
                        st.session_state.get('original_images_cache', {})
                    )
                    modified_buffer = self._create_modified_presentation(modified_content, modified_filename, image_manager)
                    
                    st.download_button(
                        label="📥 Download Reordered",
                        data=modified_buffer,
                        file_name=modified_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error creating reordered presentation: {str(e)}")
            else:
                st.markdown("ℹ️ Reorder slides above to enable download")
                st.button(
                    "📥 Download Reordered",
                    disabled=True,
                    use_container_width=True,
                )
    
    def initialize_editor_state(self, content):
        """Initialize the editor state with presentation content."""
        if not content:
            return
        
        # Get the slide count from metadata
        slide_count = content.get('metadata', {}).get('slide_count', 7)
        
        # Get all available slide keys (that have content)
        all_slide_keys = [key for key in content.keys() if key.endswith('_slide') and isinstance(content[key], dict)]
        
        # Respect the user's chosen slide count
        self.slide_keys = all_slide_keys[:slide_count]
        
        # Add to session state if needed
        if 'slide_order' not in st.session_state:
            st.session_state.slide_order = self.slide_keys.copy()
        elif len(st.session_state.slide_order) != slide_count:
            # If slide count changed, update the order with only valid slides
            st.session_state.slide_order = [
                key for key in st.session_state.slide_order 
                if key in self.slide_keys
            ]
            # Add any missing slides (should not happen normally)
            for key in self.slide_keys:
                if key not in st.session_state.slide_order:
                    st.session_state.slide_order.append(key)
    
    def _prepare_modified_content(self, original_content):
        """Prepare modified content based on slide order and edits."""
        modified_content = {}
        
        # Copy metadata and non-slide content
        for key, value in original_content.items():
            if not key.endswith('_slide'):
                modified_content[key] = value
        
        # Get the slide count from metadata
        slide_count = original_content.get('metadata', {}).get('slide_count', 7)
        
        # Use ordered slides and respect slide count
        ordered_slides = st.session_state.slide_order[:slide_count]
        
        # Copy the slides in the new order
        for slide_key in ordered_slides:
            if slide_key in original_content:
                modified_content[slide_key] = original_content[slide_key]
        
        return modified_content
    
    def _create_modified_presentation(self, content: Dict[str, Any], filename: str, image_manager) -> io.BytesIO:
        """Create a modified presentation with custom images and content."""
        from ppt_generator_custom import create_custom_presentation
        return create_custom_presentation(content, filename, st.session_state.slide_order, image_manager)

    def get_slide_order(self) -> List[str]:
        """Get the current slide order."""
        return [key for key in st.session_state.slide_order if key not in st.session_state.deleted_slides]
    
    def _update_preview_if_available(self):
        """Update the preview if a preview generator is available."""
        if hasattr(self, 'preview_generator') and self.preview_generator:
            try:
                self.preview_generator.update_preview_from_session_state(
                    st.session_state.editor_content,
                    st.session_state.slide_order,
                    st.session_state.deleted_slides
                )
            except Exception as e:
                # Silently handle preview update errors to not break the editor
                pass
    
    def _update_preview_lightweight(self):
        """Update preview order without fetching any images - for real-time reordering."""
        if hasattr(self, 'preview_generator') and self.preview_generator:
            try:
                # Only reorder existing slides in session state, don't regenerate anything
                if 'preview_slides' in st.session_state and st.session_state.preview_slides:
                    # Create a mapping of slide keys to preview slides
                    slide_key_to_preview = {}
                    slide_type_to_key = {
                        'title_slide': 0, 'problem_slide': 1, 'solution_slide': 2, 'features_slide': 3,
                        'advantage_slide': 4, 'audience_slide': 5, 'cta_slide': 6
                    }
                    
                    # Map existing preview slides by their original index
                    for preview_slide in st.session_state.preview_slides:
                        original_index = preview_slide.get('index', 0)
                        # Find the slide key that corresponds to this index
                        for key, index in slide_type_to_key.items():
                            if index == original_index:
                                slide_key_to_preview[key] = preview_slide
                                break
                    
                    # Reorder preview slides according to current slide_order
                    reordered_previews = []
                    for slide_key in st.session_state.slide_order:
                        if (slide_key not in st.session_state.deleted_slides and 
                            slide_key in slide_key_to_preview):
                            reordered_previews.append(slide_key_to_preview[slide_key])
                    
                    # Update the preview slides order
                    st.session_state.preview_slides = reordered_previews
                    
                    # Only re-render the UI layout, no image processing
                    self.preview_generator.render_all_preview_slides()
            except Exception as e:
                # Silently handle preview update errors to not break the editor
                pass
