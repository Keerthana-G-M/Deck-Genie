import streamlit as st
from typing import Dict, Any, List
import time

class SlidePreviewGenerator:
    """Generates live preview cards for slides as they are created."""
    
    def __init__(self):
        """Initialize the preview generator."""
        self.preview_container = None
        
    def create_preview_container(self):
        """Create the container for live slide previews that persists throughout the session."""
        # Create persistent container that can be cleared and rewritten
        self.preview_container = st.empty()
        
        # Initialize session state for preview slides
        if 'preview_slides' not in st.session_state:
            st.session_state.preview_slides = []
    
    def reset_preview(self):
        """Reset the preview slides."""
        st.session_state.preview_slides = []
        self.render_all_preview_slides()
    
    def add_slide_preview(self, slide_data: Dict[str, Any], slide_type: str):
        """Add a slide to the preview with sequential animation."""
        # Create a preview slide object
        icon = self._get_slide_icon(slide_type)
        title = self._get_slide_title(slide_type)
        
        slide = {
            'type': slide_type,
            'icon': icon,
            'title': title,
            'data': slide_data
        }
        
        # Append to session state list
        st.session_state.preview_slides.append(slide)
        
        # Render all slides - this will show the new slide with animation
        self.render_all_preview_slides(animate_last=True)

    def update_preview_with_content(self, content: Dict[str, Any]) -> None:
        """
        Update the preview with content from a fully generated presentation.
        This efficiently reuses cached content without regenerating slides.
        
        Args:
            content: Dictionary containing presentation content
        """
        # Reset current preview
        st.session_state.preview_slides = []
        
        # Get slide keys in the proper order
        slide_keys = []
        
        # Get slide count from metadata
        slide_count = 7  # Default if not specified
        if 'metadata' in content and 'slide_count' in content['metadata']:
            slide_count = content['metadata']['slide_count']
        
        # First, add slides based on metadata included_slides if available
        if 'metadata' in content and 'included_slides' in content['metadata']:
            # Important: Limit to the specified slide count
            slide_keys = content['metadata']['included_slides'][:slide_count]
        else:
            # Fallback to standard slide order
            standard_slides = ['title_slide', 'problem_slide', 'solution_slide', 'features_slide', 
                              'advantage_slide', 'audience_slide', 'cta_slide']
            additional_slides = ['market_slide', 'roadmap_slide', 'team_slide']
            
            # Add available slides in standard order, limited by slide count
            for slide_type in standard_slides + additional_slides:
                if slide_type in content and len(slide_keys) < slide_count:
                    slide_keys.append(slide_type)
        
        # Create preview slides without animation delays
        for slide_type in slide_keys:
            if slide_type in content:
                icon = self._get_slide_icon(slide_type)
                title = self._get_slide_title(slide_type)
                
                slide = {
                    'type': slide_type,
                    'icon': icon,
                    'title': title,
                    'data': content[slide_type]
                }
                
                st.session_state.preview_slides.append(slide)
        
        # Render all slides at once, no animation
        self.render_all_preview_slides(animate_last=False)
    
    def update_preview_from_session_state(self, editor_content: Dict[str, Any], slide_order: List[str], deleted_slides: set) -> None:
        """
        Update the preview based on current session state editor content.
        Used for real-time updates during customization.
        
        Args:
            editor_content: Dictionary containing edited slide content
            slide_order: List of slide keys in desired order
            deleted_slides: Set of slide keys that are deleted
        """
        # Reset current preview
        st.session_state.preview_slides = []
        
        # Get slide count limit from metadata if available
        slide_count = 10  # Default maximum
        if 'metadata' in editor_content and 'slide_count' in editor_content['metadata']:
            slide_count = editor_content['metadata']['slide_count']
        
        # Track how many slides we've added
        slides_added = 0
        
        # Add slides in the specified order, skipping deleted ones and respecting slide count
        for slide_type in slide_order:
            # Stop if we've reached the slide count limit
            if slides_added >= slide_count:
                break
                
            if slide_type not in deleted_slides and slide_type in editor_content:
                icon = self._get_slide_icon(slide_type)
                title = self._get_slide_title(slide_type)
                
                slide = {
                    'type': slide_type,
                    'icon': icon,
                    'title': title,
                    'data': editor_content[slide_type]
                }
                
                st.session_state.preview_slides.append(slide)
                slides_added += 1
        
        # Render all slides at once, no animation
        self.render_all_preview_slides(animate_last=False)
    
    def get_current_slide_order_for_customization(self) -> List[str]:
        """Get the current slide order from preview for customization."""
        return [slide['type'] for slide in st.session_state.preview_slides]
    
    def get_preview_slides_data(self) -> Dict[str, Any]:
        """Get the slide data from preview slides."""
        return {slide['type']: slide['data'] for slide in st.session_state.preview_slides}
    
    def get_slide_display_name(self, slide_type: str) -> str:
        """Get display name for a slide type."""
        return self._get_slide_title(slide_type)
    
    def get_slide_icon(self, slide_type: str) -> str:
        """Get icon for a slide type."""
        return self._get_slide_icon(slide_type)
    
    def render_all_preview_slides(self, animate_last=False):
        """Render all preview slides with equal dimensions."""
        if not self.preview_container:
            self.preview_container = st.empty()
        
        # Clear and rebuild the entire preview using st.empty()
        with self.preview_container.container():
            slides = st.session_state.preview_slides
            if not slides:
                return
            
            num_slides = len(slides)
            # Determine optimal number of columns based on slide count
            if num_slides <= 4:
                cols_per_row = 3  # Use 3 columns for 1-4 slides
            elif num_slides <= 8:
                cols_per_row = 4  # Use 4 columns for 5-8 slides
            else:
                cols_per_row = 5  # Use 5 columns for 9-10 slides
            
            num_rows = (num_slides + cols_per_row - 1) // cols_per_row  # Ceiling division
            
            # Use equal width for all columns
            for row in range(num_rows):
                cols = st.columns([1] * cols_per_row)  # Equal width distribution
                for col_idx in range(cols_per_row):
                    slide_idx = row * cols_per_row + col_idx
                    if slide_idx < num_slides:
                        slide = slides[slide_idx]
                        with cols[col_idx]:
                            # Apply animation class to the last slide if requested
                            animation = "slide-in" if animate_last and slide_idx == num_slides - 1 else ""
                            self.render_single_slide_preview(slide, animation)
    
    def render_single_slide_preview(self, slide: Dict[str, Any], animation=""):
        """Render a single slide preview card with animation if specified."""
        # Get content preview based on slide data
        content_preview = self.get_slide_content_preview(slide['data'])
        
        # Add animation class if requested
        animation_style = """
        @keyframes slideIn {
            0% { opacity: 0; transform: translateY(20px); }
            100% { opacity: 1; transform: translateY(0); }
        }
        .slide-in {
            animation: slideIn 0.7s ease-out;
        }
        """ if animation else ""
        
        st.markdown(f"""
        <style>
            {animation_style}
        </style>
        <div class="{animation}" style="
            border: 2px solid #0078D7; 
            border-radius: 12px; 
            padding: 10px; 
            text-align: center; 
            background: var(--background-color);
            margin: 5px 2px;
            height: 150px;
            width: 100%;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            align-items: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            transition: transform 0.2s ease;
            overflow: hidden;
        ">
            <div style="font-size: 1.8em; margin-bottom: 4px;">{slide['icon']}</div>
            <div style="font-weight: bold; font-size: 0.85em; margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 100%;">{slide['title']}</div>
            <div style="font-size: 0.7em; line-height: 1.1; overflow: hidden; text-overflow: ellipsis; max-height: 65px; padding: 0 4px;">
                {content_preview}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    def get_slide_content_preview(self, slide_data: Dict[str, Any]) -> str:
        """Extract a meaningful preview from slide data."""
        if not slide_data:
            return "Generated content"
        
        # Try to get the most relevant content for preview (with shorter text limits)
        if 'title' in slide_data and slide_data['title']:
            return slide_data['title'][:40] + ('...' if len(slide_data['title']) > 40 else '')
        elif 'subtitle' in slide_data and slide_data['subtitle']:
            return slide_data['subtitle'][:40] + ('...' if len(slide_data['subtitle']) > 40 else '')
        elif 'content' in slide_data and slide_data['content']:
            return slide_data['content'][:40] + ('...' if len(slide_data['content']) > 40 else '')
        elif 'bullet_points' in slide_data and slide_data['bullet_points']:
            first_bullet = slide_data['bullet_points'][0] if slide_data['bullet_points'] else ""
            return first_bullet[:40] + ('...' if len(first_bullet) > 40 else '')
        elif 'bullets' in slide_data and slide_data['bullets']:
            first_bullet = slide_data['bullets'][0] if slide_data['bullets'] else ""
            return first_bullet[:40] + ('...' if len(first_bullet) > 40 else '')
        else:
            return "Content pending..."
    
    def _get_slide_icon(self, slide_type: str) -> str:
        """Get the icon for a given slide type."""
        icons = {
            "title_slide": "📊",
            "problem_slide": "⚠️",
            "solution_slide": "💡",
            "features_slide": "⚙️",
            "advantage_slide": "🏆",
            "audience_slide": "👥",
            "cta_slide": "🚀",
            "market_slide": "💰",
            "roadmap_slide": "🗓️",
            "team_slide": "🧑‍🤝‍🧑"
        }
        return icons.get(slide_type, "📄")
    
    def _get_slide_title(self, slide_type: str) -> str:
        """Get a human-readable title for a slide type."""
        titles = {
            "title_slide": "Title",
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
        return titles.get(slide_type, "Slide")

# Function to use outside the class for compatibility with older code
def simulate_slide_generation_with_preview(
    content_generator_func, 
    args, 
    progress_bar, 
    progress_status, 
    progress_detail,
    preview_generator: SlidePreviewGenerator,
    slide_placeholders: List = None
):
    """Simulate the generation of slides with a live preview.
    
    Args:
        content_generator_func: The function that generates presentation content
        args: Arguments to pass to the content generator function
        progress_bar: Streamlit progress bar component
        progress_status: Streamlit component for status text
        progress_detail: Streamlit component for detail text
        preview_generator: SlidePreviewGenerator instance
        slide_placeholders: Optional list of placeholders for each slide
    
    Returns:
        Generated content
    """
    # Generate the content first
    content = content_generator_func(*args)
    
    # Get the slide count and ensure it's an integer
    try:
        slide_count = int(args[8]) if len(args) > 8 else 7
    except (ValueError, TypeError):
        slide_count = 7
    
    # Store metadata
    if 'metadata' not in content:
        content['metadata'] = {}
    content['metadata']['slide_count'] = slide_count
    
    # Define slide types in order of appearance
    standard_slides = ['title_slide', 'problem_slide', 'solution_slide', 'features_slide', 
                      'advantage_slide', 'audience_slide', 'cta_slide']
    additional_slides = ['market_slide', 'roadmap_slide', 'team_slide']
    
    # Reset preview at the beginning
    st.session_state.preview_slides = []
    
    # Set initial progress
    progress_bar.progress(15)
    
    # Calculate total slides to show based on content and slide count
    available_slides = [s for s in standard_slides + additional_slides if s in content]
    total_slides_to_show = min(slide_count, len(available_slides))
    
    # Prepare slide sequences for sequential display
    slide_sequence = []
    
    # Add standard slides first
    for slide_type in standard_slides:
        if slide_type in content and len(slide_sequence) < slide_count:
            slide_sequence.append(slide_type)
    
    # Add additional slides if needed
    for slide_type in additional_slides:
        if slide_type in content and len(slide_sequence) < slide_count:
            slide_sequence.append(slide_type)
            
    # Sequential slide preview generation with visual timing for live preview effect
    for i, slide_type in enumerate(slide_sequence):
        # Calculate progress percentage
        progress_value = 15 + int((i+1) / len(slide_sequence) * 55)
        progress_bar.progress(progress_value)
        
        # Customize message based on slide type
        if slide_type == 'title_slide':
            progress_status.markdown("### 🧠 Crafting the title slide...")
            progress_detail.markdown("_Creating an impactful introduction_")
        elif slide_type == 'problem_slide':
            progress_status.markdown("### 🧠 Articulating the problem...")
            progress_detail.markdown("_Defining the challenge your product solves_")
        elif slide_type == 'solution_slide':
            progress_status.markdown("### 🧠 Designing the solution narrative...")
            progress_detail.markdown("_Explaining how your product addresses the challenge_")
        elif slide_type == 'features_slide':
            progress_status.markdown("### 🧠 Highlighting key features...")
            progress_detail.markdown("_Emphasizing what makes your product valuable_")
        elif slide_type == 'advantage_slide':
            progress_status.markdown("### 🧠 Showcasing competitive advantages...")
            progress_detail.markdown("_Explaining why your product stands out_")
        elif slide_type == 'audience_slide':
            progress_status.markdown("### 🧠 Defining target audience...")
            progress_detail.markdown("_Identifying ideal customers for your solution_")
        elif slide_type == 'cta_slide':
            progress_status.markdown("### 🧠 Creating call to action...")
            progress_detail.markdown("_Crafting a compelling next step_")
        elif slide_type == 'market_slide':
            progress_status.markdown("### 🧠 Analyzing the market...")
            progress_detail.markdown("_Examining market size and opportunities_")
        elif slide_type == 'roadmap_slide':
            progress_status.markdown("### 🧠 Planning the roadmap...")
            progress_detail.markdown("_Outlining future development milestones_")
        elif slide_type == 'team_slide':
            progress_status.markdown("### 🧠 Introducing the team...")
            progress_detail.markdown("_Highlighting key team members and expertise_")
        
        # Add the slide to the preview with animation
        preview_generator.add_slide_preview(content[slide_type], slide_type)
        
        # Add short delay for visual effect - just enough to show sequential appearance
        # but not enough to significantly slow down the process
        time.sleep(0.4)  # Reduced from 0.8-1.2 seconds to 0.4 seconds
    
    # Store the included slides in metadata
    content['metadata']['included_slides'] = slide_sequence
    
    # Finalize progress
    progress_bar.progress(70)
    
    return content
