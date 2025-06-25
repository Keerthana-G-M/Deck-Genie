import streamlit as st
import os
import io
from datetime import datetime
from content_generator import generate_presentation_content
from ppt_generator import create_presentation
from utils import sanitize_filename
from preview_generator import SlidePreviewGenerator, simulate_slide_generation_with_preview
from slide_editor import SlideEditor
from pdf_converter import convert_to_pdf

def load_demo_data():
    """Load demo data for presentation fields."""
    return {
        "company_name": "SecureLayer",
        "product_name": "ThreatGuard AI",
        "target_audience": "IT security teams at mid-sized enterprises and managed security service providers (MSSPs)",
        "problem_statement": "Mid-sized enterprises face a 3x higher risk of undetected cyber threats due to limited resources and fragmented security tools. Security teams spend 40% of their time on manual log analysis, leading to delayed responses and increased breach costs. Compliance requirements add complexity, with 62% of organizations struggling to keep up with evolving standards.",
        "key_features": "AI-Powered Threat Detection: Identify zero-day attacks and anomalies in real time\nAutomated Incident Response: Orchestrate and remediate threats with one click\nUnified Security Dashboard: Centralize alerts, logs, and compliance status\nContinuous Compliance Monitoring: Map controls to SOC2, ISO 27001, and GDPR\nSeamless Integration: Connect with 50+ security and IT tools\n24/7 Expert Support: Access cybersecurity experts anytime",
        "competitive_advantage": "ThreatGuard AI reduces mean time to detect (MTTD) by 80% and automates 90% of incident response actions. Unlike legacy SIEMs, our platform deploys in under 30 minutes with no coding required. Customers report a 60% reduction in compliance audit preparation time and 2x faster breach containment.",
        "call_to_action": "Book a free security assessment today"
    }

def render_ui():
    """Render the Streamlit UI for the presentation generator."""
    st.title("Deck Genie: B2B SaaS Presentation Generator")
    st.subheader("Create a minimalist, executive-ready product overview in seconds")
    
    # Initialize session state for demo toggle
    if "use_demo" not in st.session_state:
        st.session_state.use_demo = False
    
    # Initialize session state for presentation generation
    if "presentation_generated" not in st.session_state:
        st.session_state.presentation_generated = False
    
    if "original_content" not in st.session_state:
        st.session_state.original_content = None
    
    # Ensure deleted_slides is always properly initialized as a set of strings
    if 'deleted_slides' not in st.session_state:
        st.session_state.deleted_slides = set()
    elif not isinstance(st.session_state.deleted_slides, set):
        st.session_state.deleted_slides = set()
    else:
        try:
            st.session_state.deleted_slides = {
                item for item in st.session_state.deleted_slides 
                if isinstance(item, str)
            }
        except (TypeError, AttributeError):
            st.session_state.deleted_slides = set()
    
    def toggle_demo():
        st.session_state.use_demo = not st.session_state.use_demo
    
    def reset_presentation():
        st.session_state.presentation_generated = False
        st.session_state.original_content = None
        # Clear editor state explicitly
        editor_keys = ['editor_content', 'slide_order', 'deleted_slides', 'has_modifications', 'uploaded_images', 'preview_slides', 'persistent_preview_generator']
        for key in editor_keys:
            if key in st.session_state:
                del st.session_state[key]
        for key in list(st.session_state.keys()):
            if key.startswith(('title_', 'bullet_', 'feature_', 'paragraph_', 'cta_', 'subtitle_', 'delete_', 'add_', 'image_', 'remove_')):
                del st.session_state[key]
    
    # Show reset button if presentation is generated
    if st.session_state.get('presentation_generated', False):
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🔄 Generate New Presentation", key="reset_btn"):
                reset_presentation()
                st.rerun()
    
    # Only show the form if no presentation is generated
    if not st.session_state.get('presentation_generated', False):
        # Checkbox outside the form to toggle demo data
        st.checkbox("Use Demo Data", value=st.session_state.use_demo, on_change=toggle_demo)
        
        with st.form("presentation_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                demo_data = load_demo_data()
                
                # Simplified form values logic
                form_values = demo_data if st.session_state.use_demo else {}

                company_name = st.text_input("Company Name", value=form_values.get("company_name", ""))
                product_name = st.text_input("Product Name", value=form_values.get("product_name", ""))
                target_audience = st.text_input("Target Audience", value=form_values.get("target_audience", ""))
                problem_statement = st.text_area("Problem Statement (max 200 words)", 
                                                   value=form_values.get("problem_statement", ""), height=150)
                key_features = st.text_area("Key Features (one per line, max 6)", 
                                              value=form_values.get("key_features", ""), height=150)
                competitive_advantage = st.text_area("Competitive Advantage (max 150 words)",
                                                       value=form_values.get("competitive_advantage", ""), height=150)
                call_to_action = st.text_input("Call to Action", value=form_values.get("call_to_action", ""))
                    
                st.markdown("### Presentation Customization")
                
                # Enhanced persona options
                persona_options = ["Generic", "Technical", "Marketing", "Executive", "Investor"]
                persona = st.selectbox(
                    "Target Persona",
                    options=persona_options,
                    index=persona_options.index(form_values.get("persona", "Generic")) if "persona" in form_values else 0,
                    help="Choose the primary audience persona for language adaptation"
                )
                    
                slide_count = st.slider(
                    "Number of Slides",
                    min_value=5, max_value=10,
                    value=form_values.get("slide_count", 7),
                    help="Choose how many slides you want in your presentation (typically 5-10)"
                )
            
            with col2:
                st.markdown("### Instructions")
                st.markdown("""
                1. Fill out all fields with your product information
                2. Choose persona and slide count
                3. Click 'Generate Presentation' to create your deck
                4. Download the finished PowerPoint file
                
                **Tips for best results:**
                - Be specific in your problem statement
                - Limit key features to the most important 4-6 items
                - Highlight measurable advantages over competitors
                - Keep your call to action simple and direct
                """)
                
                st.markdown("### Preview Scope")
                st.markdown("Your presentation will include standard slides like Title, Problem, Solution, Features, Advantage, Audience, and CTA. Additional slides (Market, Roadmap, Team) may be included based on persona and slide count.")

            submitted = st.form_submit_button("Generate Presentation")
        
        if submitted:
            # Simplified error validation
            errors = []
            if not company_name or not product_name: 
                errors.append("Company and product names are required")
            if not target_audience: 
                errors.append("Target audience is required")
            if not problem_statement: 
                errors.append("Problem statement is required")
            elif len(problem_statement.split()) > 200: 
                errors.append("Problem statement exceeds 200 words")
            if not key_features: 
                errors.append("Key features are required")
            else:
                features_list = [f for f in key_features.strip().split('\n') if f]
                if len(features_list) > 6: 
                    errors.append("Maximum of 6 key features allowed")
            if not competitive_advantage: 
                errors.append("Competitive advantage is required")
            elif len(competitive_advantage.split()) > 150: 
                errors.append("Competitive advantage exceeds 150 words")
            if not call_to_action: 
                errors.append("Call to action is required")
                
            if errors:
                for error in errors: 
                    st.error(error)
            else:
                try:
                    preview_generator = SlidePreviewGenerator()
                    
                    progress_container = st.container()
                    with progress_container:
                        progress_bar = st.progress(0)
                        progress_status = st.empty()
                        progress_detail = st.empty()
                        progress_status.markdown("### 🚀 Starting the creation process...")
                        progress_detail.markdown("_Analyzing your inputs..._")
                        progress_bar.progress(10)
                        
                        # Create the preview container before showing balloons to ensure it appears first
                        preview_container = st.container()
                        with preview_container:
                            st.markdown("### 🔍 Live Preview")
                            preview_generator.create_preview_container()
                        st.session_state.preview_shown_during_generation = True
                        
                        # Show balloons after creating container
                        st.balloons()
                        
                        features_list = [f for f in key_features.strip().split('\n') if f]
                        
                        content = simulate_slide_generation_with_preview(
                            generate_presentation_content,
                            (company_name, product_name, target_audience, problem_statement, 
                             features_list, competitive_advantage, call_to_action, 
                             persona, slide_count),
                            progress_bar, progress_status, progress_detail,
                            preview_generator,
                            None  # Pass None for slide_placeholders as it's not used
                        )
                        
                        progress_status.markdown("### 🎨 Designing your slides...")
                        progress_detail.markdown("_Creating a visually appealing presentation..._")
                        progress_bar.progress(70)
                        
                        # Store slide count and other metadata AFTER content is generated
                        if 'metadata' not in content:
                            content['metadata'] = {}
                        content['metadata']['slide_count'] = slide_count
                        content['metadata']['company_name'] = company_name
                        content['metadata']['product_name'] = product_name
                        
                        progress_status.markdown("### 📊 Assembling your presentation...")
                        progress_detail.markdown("_Putting everything together in a cohesive package_")
                        progress_bar.progress(90)
                        
                        filename = f"{sanitize_filename(company_name)}_{sanitize_filename(product_name)}_Overview.pptx"
                        presentation_buffer = create_presentation(content, filename)
                        
                        progress_status.markdown("### ✨ Polishing final details...")
                        progress_detail.markdown("_Adding final touches..._")
                        progress_bar.progress(100)
                        
                        progress_container.empty()
                        st.success("✅ Your presentation is ready!")
                        
                        st.session_state.original_content = content
                        st.session_state.presentation_generated = True
                        
                        if 'original_images_cache' not in st.session_state:
                            st.session_state.original_images_cache = {}
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.exception(e)
    
    # Render the slide editor after generation (outside the form)
    if st.session_state.get('presentation_generated', False) and st.session_state.get('original_content'):
        # Create a persistent preview generator for real-time updates
        if 'persistent_preview_generator' not in st.session_state:
            st.session_state.persistent_preview_generator = SlidePreviewGenerator()
        preview_generator = st.session_state.persistent_preview_generator
        
        # Only show preview if it hasn't been shown during generation
        if not st.session_state.get('preview_shown_during_generation', False):
            preview_generator.create_preview_container()
            if st.session_state.get('original_content'):
                preview_generator.update_preview_with_content(st.session_state.original_content)
        
        # Download original presentation section
        st.markdown("---")
        st.markdown("## 📥 Download Your Presentation")
        
        try:
            original_content_meta = st.session_state.original_content['metadata']
            original_filename = f"{sanitize_filename(original_content_meta['company_name'])}_{sanitize_filename(original_content_meta['product_name'])}_Overview.pptx"
            
            # Cache the original presentation buffer to prevent API calls on every re-render
            cache_key = f"original_buffer_{id(st.session_state.original_content)}"
            if cache_key not in st.session_state:
                st.session_state[cache_key] = create_presentation(st.session_state.original_content, original_filename)
            original_buffer = st.session_state[cache_key]
            
            st.download_button(
                label="📥 Download Original Presentation",
                data=original_buffer,
                file_name=original_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
            # Add PDF download button
            if st.button("📄 Download as PDF", use_container_width=True):
                pptx_path = os.path.abspath(original_filename)
                # Save the presentation buffer to a temporary file
                with open(pptx_path, "wb") as f:
                    f.write(original_buffer.getvalue())
    
                pdf_path = convert_to_pdf(pptx_path)
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                        st.download_button(
                            label="📥 Download PDF Version",
                            data=pdf_bytes,
                            file_name=f"{os.path.splitext(original_filename)[0]}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    # Clean up temporary files
                    os.remove(pdf_path)
                    os.remove(pptx_path)
                else:
                    st.error("Failed to convert presentation to PDF")
        except Exception as e:
            st.error(f"Error creating original presentation: {str(e)}")
        
        # Customization section
        st.markdown("---")
        st.markdown("## 🛠️ Customize Your Presentation")
        
        # Initialize slide order using our new function if not already done
        from slide_reordering import initialize_slide_order, render_slide_reordering_ui, generate_reordered_presentation
        
        # Initialize slide_order only when first showing this section
        if 'slide_order' not in st.session_state:
            st.session_state.slide_order = initialize_slide_order(st.session_state.original_content)
        
        st.markdown("### Reorder your slides using the buttons below")
        render_slide_reordering_ui(st.session_state.original_content)
        
        # Download reordered presentation if modifications have been made
        if st.session_state.get('has_modifications', False):
            st.markdown("---")
            st.markdown("## 🔄 Download Reordered Version")
            st.markdown("✅ **Slide order changed** - Download your reordered presentation")
            
            try:
                # Get metadata for filenames
                original_content_meta = st.session_state.original_content.get('metadata', {})
                company_name = original_content_meta.get('company_name', "Company")
                product_name = original_content_meta.get('product_name', "Product")
                
                # Create filename
                modified_filename = f"{sanitize_filename(company_name)}_{sanitize_filename(product_name)}_Reordered.pptx"
                
                # Generate the reordered presentation
                reordered_buffer = generate_reordered_presentation(
                    st.session_state.original_content,
                    st.session_state.slide_order,
                    modified_filename
                )
                
                # Show download button if we have valid data
                if reordered_buffer:
                    st.download_button(
                        label="📥 Download Reordered Presentation",
                        data=reordered_buffer,
                        file_name=modified_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
            except Exception as e:
                st.error(f"Error creating reordered presentation: {str(e)}")
        
# This ensures that the render_ui function is available when imported
if __name__ == "__main__":
    render_ui()
