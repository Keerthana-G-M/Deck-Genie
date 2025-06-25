import streamlit as st
from ui import render_ui
import os
import shutil

def main():
    st.set_page_config(
        page_title="Deck Genie - B2B SaaS Presentation Generator",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Check if we need to use fixed content generator
    if os.path.exists("content_generator_fixed.py"):
        try:
            # Backup original if needed
            if not os.path.exists("content_generator_original.py"):
                shutil.copy("content_generator.py", "content_generator_original.py")
            
            # Replace with fixed version
            shutil.copy("content_generator_fixed.py", "content_generator.py")
        except Exception as e:
            st.error(f"Error updating content generator: {str(e)}")
    
    # Print message to show where we are in the startup process
    print("Deck Genie is starting up - using consolidated ui.py")
    
    render_ui()

if __name__ == "__main__":
    main()