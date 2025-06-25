import os
import comtypes.client
import tempfile
import streamlit as st
import pythoncom
import platform

if platform.system() == "Windows":
    import comtypes.client
    
else:
    def convert_to_pdf(*args, **kwargs):
        raise NotImplementedError("PDF conversion is only supported on Windows.")


def convert_to_pdf(input_pptx_path):
    try:
        # Initialize COM for the current thread
        pythoncom.CoInitialize()
        
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = True
        
        # Create temporary path for PDF
        temp_dir = tempfile.gettempdir()
        pdf_path = os.path.join(temp_dir, "presentation.pdf")
        
        # Convert absolute paths
        abs_pptx_path = os.path.abspath(input_pptx_path)
        abs_pdf_path = os.path.abspath(pdf_path)
        
        try:
            deck = powerpoint.Presentations.Open(abs_pptx_path)
            deck.SaveAs(abs_pdf_path, 32)  # FileFormat = 32 for PDF
            deck.Close()
        finally:
            powerpoint.Quit()
            # Uninitialize COM
            pythoncom.CoUninitialize()
        
        return pdf_path
    except Exception as e:
        st.error(f"Error converting to PDF: {str(e)}")
        return None