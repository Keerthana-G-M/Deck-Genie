import requests
import io
import json
import google.generativeai as genai
from PIL import Image, ImageDraw, ImageFont
from typing import Optional, Dict, Any, Tuple, List
from utils import get_api_key, IMAGE_KEYWORDS
from slide_content_generators import MODEL_NAME
import random
import os
import streamlit as st

def fetch_image_for_slide(slide_type: str, context: Dict = None, use_placeholders: bool = False, search_terms: List[str] = None):
    """Fetch an appropriate image for a slide type."""
    # If using placeholders, return those directly
    if use_placeholders:
        return get_placeholder_image(slide_type)
        
    # Initialize query
    query = None
    
    # Try to get query from search terms first
    if isinstance(search_terms, str):
        # If search_terms is already a string, use it directly
        query = search_terms
    elif isinstance(search_terms, list) and search_terms:
        # If search_terms is a list, join the first 3 terms
        query = " ".join(search_terms[:3])
    
    # If no valid query yet, try other methods
    if not query:
        # Default queries based on slide type
        slide_type_base = slide_type.replace('_slide', '')  # Remove '_slide' suffix if present
        if slide_type_base in IMAGE_KEYWORDS:
            query = IMAGE_KEYWORDS[slide_type_base]
        else:
            # Generate query based on slide type and context
            try:
                if context:
                    query = generate_image_query_with_gemini(slide_type_base, context)
                else:
                    query = "business professional technology"
            except Exception as e:
                print(f"Error generating query with Gemini: {e}")
                query = get_fallback_query(slide_type_base, context)
    
    # Ensure query is a string
    if not isinstance(query, str):
        print(f"Invalid query type ({type(query)}), using fallback")
        query = get_fallback_query(slide_type.replace('_slide', ''), context)
    
    print(f"Using image search query: {query}")
    
    # Try to fetch from Unsplash
    max_retries = 3
    for attempt in range(max_retries):
        try:
            image_data = fetch_image_from_unsplash(query)
            if image_data:
                return image_data
        except Exception as e:
            print(f"Error in fetch_image_from_unsplash (attempt {attempt + 1}): {e}")
            if attempt == max_retries - 1:  # Last attempt
                print(f"Using placeholder image for {slide_type} slide after {max_retries} failed attempts")
                return get_placeholder_image(slide_type)
            
    # If we get here, all attempts failed, return a placeholder
    print(f"Using placeholder image for {slide_type} slide")
    return get_placeholder_image(slide_type)

def generate_image_query_with_gemini(slide_type: str, context: Dict[str, Any]) -> str:
    """
    Use Gemini API to generate a relevant image search query based on slide content.
    """
    try:
        # Configure the Gemini API
        genai.configure(api_key=get_api_key("GEMINI_API_KEY"))
        
        # Get slide content based on type
        slide_content = context.get(slide_type, {}) if slide_type in context else {}
        metadata = context.get('metadata', {})
        company_name = metadata.get('company_name', '')
        product_name = metadata.get('product_name', '')
        
        # Create base prompt for all slide types
        base_prompt = """
        Generate a specific, professional Unsplash search query (3-5 words) for a business presentation slide.
        Focus on high-quality, modern business imagery that would be suitable for a professional presentation.
        
        Context:
        Company: {company}
        Product: {product}
        Slide Type: {slide_type}
        
        Content to match:
        {content_details}
        
        IMPORTANT GUIDELINES:
        - Focus on B2B/Enterprise/Professional contexts
        - Avoid generic or cliché business images
        - Prefer modern, tech-focused imagery
        - Ensure visual relevance to the content
        - Consider industry-specific imagery when applicable
        
        Format your response as a JSON object with a single key "query" and its string value.
        Example: {{"query": "modern enterprise technology"}}
        """
        
        # Customize content details based on slide type
        content_details = ""
        if slide_type == "problem":
            content_details = f"""
            Title: {slide_content.get('title', 'Problem Statement')}
            Key Points:
            {' '.join(slide_content.get('bullets', slide_content.get('pain_points', [''])))}
            Focus: Business challenges and pain points visualization
            """
        
        elif slide_type == "solution":
            content_details = f"""
            Title: {slide_content.get('title', 'Our Solution')}
            Solution Description: {slide_content.get('paragraph', '')}
            Focus: Modern technology solutions and innovations
            """
        
        elif slide_type == "features":
            features = slide_content.get('features', [])
            content_details = f"""
            Title: {slide_content.get('title', 'Key Features')}
            Features: {', '.join(features[:3] if features else [])}
            Focus: Product features and capabilities visualization
            """
        
        elif slide_type == "advantage":
            content_details = f"""
            Title: {slide_content.get('title', 'Our Advantage')}
            Advantages: {' '.join(slide_content.get('bullets', slide_content.get('differentiators', [])))}
            Focus: Competitive advantage and business success
            """
        
        elif slide_type == "audience":
            content_details = f"""
            Title: {slide_content.get('title', 'Target Audience')}
            Audience: {slide_content.get('paragraph', '')}
            Focus: Professional business audience and market segments
            """
        
        elif slide_type == "market":
            content_details = f"""
            Title: {slide_content.get('title', 'Market Opportunity')}
            Market Size: {slide_content.get('market_size', '')}
            Growth: {slide_content.get('growth_rate', '')}
            Focus: Market analysis and business growth
            """
        
        elif slide_type == "roadmap":
            phases = slide_content.get('phases', [])
            phase_names = [phase.get('name', '') for phase in phases]
            content_details = f"""
            Title: {slide_content.get('title', 'Product Roadmap')}
            Phases: {', '.join(phase_names)}
            Focus: Strategic planning and product development
            """
        
        elif slide_type == "team":
            content_details = f"""
            Title: {slide_content.get('title', 'Our Team')}
            Focus: Professional leadership and team excellence
            """
        
        else:
            content_details = f"""
            Title: {slide_content.get('title', '')}
            Focus: Professional business context
            """
        
        # Format the complete prompt
        prompt = base_prompt.format(
            company=company_name,
            product=product_name,
            slide_type=slide_type,
            content_details=content_details
        )
        
        # Call Gemini API with the enhanced prompt
        model = genai.GenerativeModel(
            model_name=MODEL_NAME,
            generation_config={"temperature": 0.2, "max_output_tokens": 100}
        )
        
        response = model.generate_content([{"role": "user", "parts": [prompt]}])
        
        # Parse the JSON response
        try:
            response_text = response.text
            content_start = response_text.find('{')
            content_end = response_text.rfind('}') + 1
            
            if content_start >= 0 and content_end > content_start:
                json_content = response_text[content_start:content_end]
                result = json.loads(json_content)
                
                if "query" in result:
                    # Enhance the query with industry-specific terms if available
                    query = result["query"]
                    if "software" in product_name.lower() or "tech" in company_name.lower():
                        query += " technology"
                    elif "health" in product_name.lower() or "medical" in company_name.lower():
                        query += " healthcare"
                    elif "finance" in product_name.lower() or "bank" in company_name.lower():
                        query += " finance"
                    return query
                
        except Exception as e:
            print(f"Error parsing Gemini response: {str(e)}")
        
        # Fallback to predefined keywords if parsing fails
        return get_fallback_query(slide_type, context)
            
    except Exception as e:
        print(f"Error generating image query with Gemini: {str(e)}")
        return get_fallback_query(slide_type, context)

def get_fallback_query(slide_type: str, context: Dict[str, Any] = None) -> str:
    """Get a fallback search query for a given slide type."""
    # Base queries for each slide type
    fallback_queries = {
        'title': "professional business presentation modern",
        'problem': "business challenge problem modern",
        'solution': "business solution technology modern",
        'features': "product features technology modern",
        'advantage': "competitive advantage business success",
        'audience': "business professionals meeting modern",
        'market': "market analysis business growth chart",
        'roadmap': "strategic roadmap business planning",
        'team': "professional business team modern",
        'cta': "business call to action modern",
        'success': "business success achievement modern",
        'impact': "business impact results modern",
        'future': "future business innovation modern"
    }
    
    # Get base slide type without '_slide' suffix
    base_type = slide_type.replace('_slide', '')
    
    # Get the fallback query
    query = fallback_queries.get(base_type, "professional business modern")
    
    # Add context-specific terms if available
    if context and isinstance(context, dict):
        # Add industry context if available
        industry = context.get('metadata', {}).get('industry', '').lower()
        if industry:
            if 'tech' in industry or 'software' in industry:
                query += " technology"
            elif 'health' in industry or 'medical' in industry:
                query += " healthcare"
            elif 'finance' in industry or 'bank' in industry:
                query += " finance"
            elif 'retail' in industry:
                query += " retail"
            elif 'education' in industry:
                query += " education"
            elif 'manufacturing' in industry:
                query += " manufacturing"
        
        # Add product context if available
        product = context.get('metadata', {}).get('product_name', '').lower()
        if product and len(product.split()) == 1:  # Only add if it's a single word
            query = f"{product} {query}"
    
    # Ensure the query is clean and properly formatted
    query = " ".join(query.split())  # Remove extra spaces
    return query

def get_slide_icon(slide_type: str) -> Dict[str, Any]:
    """
    Get a default icon for a slide type.
    This is used as a fallback when images can't be fetched.
    
    Args:
        slide_type: Type of slide
        
    Returns:
        Dictionary with icon information
    """
    from utils import SLIDE_ICONS
    
    icon = SLIDE_ICONS.get(slide_type, "📊")
    
    return {
        "icon": icon,
        "size": (32, 32)
    }

def create_placeholder_image(slide_type: str) -> io.BytesIO:
    """
    Create a placeholder image when Unsplash fetch fails.
    
    Args:
        slide_type: Type of slide to create placeholder for
        
    Returns:
        BytesIO object containing a simple placeholder image
    """
    try:
        # Create a simple colored rectangle
        width, height = 800, 600
        colors = {
            "problem": "#FF6B6B",  # Red
            "solution": "#4ECDC4", # Teal  
            "advantage": "#45B7D1", # Blue
            "audience": "#96CEB4",  # Green
            "features": "#FFEAA7",  # Yellow
            "call_to_action": "#DDA0DD"  # Plum
        }
        
        color = colors.get(slide_type, "#95A5A6")  # Default gray
        
        # Create image
        image = Image.new('RGB', (width, height), color)
        draw = ImageDraw.Draw(image)
        
        # Add text
        text = f"{slide_type.upper()}\nIMAGE"
        
        try:
            # Try to use a default font
            font = ImageFont.truetype("Arial.ttf", 48)
        except:
            # Fallback to default font
            font = ImageFont.load_default()
        
        # Get text size and center it
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        
        # Draw text with outline for better visibility
        outline_color = "#FFFFFF" if slide_type != "features" else "#000000"
        for adj in range(-2, 3):
            for adj2 in range(-2, 3):
                draw.text((x+adj, y+adj2), text, font=font, fill=outline_color)
        
        # Draw main text
        main_color = "#000000" if slide_type == "features" else "#FFFFFF"
        draw.text((x, y), text, font=font, fill=main_color)
        
        # Save to BytesIO
        img_buffer = io.BytesIO()
        image.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        return img_buffer
        
    except Exception as e:
        print(f"Error creating placeholder image: {str(e)}")
        # Return a minimal BytesIO with empty content as last resort
        return io.BytesIO()

def get_placeholder_image(slide_type: str):
    """
    Return a placeholder image for a given slide type when no image can be fetched.
    
    Args:
        slide_type: The type of slide that needs a placeholder image
        
    Returns:
        BytesIO object containing a basic placeholder image
    """
    # Create a basic image with slide type as text
    width, height = 800, 600
    background_colors = {
        "problem": (230, 230, 250),  # Lavender
        "solution": (240, 255, 240),  # Honeydew
        "advantage": (255, 240, 245),  # Lavender blush
        "audience": (240, 248, 255),  # Alice blue
        "market": (255, 250, 240),    # Floral white
        "roadmap": (245, 255, 250),   # Mint cream
        "team": (255, 245, 238),      # Seashell
        "features": (240, 255, 255),  # Azure
        "cta": (255, 255, 240)        # Ivory
    }
    
    # Default background color
    bg_color = background_colors.get(slide_type.lower(), (245, 245, 245))
    
    # Create image
    img = Image.new('RGB', (width, height), color=bg_color)
    draw = ImageDraw.Draw(img)
    
    # Try to use a system font
    try:
        # Try common system fonts
        font_options = ['Arial', 'Verdana', 'Tahoma', 'Calibri', 'Georgia']
        font = None
        for font_name in font_options:
            try:
                font = ImageFont.truetype(font_name, 40)
                break
            except IOError:
                continue
                
        if font is None:
            # Fallback to default font
            font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()
    
    # Draw placeholder text and design elements
    title = f"{slide_type.capitalize()} Slide"
    
    # Add a border
    border_margin = 20
    draw.rectangle([border_margin, border_margin, width-border_margin, height-border_margin], 
                  outline=(100, 100, 100), width=2)
    
    # Add centered text
    text_width, text_height = getattr(draw, 'textsize', lambda text, font: (200, 40))(title, font=font)
    position = ((width - text_width) // 2, (height - text_height) // 2)
    draw.text(position, title, fill=(80, 80, 80), font=font)
    
    # Add some design elements based on slide type
    for i in range(10):
        x = random.randint(50, width - 50)
        y = random.randint(50, height - 50)
        size = random.randint(10, 40)
        opacity = random.randint(30, 100)
        draw.ellipse([x, y, x+size, y+size], 
                    fill=(opacity, opacity, opacity, opacity))
    
    # Convert to BytesIO
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    
    return img_bytes

def fetch_image_from_unsplash(query: str) -> Optional[io.BytesIO]:
    """
    Fetch an image from Unsplash API based on a query.
    
    Args:
        query: Search term for the image
        
    Returns:
        BytesIO object containing the image or None if fetch fails
    """
    try:
        # Get the API key from environment
        api_key = os.environ.get("UNSPLASH_API_KEY")
        if not api_key:
            api_key = st.secrets.get("UNSPLASH_API_KEY")
        
        if not api_key:
            print("No Unsplash API key found. Using placeholder image.")
            return None
        
        # Prepare the request
        headers = {
            "Authorization": f"Client-ID {api_key}"
        }
        
        # Clean up the query to make it more likely to succeed
        safe_query = query.replace("[Product Name]", "product").strip()
        if not safe_query or len(safe_query) < 3:
            safe_query = "business professional"
            
        params = {
            "query": safe_query,
            "orientation": "landscape",
            "per_page": 5  # Fetch multiple options
        }
        
        # Make the API call with a timeout
        response = requests.get(
            "https://api.unsplash.com/search/photos",
            headers=headers,
            params=params,
            timeout=5  # 5 second timeout
        )
        
        # Check if the request was successful
        if response.status_code == 200:
            data = response.json()
            results = data.get("results", [])
            
            # If we have results, pick a random one
            if results:
                chosen_image = random.choice(results)
                image_url = chosen_image["urls"]["regular"]
                
                # Fetch the actual image
                image_response = requests.get(image_url, timeout=5)
                if image_response.status_code == 200:
                    # Return as BytesIO
                    image_data = io.BytesIO(image_response.content)
                    return image_data
        
        print(f"Unsplash API error: {response.status_code}, {response.text}")
        return None
        
    except Exception as e:
        print(f"Error fetching image from Unsplash: {e}")
        return None