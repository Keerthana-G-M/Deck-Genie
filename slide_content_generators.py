import google.generativeai as genai
from typing import Dict, List, Any
import json
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Cache to store generated content
_content_cache = {}
# Cache to store fetched images (path or URL mapped to image data)
_image_cache = {}

# Always use the same model for consistency
MODEL_NAME = 'gemini-1.5-flash'

# Function to get an image from cache or fetch a new one if needed
def get_image(image_path_or_url, fetch_function=None, force_refresh=False):
    """
    Get an image from cache or fetch a new one if needed.
    
    Args:
        image_path_or_url: Path or URL of the image
        fetch_function: Function to fetch the image if not in cache
        force_refresh: Force refresh the image from source
    
    Returns:
        The image data
    """
    if not force_refresh and image_path_or_url in _image_cache:
        return _image_cache[image_path_or_url]
    
    if fetch_function:
        image_data = fetch_function(image_path_or_url)
        _image_cache[image_path_or_url] = image_data
        return image_data
    
    return None

# Function to cache all images in a presentation
def cache_presentation_images(presentation_data):
    """
    Cache all images in a presentation.
    
    Args:
        presentation_data: The presentation data containing slides with images
    """
    global _image_cache
    
    # Extract image URLs/paths from presentation_data and store in _image_cache
    if isinstance(presentation_data, dict) and 'slides' in presentation_data:
        for slide in presentation_data['slides']:
            if 'background_image' in slide and slide['background_image']:
                _image_cache[slide['background_image']] = True
            
            # Handle other image properties based on your presentation structure
            # This is a generic example - adapt to your actual data structure
            if 'images' in slide:
                for img in slide['images']:
                    if 'url' in img:
                        _image_cache[img['url']] = True

# Function to clear image cache
def clear_image_cache():
    """Clear the image cache."""
    global _image_cache
    _image_cache = {}

# Function to get a copy of the current image cache
def get_image_cache():
    """Get a copy of the current image cache."""
    return _image_cache.copy()

# Function to set the image cache with provided data
def set_image_cache(cache_data):
    """Set the image cache with provided data."""
    global _image_cache
    _image_cache = cache_data

def generate_title_slide_content(product_name: str, company_name: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate title slide content."""
    cache_key = f"title_{product_name}_{company_name}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        logger.info(f"Using cached title slide content for {product_name}")
        return _content_cache[cache_key]
    
    logger.info(f"Generating title slide content for {product_name}")
    model = genai.GenerativeModel(MODEL_NAME)
    
    prompt = f"""
    Create a professional title slide for a business presentation with these details:
    - Product name: {product_name}
    - Company name: {company_name}
    
    Return a JSON object with these fields:
    - title: The main title for the slide (should include the product name)
    - subtitle: A subtitle (should include the company name)
    - product_name: The product name for reference
    
    Make the title concise but impactful. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating title slide: {e}")
        # Provide fallback content
        content = {
            "title": f"{product_name}",
            "subtitle": f"Presented by {company_name}",
            "product_name": product_name
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{product_name}"
    if not content.get("subtitle"):
        content["subtitle"] = f"Presented by {company_name}"
    content["product_name"] = product_name
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_problem_slide_content(problem_statement: str, persona: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate problem slide content with persona-specific focus."""
    cache_key = f"problem_{problem_statement}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        logger.info(f"Using cached problem slide content for {persona} persona")
        return _content_cache[cache_key]
    
    logger.info(f"Generating problem slide content for {persona} persona")
    model = genai.GenerativeModel(MODEL_NAME)
    
    # Enhanced persona-specific prompts
    persona_prompts = {
        "Technical": """Focus on technical pain points, system limitations, performance issues, and engineering challenges. 
        Emphasize scalability problems, integration difficulties, maintenance overhead, and technical debt.""",
        
        "Marketing": """Focus on customer pain points, market gaps, user experience issues, and customer acquisition challenges. 
        Emphasize customer dissatisfaction, market demand, competitive weaknesses, and opportunities for differentiation.""",
        
        "Executive": """Focus on business impact, strategic challenges, operational inefficiencies, and competitive threats. 
        Emphasize revenue impact, market positioning, resource allocation, and strategic risks.""",
        
        "Investor": """Focus on market inefficiencies, unmet demand, scalable problems, and monetizable pain points. 
        Emphasize market size of problems, competitive gaps, and opportunities for disruption and returns.""",
        
        "Generic": """Focus on balanced mix of technical and business challenges that affect multiple stakeholders."""
    }
    
    persona_specific = persona_prompts.get(persona, persona_prompts["Generic"])
    
    prompt = f"""
    Create a problem statement slide for {persona} audience based on this description:
    "{problem_statement}"
    
    {persona_specific}
    
    Make the content highly relevant and actionable for {persona} professionals.
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the problem (tailored for {persona} audience)
    - pain_points: An array of 3-4 specific pain points extracted/derived from the problem statement
    
    Format as valid JSON only. Each bullet point should be specific to {persona} concerns.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating problem slide: {e}")
        # Provide fallback content
        content = {
            "title": "The Problem",
            "pain_points": [s.strip() + "." for s in problem_statement.split(". ") if s.strip()][:3]
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "The Problem"
    if not content.get("pain_points") or not isinstance(content.get("pain_points"), list):
        # Extract simple bullet points from the problem statement
        content["pain_points"] = [s.strip() + "." for s in problem_statement.split(". ") if s.strip()][:3]
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_solution_slide_content(product_name: str, problem_statement: str, persona: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate solution slide content with persona-specific value proposition."""
    cache_key = f"solution_{product_name}_{problem_statement}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    # Enhanced persona-specific solution focus
    persona_prompts = {
        "Technical": """Focus on technical architecture, implementation approach, engineering solutions, and technical benefits. 
        Emphasize scalability, performance, integration capabilities, and technical innovation.""",
        
        "Marketing": """Focus on customer value proposition, user benefits, market differentiation, and customer success. 
        Emphasize customer outcomes, user experience improvements, and competitive advantages.""",
        
        "Executive": """Focus on business impact, strategic value, operational improvements, and competitive positioning. 
        Emphasize ROI, efficiency gains, market opportunities, and strategic advantages.""",
        
        "Investor": """Focus on market opportunity, scalable solution, revenue potential, and competitive moats. 
        Emphasize addressable market, growth potential, monetization strategy, and sustainable advantages.""",
        
        "Generic": """Focus on balanced technical and business value that appeals to multiple stakeholders."""
    }
    
    persona_specific = persona_prompts.get(persona, persona_prompts["Generic"])
    
    prompt = f"""
    Create a solution overview slide for {persona} audience for "{product_name}" that addresses:
    "{problem_statement}"
    
    {persona_specific}
    
    Make the solution compelling and specific to {persona} decision-making criteria.
    
    Return a JSON object with these fields:
    - title: A compelling slide title introducing the solution (include product name, tailored for {persona})
    - paragraph: A focused paragraph (60-80 words) explaining how {product_name} solves the problem for {persona} audience
    
    The paragraph should clearly articulate value specific to {persona} professionals.
    Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating solution slide: {e}")
        # Provide fallback content
        content = {
            "title": f"Introducing {product_name}",
            "paragraph": f"{product_name} is designed to address these challenges with an innovative approach that helps organizations achieve better outcomes."
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"Introducing {product_name}"
    if not content.get("paragraph"):
        content["paragraph"] = f"{product_name} is designed to address these challenges with an innovative approach that helps organizations achieve better outcomes."
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_features_slide_content(key_features: List[str], persona: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate features slide content."""
    cache_key = f"features_{'_'.join(key_features[:3])}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    # Limit to 5 features
    features_text = "\n".join(f"- {f}" for f in key_features[:5])
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical capabilities and specifications."
    elif persona == "Marketing":
        persona_specific = "Focus on benefits and value to customers."
    elif persona == "Executive":
        persona_specific = "Focus on business impact and strategic advantages."
    elif persona == "Investor":
        persona_specific = "Focus on market differentiation and competitive advantages."
    prompt = f"""
    Create a key features slide based on these features:
    {features_text}
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A slide title for key features
    - features: A refined list of the key features (keep the same features but enhance their descriptions)
    
    Each feature should be concise but clear. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating features slide: {e}")
        # Provide fallback content
        content = {
            "title": "Key Features",
            "features": key_features[:5]
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Key Features"
    if not content.get("features") or not isinstance(content.get("features"), list):
        content["features"] = key_features[:5]
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_advantage_slide_content(competitive_advantage: str, persona: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate competitive advantage slide content."""
    cache_key = f"advantage_{competitive_advantage}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical superiority and unique capabilities."
    elif persona == "Marketing":
        persona_specific = "Focus on market differentiation and customer benefits."
    elif persona == "Executive":
        persona_specific = "Focus on business impact, ROI, and strategic advantages."
    elif persona == "Investor":
        persona_specific = "Focus on market positioning, barriers to entry, and sustainable competitive edges."
    
    prompt = f"""
    Create a competitive advantage slide based on this statement:
    "{competitive_advantage}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about competitive advantages
    - differentiators: An array of 3-4 key differentiators or advantages
    
    Make each differentiator concise but specific. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating advantage slide: {e}")
        # Provide fallback content
        content = {
            "title": "Our Competitive Advantage",
            "differentiators": [s.strip() + "." for s in competitive_advantage.split(". ") if s.strip()][:3]
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Our Competitive Advantage"
    if not content.get("differentiators") or not isinstance(content.get("differentiators"), list):
        # Extract simple bullet points from the competitive advantage
        content["differentiators"] = [s.strip() + "." for s in competitive_advantage.split(". ") if s.strip()][:3]
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_audience_slide_content(target_audience: str, persona: str, use_cache: bool = False) -> Dict[str, Any]:
    """Generate target audience slide content."""
    cache_key = f"audience_{target_audience}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical roles, use cases, and technical needs."
    elif persona == "Marketing":
        persona_specific = "Focus on demographics, psychographics, and audience pain points."
    elif persona == "Executive":
        persona_specific = "Focus on key decision-makers, organizational impact, and business needs."
    elif persona == "Investor":
        persona_specific = "Focus on market segments, target market size, and customer acquisition strategy."
    
    prompt = f"""
    Create a target audience slide based on this description:
    "{target_audience}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the target audience
    - paragraph: A concise paragraph (60-80 words) that provides a rich description of the target audience
    
    Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating audience slide: {e}")
        # Provide fallback content
        content = {
            "title": "Target Audience",
            "paragraph": target_audience
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Target Audience"
    if not content.get("paragraph"):
        content["paragraph"] = target_audience
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_cta_slide_content(call_to_action: str, product_name: str, persona: str, use_cache: bool = False) -> Dict[str, Any]:
    """Generate call to action slide content."""
    cache_key = f"cta_{call_to_action}_{product_name}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical next steps like demos, trials, or technical discussions."
    elif persona == "Marketing":
        persona_specific = "Focus on engagement, lead generation, and customer journey."
    elif persona == "Executive":
        persona_specific = "Focus on high-level business discussions, partnerships, and strategic decisions."
    elif persona == "Investor":
        persona_specific = "Focus on investment opportunities, funding rounds, or partnership discussions."
    
    prompt = f"""
    Create a call-to-action slide based on this CTA:
    "{call_to_action}"
    
    Product name: {product_name}
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: An engaging slide title for the call to action
    - cta_text: A clear, compelling call-to-action statement
    - bullets: (optional) An array of 2-3 supporting points or next steps
    
    Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating CTA slide: {e}")
        # Provide fallback content
        content = {
            "title": "Get Started Today",
            "cta_text": call_to_action or f"Contact us to learn more about {product_name}",
            "bullets": ["Schedule a demo", "Start your free trial", "Speak with our experts"]
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Get Started Today"
    if not content.get("cta_text"):
        content["cta_text"] = call_to_action or f"Contact us to learn more about {product_name}"
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_market_slide_content(target_audience: str, persona: str, use_cache: bool = False) -> Dict[str, Any]:
    """Generate market opportunity slide content."""
    cache_key = f"market_{target_audience}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    persona_specific = ""
    if persona == "Investor" or persona == "Executive":
        persona_specific = "Include specific market size numbers, growth rates, and market opportunity details."
    
    prompt = f"""
    Create a market opportunity slide for a product targeting this audience:
    "{target_audience}"
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A compelling slide title about the market opportunity
    - market_size: A specific market size figure with dollar amount and timeframe (e.g. "$25B by 2025")
    - growth_rate: Annual growth rate with percentage (e.g. "15% CAGR")
    - description: A brief paragraph about the market opportunity and trends
    
    Be specific and data-driven with realistic but impressive figures. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating market slide: {e}")
        # Provide fallback content
        content = {
            "title": "Market Opportunity",
            "market_size": "$25B by 2025",
            "growth_rate": "15% CAGR",
            "description": "The market is experiencing significant growth as organizations increasingly recognize the need for innovative solutions in this space."
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = "Market Opportunity"
    if not content.get("market_size"):
        content["market_size"] = "$25B by 2025"
    if not content.get("growth_rate"):
        content["growth_rate"] = "15% CAGR"
    if not content.get("description"):
        content["description"] = "The market is experiencing significant growth as organizations increasingly recognize the need for innovative solutions in this space."
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_roadmap_slide_content(product_name: str, persona: str, use_cache: bool = False) -> Dict[str, Any]:
    """Generate product roadmap slide content."""
    cache_key = f"roadmap_{product_name}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    persona_specific = ""
    if persona == "Technical":
        persona_specific = "Focus on technical features, integrations, and architecture evolution."
    elif persona == "Marketing":
        persona_specific = "Focus on customer-facing features and market expansion."
    elif persona == "Executive":
        persona_specific = "Focus on strategic milestones, market expansion, and business objectives."
    elif persona == "Investor":
        persona_specific = "Focus on growth milestones, market expansion, and revenue phases."
    
    prompt = f"""
    Create a product roadmap slide for "{product_name}" with three clear phases.
    
    {persona_specific}
    
    Return a JSON object with these fields:
    - title: A slide title for the product roadmap (include product name)
    - phases: An array of 3 phase objects, each containing:
      - name: The phase name (e.g., "Phase 1: Foundation")
      - items: An array of 3-4 key deliverables or milestones for that phase
    
    Make the roadmap realistic and strategic. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating roadmap slide: {e}")
        # Provide fallback content
        content = {
            "title": f"{product_name}: Product Roadmap",
            "phases": [
                {
                    "name": "Phase 1: Foundation",
                    "items": ["Initial launch", "Core features", "First customers"]
                },
                {
                    "name": "Phase 2: Expansion",
                    "items": ["Advanced analytics", "Integration APIs", "Expanded support"]
                },
                {
                    "name": "Phase 3: Evolution",
                    "items": ["Enterprise features", "Global expansion", "Industry partnerships"]
                }
            ]
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{product_name}: Product Roadmap"
    
    # Ensure phases exist and have the right structure
    if not content.get("phases") or not isinstance(content.get("phases"), list):
        content["phases"] = [
            {
                "name": "Phase 1",
                "items": ["Initial launch", "Core features", "First customers"]
            },
            {
                "name": "Phase 2",
                "items": ["Advanced analytics", "Integration APIs", "Expanded support"]
            },
            {
                "name": "Phase 3",
                "items": ["Enterprise features", "Global expansion", "Industry partnerships"]
            }
        ]
    else:
        # Fix any phases that don't have the right structure
        for i, phase in enumerate(content["phases"]):
            if not isinstance(phase, dict):
                content["phases"][i] = {
                    "name": f"Phase {i+1}",
                    "items": ["Feature 1", "Feature 2", "Feature 3"]
                }
            elif "name" not in phase:
                phase["name"] = f"Phase {i+1}"
            elif "items" not in phase or not isinstance(phase["items"], list):
                phase["items"] = ["Feature 1", "Feature 2", "Feature 3"]
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def generate_team_slide_content(company_name: str, persona: str, use_cache: bool = True) -> Dict[str, Any]:
    """Generate team slide content."""
    cache_key = f"team_{company_name}_{persona}"
    
    # Return cached content if available and use_cache is True
    if use_cache and cache_key in _content_cache:
        return _content_cache[cache_key]
    
    model = genai.GenerativeModel(MODEL_NAME)
    
    prompt = f"""
    Create a team slide for "{company_name}" with fictional leadership team members.
    
    Return a JSON object with these fields:
    - title: A slide title for the team/leadership slide
    - team_members: An array of 4 team member objects, each containing:
      - name: The team member's name
      - role: Their role or title
    - tagline: A brief company tagline or mission statement
    
    Be professional and diverse. Format as valid JSON only.
    """
    
    # Make API call directly without rate limiting
    try:
        response = model.generate_content(prompt)
        content = extract_json_from_response(response.text)
    except Exception as e:
        logger.error(f"Error generating team slide: {e}")
        # Provide fallback content
        content = {
            "title": f"{company_name} Leadership Team",
            "team_members": [
                {"name": "Alex Johnson", "role": "Chief Executive Officer"},
                {"name": "Sam Washington", "role": "Chief Technology Officer"},
                {"name": "Jordan Smith", "role": "VP of Product"},
                {"name": "Taylor Rivera", "role": "VP of Customer Success"}
            ],
            "tagline": "Building innovative solutions for tomorrow's challenges"
        }
    
    # Ensure correct fields exist
    if not content.get("title"):
        content["title"] = f"{company_name} Leadership Team"
    
    # Ensure team_members exist and have the right structure
    if not content.get("team_members") or not isinstance(content.get("team_members"), list):
        content["team_members"] = [
            {"name": "Alex Johnson", "role": "Chief Executive Officer"},
            {"name": "Sam Washington", "role": "Chief Technology Officer"},
            {"name": "Jordan Smith", "role": "VP of Product"},
            {"name": "Taylor Rivera", "role": "VP of Customer Success"}
        ]
    
    # Ensure tagline exists
    if not content.get("tagline"):
        content["tagline"] = "Building innovative solutions for tomorrow's challenges"
    
    # Cache the result
    _content_cache[cache_key] = content
    
    return content

def extract_json_from_response(response_text: str) -> Dict[str, Any]:
    """Extract JSON from model response, handling various formats."""
    # If the response is already a dict, return it
    if isinstance(response_text, dict):
        return response_text
    
    # Handle string response
    try:
        import re
        import json
        
        if not isinstance(response_text, str):
            response_text = str(response_text)
        
        # Try to find JSON content within the string
        json_pattern = r'({[\s\S]*})'
        match = re.search(json_pattern, response_text)
        if match:
            json_str = match.group(1)
            return json.loads(json_str)
        else:
            # If no match, try to parse the whole string as JSON
            return json.loads(response_text)
    except Exception as e:
        logger.warning(f"Failed to extract JSON from response: {e}")
        # Return empty dict if parsing fails
        return {}

def clear_content_cache():
    """Clear the content cache."""
    global _content_cache
    _content_cache = {}

def get_cached_content():
    """Get a copy of the current content cache."""
    return _content_cache.copy()

def set_content_cache(cache_data):
    """Set the content cache with provided data."""
    global _content_cache
    _content_cache = cache_data
