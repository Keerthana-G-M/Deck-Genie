from typing import Dict, Any, List
import re

def validate_and_fix_slide_content(slide_type: str, slide_content: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validate and fix slide content to ensure it has all required fields.
    This prevents errors during PowerPoint generation and ensures proper content sizing.
    """
    if not isinstance(slide_content, dict):
        slide_content = {}
    
    # Ensure title exists for all slides with proper length limits
    if 'title' not in slide_content or not slide_content['title']:
        default_titles = {
            'title_slide': 'Presentation Title',
            'problem_slide': 'The Problem',
            'solution_slide': 'Our Solution',
            'features_slide': 'Key Features',
            'advantage_slide': 'Competitive Advantage',
            'audience_slide': 'Target Audience',
            'market_slide': 'Market Opportunity',
            'roadmap_slide': 'Product Roadmap',
            'team_slide': 'Our Team',
            'cta_slide': 'Get Started'
        }
        slide_content['title'] = default_titles.get(slide_type, 'Slide Title')
    else:
        # Ensure title doesn't exceed PowerPoint display limits
        slide_content['title'] = truncate_smart(slide_content['title'], 60)
    
    # Fix specific slide types with enhanced content handling
    if slide_type == 'title_slide':
        if 'subtitle' not in slide_content or not slide_content['subtitle']:
            slide_content['subtitle'] = 'Professional Presentation'
        else:
            # Limit subtitle length to prevent overflow
            slide_content['subtitle'] = truncate_smart(slide_content['subtitle'], 80)
            
        if 'product_name' not in slide_content:
            slide_content['product_name'] = ''
    
    elif slide_type == 'problem_slide':
        # Enhanced problem slide content handling
        if 'pain_points' not in slide_content and 'bullets' not in slide_content:
            slide_content['pain_points'] = [
                'Current solutions are inefficient and costly',
                'Market lacks comprehensive options',
                'Existing tools don\'t meet user needs'
            ]
        else:
            # Process existing pain points to ensure proper formatting
            pain_points = slide_content.get('pain_points', slide_content.get('bullets', []))
            if pain_points:
                # Limit to 4 bullets max and ensure proper length
                processed_points = []
                for point in pain_points[:4]:
                    clean_point = clean_bullet_text(str(point))
                    if clean_point:
                        processed_points.append(truncate_smart(clean_point, 85))
                
                slide_content['pain_points'] = processed_points if processed_points else [
                    'Current market challenges need addressing'
                ]
    
    elif slide_type == 'solution_slide':
        # Enhanced solution content with length management
        if 'paragraph' not in slide_content and 'description' not in slide_content and 'value_proposition' not in slide_content:
            slide_content['paragraph'] = 'Our innovative solution addresses these challenges with a comprehensive approach that delivers measurable results and improved efficiency for our users.'
        else:
            # Use the best available content and ensure proper length
            content_text = (slide_content.get('paragraph') or 
                          slide_content.get('description') or 
                          slide_content.get('value_proposition') or '')
            
            if content_text:
                # Ensure content fits well in slide layout (400 chars max for good readability)
                slide_content['paragraph'] = truncate_smart(content_text, 400)
            else:
                slide_content['paragraph'] = 'Our solution provides significant value to users through innovative features and reliable performance.'
    
    elif slide_type == 'features_slide':
        # Enhanced features with better formatting
        if 'features' not in slide_content or not slide_content['features']:
            slide_content['features'] = [
                'Advanced Analytics Dashboard',
                'Real-time Data Processing',
                'Seamless Integration APIs',
                'Enterprise-grade Security',
                'Intuitive User Interface'
            ]
        else:
            # Process and limit features
            features = slide_content['features']
            processed_features = []
            
            for feature in features[:5]:  # Max 5 features for good layout
                if isinstance(feature, dict):
                    feature_text = feature.get('feature', feature.get('name', feature.get('title', str(feature))))
                else:
                    feature_text = str(feature)
                
                clean_feature = clean_bullet_text(feature_text)
                if clean_feature:
                    processed_features.append(truncate_smart(clean_feature, 100))
            
            slide_content['features'] = processed_features if processed_features else [
                'Core functionality', 'User-friendly design', 'Reliable performance'
            ]
    
    elif slide_type == 'advantage_slide':
        # Enhanced competitive advantages
        if 'differentiators' not in slide_content and 'bullets' not in slide_content:
            slide_content['differentiators'] = [
                'Superior performance compared to competitors',
                'Unique features not available elsewhere',
                'Cost-effective pricing model',
                'Exceptional customer support'
            ]
        else:
            # Process existing differentiators
            differentiators = slide_content.get('differentiators', slide_content.get('bullets', []))
            processed_diffs = []
            
            for diff in differentiators[:4]:  # Max 4 for good spacing
                if isinstance(diff, dict):
                    diff_text = diff.get('point', str(diff))
                else:
                    diff_text = str(diff)
                
                clean_diff = clean_bullet_text(diff_text)
                if clean_diff:
                    processed_diffs.append(truncate_smart(clean_diff, 85))
            
            slide_content['differentiators'] = processed_diffs if processed_diffs else [
                'Market-leading solution', 'Proven results', 'Superior user experience'
            ]
    
    elif slide_type == 'audience_slide':
        # Enhanced audience description
        if not any(key in slide_content for key in ['paragraph', 'description', 'content']):
            slide_content['paragraph'] = 'Our target audience includes decision-makers and professionals who value innovative solutions that improve efficiency and deliver measurable business results.'
        else:
            # Use best available content
            content_text = (slide_content.get('paragraph') or 
                          slide_content.get('description') or 
                          slide_content.get('content') or '')
            
            if content_text:
                slide_content['paragraph'] = truncate_smart(content_text, 400)
            else:
                slide_content['paragraph'] = 'Professionals seeking efficient, reliable solutions for their business challenges.'
    
    elif slide_type == 'market_slide':
        # Enhanced market data with realistic figures
        if 'market_size' not in slide_content or not slide_content['market_size']:
            slide_content['market_size'] = '$50B by 2025'
        else:
            slide_content['market_size'] = truncate_smart(slide_content['market_size'], 25)
            
        if 'growth_rate' not in slide_content or not slide_content['growth_rate']:
            slide_content['growth_rate'] = '15% CAGR'
        else:
            slide_content['growth_rate'] = truncate_smart(slide_content['growth_rate'], 20)
            
        if 'description' not in slide_content or not slide_content['description']:
            slide_content['description'] = 'The market shows strong growth potential with increasing demand for innovative solutions. Industry trends support sustained expansion and new opportunities for market leaders.'
        else:
            slide_content['description'] = truncate_smart(slide_content['description'], 500)
    
    elif slide_type == 'roadmap_slide':
        # Enhanced roadmap with better structure
        if 'phases' not in slide_content and 'milestones' not in slide_content:
            slide_content['phases'] = [
                {
                    'name': 'Phase 1: Foundation',
                    'items': ['Product launch', 'Core features deployment', 'Initial customer acquisition']
                },
                {
                    'name': 'Phase 2: Expansion', 
                    'items': ['Advanced analytics', 'API integrations', 'Enterprise features']
                },
                {
                    'name': 'Phase 3: Scale',
                    'items': ['Global expansion', 'AI capabilities', 'Strategic partnerships']
                }
            ]
        else:
            # Process existing phases
            phases = slide_content.get('phases', slide_content.get('milestones', []))
            processed_phases = []
            
            for i, phase in enumerate(phases[:3]):  # Max 3 phases for layout
                if isinstance(phase, dict):
                    phase_name = phase.get('name', f'Phase {i+1}')
                    phase_items = phase.get('items', [])
                else:
                    phase_name = f'Phase {i+1}'
                    phase_items = []
                
                # Limit phase name length
                phase_name = truncate_smart(phase_name, 25)
                
                # Process items
                processed_items = []
                for item in phase_items[:4]:  # Max 4 items per phase
                    clean_item = clean_bullet_text(str(item))
                    if clean_item:
                        processed_items.append(truncate_smart(clean_item, 40))
                
                if not processed_items:
                    processed_items = ['Key deliverables', 'Important milestones']
                
                processed_phases.append({
                    'name': phase_name,
                    'items': processed_items
                })
            
            slide_content['phases'] = processed_phases if processed_phases else [
                {'name': 'Phase 1', 'items': ['Launch', 'Initial features']},
                {'name': 'Phase 2', 'items': ['Growth', 'Advanced features']},
                {'name': 'Phase 3', 'items': ['Scale', 'Enterprise features']}
            ]
    
    elif slide_type == 'team_slide':
        # Enhanced team information
        if 'team_members' not in slide_content or not slide_content['team_members']:
            slide_content['team_members'] = [
                {'name': 'Alex Johnson', 'role': 'Chief Executive Officer'},
                {'name': 'Sam Chen', 'role': 'Chief Technology Officer'},
                {'name': 'Jordan Smith', 'role': 'VP of Product'},
                {'name': 'Taylor Rivera', 'role': 'VP of Sales & Marketing'}
            ]
        else:
            # Process existing team members
            team_members = slide_content['team_members']
            processed_members = []
            
            for i, member in enumerate(team_members[:4]):  # Max 4 members for layout
                if isinstance(member, dict):
                    name = member.get('name', f'Team Member {i+1}')
                    role = member.get('role', member.get('title', 'Team Member'))
                else:
                    name = str(member)
                    role = 'Team Member'
                
                # Limit lengths to prevent overflow
                name = truncate_smart(name, 25)
                role = truncate_smart(role, 50)
                
                processed_members.append({'name': name, 'role': role})
            
            slide_content['team_members'] = processed_members if processed_members else [
                {'name': 'Leadership Team', 'role': 'Experienced professionals driving innovation'}
            ]
            
        # Add tagline if not present
        if 'tagline' not in slide_content:
            slide_content['tagline'] = 'Building the future with innovative solutions'
        else:
            slide_content['tagline'] = truncate_smart(slide_content['tagline'], 80)
    
    elif slide_type == 'cta_slide':
        # Enhanced CTA with compelling messaging
        if 'cta_text' not in slide_content and 'call_to_action' not in slide_content:
            slide_content['cta_text'] = 'Ready to transform your business? Contact us today to get started!'
        else:
            cta_text = slide_content.get('cta_text', slide_content.get('call_to_action', ''))
            if cta_text:
                slide_content['cta_text'] = truncate_smart(cta_text, 120)
            else:
                slide_content['cta_text'] = 'Contact us today to learn more!'
        
        # Ensure bullets field exists
        if 'bullets' not in slide_content:
            slide_content['bullets'] = []
        elif slide_content['bullets']:
            # Process CTA bullets
            processed_bullets = []
            for bullet in slide_content['bullets'][:2]:  # Max 2 bullets for CTA
                clean_bullet = clean_bullet_text(str(bullet))
                if clean_bullet:
                    processed_bullets.append(truncate_smart(clean_bullet, 60))
            slide_content['bullets'] = processed_bullets
    
    return slide_content

def truncate_smart(text: str, max_length: int) -> str:
    """
    Smart truncation that tries to break at word boundaries and adds ellipsis if needed.
    """
    if not text or len(text) <= max_length:
        return text
    
    # Try to break at word boundary
    if max_length > 3:
        truncated = text[:max_length-3]
        last_space = truncated.rfind(' ')
        
        if last_space > max_length * 0.7:  # If we can break reasonably close to the limit
            return truncated[:last_space] + '...'
    
    # Otherwise, hard truncate
    return text[:max_length-3] + '...'

def clean_bullet_text(text: str) -> str:
    """
    Clean bullet point text by removing unwanted characters and formatting.
    """
    if not text:
        return ""
    
    # Remove bullet characters and extra whitespace
    cleaned = re.sub(r'^[\s•\-\*\+\>]+', '', str(text)).strip()
    
    # Remove multiple spaces
    cleaned = re.sub(r'\s+', ' ', cleaned)
    
    # Ensure it ends with proper punctuation for readability
    if cleaned and not cleaned.endswith(('.', '!', '?', ':')):
        cleaned += '.'
    
    return cleaned

def validate_presentation_content(content: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validate and fix entire presentation content structure.
    """
    validated_content = {}
    
    # Process each slide
    for slide_key, slide_data in content.items():
        if slide_key == 'metadata':
            validated_content[slide_key] = slide_data
        else:
            validated_content[slide_key] = validate_and_fix_slide_content(slide_key, slide_data)
    
    return validated_content

def enhance_content_with_context(content: Dict[str, Any], additional_info: Dict[str, str]) -> Dict[str, Any]:
    """
    Enhance existing content with additional user-provided information.
    This function allows the model to create richer content when more details are available.
    """
    enhanced_content = content.copy()
    
    # Extract useful context from additional info
    company_context = additional_info.get('company_background', '')
    product_details = additional_info.get('product_details', '')
    market_info = additional_info.get('market_context', '')
    team_info = additional_info.get('team_background', '')
    
    # Enhance specific slides based on additional context
    if company_context and 'title_slide' in enhanced_content:
        # Could enhance subtitle with company context
        pass
    
    if product_details and 'solution_slide' in enhanced_content:
        # Enhanced solution description using product details
        current_para = enhanced_content['solution_slide'].get('paragraph', '')
        if len(product_details) > 20:  # Only if substantial details provided
            enhanced_para = f"{current_para} {product_details}"
            enhanced_content['solution_slide']['paragraph'] = truncate_smart(enhanced_para, 400)
    
    if market_info and 'market_slide' in enhanced_content:
        # Enhanced market description
        current_desc = enhanced_content['market_slide'].get('description', '')
        if len(market_info) > 20:
            enhanced_desc = f"{current_desc} {market_info}"
            enhanced_content['market_slide']['description'] = truncate_smart(enhanced_desc, 500)
    
    return enhanced_content
