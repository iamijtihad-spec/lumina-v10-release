import re
from docx import Document
from docx.shared import RGBColor
import os
import json
import pdfplumber
import pypdfium2 as pdfium

def hex_to_rgb(hex_code):
    """Converts a hex color code (e.g., #FF0000) to a docx RGBColor object."""
    hex_code = hex_code.lstrip('#')
    return RGBColor(int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16))

def is_valid_hex(hex_code):
    """Checks if a string is a valid 6-character hex color code."""
    hex_code = hex_code.lstrip('#')
    return len(hex_code) == 6 and all(c in '0123456789ABCDEFabcdef' for c in hex_code)

def is_arabic(text):
    """Checks if the keyword contains Arabic characters using Unicode."""
    # Returns True if any character in the word falls inside the Arabic Unicode block
    return any('\u0600' <= c <= '\u06FF' for c in text)

def save_categories(categories_data, filename="categories.json"):
    """Saves categories data to a JSON file."""
    serializable_data = []
    for cat in categories_data:
        rgb = cat['color']
        hex_color = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        serializable_data.append({
            'name': cat['name'],
            'color': hex_color,
            'keywords': cat['keywords']
        })
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(serializable_data, f, indent=4, ensure_ascii=False)
    print(f"Categories saved to '{filename}'.")

def load_categories(filename="categories.json"):
    """Loads categories data from a JSON file."""
    if not os.path.exists(filename):
        return None
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        categories_data = []
        for item in data:
            categories_data.append({
                'name': item['name'],
                'color': hex_to_rgb(item['color']),
                'keywords': item['keywords']
            })
        return categories_data
    except Exception as e:
        print(f"Error loading '{filename}': {e}")
        return None

# --- Configuration ---
COLOR_PRESETS = {
    "1": ("Deep Blue", "#1A237E"),
    "2": ("Forest Green", "#1B5E20"),
    "3": ("Goldenrod", "#B8860B"),
    "4": ("Royal Purple", "#4A148C"),
    "5": ("Dark Orange", "#E65100"),
    "6": ("Saddle Brown", "#5D4037"),
    "7": ("Crimson Red", "#B71C1C")
}

def gather_user_inputs():
    """Prompts the user for categories, colors, and keywords."""
    categories_data = load_categories()
    if categories_data:
        use_existing = input(f"Found existing rules in 'categories.json'. Use them? (y/n): ").strip().lower()
        if use_existing == 'y':
            return categories_data
            
    categories_data = []
    print("\n--- Quranic Study Guide Color Coder ---")
    print("(Type 'done' at any prompt to finish and process the document)")
    
    # Create a reverse mapping for name-based selection
    name_to_hex = {name.lower(): hex_val for _, (name, hex_val) in COLOR_PRESETS.items()}
    
    while True:
        try:
            category_name = input("\nEnter category name (e.g., Attributes of Allah): ").strip()
            if category_name.lower() == 'done':
                break
                
            print(f"\nSelect a color for '{category_name}':")
            for key, (name, hex_val) in COLOR_PRESETS.items():
                print(f"  {key}: {name}")
            print("  C: Custom Hex Code")
            
            color_hex = ""
            while True:
                choice = input("Choice (e.g., 1, Blue, or C): ").strip().lower()
                if choice == 'done':
                    if categories_data:
                        save_choice = input("Save these rules to 'categories.json' for next time? (y/n): ").strip().lower()
                        if save_choice == 'y':
                            save_categories(categories_data)
                    return categories_data
                
                # Check for Number, Name, or Custom
                if choice in COLOR_PRESETS:
                    color_hex = COLOR_PRESETS[choice][1]
                    break
                elif choice in name_to_hex:
                    color_hex = name_to_hex[choice]
                    break
                elif choice.upper() == 'C':
                    while True:
                        color_hex = input("Enter custom hex code (e.g., #808080): ").strip()
                        if is_valid_hex(color_hex):
                            break
                        print("Invalid hex code. Please use format #RRGGBB.")
                    break
                else:
                    print(f"Invalid choice '{choice}'. Please pick a number, name (Blue, Green, etc.), or 'C'.")
            
            print(f"Color set to: {color_hex}")
            print(f"Enter the exact words/phrases for '{category_name}', separated by commas:")
            keywords_input = input().strip()
            if keywords_input.lower() == 'done':
                break
                
            keywords = [kw.strip() for kw in keywords_input.split(',') if kw.strip()]
            if not keywords:
                print("No keywords entered. Skipping category.")
                continue
                
            keywords.sort(key=len, reverse=True)
            categories_data.append({
                'name': category_name,
                'color': hex_to_rgb(color_hex),
                'keywords': keywords
            })
            
        except EOFError:
            break
    
    if categories_data:
        save_choice = input("Save these rules to 'categories.json' for next time? (y/n): ").strip().lower()
        if save_choice == 'y':
            save_categories(categories_data)
        
    return categories_data

def extract_text_pypdfium(input_path):
    """Fallback extraction using pypdfium2 (often more robust for Arabic)."""
    paragraphs = []
    try:
        pdf = pdfium.PdfDocument(input_path)
        for page in pdf:
            text_page = page.get_textpage()
            text = text_page.get_text_range()
            if text:
                paragraphs.extend(text.split('\n'))
    except Exception as e:
        print(f"pypdfium2 error: {e}")
    return paragraphs

def process_text_to_docx(input_path, output_docx_path, categories_data):
    """Reads a text, Word, or PDF file, applies color coding, and saves a .docx"""
    doc_out = Document()
    paragraphs = []
    ext = os.path.splitext(input_path)[1].lower()
    
    try:
        if ext == '.docx':
            doc_in = Document(input_path)
            paragraphs = [p.text for p in doc_in.paragraphs]
        elif ext == '.pdf':
            print(f"Extracting text from PDF: {input_path}...")
            # Try pdfplumber first
            with pdfplumber.open(input_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        paragraphs.extend(text.split('\n'))
            
            # If pdfplumber failed (e.g. scanned PDF), try pypdfium2 as fallback
            if not paragraphs:
                print("pdfplumber yielded no text. Trying pypdfium2...")
                paragraphs = extract_text_pypdfium(input_path)
            
            # If still no text, it's likely a scanned PDF
            if not paragraphs:
                print(f"WARNING: No text could be extracted from '{input_path}'. It may be a scanned image.")
                print("Looking for a matching text file fallback...")
                txt_fallback = os.path.splitext(input_path)[0] + ".txt"
                if os.path.exists(txt_fallback):
                    print(f"Found fallback: {txt_fallback}")
                    return process_text_to_docx(txt_fallback, output_docx_path, categories_data)
                elif os.path.exists("manuscript_draft.txt"):
                    print("Found fallback: manuscript_draft.txt")
                    return process_text_to_docx("manuscript_draft.txt", output_docx_path, categories_data)
                else:
                    print("No text fallback found. Aborting PDF processing.")
                    return
        else:
            with open(input_path, 'r', encoding='utf-8') as file:
                paragraphs = file.readlines()
    except Exception as e:
        print(f"Error reading file '{input_path}': {e}")
        return

    # Build a gated master regex pattern
    pattern_parts = []
    keyword_to_color = {}
    
    for cat in categories_data:
        for kw in cat['keywords']:
            keyword_to_color[kw.lower()] = cat['color']
            escaped_kw = re.escape(kw)
            
            # THE GATE: Decide how to parse based on language
            if is_arabic(kw):
                # Arabic: No boundaries. Allows finding root words attached to prefixes/suffixes.
                pattern_parts.append(escaped_kw)
            else:
                # English: Add boundaries (\b). Prevents "Ali" from matching inside "reality".
                pattern_parts.append(r'\b' + escaped_kw + r'\b')
            
    if not pattern_parts:
        print("No keywords to process.")
        for text in paragraphs:
            doc_out.add_paragraph(text.strip())
        doc_out.save(output_docx_path)
        return

    # Join all the mixed patterns together
    pattern_str = r'(' + '|'.join(pattern_parts) + r')'
    master_pattern = re.compile(pattern_str, re.IGNORECASE)

    print(f"\nProcessing manuscript from {input_path}...")
    
    for text in paragraphs:
        text = text.strip()
        if not text:
            doc_out.add_paragraph()
            continue
            
        p = doc_out.add_paragraph()
        last_index = 0
        
        for match in master_pattern.finditer(text):
            start, end = match.span()
            matched_word = match.group()
            
            if start > last_index:
                p.add_run(text[last_index:start])
                
            run = p.add_run(matched_word)
            run.font.color.rgb = keyword_to_color[matched_word.lower()]
            run.font.bold = True
            last_index = end
            
        if last_index < len(text):
            p.add_run(text[last_index:])

    doc_out.save(output_docx_path)
    print(f"\nSuccess! Color-coded document saved as: {output_docx_path}")

if __name__ == "__main__":
    yusuf_ali_pdf = "695084765-Holy-Qur-an-Yusuf-Ali-Translation-1946-Edition.pdf"
    
    if os.path.exists(yusuf_ali_pdf):
        input_file = yusuf_ali_pdf
    elif os.path.exists("manuscript_draft.docx"):
        input_file = "manuscript_draft.docx"
    elif os.path.exists("manuscript_draft.pdf"):
        input_file = "manuscript_draft.pdf"
    else:
        input_file = "manuscript_draft.txt"
        
    output_file = "Quran_StudyGuide_ColorCoded.docx"
    rules = gather_user_inputs()
    
    if rules:
        if not os.path.exists(input_file):
            input_file = "manuscript_draft.txt"
            if not os.path.exists(input_file):
                with open(input_file, 'w', encoding='utf-8') as f:
                    f.write("In the name of Allah, the Most Gracious, the Most Merciful.\n\n")
                    f.write("Establish regular prayer and give Zakat to the poor.\n\n")
                    f.write("And remember the story of Musa when he spoke to his people.\n")
                print(f"Created a sample '{input_file}' for testing.")

        process_text_to_docx(input_file, output_file, rules)
    else:
        print("No categories entered. Exiting.")
