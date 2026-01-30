import os
import glob
from pptx import Presentation
import logging
from logger_config import setup_logging
from translator import translate_all
from file_utils import save_presentation

setup_logging()

def has_hlink(paragraph):
    """Проверяет наличие гиперссылок в параграфе с защитой от пустых rId."""
    for run in paragraph.runs:
        try:
            hlink = run.hyperlink
            if hlink is not None:
                if hasattr(hlink, 'rId') and hlink.rId:
                    return True
                if hlink.address is not None:
                    return True
        except Exception:
            return True
    return False

def extract_table_texts(shape):
    """Extract texts from a table shape."""
    texts = []
    locations = []
    
    if not shape.has_table:
        return texts, locations
        
    for row_idx, row in enumerate(shape.table.rows):
        for cell_idx, cell in enumerate(row.cells):
            if cell.text.strip():
                for para_idx, para in enumerate(cell.text_frame.paragraphs):
                    if para.text.strip() and not has_hlink(para):
                        texts.append(para.text.strip())
                        locations.append((row_idx, cell_idx, para_idx))
                    elif has_hlink(para):
                        logging.info(f"Skipping table cell paragraph with hyperlink: {para.text[:30]}...")
    
    return texts, locations

def process_presentation(input_file):
    """Process a PowerPoint presentation, translating text from English to Simplified Chinese."""
    logging.info(f"Processing {input_file}")
    print(f"\n Начало обработки файла: {os.path.basename(input_file)}")
    
    try:
        prs = Presentation(input_file)
        all_texts = []
        text_locations = []
        
        # Extract all texts that need translation
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                # Handle regular text shapes
                if hasattr(shape, "text_frame") and shape.text.strip():
                    for para_idx, para in enumerate(shape.text_frame.paragraphs):
                        if para.text.strip():
                            if has_hlink(para):
                                logging.info(f"Skipping paragraph with hyperlink: {para.text[:30]}...")
                                continue
                                
                            all_texts.append(para.text.strip())
                            text_locations.append(("paragraph", slide_idx, shape_idx, para_idx))
                
                # Handle tables
                if shape.has_table:
                    table_texts, table_locations = extract_table_texts(shape)
                    all_texts.extend(table_texts)
                    for (row_idx, cell_idx, para_idx) in table_locations:
                        text_locations.append(("table", slide_idx, shape_idx, row_idx, cell_idx, para_idx))
        
        if not all_texts:
            logging.info(f"No text found in {input_file}")
            return
        
        translated_texts = translate_all(all_texts)
        
        # Update presentation with translations
        for location, translated_text in zip(text_locations, translated_texts):
            if location[0] == "paragraph":
                _, slide_idx, shape_idx, para_idx = location
                shape = prs.slides[slide_idx].shapes[shape_idx]
                if hasattr(shape, "text_frame") and para_idx < len(shape.text_frame.paragraphs):
                    paragraph = shape.text_frame.paragraphs[para_idx]
                    
                    # Store original formatting
                    original_alignment = paragraph.alignment
                    original_level = paragraph.level
                    has_bullet = False
                    if hasattr(paragraph, "format") and hasattr(paragraph.format, "bullet"):
                        has_bullet = True

                    # Extract original color before clearing text
                    orig_color = None
                    if paragraph.runs:
                        try:
                            if hasattr(paragraph.runs[0].font.color, 'rgb'):
                                orig_color = paragraph.runs[0].font.color.rgb
                        except:
                            pass

                    # Store original font sizes before updating text
                    original_font_sizes = []
                    for run in paragraph.runs:
                        if hasattr(run, "font") and hasattr(run.font, "size"):
                            original_font_sizes.append(run.font.size)
                        else:
                            original_font_sizes.append(None)  # None means use default
                    
                    paragraph.text = translated_text
                    
                    # Set font to Microsoft YaHei for all runs in the paragraph while keeping original size
                    for idx, run in enumerate(paragraph.runs):
                        run.font.name = "Microsoft YaHei"
                        # If we have stored a font size and have enough runs, use the original
                        if idx < len(original_font_sizes) and original_font_sizes[idx] is not None:
                            run.font.size = original_font_sizes[idx]
                        
                        # Apply original color if found
                        if orig_color:
                            run.font.color.rgb = orig_color
                    
                    # Restore original formatting
                    paragraph.alignment = original_alignment
                    paragraph.level = original_level
                    if has_bullet and hasattr(paragraph, "format"):
                        try:
                            paragraph.format.bullet.enable = True
                        except:
                            # If bullet restoration fails, log but continue
                            logging.warning("Failed to restore bullet formatting")
            
            elif location[0] == "text":
                _, slide_idx, shape_idx, _ = location
                shape = prs.slides[slide_idx].shapes[shape_idx]
                shape.text = translated_text
                if hasattr(shape, "text_frame"):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = "Microsoft YaHei"
                
            elif location[0] == "table":
                _, slide_idx, shape_idx, row_idx, cell_idx, para_idx = location
                shape = prs.slides[slide_idx].shapes[shape_idx]
                if shape.has_table:
                    cell = shape.table.rows[row_idx].cells[cell_idx]
                    if hasattr(cell, "text_frame") and para_idx < len(cell.text_frame.paragraphs):
                        paragraph = cell.text_frame.paragraphs[para_idx]
                        
                        # Store original formatting
                        original_alignment = paragraph.alignment
                        original_level = paragraph.level
                        has_bullet = False
                        if hasattr(paragraph, "format") and hasattr(paragraph.format, "bullet"):
                            has_bullet = True
                        
                        # Extract original color for tables
                        orig_color = None
                        if paragraph.runs:
                            try:
                                if hasattr(paragraph.runs[0].font.color, 'rgb'):
                                    orig_color = paragraph.runs[0].font.color.rgb
                            except:
                                pass

                        # Store original font sizes before updating text
                        original_font_sizes = []
                        for run in paragraph.runs:
                            if hasattr(run, "font") and hasattr(run.font, "size"):
                                original_font_sizes.append(run.font.size)
                            else:
                                original_font_sizes.append(None)  # None means use default
                        
                        paragraph.text = translated_text
                        
                        # Set font to Microsoft YaHei for all runs in the paragraph while keeping original size
                        for idx, run in enumerate(paragraph.runs):
                            run.font.name = "Microsoft YaHei"
                            # If we have stored a font size and have enough runs, use the original
                            if idx < len(original_font_sizes) and original_font_sizes[idx] is not None:
                                run.font.size = original_font_sizes[idx]
                            
                            # Apply color to table text
                            if orig_color:
                                run.font.color.rgb = orig_color
                        
                        # Restore original formatting
                        paragraph.alignment = original_alignment
                        paragraph.level = original_level
                        if has_bullet and hasattr(paragraph, "format"):
                            try:
                                paragraph.format.bullet.enable = True
                            except:
                                # If bullet restoration fails, log but continue
                                logging.warning("Failed to restore bullet formatting")
        
        # Save translated presentation with error handling
        save_presentation(prs, input_file)
        
    except Exception as e:
        logging.error(f"Error processing presentation {input_file}: {str(e)}")
        raise

def main():
    # Find all PPTX files in the input directory
    input_files = glob.glob('input/*.pptx')
    
    if not input_files:
        logging.warning("No PowerPoint files found in the input directory")
        return
    
    for input_file in input_files:
        logging.info(f"\n=== Processing file: {input_file} ===")
        process_presentation(input_file)
        logging.info(f"Completed translation of {input_file}")

if __name__ == "__main__":
    main()