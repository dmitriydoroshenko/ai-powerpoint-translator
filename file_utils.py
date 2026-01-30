import os
import logging

def save_presentation(prs, original_filename):
    """Save presentation with error handling and unique filename."""
    # Create output directory if it doesn't exist
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate base output filename
    base_name = os.path.basename(original_filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    # Try to save with different names if file exists or is locked
    counter = 1
    while True:
        if counter == 1:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_cn.pptx")
        else:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_cn_{counter}.pptx")
        
        try:
            prs.save(output_filename)
            logging.info(f"Successfully saved presentation to {output_filename}")
            print(f"✅ Файл сохранен: {output_filename}")
            return output_filename
        except PermissionError:
            logging.warning(f"Permission denied when saving to {output_filename}. File might be open in PowerPoint.")
            logging.info("Please close the file in PowerPoint if it's open.")
            counter += 1
            if counter > 5:  # Limit number of attempts
                raise Exception(f"Failed to save presentation after {counter-1} attempts. Please ensure the file is not open in PowerPoint.")
        except Exception as e:
            logging.error(f"Error saving presentation: {str(e)}")
            raise