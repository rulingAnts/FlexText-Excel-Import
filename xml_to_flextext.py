import argparse
import os
import sys
import traceback
from xml.etree.ElementTree import Element, SubElement, tostring, parse
from xml.dom import minidom 

def transform_to_flextext_dom(xml_root_in, ws_vernacular, ws_gloss, ws_freetrans):
    """
    [MAIN CONVERSION FUNCTION]
    Transforms the custom interlinear XML DOM object into the FLExText XML format, 
    using the user-provided writing system codes for the 'lang' attributes.

    Args:
        xml_root_in (xml.etree.ElementTree.Element): The root XML element (<text> object).
        ws_vernacular (str): Writing system code for the vernacular text (e.g., 'xku').
        ws_gloss (str): Writing system code for the word glosses (e.g., 'gls').
        ws_freetrans (str): Writing system code for the free translation (e.g., 'en').

    Returns:
        xml.etree.ElementTree.Element: The root <document> element (FLExText object).
    """
    
    # Nested Helper Function (Now scoped locally)
    def create_languages_block():
        """
        Creates the essential <languages> element for the FLExText file, 
        using the writing system codes from the parent function's scope.
        """
        languages = Element('languages')
        
        # 1. Vernacular Language (The language of the text, marked as vernacular)
        SubElement(languages, 'language', lang=ws_vernacular, font="Charis SIL", vernacular="true")
        
        # 2. Analysis/Gloss Language (Used for word glosses and often for the title)
        SubElement(languages, 'language', lang=ws_gloss, font="Times New Roman")
        
        # 3. Free Translation/Title Language (If different from the gloss language)
        if ws_freetrans != ws_gloss:
            SubElement(languages, 'language', lang=ws_freetrans, font="Times New Roman")
        
        return languages

    
    # 1. Create the outermost root: <document>
    document_root = Element('document')
    document_root.set('version', '2')
    
    # 2. Create the <interlinear-text> container
    flextext_root = SubElement(document_root, 'interlinear-text')
    
    # 3. Extract and create the Title
    title_element = xml_root_in.find('.//title')
    title_text = title_element.text if title_element is not None and title_element.text else "Untitled Text"
    
    title_item = SubElement(flextext_root, 'item')
    title_item.set('type', 'title')
    title_item.set('lang', ws_freetrans) 
    title_item.text = title_text
    
    # 4. Process the Body (Paragraphs and Phrases)
    paragraphs_container = SubElement(flextext_root, 'paragraphs')
    
    for paragraph_in in xml_root_in.findall('.//paragraph'):
        
        # Create a new <paragraph>
        paragraph_out = SubElement(paragraphs_container, 'paragraph')
        phrases_container = SubElement(paragraph_out, 'phrases')
        
        # Iterate through lines (which become phrases)
        for line in paragraph_in.findall('./line'):
            
            phrase_element = SubElement(phrases_container, 'phrase')

            # 4.1. Add Free Translation (<item type="gls"> inside <phrase>)
            free_element = line.find('./free')
            free_translation_text = free_element.text if free_element is not None and free_element.text else ""
            
            if free_translation_text:
                free_item = SubElement(phrase_element, 'item')
                free_item.set('type', 'gls')
                free_item.set('lang', ws_freetrans) # Use Free Translation WS code
                free_item.text = free_translation_text

            # 4.2. Prepare the <words> container
            words_container = SubElement(phrase_element, 'words')
            
            # Find the interlinear lines within the source XML
            vern_line = line.find('./il-lines/vernacular-line')
            gloss_line = line.find('./il-lines/gloss-line')
            
            if vern_line is None or gloss_line is None:
                continue 

            vern_words = vern_line.findall('./wrd')
            glosses = gloss_line.findall('./gls')
            
            # Process words and glosses 1:1
            for i in range(min(len(vern_words), len(glosses))):
                vern_word = vern_words[i]
                word_gloss = glosses[i]
                
                word_element = SubElement(words_container, 'word')
                
                # Add Vernacular Word (<item type="txt">)
                txt_item = SubElement(word_element, 'item')
                txt_item.set('type', 'txt')
                txt_item.set('lang', ws_vernacular) # Use Vernacular WS code
                txt_item.text = vern_word.text if vern_word.text else ""

                # Add Word Gloss (<item type="gls">)
                gls_item = SubElement(word_element, 'item')
                gls_item.set('type', 'gls')
                gls_item.set('lang', ws_gloss) # Use Gloss WS code
                gls_item.text = word_gloss.text if word_gloss.text else ""
                
    # 5. Add the mandatory <languages> block
    languages_block = create_languages_block() # Called without arguments
    flextext_root.append(languages_block)
    
    return document_root 

# ======================================================================
# --- HELPER FUNCTIONS (Outside main conversion) ---
# ======================================================================

def prettify_xml(element):
    """Return a pretty-printed XML string for the given element."""
    rough_string = tostring(element, encoding='utf-8')
    reparsed = minidom.parseString(rough_string)
    # Ensure UTF-8 output
    return reparsed.toprettyxml(encoding='utf-8', indent="  ").decode('utf-8')

# ======================================================================
# --- CLI WRAPPER (Execution Block) ---
# ======================================================================

def cli_wrapper():
    """Handles command-line arguments, user input, I/O, and error logging."""
    parser = argparse.ArgumentParser(
        description="Transforms a custom XML format into a FLExText XML format."
    )
    parser.add_argument(
        "input_xml_file", 
        help="The path to the input XML document (e.g., output_text.xml)."
    )
    args = parser.parse_args()
    
    input_path = os.path.abspath(args.input_xml_file)
    base_name, _ = os.path.splitext(input_path)
    output_flextext_path = base_name + ".flextext"
    error_log_path = base_name + "_error.log"
    
    # 1. Input File Validation and Prompts
    if not os.path.exists(input_path):
        error_message = f"FATAL ERROR: Input XML file not found at path: {input_path}\n"
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write(error_message)
        print(f"FATAL ERROR: Input XML file not found. Details logged to {os.path.basename(error_log_path)}")
        sys.exit(1)

    print("\n--- FLEx Writing System Configuration ---")
    print("Please enter the exact writing system codes (WS Codes) used in your FLEx project.")
    print("This ensures the text imports correctly into the corresponding fields.")
    print("(You can find these under Tools -> Configure -> Writing Systems...)\n")
    
    ws_vernacular = input("1. Enter Vernacular (Baseline) WS Code (e.g., 'fau' or 'v'): ").strip()
    ws_gloss = input("2. Enter Word Gloss (Analysis) WS Code (e.g., 'en' or 'gls'): ").strip()
    ws_freetrans = input("3. Enter Free Translation WS Code (e.g., 'en' or 'ft'): ").strip()
    
    if not (ws_vernacular and ws_gloss and ws_freetrans):
        error_message = "\nFATAL ERROR: All three writing system codes must be provided."
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write(error_message + '\n')
        print(error_message)
        sys.exit(1)

    input_root = None # XML object placeholder

    # 2. Parse Input XML
    try:
        print("\n1. Parsing Input XML...")
        xml_tree = parse(input_path)
        input_root = xml_tree.getroot()
        if input_root.tag != 'text':
             raise ValueError(f"Root tag expected to be 'text', found '{input_root.tag}'")
        print("   - Input XML successfully parsed.")
    except Exception:
        error_message = f"\nFATAL ERROR during XML Parsing:\n{traceback.format_exc()}"
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write(error_message)
        print(f"\nFATAL ERROR: Could not parse XML file. Details logged to {os.path.basename(error_log_path)}")
        sys.exit(1)
        
    # 3. Perform Conversion (Calls the main modular function)
    try:
        print("2. Transforming XML to FLExText object...")
        document_root = transform_to_flextext_dom(input_root, ws_vernacular, ws_gloss, ws_freetrans)
        print("   - Transformation successful.")
    except Exception:
        error_message = f"\nFATAL ERROR during XML Transformation:\n{traceback.format_exc()}"
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write(error_message)
        print(f"\nFATAL ERROR: Conversion failed. Details logged to {os.path.basename(error_log_path)}")
        sys.exit(1)
        
    # 4. Write Output XML (FlexText)
    try:
        print("3. Writing output FlexText file...")
        pretty_xml = prettify_xml(document_root)
        
        with open(output_flextext_path, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
            
        print(f"\nCOMPLETED SUCCESSFULLY.")
        print(f"   - FlexText output saved to: '{os.path.basename(output_flextext_path)}'")
        
        # Clean up the error log if the entire process was successful
        if os.path.exists(error_log_path):
            os.remove(error_log_path) 
            
    except Exception:
        # This error handles failure during the final writing phase
        error_message = f"\nERROR during file writing. Output file may be incomplete.\n{traceback.format_exc()}"
        with open(error_log_path, 'a', encoding='utf-8') as f: 
            f.write(error_message)
        print(f"ERROR: Could not write FlexText file. Details logged to {os.path.basename(error_log_path)}")
        sys.exit(1)

if __name__ == "__main__":
    cli_wrapper()