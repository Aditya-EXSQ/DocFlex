import zipfile
import xmltodict
import json
import os
import shutil
from pathlib import Path

def docx_to_xml(docx_path, output_dir):
    """Extract DOCX as XML files with error handling"""
    # Clean output directory
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        docx_zip.extractall(output_dir)
    
    return output_dir

def xml_to_json(xml_dir, json_output_path):
    """Convert XML files to JSON with robust error handling"""
    docx_structure = {}
    
    for root, dirs, files in os.walk(xml_dir):
        for file in files:
            file_path = os.path.join(root, file)
            rel_path = os.path.relpath(file_path, xml_dir)
            
            # Skip non-XML files and binary content
            if not (file.endswith('.xml') or file.endswith('.rels')):
                # Store binary files as base64
                try:
                    with open(file_path, 'rb') as bin_file:
                        content = bin_file.read()
                    docx_structure[rel_path] = {
                        '_binary': True,
                        'content': content.hex()  # Store as hex for JSON compatibility
                    }
                except Exception as e:
                    print(f"Warning: Could not read binary file {rel_path}: {e}")
                continue
            
            # Handle XML files
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as xml_file:
                    xml_content = xml_file.read().strip()
                    
                if not xml_content:
                    docx_structure[rel_path] = {'_empty': True}
                    continue
                    
                # Try to parse as XML
                xml_dict = xmltodict.parse(xml_content)
                docx_structure[rel_path] = xml_dict
                
            except Exception as e:
                print(f"Warning: Could not parse XML file {rel_path}: {e}")
                # Store as raw text if XML parsing fails
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        raw_content = f.read()
                    docx_structure[rel_path] = {
                        '_raw': True,
                        'content': raw_content
                    }
                except Exception as read_error:
                    print(f"Error reading file {rel_path}: {read_error}")
                    docx_structure[rel_path] = {'_error': str(read_error)}
    
    # Save JSON
    with open(json_output_path, 'w', encoding='utf-8') as json_file:
        json.dump(docx_structure, json_file, indent=2, ensure_ascii=False)
    
    return json_output_path

def json_to_xml(json_path, xml_output_dir):
    """Convert JSON back to XML files"""
    if os.path.exists(xml_output_dir):
        shutil.rmtree(xml_output_dir)
    os.makedirs(xml_output_dir)
    
    with open(json_path, 'r', encoding='utf-8') as json_file:
        docx_structure = json.load(json_file)
    
    for rel_path, content_data in docx_structure.items():
        file_path = os.path.join(xml_output_dir, rel_path)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        if isinstance(content_data, dict):
            # Handle special cases
            if '_binary' in content_data:
                # Binary content
                hex_content = content_data['content']
                with open(file_path, 'wb') as bin_file:
                    bin_file.write(bytes.fromhex(hex_content))
            elif '_raw' in content_data:
                # Raw text content
                with open(file_path, 'w', encoding='utf-8') as text_file:
                    text_file.write(content_data['content'])
            elif '_empty' in content_data:
                # Empty file
                open(file_path, 'w').close()
            else:
                # Regular XML content
                try:
                    xml_content = xmltodict.unparse(content_data, pretty=True)
                    with open(file_path, 'w', encoding='utf-8') as xml_file:
                        xml_file.write(xml_content)
                except Exception as e:
                    print(f"Warning: Could not unparse XML for {rel_path}: {e}")
                    # Fallback: store as pretty-printed JSON
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(content_data, f, indent=2, ensure_ascii=False)
        else:
            # Simple content
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(str(content_data))
    
    return xml_output_dir

def xml_to_docx(xml_dir, docx_output_path):
    """Recreate DOCX from XML files"""
    if os.path.exists(docx_output_path):
        os.remove(docx_output_path)
    
    with zipfile.ZipFile(docx_output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for root, dirs, files in os.walk(xml_dir):
            for file in files:
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, xml_dir)
                docx_zip.write(file_path, rel_path)
    
    return docx_output_path

def convert_docx_roundtrip(docx_path, output_base="output"):
    """Complete roundtrip conversion"""
    # Create output directories
    xml_dir = os.path.join(output_base, "xml_extracted")
    json_path = os.path.join(output_base, "document.json")
    reconstructed_xml_dir = os.path.join(output_base, "xml_reconstructed")
    reconstructed_docx = os.path.join(output_base, "reconstructed.docx")
    
    print("Step 1: Extracting DOCX to XML...")
    docx_to_xml(docx_path, xml_dir)
    
    print("Step 2: Converting XML to JSON...")
    xml_to_json(xml_dir, json_path)
    
    print("Step 3: Converting JSON back to XML...")
    json_to_xml(json_path, reconstructed_xml_dir)
    
    print("Step 4: Recreating DOCX from XML...")
    xml_to_docx(reconstructed_xml_dir, reconstructed_docx)
    
    print(f"Roundtrip complete! Output: {reconstructed_docx}")
    return reconstructed_docx

# Usage
if __name__ == "__main__":
    convert_docx_roundtrip("Data\\DOCX Files\\Master Approval Letter.docx")

# json_to_xml("Output\\DOCX Files\\Master Approval Letter\\Document.json", "Data\\DOCX Files\\Master Approval Letter Reconstructed")
# xml_to_docx("Data\\DOCX Files\\Master Approval Letter Reconstructed", "Data\\DOCX Files\\Master Approval Letter Reconstructed.docx")