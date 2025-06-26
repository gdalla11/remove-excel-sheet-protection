import zipfile
import os
import shutil
import re
from tempfile import mkdtemp

def extract_archive(xlsm_path, extract_dir):
    """Extract Excel file to directory."""
    with zipfile.ZipFile(xlsm_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def rezip_folder(folder_path, output_file):
    """Rezip folder back to Excel format."""
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root_dir, dirs, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, folder_path)
                zipf.write(full_path, rel_path)

def remove_sheet_protection_string_method(sheet_path):
    """Remove sheet protection using safe string replacement."""
    try:
        # Read the file as text
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # Pattern to match sheetProtection elements
        # This matches both self-closing and regular tags
        protection_patterns = [
            r'<sheetProtection[^>]*/>',  # Self-closing tags
            r'<sheetProtection[^>]*>.*?</sheetProtection>',  # Regular tags with content
        ]
        
        removals_made = 0
        
        for pattern in protection_patterns:
            matches = re.findall(pattern, content, re.DOTALL | re.IGNORECASE)
            if matches:
                print(f"🔍 Found {len(matches)} protection element(s) in {os.path.basename(sheet_path)}")
                for match in matches:
                    print(f"   Removing: {match[:100]}{'...' if len(match) > 100 else ''}")
                
                content = re.sub(pattern, '', content, flags=re.DOTALL | re.IGNORECASE)
                removals_made += len(matches)
        
        if removals_made == 0:
            return False
        
        # Write back the modified content
        with open(sheet_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✅ Removed {removals_made} protection element(s) from {os.path.basename(sheet_path)}")
        return True
        
    except Exception as e:
        print(f"❌ Error processing {sheet_path}: {e}")
        # Try to restore original content
        try:
            with open(sheet_path, 'w', encoding='utf-8') as f:
                f.write(original_content)
        except:
            pass
        return False

def inspect_sheet_content(sheet_path):
    """Inspect sheet content to see what protection elements exist."""
    try:
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"\n🔍 Inspecting {os.path.basename(sheet_path)}:")
        
        # Look for any protection-related elements
        protection_patterns = [
            r'<sheetProtection[^>]*/>',
            r'<sheetProtection[^>]*>.*?</sheetProtection>',
            r'<protectedRanges[^>]*>.*?</protectedRanges>',
            r'<protectedRange[^>]*/>',
        ]
        
        found_any = False
        for i, pattern in enumerate(protection_patterns):
            matches = re.findall(pattern, content, re.DOTALL | re.IGNORECASE)
            if matches:
                found_any = True
                pattern_names = ['sheetProtection (self-closing)', 'sheetProtection (with content)', 
                               'protectedRanges', 'protectedRange']
                print(f"   📋 {pattern_names[i]}: {len(matches)} found")
                for match in matches[:3]:  # Show first 3 matches
                    print(f"      {match[:150]}{'...' if len(match) > 150 else ''}")
                if len(matches) > 3:
                    print(f"      ... and {len(matches) - 3} more")
        
        if not found_any:
            print(f"   ✅ No protection elements found")
            
        return found_any
        
    except Exception as e:
        print(f"   ❌ Error inspecting {sheet_path}: {e}")
        return False

def remove_protection_from_all_sheets(temp_dir, inspect_only=False):
    """Remove protection from all worksheet files."""
    sheets_path = os.path.join(temp_dir, 'xl', 'worksheets')
    
    if not os.path.exists(sheets_path):
        print("❌ Worksheets directory not found")
        return []
    
    modified_sheets = []
    sheet_files = [f for f in os.listdir(sheets_path) if f.endswith('.xml')]
    
    print(f"📊 Found {len(sheet_files)} worksheet files")
    
    for file_name in sheet_files:
        full_path = os.path.join(sheets_path, file_name)
        
        if inspect_only:
            inspect_sheet_content(full_path)
        else:
            if remove_sheet_protection_string_method(full_path):
                modified_sheets.append(file_name)
    
    return modified_sheets

def map_sheet_names(temp_dir):
    """Map worksheet file names to human-readable sheet names."""
    try:
        workbook_xml = os.path.join(temp_dir, 'xl', 'workbook.xml')
        if not os.path.exists(workbook_xml):
            return {}

        with open(workbook_xml, 'r', encoding='utf-8') as f:
            workbook_content = f.read()

        # Extract sheet information using regex
        sheet_pattern = r'<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"'
        sheets = re.findall(sheet_pattern, workbook_content)
        
        sheet_map = {rid: name for name, rid in sheets}

        # Map relationship IDs to file names
        rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        if not os.path.exists(rels_path):
            return {}
            
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_content = f.read()

        rel_pattern = r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="worksheets/([^"]+)"'
        relationships = re.findall(rel_pattern, rels_content)
        
        id_to_file = {rid: filename for rid, filename in relationships}

        # Create final mapping
        file_to_name = {}
        for rid, filename in id_to_file.items():
            if rid in sheet_map:
                file_to_name[filename] = sheet_map[rid]

        return file_to_name
    except Exception as e:
        print(f"❌ Error mapping sheet names: {e}")
        return {}

def main():
    print("🔓 Excel Sheet Protection Remover (Safe String Method)")
    print("=" * 55)
    
    original_file = input("📂 Enter full path to the .xlsm/.xlsx file: ").strip('"')
    
    if not os.path.exists(original_file):
        print("❌ File not found.")
        return
    
    if not original_file.lower().endswith(('.xlsx', '.xlsm')):
        print("❌ File must be .xlsx or .xlsm format.")
        return

    # Ask if user wants to inspect first
    inspect_first = input("🔍 Do you want to inspect the file first to see what protection exists? (y/n): ").lower().startswith('y')

    # Create backup
    backup_file = original_file + '.backup'
    try:
        shutil.copy2(original_file, backup_file)
        print(f"💾 Backup created: {backup_file}")
    except Exception as e:
        print(f"❌ Could not create backup: {e}")
        return

    base_name = os.path.splitext(os.path.basename(original_file))[0]
    temp_dir = mkdtemp(prefix=f"{base_name}_temp_")
    extension = '.xlsm' if original_file.lower().endswith('.xlsm') else '.xlsx'
    new_file = os.path.join(os.path.dirname(original_file), f"{base_name}_unprotected{extension}")

    try:
        print("📦 Extracting Excel archive...")
        extract_archive(original_file, temp_dir)

        print("🧭 Mapping sheet names...")
        sheet_name_map = map_sheet_names(temp_dir)

        if inspect_first:
            print("\n🔍 INSPECTION MODE - Analyzing protection elements:")
            remove_protection_from_all_sheets(temp_dir, inspect_only=True)
            
            proceed = input("\n🤔 Do you want to proceed with removing the protection? (y/n): ").lower().startswith('y')
            if not proceed:
                print("❌ Operation cancelled by user.")
                return

        print("\n🔓 Removing protection from sheets...")
        modified = remove_protection_from_all_sheets(temp_dir, inspect_only=False)

        if not modified:
            print("⚠️  No sheet protection found - file may already be unprotected.")
        else:
            print(f"\n📊 Successfully unprotected {len(modified)} sheet(s):")
            for fname in modified:
                readable_name = sheet_name_map.get(fname, fname.replace('.xml', ''))
                print(f"   • {readable_name} ({fname})")

        print(f"\n💾 Creating unprotected file...")
        rezip_folder(temp_dir, new_file)
        print(f"✅ Done! New file created: {new_file}")
        
        # Test the new file by trying to open it as a zip
        try:
            with zipfile.ZipFile(new_file, 'r') as test_zip:
                test_zip.testzip()
            print("✅ File integrity verified")
        except Exception as e:
            print(f"❌ File integrity check failed: {e}")
            print("⚠️  The unprotected file may be corrupted")

    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Clean up temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()