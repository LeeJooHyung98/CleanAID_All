
import re

file_path = r'c:\WorkSpace\CleadAid\Source\ASIS_Sources\가맹점프로그램\Form\frm세탁물인도문자1.frm'

try:
    with open(file_path, 'r', encoding='cp949', errors='ignore') as f:
        content = f.read()

    # Find the button with the specific caption
    # VB6 form structure: 
    # Begin VB.CommandButton cmdSend 
    #    Caption = "문자..."
    # End
    
    # We will look for blocks
    lines = content.split('\n')
    current_object = None
    object_stack = []
    
    for i, line in enumerate(lines):
        line = line.strip()
        if line.startswith('Begin '):
            parts = line.split()
            if len(parts) >= 3:
                obj_type = parts[1]
                obj_name = parts[2]
                object_stack.append({'type': obj_type, 'name': obj_name, 'caption': ''})
        elif line.startswith('End'):
            if object_stack:
                obj = object_stack.pop()
                if '문자' in obj['caption']:
                    print(f"Found match: Type={obj['type']}, Name={obj['name']}, Caption={obj['caption']}")
        elif line.startswith('Caption'):
            if object_stack:
                # Caption = "..."
                if '"' in line:
                    caption = line.split('"')[1]
                    object_stack[-1]['caption'] = caption

    # Also search for Sub that might be relevant
    print("\n--- Subs related to Send/Click ---")
    for line in lines:
        if 'Sub ' in line and ('Click' in line or 'Send' in line):
            print(line.strip())

except Exception as e:
    print(f"Error: {e}")
