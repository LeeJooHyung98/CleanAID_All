
import sys
import codecs

def convert_encoding(filename, source_enc='cp949', target_enc='utf-8'):
    try:
        with open(filename, 'rb') as f:
            content = f.read()
        
        # Decode and re-encode
        unicode_content = content.decode(source_enc)
        utf8_content = unicode_content.encode(target_enc)
        
        with open(filename, 'wb') as f:
            f.write(utf8_content)
            
        print(f"Successfully converted {filename} to {target_enc}")
    except Exception as e:
        print(f"Error converting {filename}: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        convert_encoding(sys.argv[1])
