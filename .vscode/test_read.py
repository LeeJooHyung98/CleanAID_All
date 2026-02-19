
try:
    with open(r'c:\WorkSpace\CleadAid\Source\ASIS_Sources\가맹점프로그램\Form\frm세탁물인도문자1.frm', 'rb') as f:
        data = f.read()
    print(f"Read {len(data)} bytes")
    text = data.decode('cp949', errors='ignore')
    print("Decoded successfully")
    for line in text.splitlines():
        if "문자" in line and "Caption" in line:
            print(line)
except Exception as e:
    print(f"Error: {e}")
