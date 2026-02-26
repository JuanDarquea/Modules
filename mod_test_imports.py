# test_imports.py
import sys

# modules = input("Please write the module to import: ")
# modules_to_test = str(modules)
modules_to_test = ['docx', 'deepl', 'googletrans']

for module in modules_to_test:
    try:
        __import__(module)
        print(f"Module {module} imported successfully")
    except ImportError as e:
        print(f"Module {module} failed to import")
        print(f"Reason: {e}")