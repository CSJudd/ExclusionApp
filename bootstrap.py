from pathlib import Path

BASE_DIR = Path.cwd()

folders = [
    "engine",
    "clients",
    "assets"
]

files = {
    "engine/__init__.py": "",
    "engine/config_loader.py": "",
    "engine/normalizer.py": "",
    "engine/vendor_classifier.py": "",
    "engine/reference_cache.py": "",
    "engine/matcher_people.py": "",
    "engine/matcher_entity.py": "",
    "engine/runner.py": "",
    "engine/pdf_reports.py": "",
    "engine/audit_xlsx.py": "",
    "engine/history.py": "",
    "app_gui.py": "",
    "version.py": "",
    "clients/tri_area.yaml": ""
}

def create_structure():
    for folder in folders:
        path = BASE_DIR / folder
        path.mkdir(parents=True, exist_ok=True)

    for file_path, content in files.items():
        path = BASE_DIR / file_path
        path.parent.mkdir(parents=True, exist_ok=True)
        if not path.exists():
            path.write_text(content)

    print("ExclusionApp project structure created successfully.")

if __name__ == "__main__":
    create_structure()
