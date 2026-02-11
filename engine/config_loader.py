# engine/config_loader.py

import yaml
from pathlib import Path

class ClientConfig:
    def __init__(self, config_path):
        self.path = Path(config_path)
        self.data = self._load_yaml()

    def _load_yaml(self):
        if not self.path.exists():
            raise FileNotFoundError(f"Client config not found: {self.path}")
        with open(self.path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)

    @property
    def client_name(self):
        return self.data.get("client_name")

    def section(self, section_name):
        return self.data.get(section_name, {})
