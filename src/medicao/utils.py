import unicodedata
import os
import json


def normalize(s):
    if not s:
        return ''
    s = str(s).lower().strip()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s


def load_config():
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    with open(config_path, encoding='utf-8') as f:
        return json.load(f)


def find_folder(parent_path, patterns):
    if not os.path.isdir(parent_path):
        return None
    norm_patterns = [normalize(p) for p in patterns]
    for item in os.listdir(parent_path):
        item_path = os.path.join(parent_path, item)
        if os.path.isdir(item_path):
            item_norm = normalize(item)
            for pat in norm_patterns:
                if pat in item_norm or item_norm == pat:
                    return item_path
    return None


def find_files_by_keywords(folder_path, keywords, recurse=False):
    if not os.path.isdir(folder_path):
        return []
    norm_kws = [normalize(k) for k in keywords]
    matches = []
    items = os.listdir(folder_path)
    for item in items:
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            item_norm = normalize(item)
            if any(kw in item_norm for kw in norm_kws):
                matches.append(item_path)
        elif recurse and os.path.isdir(item_path):
            matches.extend(find_files_by_keywords(item_path, keywords, recurse=True))
    return matches


def list_files(folder_path, recurse=False):
    if not os.path.isdir(folder_path):
        return []
    result = []
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            result.append(item_path)
        elif recurse and os.path.isdir(item_path):
            result.extend(list_files(item_path, recurse=True))
    return result


def list_subfolders(folder_path):
    if not os.path.isdir(folder_path):
        return []
    return [
        os.path.join(folder_path, item)
        for item in os.listdir(folder_path)
        if os.path.isdir(os.path.join(folder_path, item))
    ]


def employee_matches_folder(emp_name, folder_name, loose=False):
    emp_norm = normalize(emp_name)
    folder_norm = normalize(folder_name)
    emp_words = [w for w in emp_norm.split() if len(w) > 2]
    if not emp_words:
        return False
    matches = sum(1 for w in emp_words if w in folder_norm)
    # Loose mode: accept if the first name (first word) alone is in the folder
    if loose and emp_words and emp_words[0] in folder_norm:
        return True
    threshold = 2 if len(emp_words) >= 3 else 1
    return matches >= threshold


def has_any_file(folder_path, extensions=None):
    if not os.path.isdir(folder_path):
        return False
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            if extensions is None:
                return True
            _, ext = os.path.splitext(item)
            if ext.lower() in extensions:
                return True
    return False
