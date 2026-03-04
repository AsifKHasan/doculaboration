import yaml

def transform_to_any_parent(data, mapping_schema):
    """
    mapping_schema: {
        (child_full_path): (parent_full_path, "new_key_name")
    }
    """
    moved_values = {} # Stores values to be injected: { parent_path: {new_key: value} }
    paths_to_delete = set(mapping_schema.keys())

    # --- PASS 1: Collection ---
    def collect_moves(obj, current_path):
        if not isinstance(obj, dict):
            return
        
        for key, value in obj.items():
            path = current_path + (key,)
            if path in mapping_schema:
                target_parent_path, new_key_name = mapping_schema[path]
                if target_parent_path not in moved_values:
                    moved_values[target_parent_path] = {}
                # Store the value (and recurse into it in case it's a dict)
                moved_values[target_parent_path][new_key_name] = value
            
            if isinstance(value, dict):
                collect_moves(value, path)

    collect_moves(data, ())

    # --- PASS 2: Reconstruction & Injection ---
    def rebuild(obj, current_path):
        if not isinstance(obj, dict):
            return obj

        new_obj = {}
        
        # 1. Add existing keys (unless they were marked for deletion/moving)
        for key, value in obj.items():
            path = current_path + (key,)
            if path in paths_to_delete:
                continue
            
            if isinstance(value, dict):
                processed_dict = rebuild(value, path)
                # Only keep the dictionary if it's not empty after moves
                if processed_dict:
                    new_obj[key] = processed_dict
            else:
                new_obj[key] = value

        # 2. Inject moved keys belonging to this parent
        if current_path in moved_values:
            for new_key, val in moved_values[current_path].items():
                # Recurse into the moved value in case it's a nested dict itself
                new_obj[new_key] = rebuild(val, current_path + (new_key,))

        return new_obj

    return rebuild(data, ())

# --- EXAMPLE CASE ---
style_mapping = {
    ("text-properties", "font", "family")           : (("text-properties",), "fontfamily"),
    ("text-properties", "font", "size")             : (("text-properties",), "fontsize"),
    ("text-properties", "font", "style")            : (("text-properties",), "fontstyle"),
    ("text-properties", "font", "weight")           : (("text-properties",), "fontweight"),

    ("paragraph-properties", "margin", "all")       : (("paragraph-properties",), "margin"),
    ("paragraph-properties", "margin", "bottom")    : (("paragraph-properties",), "marginbottom"),
    ("paragraph-properties", "margin", "left")      : (("paragraph-properties",), "marginleft"),
    ("paragraph-properties", "margin", "top")       : (("paragraph-properties",), "margintop"),
    ("paragraph-properties", "margin", "right")     : (("paragraph-properties",), "marginright"),

    ("paragraph-properties", "padding", "all")      : (("paragraph-properties",), "padding"),
    ("paragraph-properties", "padding", "bottom")   : (("paragraph-properties",), "paddingbottom"),
    ("paragraph-properties", "padding", "left")     : (("paragraph-properties",), "paddingleft"),
    ("paragraph-properties", "padding", "top")      : (("paragraph-properties",), "paddingtop"),
    ("paragraph-properties", "padding", "right")    : (("paragraph-properties",), "paddingright"),
}

original = {
            "separator-02": {
                "comment": "example style for a cell - acts like a page separator to be applied on a heading level 2",
                "text-properties": {
                    "color": "#0c1130",
                    "font": {
                        "fontfamily": "Droid Sans",
                        "fontsize": "14pt",
                        "fontstyle": "italic",
                        "fontweight": "bold"
                    }
                },
                "paragraph-properties": {
                    "textalign": "center",
                    "backgroundcolor": "#D8D8D8",
                    "verticalalign": "middle",
                    "border": {
                        "style": {
                            "left": "1pt double #222222 0.25in"
                        },
                        "linewidth": {
                            "left": "0.5pt 2.0pt 0.5pt"
                        }
                    },
                    "margin": {
                        "bottom": "1.00in",
                        "left": "1.00in",
                        "top": "1.00in",
                        "right": "1.00in"
                    },
                    "padding": {
                        "bottom": "0.25in",
                        "left": "0.25in",
                        "top": "0.25in",
                        "right": "0.25in"
                    }
                }
            }
}

transformed = transform_to_any_parent(original, style_mapping)
print(yaml.dump(transformed, sort_keys=False))