from odf.style import TextProperties
import odf.style

def is_attribute_allowed(class_name: str, attribute_name: str) -> bool:
    # Get the class dynamically from odf.style
    cls = getattr(odf.style, class_name, None)
    if cls is None:
        raise ValueError(f"No such class: {class_name}")

    # Create instance
    element = cls()

    # Allowed attributes are stored in element.allowed_attributes
    allowed = element.allowed_attributes()
    print(allowed)

    return any(local == attribute_name for ns, local in allowed)

# Example usage
print(is_attribute_allowed("TextProperties", "font-size"))