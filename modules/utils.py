def convertir_hex_a_argb(hex_color):
    if not hex_color:
        return ""
    hex_color = hex_color.strip().lstrip("#")
    return f"FF{hex_color.upper()}"
