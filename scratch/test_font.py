import fitz

font_path = "C:/Windows/Fonts/cambriaz.ttf"
font = fitz.Font(fontfile=font_path)
width = font.text_length("TEST NAME", fontsize=14)

print(f"Font loaded: {font.name}")
print(f"Text width for 'TEST NAME' at 14pt: {width:.2f} points")
