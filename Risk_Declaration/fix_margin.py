import re

with open('Risk_Disclosure_Notice.tex', 'r', encoding='utf-8') as f:
    tex = f.read()

# Remove the negative top vertical space that's pulling the title up
tex = tex.replace(r'\vspace*{-1.5cm}', r'\vspace*{0cm}')

# Optionally adjust the geometry top margin slightly if needed, but the negative space is the main issue.
tex = tex.replace(r'top=2cm, bottom=1cm', r'top=2.5cm, bottom=2cm')

with open('Risk_Disclosure_Notice.tex', 'w', encoding='utf-8') as f:
    f.write(tex)

print("Margin fixed.")
