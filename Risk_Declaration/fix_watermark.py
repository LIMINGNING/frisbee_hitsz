import re

with open('Risk_Disclosure_Notice.tex', 'r', encoding='utf-8') as f:
    tex = f.read()

# Add tikz package
if r'\usepackage{tikz}' not in tex:
    tex = tex.replace(r'\usepackage{eso-pic}', r'\usepackage{eso-pic}' + '\n' + r'\usepackage{tikz}')

# Replace the AddToShipoutPictureBG block
old_watermark = r"""\AddToShipoutPictureBG{%
  \AtPageCenter{%
    \makebox(0,0)[c]{%
      \textcolor{black!10}{\includegraphics[width=0.5\paperwidth]{协会徽章_淡.png}}%
    }%
  }%
}"""

new_watermark = r"""\AddToShipoutPictureBG{%
  \AtPageCenter{%
    \makebox(0,0)[c]{%
      \begin{tikzpicture}[remember picture, overlay]
        \node[opacity=0.15] at (current page.center) {\includegraphics[width=0.5\paperwidth]{协会徽章_淡.png}};
      \end{tikzpicture}%
    }%
  }%
}"""

tex = tex.replace(old_watermark, new_watermark)

with open('Risk_Disclosure_Notice.tex', 'w', encoding='utf-8') as f:
    f.write(tex)

print("Watermark fixed.")
