#!/usr/bin/env python
# coding: utf-8

# In[42]:


import matplotlib
import matplotlib.pyplot as plt
import matplotlib.path as mpath
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm

import re
import os
import numpy as np


# In[43]:


characters = ["✔","✘","?","✚","⚊","⋯"]
names = ["IconCheck","IconCross","IconQuestion","IconPlus","IconMinus","IconEllipsis"]
fonts = [None, None, r"D:\Seafile\Personal\projects\decksterity\assets\generators\ariblk.ttf", None, r"D:\Seafile\Personal\projects\decksterity\assets\generators\NotoSansSymbols2-Regular.ttf", None]


# In[44]:


font_path = r"D:\Seafile\Personal\projects\decksterity\assets\generators\arial.ttf"
font = fm.FontProperties(fname=font_path)


# In[45]:


for char, name, font_name in zip(characters, names, fonts):
    fig, ax = plt.subplots(figsize=(1, 1), dpi=300)

    font = fm.FontProperties(fname=font_name) if font_name else font_name
    
    ax.text(0.5, 0.5, char, fontsize=200, fontproperties=font, ha='center', va='center')
    # ax.text(0.5, 0.5, char, fontsize=200, ha='center', va='center')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

    filename = os.path.join("..", f"{name}.png")
    
    # plt.show()
    fig.savefig(filename, dpi=300, transparent=True, bbox_inches='tight', pad_inches=0)
    plt.close(fig)


# In[46]:


from PIL import Image
import os

def autocrop_image(path_in, path_out=None):
    with Image.open(path_in) as im:
        im = im.convert("RGBA")  # Ensure alpha channel
        bbox = im.getbbox()
        if bbox:
            cropped = im.crop(bbox)
            path_out = path_out or path_in
            cropped.save(path_out)

for name in names:
    autocrop_image(f"../{name}.png")

