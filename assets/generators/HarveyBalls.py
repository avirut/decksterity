#!/usr/bin/env python
# coding: utf-8

# Code sourced from https://github.com/emorisse/HarveyBalls, used for icon generation with modifications to remove border / whitespace around saved Harvey ball images completely

# In[39]:


import matplotlib.pyplot as plt
import matplotlib.path as mpath
import matplotlib.patches as mpatches
import re
import numpy as np


# In[40]:


color = "black"
colors = [ color, color ]

sets = [0, 0.25, 0.5, 0.75, 1, 0.33333333333]


# In[41]:


def cutwedge(wedge, r=0.95):
    path = wedge.get_path()
    verts = path.vertices[:-3]
    codes = path.codes[:-3]
    new_verts = np.vstack((verts , verts[::-1]*r, verts[0,:]))
    new_codes =  np.concatenate((codes , codes[::-1], np.array([79])) )
    new_codes[len(codes)] = 2
    new_path = mpath.Path(new_verts, new_codes)
    new_patch = mpatches.PathPatch(new_path)
    new_patch.update_from(wedge)
    wedge.set_visible(False)
    wedge.axes.add_patch(new_patch)
    return new_patch


# In[42]:


for n in sets:
    fig, ax = plt.subplots(figsize=(4, 4), dpi=100)  # Ensure it's square
    sizes = [n, 1 - n]
    perc = re.sub(r"\.", "", "%0.2f" % n) + "-" + color + ".png"
    wedges, text = ax.pie(
        sizes, 
        colors=colors, 
        startangle=90, 
        counterclock=False, 
        wedgeprops=dict(linewidth=0)  # Optional: remove wedge borders
    )

    cutwedge(wedges[1])

    plt.axis('equal')  # Keep the pie circular
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)  # Remove subplot margins
    plt.margins(0)  # No margins

    plt.savefig(perc, bbox_inches='tight', pad_inches=0, transparent=True, dpi=37.5)
    plt.close()


# In[ ]:




