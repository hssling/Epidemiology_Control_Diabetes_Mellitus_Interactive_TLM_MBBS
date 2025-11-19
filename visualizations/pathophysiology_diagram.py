import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch, ConnectionPatch

# Create figure and axis
fig, ax = plt.subplots(figsize=(16, 10))
ax.set_xlim(0, 16)
ax.set_ylim(0, 10)
ax.axis('off')

# Define colors
bg_colors = ['#E8F5E8', '#FFF3E0', '#FCE4EC', '#F3E5F5', '#E0F2F1', '#FFF8E1']
text_colors = ['#2E7D32', '#E65100', '#C2185B', '#7B1FA2', '#00695C', '#F57F17']

# Type 1 DM Box
type1_box = FancyBboxPatch((0.5, 7), 4, 2, boxstyle="round,pad=0.1",
                          facecolor=bg_colors[0], edgecolor=text_colors[0], linewidth=2)
ax.add_patch(type1_box)
ax.text(2.5, 8.5, 'TYPE 1 DIABETES', ha='center', va='center', fontsize=14, fontweight='bold', color=text_colors[0])
ax.text(2.5, 7.7, 'Autoimmune destruction\nof β-cells', ha='center', va='center', fontsize=10, color=text_colors[0])
ax.text(2.5, 7.3, 'Absolute insulin deficiency', ha='center', va='center', fontsize=10, color=text_colors[0])

# Type 2 DM Box
type2_box = FancyBboxPatch((5.8, 7), 4, 2, boxstyle="round,pad=0.1",
                          facecolor=bg_colors[1], edgecolor=text_colors[1], linewidth=2)
ax.add_patch(type2_box)
ax.text(7.8, 8.5, 'TYPE 2 DIABETES', ha='center', va='center', fontsize=14, fontweight='bold', color=text_colors[1])
ax.text(7.8, 7.7, 'Insulin resistance +\nβ-cell dysfunction', ha='center', va='center', fontsize=10, color=text_colors[1])
ax.text(7.8, 7.3, 'Relative insulin deficiency', ha='center', va='center', fontsize=10, color=text_colors[1])

# Other Types Box
other_box = FancyBboxPatch((11.1, 7.5), 4, 2.5, boxstyle="round,pad=0.1",
                          facecolor=bg_colors[2], edgecolor=text_colors[2], linewidth=2)
ax.add_patch(other_box)
ax.text(13.1, 8.8, 'OTHER TYPES', ha='center', va='center', fontsize=14, fontweight='bold', color=text_colors[2])
ax.text(13.1, 8.4, '• MODY', ha='center', va='center', fontsize=9, color=text_colors[2])
ax.text(13.1, 8.0, '• Mitochondrial', ha='center', va='center', fontsize=9, color=text_colors[2])
ax.text(13.1, 7.6, '• Pancreatic', ha='center', va='center', fontsize=9, color=text_colors[2])
ax.text(13.1, 7.2, '• Drug-induced', ha='center', va='center', fontsize=9, color=text_colors[2])

# Insulin Resistance Box
resistance_box = FancyBboxPatch((1, 4), 3.5, 1.5, boxstyle="round,pad=0.05",
                               facecolor=bg_colors[3], edgecolor=text_colors[3], linewidth=2)
ax.add_patch(resistance_box)
ax.text(2.75, 4.75, 'INSULIN RESISTANCE', ha='center', va='center', fontsize=12, fontweight='bold', color=text_colors[3])
ax.text(2.75, 4.25, '• Obesity', ha='center', va='center', fontsize=9, color=text_colors[3])
ax.text(2.75, 4.05, '• Physical inactivity', ha='center', va='center', fontsize=9, color=text_colors[3])
ax.text(2.75, 3.85, '• Genetic factors', ha='center', va='center', fontsize=9, color=text_colors[3])

# Beta Cell Dysfunction Box
betacell_box = FancyBboxPatch((6.3, 4), 3.5, 1.5, boxstyle="round,pad=0.05",
                             facecolor=bg_colors[4], edgecolor=text_colors[4], linewidth=2)
ax.add_patch(betacell_box)
ax.text(8.05, 4.75, 'β-CELL DYSFUNCTION', ha='center', va='center', fontsize=12, fontweight='bold', color=text_colors[4])
ax.text(8.05, 4.25, '• Reduced secretion', ha='center', va='center', fontsize=9, color=text_colors[4])
ax.text(8.05, 4.05, '• Cellular stress', ha='center', va='center', fontsize=9, color=text_colors[4])
ax.text(8.05, 3.85, '• Apoptosis', ha='center', va='center', fontsize=9, color=text_colors[4])

# Pathogenic Factors Box
pathogenic_box = FancyBboxPatch((11.6, 4), 3.5, 1.5, boxstyle="round,pad=0.05",
                               facecolor=bg_colors[5], edgecolor=text_colors[5], linewidth=2)
ax.add_patch(pathogenic_box)
ax.text(13.35, 4.75, 'PATHOGENIC FACTORS', ha='center', va='center', fontsize=12, fontweight='bold', color=text_colors[5])
ax.text(13.35, 4.25, '• Inflammation', ha='center', va='center', fontsize=9, color=text_colors[5])
ax.text(13.35, 4.05, '• Oxidative stress', ha='center', va='center', fontsize=9, color=text_colors[5])
ax.text(13.35, 3.85, '• Lipotoxicity', ha='center', va='center', fontsize=9, color=text_colors[5])

# Final Outcome Box
outcome_box = FancyBboxPatch((6, 1), 4, 1.5, boxstyle="round,pad=0.1",
                            facecolor='#FFEBEE', edgecolor='#C62828', linewidth=3)
ax.add_patch(outcome_box)
ax.text(8, 2.1, 'HYPERGLYCEMIA', ha='center', va='center', fontsize=16, fontweight='bold', color='#C62828')
ax.text(8, 1.5, 'Carbohydrate, fat, protein metabolism disturbances', ha='center', va='center', fontsize=10, color='#C62828')

# Connecting arrows
# From insulin resistance to type 2 DM
arrow1 = ConnectionPatch((4.85, 4.75), (5.8, 7.5), "data", "data",
                        arrowstyle="->", shrinkA=5, shrinkB=5,
                        mutation_scale=20, fc=text_colors[1], color=text_colors[1])
ax.add_artist(arrow1)

# From beta cell dysfunction to type 2 DM
arrow2 = ConnectionPatch((9.8, 4.75), (9.8, 7.5), "data", "data",
                        arrowstyle="->", shrinkA=5, shrinkB=5,
                        mutation_scale=20, fc=text_colors[1], color=text_colors[1])
ax.add_artist(arrow2)

# From pathogenic factors to type 1 and 2 DM
arrow3 = ConnectionPatch((13.35, 5.5), (12.1, 7), "data", "data",
                        arrowstyle="->", shrinkA=5, shrinkB=5,
                        mutation_scale=20, fc=text_colors[2], color=text_colors[2])
ax.add_artist(arrow3)

# From type 1 DM to hyperglycemia
arrow4 = ConnectionPatch((4.5, 7.5), (8, 2.6), "data", "data",
                        arrowstyle="->", shrinkA=5, shrinkB=5,
                        mutation_scale=25, fc=text_colors[0], color=text_colors[0])
ax.add_artist(arrow4)

# From type 2 DM to hyperglycemia
arrow5 = ConnectionPatch((9.8, 7.5), (8, 2.6), "data", "data",
                        arrowstyle="->", shrinkA=5, shrinkB=5,
                        mutation_scale=25, fc=text_colors[1], color=text_colors[1])
ax.add_artist(arrow5)

# Title
ax.text(8, 9.5, 'PATHOPHYSIOLOGY OF DIABETES MELLITUS', ha='center', va='center',
        fontsize=18, fontweight='bold', color='#263238')

# Subtitle
ax.text(8, 9.0, 'Mechanisms leading to hyperglycemia in different types', ha='center', va='center',
        fontsize=12, color='#546E7A')

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/pathophysiology_diagram.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("Pathophysiology diagram saved as pathophysiology_diagram.png")
