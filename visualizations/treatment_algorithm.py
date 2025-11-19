import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Create figure and axis
fig, ax = plt.subplots(figsize=(16, 10))
ax.set_xlim(0, 16)
ax.set_ylim(0, 12)
ax.axis('off')

# Define colors and styles
decision_colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6']
action_colors = ['#ecf0f1', '#d5dbdb', '#a8e6cf', '#ffd3a5', '#98d8c8']
text_color = '#2c3e50'

# Starting point
ax.add_patch(patches.FancyBboxPatch((7, 10.5), 2, 0.8, boxstyle="round,pad=0.2",
                                   facecolor=decision_colors[0], edgecolor=decision_colors[0]))
ax.text(8, 10.9, 'PATIENT WITH DIABETES', ha='center', va='center', fontsize=12, fontweight='bold', color='white')

# Arrow down
ax.arrow(8, 10.3, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# First decision: Newly diagnosed?
ax.add_patch(patches.FancyBboxPatch((6.5, 8.5), 3, 1, boxstyle="round,pad=0.2",
                                   facecolor=decision_colors[1], edgecolor=decision_colors[1]))
ax.text(8, 9, 'NEWLY DIAGNOSED\nTYPE 2 DM?', ha='center', va='center', fontsize=11, fontweight='bold', color='white')

# Left arrow (No - established)
ax.arrow(8, 8.5, -2.5, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(5.5, 8.7, 'No', ha='center', va='center', fontsize=10, fontweight='bold')

# Right arrow (Yes - newly diagnosed)
ax.arrow(8, 8.5, 2.5, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(10.5, 8.7, 'Yes', ha='center', va='center', fontsize=10, fontweight='bold')

# Newly diagnosed pathway (right side)
ax.add_patch(patches.FancyBboxPatch((10.5, 7), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[0], edgecolor=decision_colors[2]))
ax.text(11.75, 7.4, 'Lifestyle + Metformin', ha='center', va='center', fontsize=10, fontweight='bold')

ax.arrow(11.75, 6.9, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

ax.add_patch(patches.FancyBboxPatch((10.5, 5.5), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=decision_colors[3], edgecolor=decision_colors[3]))
ax.text(11.75, 6, 'Target achieved?', ha='center', va='center', fontsize=10, fontweight='bold', color='white')

# Target not achieved - add second agent
ax.arrow(11.75, 5.5, -1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.arrow(11.75, 5.5, 1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(10.5, 5.7, 'No', ha='center', va='center', fontsize=9, fontweight='bold')
ax.text(13, 5.7, 'Yes', ha='center', va='center', fontsize=9, fontweight='bold')

# Dual therapy
ax.add_patch(patches.FancyBboxPatch((8.5, 4), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[1], edgecolor=decision_colors[4]))
ax.text(9.75, 4.4, 'Add second oral agent\n(DPP-4i/SGLT2i/GLP-1RA)', ha='center', va='center', fontsize=9, fontweight='bold')

ax.arrow(9.75, 3.9, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Target check again
ax.add_patch(patches.FancyBboxPatch((8.5, 2.5), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=decision_colors[2], edgecolor=decision_colors[2]))
ax.text(9.75, 3, 'Target achieved?', ha='center', va='center', fontsize=9, fontweight='bold', color='white')

# Continue to triple therapy or insulin
ax.arrow(9.75, 2.5, -1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.arrow(9.75, 2.5, 1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(8.5, 2.7, 'No', ha='center', va='center', fontsize=8, fontweight='bold')
ax.text(11, 2.7, 'Yes', ha='center', va='center', fontsize=8, fontweight='bold')

# Triple therapy
ax.add_patch(patches.FancyBboxPatch((6.5, 1), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[2], edgecolor=decision_colors[1]))
ax.text(7.75, 1.4, 'Triple oral therapy', ha='center', va='center', fontsize=9, fontweight='bold')

# Target achieved - continue therapy
ax.add_patch(patches.FancyBboxPatch((12.5, 1), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[4], edgecolor=decision_colors[4]))
ax.text(13.75, 1.4, 'Continue current therapy', ha='center', va='center', fontsize=9, fontweight='bold')

# Established diabetes pathway (left side)
ax.add_patch(patches.FancyBboxPatch((3.5, 7), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[3], edgecolor=decision_colors[3]))
ax.text(4.75, 7.4, 'Review current therapy', ha='center', va='center', fontsize=10, fontweight='bold')

ax.arrow(4.75, 6.9, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

# Insulin needed?
ax.add_patch(patches.FancyBboxPatch((3.5, 5.5), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=decision_colors[4], edgecolor=decision_colors[4]))
ax.text(4.75, 6, 'Insulin required?', ha='center', va='center', fontsize=9, fontweight='bold', color='white')

ax.arrow(4.75, 5.5, -1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.arrow(4.75, 5.5, 1.25, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(3.5, 5.7, 'Yes', ha='center', va='center', fontsize=9, fontweight='bold')
ax.text(6, 5.7, 'No', ha='center', va='center', fontsize=9, fontweight='bold')

# Insulin initiation
ax.add_patch(patches.FancyBboxPatch((1.5, 4), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=action_colors[0], edgecolor=decision_colors[1]))
ax.text(2.75, 4.4, 'Start insulin\n(Basal ± prandial)', ha='center', va='center', fontsize=9, fontweight='bold')

ax.arrow(2.75, 3.9, 0, -0.5, head_width=0.1, head_length=0.1, fc='black', ec='black')

ax.add_patch(patches.FancyBboxPatch((1.5, 2.5), 2.5, 0.8, boxstyle="round,pad=0.1",
                                   facecolor=decision_colors[2], edgecolor=decision_colors[2]))
ax.text(2.75, 3, 'Target achieved?', ha='center', va='center', fontsize=9, fontweight='bold', color='white')

ax.arrow(2.75, 2.5, 0.75, 0, head_width=0.1, head_length=0.1, fc='black', ec='black')
ax.text(4, 2.7, 'Yes/Continue', ha='center', va='center', fontsize=8, fontweight='bold')

# Title
ax.text(8, 11.5, 'DIABETES MELLITUS TREATMENT ALGORITHM', ha='center', va='center',
        fontsize=18, fontweight='bold', color=text_color)

# Subtitle
ax.text(8, 11.0, 'ADA 2024 Evidence-Based Treatment Pathway', ha='center', va='center',
        fontsize=12, color='#7f8c8d')

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/treatment_algorithm.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("Treatment algorithm saved as treatment_algorithm.png")
