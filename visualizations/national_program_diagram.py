import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle

# Create figure and axis
fig, ax = plt.subplots(figsize=(16, 12))
ax.set_xlim(0, 16)
ax.set_ylim(0, 12)
ax.axis('off')

# Colors for different components
colors = {
    'primary': '#e74c3c',
    'secondary': '#f39c12',
    'tertiary': '#27ae60',
    'monitoring': '#3498db',
    'policy': '#9b59b6',
    'infrastructure': '#34495e'
}

# Title
ax.text(8, 11.5, 'INDIA NATIONAL DIABETES CONTROL PROGRAMME', ha='center', va='center',
        fontsize=22, fontweight='bold', color='#2c3e50')
ax.text(8, 11.0, 'NPCDCS - National Programme for Prevention and Control of Cancer, Diabetes, CVD and Stroke',
        ha='center', va='center', fontsize=12, color='#7f8c8d')

# Central coordinating mechanism
central_box = FancyBboxPatch((6.5, 9), 3, 1, boxstyle="round,pad=0.1",
                           facecolor=colors['policy'], alpha=0.9)
ax.add_patch(central_box)
ax.text(8, 9.5, 'NATIONAL LEVEL', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(8, 9.25, 'Ministry of Health & Family Welfare', ha='center', va='center', fontsize=10, color='white')
ax.text(8, 9.05, 'Policy framework • Funding • Guidelines', ha='center', va='center', fontsize=9, color='white')

# State level implementation
state_box = FancyBboxPatch((6.5, 7.5), 3, 1, boxstyle="round,pad=0.1",
                          facecolor=colors['infrastructure'], alpha=0.9)
ax.add_patch(state_box)
ax.text(8, 8, 'STATE LEVEL', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(8, 7.75, 'State Health Departments', ha='center', va='center', fontsize=10, color='white')
ax.text(8, 7.55, 'Implementation • Monitoring • Resource allocation', ha='center', va='center', fontsize=9, color='white')

# Primary Prevention (Left side)
primary_box = FancyBboxPatch((0.5, 6), 2.5, 1.5, boxstyle="round,pad=0.1",
                           facecolor=colors['primary'], alpha=0.8)
ax.add_patch(primary_box)
ax.text(1.75, 6.9, 'PRIMARY PREVENTION', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(1.75, 6.6, '• Health promotion', ha='center', va='center', fontsize=10, color='white')
ax.text(1.75, 6.35, '• Lifestyle education', ha='center', va='center', fontsize=10, color='white')
ax.text(1.75, 6.1, '• Risk factor screening', ha='center', va='center', fontsize=10, color='white')

# Secondary Prevention (Left-middle)
secondary_box = FancyBboxPatch((0.5, 4), 2.5, 1.5, boxstyle="round,pad=0.1",
                              facecolor=colors['secondary'], alpha=0.8)
ax.add_patch(secondary_box)
ax.text(1.75, 4.9, 'SECONDARY PREVENTION', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(1.75, 4.6, '• Early diagnosis', ha='center', va='center', fontsize=10, color='white')
ax.text(1.75, 4.35, '• Prediabetes management', ha='center', va='center', fontsize=10, color='white')
ax.text(1.75, 4.1, '• Lifestyle intervention', ha='center', va='center', fontsize=10, color='white')

# Tertiary Prevention (Right side)
tertiary_box = FancyBboxPatch((4, 5), 2.5, 1.5, boxstyle="round,pad=0.1",
                             facecolor=colors['tertiary'], alpha=0.8)
ax.add_patch(tertiary_box)
ax.text(5.25, 5.9, 'TERTIARY PREVENTION', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(5.25, 5.6, '• Complication screening', ha='center', va='center', fontsize=10, color='white')
ax.text(5.25, 5.35, '• Optimal glycemic control', ha='center', va='center', fontsize=10, color='white')
ax.text(5.25, 5.1, '• Multi-disciplinary care', ha='center', va='center', fontsize=10, color='white')

# District level operations
district_box = FancyBboxPatch((8.5, 6), 3.5, 1.5, boxstyle="round,pad=0.1",
                             facecolor=colors['monitoring'], alpha=0.8)
ax.add_patch(district_box)
ax.text(10.25, 6.9, 'DISTRICT LEVEL OPERATIONS', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(10.25, 6.6, '• NCD Clinics establishment', ha='center', va='center', fontsize=10, color='white')
ax.text(10.25, 6.35, '• Screening camps organization', ha='center', va='center', fontsize=10, color='white')
ax.text(10.25, 6.1, '• Training of healthcare workers', ha='center', va='center', fontsize=10, color='white')

# Community level
community_box = FancyBboxPatch((12.5, 6), 3, 1.5, boxstyle="round,pad=0.1",
                              facecolor='#e67e22', alpha=0.8)
ax.add_patch(community_box)
ax.text(14, 6.9, 'COMMUNITY LEVEL', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(14, 6.6, 'ASHA workers • ANM', ha='center', va='center', fontsize=10, color='white')
ax.text(14, 6.35, 'Screening • Education', ha='center', va='center', fontsize=10, color='white')
ax.text(14, 6.1, 'Referral services', ha='center', va='center', fontsize=10, color='white')

# Key components (bottom level)
components_box = FancyBboxPatch((2, 2), 12, 1.5, boxstyle="round,pad=0.1",
                               facecolor='#34495e', alpha=0.9)
ax.add_patch(components_box)
ax.text(8, 3, 'KEY PROGRAM COMPONENTS', ha='center', va='center', fontsize=16, fontweight='bold', color='white')
key_components = [
    '• Capacity building • Drug procurement • Monitoring & evaluation',
    '• Public-private partnerships • IEC activities • Surveillance system',
    '• Research & innovation • Quality assurance • Financial protection'
]
for i, component in enumerate(key_components):
    ax.text(8, 2.7 - i*0.15, component, ha='center', va='center', fontsize=10, color='white')

# Connecting arrows
# Central to State
plt.arrow(8, 9, 0, -0.8, head_width=0.1, head_length=0.1, fc=colors['policy'], ec=colors['policy'], alpha=0.7)

# State to District
plt.arrow(8, 7.5, 0, -0.8, head_width=0.1, head_length=0.1, fc=colors['infrastructure'], ec=colors['infrastructure'], alpha=0.7)

# District to Community
plt.arrow(10.25, 6, 2.25, 0, head_width=0.1, head_length=0.1, fc=colors['monitoring'], ec=colors['monitoring'], alpha=0.7)

# District to prevention components
plt.arrow(8.5, 6, -3, -0.8, head_width=0.1, head_length=0.1, fc=colors['monitoring'], ec=colors['monitoring'], alpha=0.7, linestyle='--')
plt.arrow(8.5, 6, -2.25, -1.8, head_width=0.1, head_length=0.1, fc=colors['monitoring'], ec=colors['monitoring'], alpha=0.7, linestyle='--')

# Outcomes (bottom right)
outcomes_box = FancyBboxPatch((12, 2), 3.5, 1, boxstyle="round,pad=0.1",
                             facecolor='#27ae60', alpha=0.9)
ax.add_patch(outcomes_box)
ax.text(13.75, 2.5, 'EXPECTED OUTCOMES', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(13.75, 2.25, '• 25% reduction in mortality', ha='center', va='center', fontsize=9, color='white')
ax.text(13.75, 2.05, '• Improved quality of life', ha='center', va='center', fontsize=9, color='white')

# Bottom explanatory text
ax.text(8, 0.8, 'India launched NPCDCS in 2010, expanded to all states. Focus: Early detection, treatment, and prevention of NCDs',
        ha='center', va='center', fontsize=11, color='#7f8c8d', style='italic')
ax.text(8, 0.5, 'Components: Population-based screening • Capacity building • Referral system • Health promotion',
        ha='center', va='center', fontsize=10, color='#95a5a6')

# Legend (top right)
legend_box = FancyBboxPatch((12.5, 9), 3, 1, boxstyle="round,pad=0.05",
                           facecolor='white', edgecolor='#bdc3c7', alpha=0.9)
ax.add_patch(legend_box)
ax.text(12.8, 9.9, 'LEVELS OF PREVENTION', ha='left', va='center', fontsize=10, fontweight='bold')
# Color indicators
ax.add_patch(Circle((13.2, 9.6), 0.08, facecolor=colors['primary']))
ax.text(13.4, 9.6, 'Primary', ha='left', va='center', fontsize=9)
ax.add_patch(Circle((13.2, 9.4), 0.08, facecolor=colors['secondary']))
ax.text(13.4, 9.4, 'Secondary', ha='left', va='center', fontsize=9)
ax.add_patch(Circle((13.2, 9.2), 0.08, facecolor=colors['tertiary']))
ax.text(13.4, 9.2, 'Tertiary', ha='left', va='center', fontsize=9)

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/national_program_diagram.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("National program diagram saved as national_program_diagram.png")
