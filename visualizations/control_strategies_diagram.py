import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch, Rectangle

# Create figure and axis
fig, ax = plt.subplots(figsize=(16, 12))
ax.set_xlim(0, 16)
ax.set_ylim(0, 12)
ax.axis('off')

# Colors for different categories
colors = {
    'monitoring': '#3498db',
    'targets': '#2ecc71',
    'lifestyle': '#e74c3c',
    'medication': '#f39c12',
    'technology': '#9b59b6',
    'behavioral': '#1abc9c'
}

# Title
ax.text(8, 11.5, 'DIABETES CONTROL STRATEGIES PYRAMID', ha='center', va='center',
        fontsize=20, fontweight='bold', color='#2c3e50')

# Glycemic Control Targets
targets_box = FancyBboxPatch((1, 9), 5, 1.5, boxstyle="round,pad=0.1",
                           facecolor=colors['targets'], alpha=0.8)
ax.add_patch(targets_box)
ax.text(3.5, 9.75, 'GLYCEMIC TARGETS', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(3.5, 9.45, 'HbA1c <7.0%', ha='center', va='center', fontsize=12, color='white')
ax.text(3.5, 9.15, 'Individualized based on patient factors', ha='center', va='center', fontsize=10, color='white')

# Monitoring Strategies
monitoring_box = FancyBboxPatch((7, 9), 4, 1.5, boxstyle="round,pad=0.1",
                              facecolor=colors['monitoring'], alpha=0.8)
ax.add_patch(monitoring_box)
ax.text(9, 9.75, 'MONITORING', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(9, 9.45, 'SMBG • CGM • HbA1c', ha='center', va='center', fontsize=12, color='white')
ax.text(9, 9.15, 'Every 3-6 months', ha='center', va='center', fontsize=10, color='white')

# Risk Factor Control
risk_control_box = FancyBboxPatch((11.5, 9), 3.5, 1.5, boxstyle="round,pad=0.1",
                                facecolor=colors['lifestyle'], alpha=0.8)
ax.add_patch(risk_control_box)
ax.text(13.25, 9.75, 'RISK FACTOR CONTROL', ha='center', va='center', fontsize=12, fontweight='bold', color='white')
ax.text(13.25, 9.45, 'BP • Lipids • Weight', ha='center', va='center', fontsize=10, color='white')
ax.text(13.25, 9.15, 'Annual assessment', ha='center', va='center', fontsize=10, color='white')

# Lifestyle Management
lifestyle_box = FancyBboxPatch((1, 6.5), 6, 1.8, boxstyle="round,pad=0.1",
                              facecolor=colors['lifestyle'], alpha=0.9)
ax.add_patch(lifestyle_box)
ax.text(4, 7.65, 'LIFESTYLE MANAGEMENT', ha='center', va='center', fontsize=16, fontweight='bold', color='white')
ax.text(4, 7.25, 'Medical Nutritional Therapy • 150 min/week exercise', ha='center', va='center', fontsize=12, color='white')
ax.text(4, 6.9, '5-7% weight loss • Smoking cessation • Stress management', ha='center', va='center', fontsize=10, color='white')

# Pharmacological Management
medication_box = FancyBboxPatch((8, 6.5), 6, 1.8, boxstyle="round,pad=0.1",
                               facecolor=colors['medication'], alpha=0.9)
ax.add_patch(medication_box)
ax.text(11, 7.65, 'PHARMACOLOGICAL MANAGEMENT', ha='center', va='center', fontsize=16, fontweight='bold', color='white')
ax.text(11, 7.25, 'Metformin first-line • DPP-4i/SGLT2i/GLP-1RA • Insulin when needed', ha='center', va='center', fontsize=12, color='white')
ax.text(11, 6.9, 'Individualize regimen • Monitor hypoglycemia • Cost considerations', ha='center', va='center', fontsize=10, color='white')

# Technology Integration
tech_box = FancyBboxPatch((2, 4), 5, 1.5, boxstyle="round,pad=0.1",
                         facecolor=colors['technology'], alpha=0.8)
ax.add_patch(tech_box)
ax.text(4.5, 4.75, 'TECHNOLOGY INTEGRATION', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(4.5, 4.45, 'CGM • Insulin pumps • Digital health', ha='center', va='center', fontsize=12, color='white')
ax.text(4.5, 4.15, 'Mobile apps • Telemedicine • AI support', ha='center', va='center', fontsize=10, color='white')

# Behavioral Support
behavioral_box = FancyBboxPatch((8.5, 4), 5, 1.5, boxstyle="round,pad=0.1",
                              facecolor=colors['behavioral'], alpha=0.8)
ax.add_patch(behavioral_box)
ax.text(11, 4.75, 'BEHAVIORAL SUPPORT', ha='center', va='center', fontsize=14, fontweight='bold', color='white')
ax.text(11, 4.45, 'Motivational interviewing • Self-management', ha='center', va='center', fontsize=12, color='white')
ax.text(11, 4.15, 'Education • Counseling • Family support', ha='center', va='center', fontsize=10, color='white')

# Comprehensive Care Foundation
foundation_box = FancyBboxPatch((5, 1.5), 6, 1.5, boxstyle="round,pad=0.1",
                               facecolor='#34495e', alpha=0.9)
ax.add_patch(foundation_box)
ax.text(8, 2.5, 'COMPREHENSIVE MULTIDISCIPLINARY CARE', ha='center', va='center', fontsize=16, fontweight='bold', color='white')
ax.text(8, 2.1, 'Primary care + Specialists + Educators + Nutritionists', ha='center', va='center', fontsize=12, color='white')
ax.text(8, 1.8, 'Regular monitoring • Complication screening • Team approach', ha='center', va='center', fontsize=10, color='white')

# Connecting arrows showing hierarchy
plt.arrow(3.5, 9, 0, -1, head_width=0.1, head_length=0.1, fc=colors['targets'], ec=colors['targets'], alpha=0.7)
plt.arrow(9, 9, 0, -1, head_width=0.1, head_length=0.1, fc=colors['monitoring'], ec=colors['monitoring'], alpha=0.7)
plt.arrow(13.25, 9, 0, -1, head_width=0.1, head_length=0.1, fc=colors['lifestyle'], ec=colors['lifestyle'], alpha=0.7)

plt.arrow(4, 6.5, 0, -0.8, head_width=0.15, head_length=0.15, fc=colors['lifestyle'], ec=colors['lifestyle'], alpha=0.8)
plt.arrow(11, 6.5, 0, -0.8, head_width=0.15, head_length=0.15, fc=colors['medication'], ec=colors['medication'], alpha=0.8)

plt.arrow(4.5, 4, 0, -0.8, head_width=0.15, head_length=0.15, fc=colors['technology'], ec=colors['technology'], alpha=0.8)
plt.arrow(11, 4, 0, -0.8, head_width=0.15, head_length=0.15, fc=colors['behavioral'], ec=colors['behavioral'], alpha=0.8)

# Bottom text
ax.text(8, 0.5, 'Control strategies build upon each other: Start with fundamentals, add technology and behavioral support as needed', ha='center', va='center', fontsize=12, color='#7f8c8d', style='italic')

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/control_strategies_diagram.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("Control strategies diagram saved as control_strategies_diagram.png")
