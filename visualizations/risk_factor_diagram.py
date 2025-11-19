import matplotlib.pyplot as plt
import numpy as np

# Risk factor data
categories = ['Obesity', 'Physical\nInactivity', 'Family\nHistory', 'Age\n>45', 'Urban\nLiving', 'High BMI\nDiet']
outer_values = [85, 7, 40, 60, 30, 25]  # Population attributable risk (%)
inner_values = [50, 25, 30, 40, 45, 35]  # Individual risk multiplier

# Colors for different categories
colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', '#FF9FF3']

# Create figure with polar plot
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
fig.suptitle('DIABETES RISK FACTOR ANALYSIS', fontsize=16, fontweight='bold')

# Population Attributable Risk - Outer ring
ax1.bar(categories, outer_values, color=colors, alpha=0.8, width=0.6)
ax1.set_title('Population Attributable Risk (%)', fontsize=14, fontweight='bold')
ax1.set_ylabel('Percentage (%)')
ax1.grid(True, alpha=0.3, axis='y')

for i, v in enumerate(outer_values):
    ax1.text(i, v + 2, f'{v}%', ha='center', va='bottom', fontweight='bold', fontsize=11)

# Risk Multiplier - Inner circle
bars = ax2.bar(categories, inner_values, color=colors, alpha=0.7, width=0.4)
ax2.set_title('Individual Risk Multiplier', fontsize=14, fontweight='bold')
ax2.set_ylabel('Risk Multiplier')
ax2.grid(True, alpha=0.3, axis='y')

for bar, value in zip(bars, inner_values):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
             f'{value}x', ha='center', va='bottom', fontweight='bold', fontsize=11)

# Rotate x-axis labels
ax1.tick_params(axis='x', rotation=45)
ax2.tick_params(axis='x', rotation=45)

# Add legend
legend_elements = [plt.Rectangle((0,0),1,1, facecolor=color, alpha=0.8)
                  for color in colors]
ax1.legend(legend_elements, categories, bbox_to_anchor=(1.05, 1), loc='upper left')

plt.tight_layout()
plt.savefig('TLM_Diabetes_Mellitus/visualizations/risk_factor_diagram.png', dpi=300, bbox_inches='tight', pad_inches=0.5)
plt.close()

print("Risk factor diagram saved as risk_factor_diagram.png")
